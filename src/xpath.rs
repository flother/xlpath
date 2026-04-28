//! Namespace registry + compiled XPath query evaluator.

use sxd_document::parser;
use sxd_xpath::{nodeset::Node, Context, Factory, Value, XPath};
use thiserror::Error;

/// Well-known OOXML namespace URIs, paired with the prefixes that users
/// conventionally write. These are pre-registered so that queries like
/// `//c:chart` or `//x:sheet/@name` work against any OOXML document without
/// requiring `--ns` declarations.
pub const OOXML_DEFAULTS: &[(&str, &str)] = &[
    (
        "x",
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    ),
    (
        "r",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    ),
    (
        "c",
        "http://schemas.openxmlformats.org/drawingml/2006/chart",
    ),
    ("a", "http://schemas.openxmlformats.org/drawingml/2006/main"),
    (
        "xdr",
        "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    ),
    (
        "mc",
        "http://schemas.openxmlformats.org/markup-compatibility/2006",
    ),
    (
        "rel",
        "http://schemas.openxmlformats.org/package/2006/relationships",
    ),
    (
        "ct",
        "http://schemas.openxmlformats.org/package/2006/content-types",
    ),
    (
        "x14",
        "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    ),
    (
        "x15",
        "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
    ),
    (
        "xr",
        "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    ),
    (
        "xp",
        "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
    ),
];

/// A prefix → URI registry applied to every XPath evaluation. Built from the
/// OOXML defaults and optionally extended with user-supplied overrides.
#[derive(Debug, Clone, Default)]
pub struct Namespaces {
    bindings: Vec<(String, String)>,
}

impl Namespaces {
    /// Seed with the well-known OOXML bindings (see [`OOXML_DEFAULTS`]).
    pub fn with_defaults() -> Self {
        let bindings = OOXML_DEFAULTS
            .iter()
            .map(|(p, u)| ((*p).to_string(), (*u).to_string()))
            .collect();
        Self { bindings }
    }

    /// Look up a prefix's URI. `bindings` is kept deduplicated, so a plain
    /// linear search is sufficient.
    pub fn get(&self, prefix: &str) -> Option<&str> {
        self.bindings
            .iter()
            .find(|(p, _)| p == prefix)
            .map(|(_, u)| u.as_str())
    }

    /// Add a prefix binding. Later calls shadow earlier ones, so user `--ns`
    /// flags override the OOXML defaults. The invariant maintained here is that
    /// `bindings` contains at most one entry per prefix and lists them in the
    /// order they should be surfaced to downstream code (so [`Self::effective`]
    /// can be a plain iterator with no allocation).
    pub fn override_with(&mut self, prefix: &str, uri: &str) {
        self.bindings.retain(|(p, _)| p != prefix);
        self.bindings.push((prefix.to_string(), uri.to_string()));
    }

    /// Iterate over the effective bindings. Called on the hot path once per XML
    /// part, so this is a zero-allocation view of the internal vector.
    pub fn effective(&self) -> impl Iterator<Item = (&str, &str)> + '_ {
        self.bindings.iter().map(|(p, u)| (p.as_str(), u.as_str()))
    }
}

/// How an XPath match presented itself. Informs the final output formatting.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MatchKind {
    Element,
    Attribute,
    Text,
    /// String, number, or boolean produced by an XPath function (e.g.
    /// `count(...)`).
    Atomic,
}

/// A single hit from an XPath evaluation. `value` holds the text/attribute
/// content. For element matches with `EvalOptions::as_tag`, `tag` holds the
/// synthetic self-closing tag for use as a prefix component in output lines.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Match {
    pub kind: MatchKind,
    pub value: String,
    /// Synthetic self-closing element tag (e.g. `<c:chart/>`). Populated only
    /// when `EvalOptions::as_tag` is true and the match is an element node.
    pub tag: Option<String>,
}

/// Per-evaluation rendering switches. The defaults produce plain text-content
/// values with no tag.
#[derive(Debug, Clone, Copy, Default)]
pub struct EvalOptions {
    /// When true, node matches populate `tag` with a synthetic self-closing
    /// opening tag. For element matches the tag is the element itself; for
    /// attribute and text matches it is the parent element. `value` is always
    /// the text or attribute content. No effect on atomic matches.
    pub as_tag: bool,
}

#[derive(Debug, Error)]
pub enum QueryError {
    #[error("invalid XPath expression: {0}")]
    InvalidExpression(String),
    #[error("unsupported XPath expression")]
    UnsupportedExpression,
    #[error("unknown namespace prefix: {0}")]
    UnknownNamespacePrefix(String),
    #[error("malformed XML: {0}")]
    MalformedXml(String),
}

/// A validated XPath query plus the namespace bindings it should be evaluated
/// against. Reuse the same `Query` across many files.
///
/// The raw expression string is stored and re-parsed once per document because
/// the `sxd-xpath` expression tree is neither `Send` nor `Sync`: sharing it
/// across rayon workers would require `unsafe impl Sync`. Parsing is cheap
/// relative to reading the part and parsing the XML, so the trade-off
/// favours simplicity here.
#[derive(Debug)]
pub struct Query {
    expression: String,
    namespaces: Namespaces,
}

impl Query {
    /// Validate an XPath 1.0 expression and attach a namespace registry seeded
    /// with the OOXML defaults and extended with the caller's overrides. The
    /// expression is stored as a string; each evaluation re-parses it (see the
    /// struct-level doc for why).
    pub fn compile(expr: &str, user_ns: &[(String, String)]) -> Result<Self, QueryError> {
        // Parse once up-front to reject bad expressions early, then discard the
        // parsed tree; each evaluation will parse afresh.
        let factory = Factory::new();
        factory
            .build(expr)
            .map_err(|e| QueryError::InvalidExpression(e.to_string()))?
            .ok_or(QueryError::UnsupportedExpression)?;

        let mut namespaces = Namespaces::with_defaults();
        for (p, u) in user_ns {
            namespaces.override_with(p, u);
        }

        // sxd-xpath panics when it has no URI for a namespace prefix. Catch
        // that here callers get a clean error rather than a thread panic.
        for prefix in extract_namespace_prefixes(expr) {
            if namespaces.get(&prefix).is_none() {
                return Err(QueryError::UnknownNamespacePrefix(prefix));
            }
        }

        Ok(Self {
            expression: expr.to_string(),
            namespaces,
        })
    }

    /// Parse `xml` and run the compiled query over it, returning one [`Match`]
    /// per hit.
    pub fn evaluate_xml(&self, xml: &str) -> Result<Vec<Match>, QueryError> {
        self.evaluate_xml_with(xml, EvalOptions::default())
    }

    /// Parse `xml` and run the compiled query over it, applying the
    /// per-evaluation rendering options in `opts`.
    pub fn evaluate_xml_with(
        &self,
        xml: &str,
        opts: EvalOptions,
    ) -> Result<Vec<Match>, QueryError> {
        // Re-parse the XPath on each call: the compiled tree is not
        // `Send`/`Sync`, so per-call parsing lets us share a `Query` across
        // rayon workers without any `unsafe`.
        let factory = Factory::new();
        let xpath: XPath = factory
            .build(&self.expression)
            .map_err(|e| QueryError::InvalidExpression(e.to_string()))?
            .ok_or(QueryError::UnsupportedExpression)?;

        let pkg = parser::parse(xml).map_err(|e| QueryError::MalformedXml(e.to_string()))?;
        let doc = pkg.as_document();

        let mut ctx = Context::new();
        for (prefix, uri) in self.namespaces.effective() {
            ctx.set_namespace(prefix, uri);
        }

        let uri_to_prefix: Vec<(String, String)> = self
            .namespaces
            .effective()
            .map(|(p, u)| (u.to_string(), p.to_string()))
            .collect();

        let value = xpath
            .evaluate(&ctx, doc.root())
            .map_err(|e| QueryError::InvalidExpression(e.to_string()))?;

        Ok(collect_matches(value, opts, &uri_to_prefix))
    }
}

fn collect_matches(
    value: Value<'_>,
    opts: EvalOptions,
    uri_to_prefix: &[(String, String)],
) -> Vec<Match> {
    match value {
        Value::Nodeset(nodeset) => nodeset
            .document_order()
            .into_iter()
            .map(|node| node_to_match(node, opts, uri_to_prefix))
            .collect(),
        Value::String(s) => vec![Match {
            kind: MatchKind::Atomic,
            value: s,
            tag: None,
        }],
        Value::Number(n) => vec![Match {
            kind: MatchKind::Atomic,
            value: format_number(n),
            tag: None,
        }],
        Value::Boolean(b) => vec![Match {
            kind: MatchKind::Atomic,
            value: b.to_string(),
            tag: None,
        }],
    }
}

fn node_to_match(node: Node<'_>, opts: EvalOptions, uri_to_prefix: &[(String, String)]) -> Match {
    match node {
        Node::Element(e) => Match {
            kind: MatchKind::Element,
            value: collapse_whitespace(&node.string_value()),
            tag: if opts.as_tag {
                Some(render_opening_tag(e, uri_to_prefix))
            } else {
                None
            },
        },
        Node::Attribute(a) => Match {
            kind: MatchKind::Attribute,
            value: a.value().to_string(),
            tag: if opts.as_tag {
                // An attribute always has an owner element in well-formed XML.
                a.parent().map(|e| render_opening_tag(e, uri_to_prefix))
            } else {
                None
            },
        },
        Node::Text(t) => Match {
            kind: MatchKind::Text,
            value: collapse_whitespace(t.text()),
            tag: if opts.as_tag {
                t.parent().map(|e| render_opening_tag(e, uri_to_prefix))
            } else {
                None
            },
        },
        Node::Comment(c) => Match {
            kind: MatchKind::Text,
            value: collapse_whitespace(c.text()),
            tag: None,
        },
        Node::ProcessingInstruction(p) => Match {
            kind: MatchKind::Text,
            value: p.value().unwrap_or("").to_string(),
            tag: None,
        },
        Node::Namespace(n) => Match {
            kind: MatchKind::Atomic,
            value: n.uri().to_string(),
            tag: None,
        },
        Node::Root(_) => Match {
            kind: MatchKind::Element,
            value: collapse_whitespace(&node.string_value()),
            tag: None,
        },
    }
}

/// Render an element as a synthetic self-closing opening tag. Uses the
/// canonical prefix from the namespace registry, emits every attribute in the
/// order sxd-document exposes them, and never includes children or xmlns
/// declarations — the output is a reporting artefact, not a round-trippable XML
/// fragment.
fn render_opening_tag(
    element: sxd_document::dom::Element<'_>,
    uri_to_prefix: &[(String, String)],
) -> String {
    let mut out = String::from("<");
    let name = element.name();
    if let Some(uri) = name.namespace_uri() {
        if let Some(prefix) = lookup_prefix(uri_to_prefix, uri) {
            out.push_str(prefix);
            out.push(':');
        }
    }
    out.push_str(name.local_part());

    for attr in element.attributes() {
        out.push(' ');
        let aname = attr.name();
        if let Some(uri) = aname.namespace_uri() {
            if let Some(prefix) = lookup_prefix(uri_to_prefix, uri) {
                out.push_str(prefix);
                out.push(':');
            }
        }
        out.push_str(aname.local_part());
        out.push_str("=\"");
        out.push_str(&escape_attr_value(attr.value()));
        out.push('"');
    }

    out.push_str("/>");
    out
}

/// Escape the minimal set of characters that must not appear raw inside a
/// double-quoted XML attribute value.
fn escape_attr_value(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '"' => out.push_str("&quot;"),
            _ => out.push(c),
        }
    }
    out
}

fn lookup_prefix<'a>(uri_to_prefix: &'a [(String, String)], uri: &str) -> Option<&'a str> {
    // Forward search --- the first registered prefix for a URI wins.
    uri_to_prefix
        .iter()
        .find(|(u, _)| u == uri)
        .map(|(_, p)| p.as_str())
}

/// XPath 1.0 renders integer-valued numbers without a trailing `.0`; keep that
/// convention so `count(...)` outputs a clean integer.
fn format_number(n: f64) -> String {
    if n.is_nan() {
        "NaN".to_string()
    } else if n.is_infinite() {
        if n.is_sign_negative() {
            "-Infinity".to_string()
        } else {
            "Infinity".to_string()
        }
    } else if n == n.trunc() && n.abs() < 1e15 {
        format!("{}", n as i64)
    } else {
        format!("{n}")
    }
}

/// Extract every namespace prefix used in an XPath expression. Skips over
/// string literals to avoid false positives from text like `"foo:bar"`. Matches
/// `NCName:` only when the colon is not doubled (which would indicate an axis
/// separator such as `child::`).
fn extract_namespace_prefixes(expr: &str) -> Vec<String> {
    let mut prefixes = Vec::new();
    let mut iter = expr.chars().peekable();

    while let Some(ch) = iter.next() {
        // Skip string literals so "foo:bar" doesn't look like a prefix.
        if ch == '"' || ch == '\'' {
            for c in iter.by_ref() {
                if c == ch {
                    break;
                }
            }
            continue;
        }

        if is_ncname_start(ch) {
            let mut name = String::new();
            name.push(ch);
            while iter.peek().is_some_and(|&c| is_ncname_continue(c)) {
                name.push(iter.next().unwrap());
            }
            // A single `:` (not `::`) means this NCName is a namespace prefix.
            if iter.peek() == Some(&':') {
                iter.next(); // consume the `:`
                if iter.peek() != Some(&':') && !prefixes.contains(&name) {
                    prefixes.push(name);
                }
            }
        }
    }

    prefixes
}

fn is_ncname_start(c: char) -> bool {
    c.is_ascii_alphabetic() || c == '_'
}

fn is_ncname_continue(c: char) -> bool {
    c.is_ascii_alphanumeric() || matches!(c, '_' | '-' | '.')
}

fn collapse_whitespace(s: &str) -> String {
    let collapsed: String = s
        .chars()
        .map(|c| {
            if c == '\n' || c == '\r' || c == '\t' {
                ' '
            } else {
                c
            }
        })
        .collect();
    let mut out = String::with_capacity(collapsed.len());
    let mut prev_space = false;
    for c in collapsed.chars() {
        if c == ' ' {
            if !prev_space {
                out.push(c);
            }
            prev_space = true;
        } else {
            out.push(c);
            prev_space = false;
        }
    }
    out.trim().to_string()
}

#[cfg(test)]
mod tests {
    use super::Namespaces;

    #[test]
    fn user_overrides_beat_defaults() {
        let mut ns = Namespaces::with_defaults();
        ns.override_with("c", "urn:example:custom-chart");

        assert_eq!(ns.get("c"), Some("urn:example:custom-chart"));
    }

    #[test]
    fn overrides_can_add_new_prefixes() {
        let mut ns = Namespaces::with_defaults();
        ns.override_with("custom", "urn:example:thing");

        assert_eq!(ns.get("custom"), Some("urn:example:thing"));
    }

    #[test]
    fn evaluates_a_namespaced_query_against_an_ooxml_document() {
        use super::Query;

        let xml = r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets>
    <sheet name="Alpha" sheetId="1" />
    <sheet name="Beta" sheetId="2" />
  </sheets>
</workbook>"#;
        let q = Query::compile("//x:sheet/@name", &[]).unwrap();
        let matches: Vec<String> = q
            .evaluate_xml(xml)
            .unwrap()
            .into_iter()
            .map(|m| m.value)
            .collect();

        assert_eq!(matches, vec!["Alpha".to_string(), "Beta".to_string()]);
    }

    #[test]
    fn evaluates_element_matches_by_concatenated_text() {
        use super::Query;

        let xml = "<root><a>hello</a><a>world</a></root>";
        let q = Query::compile("//a", &[]).unwrap();
        let values: Vec<String> = q
            .evaluate_xml(xml)
            .unwrap()
            .into_iter()
            .map(|m| m.value)
            .collect();

        assert_eq!(values, vec!["hello", "world"]);
    }

    #[test]
    fn user_namespace_argument_is_applied_to_the_query() {
        use super::Query;

        let xml = r#"<r:thing xmlns:r="urn:example:custom"><r:name>ok</r:name></r:thing>"#;
        let bindings = vec![("my".to_string(), "urn:example:custom".to_string())];
        let q = Query::compile("//my:name", &bindings).unwrap();
        let values: Vec<String> = q
            .evaluate_xml(xml)
            .unwrap()
            .into_iter()
            .map(|m| m.value)
            .collect();

        assert_eq!(values, vec!["ok"]);
    }

    #[test]
    fn atomic_result_is_returned_as_a_single_match() {
        use super::Query;

        let xml = "<root><a/><a/><a/></root>";
        let q = Query::compile("count(//a)", &[]).unwrap();
        let values: Vec<String> = q
            .evaluate_xml(xml)
            .unwrap()
            .into_iter()
            .map(|m| m.value)
            .collect();

        assert_eq!(values, vec!["3"]);
    }

    #[test]
    fn invalid_xpath_is_rejected_at_compile_time() {
        use super::Query;

        let err = Query::compile("//[", &[]);
        assert!(err.is_err());
    }

    #[test]
    fn unregistered_namespace_prefix_is_rejected_at_compile_time() {
        use super::{Query, QueryError};

        let err = Query::compile("//doesnotexist:checksum", &[]).unwrap_err();
        assert!(
            matches!(err, QueryError::UnknownNamespacePrefix(ref p) if p == "doesnotexist"),
            "expected UnknownNamespacePrefix(\"doesnotexist\"), got: {err:?}"
        );
    }

    #[test]
    fn user_supplied_namespace_prefix_is_accepted() {
        use super::Query;

        let bindings = vec![("custom".to_string(), "urn:example:ns".to_string())];
        assert!(Query::compile("//custom:thing", &bindings).is_ok());
    }

    #[test]
    fn prefix_lookalike_inside_string_literal_is_not_rejected() {
        use super::Query;

        // "unknown:thing" is a string value, not a name test — no namespace
        // lookup should occur and compilation should succeed.
        assert!(Query::compile(r#"//x:sheet[@name = "unknown:thing"]"#, &[]).is_ok());
    }

    #[test]
    fn malformed_xml_is_reported() {
        use super::Query;

        let q = Query::compile("//a", &[]).unwrap();
        let err = q.evaluate_xml("<broken");
        assert!(err.is_err());
    }

    #[test]
    fn tag_mode_populates_tag_field_for_element_matches() {
        use super::{EvalOptions, Query};

        let xml = r#"<root><a x="1"/></root>"#;
        let q = Query::compile("//a", &[]).unwrap();
        let matches = q
            .evaluate_xml_with(xml, EvalOptions { as_tag: true })
            .unwrap();

        assert_eq!(matches.len(), 1);
        assert_eq!(matches[0].tag, Some(r#"<a x="1"/>"#.to_string()));
        assert_eq!(matches[0].value, "");
    }

    #[test]
    fn tag_mode_uses_registered_prefix() {
        use super::{EvalOptions, Query};

        let xml = r#"<?xml version="1.0"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart />
</c:chartSpace>"#;
        let q = Query::compile("//c:chart", &[]).unwrap();
        let matches = q
            .evaluate_xml_with(xml, EvalOptions { as_tag: true })
            .unwrap();

        assert_eq!(matches.len(), 1);
        assert_eq!(matches[0].tag, Some("<c:chart/>".to_string()));
    }

    #[test]
    fn tag_mode_escapes_attribute_values() {
        use super::{EvalOptions, Query};

        let xml = r#"<root><a note="a &amp; b &lt; c &quot;d&quot;"/></root>"#;
        let q = Query::compile("//a", &[]).unwrap();
        let matches = q
            .evaluate_xml_with(xml, EvalOptions { as_tag: true })
            .unwrap();

        assert_eq!(matches.len(), 1);
        assert_eq!(
            matches[0].tag,
            Some(r#"<a note="a &amp; b &lt; c &quot;d&quot;"/>"#.to_string())
        );
    }

    #[test]
    fn tag_mode_shows_parent_element_for_attribute_matches() {
        use super::{EvalOptions, Query};

        let xml = r#"<?xml version="1.0"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets><sheet name="Alpha" sheetId="1"/></sheets>
</workbook>"#;
        let q = Query::compile("//x:sheet/@name", &[]).unwrap();
        let matches = q
            .evaluate_xml_with(xml, EvalOptions { as_tag: true })
            .unwrap();

        assert_eq!(matches.len(), 1);
        assert_eq!(matches[0].value, "Alpha");
        // tag shows the parent element, not None
        assert_eq!(
            matches[0].tag,
            Some(r#"<x:sheet name="Alpha" sheetId="1"/>"#.to_string())
        );
    }

    #[test]
    fn tag_mode_ignores_children_and_always_self_closes() {
        use super::{EvalOptions, Query};

        let xml = r#"<root><a><b/>text<c/></a></root>"#;
        let q = Query::compile("//a", &[]).unwrap();
        let matches = q
            .evaluate_xml_with(xml, EvalOptions { as_tag: true })
            .unwrap();

        assert_eq!(matches.len(), 1);
        assert_eq!(matches[0].tag, Some("<a/>".to_string()));
        assert_eq!(matches[0].value, "text");
    }

    #[test]
    fn default_registry_includes_well_known_ooxml_prefixes() {
        let ns = Namespaces::with_defaults();

        assert_eq!(
            ns.get("x"),
            Some("http://schemas.openxmlformats.org/spreadsheetml/2006/main")
        );
        assert_eq!(
            ns.get("r"),
            Some("http://schemas.openxmlformats.org/officeDocument/2006/relationships")
        );
        assert_eq!(
            ns.get("c"),
            Some("http://schemas.openxmlformats.org/drawingml/2006/chart")
        );
        assert_eq!(
            ns.get("ct"),
            Some("http://schemas.openxmlformats.org/package/2006/content-types")
        );
        assert_eq!(ns.get("nope"), None);
    }

    #[test]
    fn ct_prefix_matches_the_content_types_part() {
        use super::Query;

        let xml = r#"<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/xl/workbook.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/charts/chart1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>"#;
        let q = Query::compile("//ct:Override/@PartName", &[]).unwrap();
        let values: Vec<String> = q
            .evaluate_xml(xml)
            .unwrap()
            .into_iter()
            .map(|m| m.value)
            .collect();

        assert_eq!(values, vec!["/xl/workbook.xml", "/xl/charts/chart1.xml"]);
    }
}
