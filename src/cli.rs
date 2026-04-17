use std::path::PathBuf;

use clap::{ArgAction, Parser};

/// Query Excel (OOXML) files with XPath.
///
/// `xlpath` opens each supplied workbook as a ZIP archive, evaluates the given
/// XPath expression against each internal XML part (optionally filtered by
/// include/exclude globs), and prints matches to stdout in a grep-like form.
#[derive(Debug, Parser)]
#[command(name = "xlpath", version, about, long_about = None)]
pub struct Cli {
    /// XPath 1.0 expression to evaluate against each XML part.
    ///
    /// Note: XPath 1.0 cannot match an unprefixed default namespace. Either use
    /// a registered prefix (e.g. `//x:workbook`) or pass `--default-ns`.
    #[arg(value_name = "XPATH", long_help = XPATH_LONG_HELP)]
    pub xpath: String,

    /// Files and/or directories to scan. Use `-` to read newline-separated
    /// paths from stdin. Directories are walked recursively.
    #[arg(value_name = "PATH", required = true)]
    pub paths: Vec<PathBuf>,

    /// Only consider XML parts matching this glob. Repeatable. Globs apply to
    /// the zip-internal path (e.g. `xl/charts/*.xml`).
    #[arg(long = "include", value_name = "GLOB", action = ArgAction::Append)]
    pub includes: Vec<String>,

    /// Skip XML parts matching this glob. Repeatable.
    #[arg(long = "exclude", value_name = "GLOB", action = ArgAction::Append)]
    pub excludes: Vec<String>,

    /// Register a namespace prefix. Repeatable. Format: `prefix=uri`. Applied
    /// after the auto-registered OOXML defaults (last-wins).
    #[arg(long = "ns", value_name = "PREFIX=URI", action = ArgAction::Append)]
    pub namespaces: Vec<String>,

    /// Bind the root element's default xmlns (if any) to this prefix for each
    /// document.
    #[arg(long = "default-ns", value_name = "PREFIX")]
    pub default_ns: Option<String>,

    /// Only print a count of matches per file.
    #[arg(long, conflicts_with = "files_only")]
    pub count: bool,

    /// Only print the paths of files containing at least one match.
    #[arg(long = "files-only")]
    pub files_only: bool,

    /// Include an XPath-like location for each match in the output.
    #[arg(long = "with-path")]
    pub with_path: bool,

    /// Render each element match as a synthetic self-closing opening tag (e.g.
    /// `<c:lineChart val="1"/>`) instead of its text content. No effect on
    /// attribute, text, or atomic matches. Ignored under `--count` and
    /// `--files-only`.
    #[arg(long = "tag")]
    pub tag: bool,

    /// Print only the match value on each line — no `file:part:` prefix.
    /// Requires `--tag` (so element matches render as their synthetic
    /// self-closing tag). Conflicts with `--count`, `--files-only`, and
    /// `--with-path`.
    #[arg(
        long = "tag-only",
        requires = "tag",
        conflicts_with_all = ["count", "files_only", "with_path"],
    )]
    pub tag_only: bool,

    /// Number of worker threads. Defaults to the number of logical CPUs. Use
    /// `-j 1` for deterministic output order.
    #[arg(short = 'j', long = "threads", value_name = "N")]
    pub threads: Option<usize>,
}

/// Which output style to use.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OutputMode {
    /// Default: one line per match, `file:part: value` (plus optional location
    /// when `--with-path` is set).
    Minimal,
    /// `--count`: one line per matching file, `file:N`.
    Count,
    /// `--files-only`: one line per matching file, `file`.
    FilesOnly,
    /// `--tag-only`: one line per match containing just the match value (which,
    /// with `--tag`, is the synthetic self-closing tag). No file or part
    /// prefix.
    TagOnly,
}

impl Cli {
    pub fn output_mode(&self) -> OutputMode {
        if self.count {
            OutputMode::Count
        } else if self.files_only {
            OutputMode::FilesOnly
        } else if self.tag_only {
            OutputMode::TagOnly
        } else {
            OutputMode::Minimal
        }
    }
}

/// Long-form help for the positional XPath argument, shown under `--help` but
/// not `-h`. Includes the table of OOXML namespace prefixes that are
/// pre-registered for every query.
///
/// The table is maintained by hand rather than built from
/// [`crate::xpath::OOXML_DEFAULTS`] at runtime because clap's derive API only
/// accepts `&'static str` for `long_help`. The unit test
/// `long_help_covers_every_default_namespace` guards against the two lists
/// drifting apart.
const XPATH_LONG_HELP: &str = "\
XPath 1.0 expression to evaluate against each XML part.

Note: XPath 1.0 cannot match an unprefixed default namespace. Either use a
registered prefix (e.g. `//x:workbook`) or pass `--default-ns`.

Pre-registered namespace prefixes (override or extend with `--ns`):

  x, main  http://schemas.openxmlformats.org/spreadsheetml/2006/main
  r        http://schemas.openxmlformats.org/officeDocument/2006/relationships
  c        http://schemas.openxmlformats.org/drawingml/2006/chart
  a        http://schemas.openxmlformats.org/drawingml/2006/main
  xdr      http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing
  mc       http://schemas.openxmlformats.org/markup-compatibility/2006
  rel      http://schemas.openxmlformats.org/package/2006/relationships
  ct       http://schemas.openxmlformats.org/package/2006/content-types
  x14      http://schemas.microsoft.com/office/spreadsheetml/2009/9/main
  x15      http://schemas.microsoft.com/office/spreadsheetml/2010/11/main
  xr       http://schemas.microsoft.com/office/spreadsheetml/2014/revision
";

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xpath::OOXML_DEFAULTS;

    #[test]
    fn long_help_covers_every_default_namespace() {
        for (prefix, uri) in OOXML_DEFAULTS {
            assert!(
                XPATH_LONG_HELP.contains(prefix),
                "XPATH_LONG_HELP is missing prefix `{prefix}`"
            );
            assert!(
                XPATH_LONG_HELP.contains(uri),
                "XPATH_LONG_HELP is missing URI `{uri}`"
            );
        }
    }
}
