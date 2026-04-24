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
    /// Note: XPath 1.0 cannot match an unprefixed default namespace. Use a
    /// registered prefix (e.g. `//x:workbook`) or `--ns` to add one.
    #[arg(value_name = "XPATH", long_help = XPATH_LONG_HELP)]
    pub xpath: String,

    /// Files and/or directories to scan. Use `-` to read newline-separated
    /// paths from stdin. Directories are walked recursively. Defaults to the
    /// current working directory if omitted.
    #[arg(value_name = "PATH")]
    pub paths: Vec<PathBuf>,

    /// Follow symbolic links when walking directories. Off by default.
    #[arg(short = 'L', long = "follow")]
    pub follow: bool,

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

    /// Only print a count of matches per file.
    #[arg(short = 'c', long, conflicts_with_all = ["only_filenames", "json"])]
    pub count: bool,

    /// Only print the paths of files containing at least one match.
    #[arg(long = "only-filenames", conflicts_with_all = ["count", "json"])]
    pub only_filenames: bool,

    /// Output matches as newline-delimited JSON (one object per match).
    ///
    /// Each object can have `file`, `part`, `tag`, and `value` string fields,
    /// as controlled by flags (`--no-filename`, `--no-part`, `--tag`).
    #[arg(long = "json", conflicts_with_all = ["count", "only_filenames"])]
    pub json: bool,

    /// Render each element match as a synthetic self-closing opening tag (e.g.
    /// `<c:lineChart val="1"/>`) in the output prefix, alongside its text
    /// content. No effect on attribute, text, or atomic matches.
    #[arg(
        long = "tag",
        conflicts_with_all = ["count", "only_filenames"],
    )]
    pub tag: bool,

    /// Omit the filename from each output line.
    #[arg(
        long = "no-filename",
        conflicts_with_all = ["count", "only_filenames"],
    )]
    pub no_filename: bool,

    /// Omit the zip-internal part path from each output line.
    #[arg(
        long = "no-part",
        conflicts_with_all = ["count", "only_filenames"],
    )]
    pub no_part: bool,

    /// Number of worker threads. Defaults to the number of logical CPUs. Use
    /// `-j 1` for deterministic output order.
    #[arg(short = 'j', long = "threads", value_name = "N")]
    pub threads: Option<usize>,
}

/// Which output style to use.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OutputMode {
    /// Default: one line per match, `[file:][part:][tag: ]value`.
    Minimal,
    /// `--count`: one line per matching file, `file:N`.
    Count,
    /// `--only-filenames`: one line per matching file, `file`.
    OnlyFilenames,
    /// `--json`: one JSON object per match, newline-delimited.
    Json,
}

impl Cli {
    pub fn output_mode(&self) -> OutputMode {
        if self.count {
            OutputMode::Count
        } else if self.only_filenames {
            OutputMode::OnlyFilenames
        } else if self.json {
            OutputMode::Json
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

Note: XPath 1.0 cannot match an unprefixed default namespace. Use a
registered prefix (e.g. `//x:workbook`) or `--ns` to add one.

Pre-registered namespace prefixes (override or extend with `--ns`):

  x    http://schemas.openxmlformats.org/spreadsheetml/2006/main
  r    http://schemas.openxmlformats.org/officeDocument/2006/relationships
  c    http://schemas.openxmlformats.org/drawingml/2006/chart
  a    http://schemas.openxmlformats.org/drawingml/2006/main
  xdr  http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing
  mc   http://schemas.openxmlformats.org/markup-compatibility/2006
  rel  http://schemas.openxmlformats.org/package/2006/relationships
  ct   http://schemas.openxmlformats.org/package/2006/content-types
  x14  http://schemas.microsoft.com/office/spreadsheetml/2009/9/main
  x15  http://schemas.microsoft.com/office/spreadsheetml/2010/11/main
  xr   http://schemas.microsoft.com/office/spreadsheetml/2014/revision
  xp   http://schemas.openxmlformats.org/officeDocument/2006/extended-properties
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

    #[test]
    fn output_mode_only_filenames() {
        let cli = Cli::try_parse_from(["xlpath", "--only-filenames", "//foo"]).unwrap();
        assert!(matches!(cli.output_mode(), OutputMode::OnlyFilenames));
    }

    #[test]
    fn no_filename_and_no_part_flags_parse() {
        let cli = Cli::try_parse_from(["xlpath", "--no-filename", "--no-part", "//foo"]).unwrap();
        assert!(cli.no_filename);
        assert!(cli.no_part);
    }

    #[test]
    fn tag_conflicts_with_count() {
        let result = Cli::try_parse_from(["xlpath", "--tag", "--count", "//foo"]);
        assert!(result.is_err());
    }

    #[test]
    fn tag_conflicts_with_only_filenames() {
        let result = Cli::try_parse_from(["xlpath", "--tag", "--only-filenames", "//foo"]);
        assert!(result.is_err());
    }

    #[test]
    fn no_filename_conflicts_with_count() {
        let result = Cli::try_parse_from(["xlpath", "--no-filename", "--count", "//foo"]);
        assert!(result.is_err());
    }

    #[test]
    fn no_part_conflicts_with_only_filenames() {
        let result = Cli::try_parse_from(["xlpath", "--no-part", "--only-filenames", "//foo"]);
        assert!(result.is_err());
    }
}
