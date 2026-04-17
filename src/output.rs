//! Per-file output buffering and formatters for each output mode.

use std::io::{self, BufWriter, Stderr, Stdout, Write};
use std::path::Path;
use std::sync::Mutex;

use crate::cli::OutputMode;
use crate::xpath::Match;

/// Per-part match batch as produced by the main worker loop.
pub type PartMatches = (String, Vec<Match>);

/// Build the full output string for a single workbook, ready to be written to
/// stdout in one shot (keeps each file's lines contiguous even under parallel
/// execution).
pub fn format_file(
    mode: OutputMode,
    no_filename: bool,
    no_path: bool,
    path: &Path,
    parts: &[PartMatches],
) -> String {
    match mode {
        OutputMode::Minimal => format_minimal(no_filename, no_path, path, parts),
        OutputMode::Count => format_count(path, parts),
        OutputMode::OnlyFilenames => format_only_filenames(path, parts),
    }
}

fn format_minimal(no_filename: bool, no_path: bool, path: &Path, parts: &[PartMatches]) -> String {
    let mut out = String::new();
    for (part, matches) in parts {
        for m in matches {
            let mut prefix = String::new();
            if !no_filename {
                prefix.push_str(&path.display().to_string());
                prefix.push(':');
            }
            if !no_path {
                prefix.push_str(part);
                prefix.push(':');
            }
            if let Some(tag) = &m.tag {
                prefix.push_str(tag);
                prefix.push(':');
            }
            if prefix.is_empty() {
                out.push_str(&m.value);
            } else {
                out.push_str(&prefix);
                out.push(' ');
                out.push_str(&m.value);
            }
            out.push('\n');
        }
    }
    out
}

fn format_count(path: &Path, parts: &[PartMatches]) -> String {
    let total: usize = parts.iter().map(|(_, ms)| ms.len()).sum();
    if total == 0 {
        return String::new();
    }
    format!("{}:{}\n", path.display(), total)
}

fn format_only_filenames(path: &Path, parts: &[PartMatches]) -> String {
    let any = parts.iter().any(|(_, ms)| !ms.is_empty());
    if any {
        format!("{}\n", path.display())
    } else {
        String::new()
    }
}

/// Thread-safe writer that all workers share. Each worker builds its per-file
/// output as a `String`, then calls `emit_out` once; the mutex guarantees the
/// lines of one file are not interleaved with another's.
pub struct Writer {
    stdout: Mutex<BufWriter<Stdout>>,
    stderr: Mutex<BufWriter<Stderr>>,
}

impl Writer {
    pub fn new() -> Self {
        Self {
            stdout: Mutex::new(BufWriter::new(io::stdout())),
            stderr: Mutex::new(BufWriter::new(io::stderr())),
        }
    }

    pub fn emit_out(&self, s: &str) {
        if s.is_empty() {
            return;
        }
        if let Ok(mut g) = self.stdout.lock() {
            let _ = g.write_all(s.as_bytes());
        }
    }

    pub fn emit_err(&self, s: &str) {
        if s.is_empty() {
            return;
        }
        if let Ok(mut g) = self.stderr.lock() {
            let _ = g.write_all(s.as_bytes());
        }
    }

    pub fn flush(&self) {
        if let Ok(mut g) = self.stdout.lock() {
            let _ = g.flush();
        }
        if let Ok(mut g) = self.stderr.lock() {
            let _ = g.flush();
        }
    }
}

impl Default for Writer {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use std::path::PathBuf;

    use crate::cli::OutputMode;
    use crate::xpath::{Match, MatchKind};

    use super::format_file;

    fn m(kind: MatchKind, value: &str) -> Match {
        Match {
            kind,
            value: value.into(),
            tag: None,
        }
    }

    fn m_tag(kind: MatchKind, value: &str, tag: &str) -> Match {
        Match {
            kind,
            value: value.into(),
            tag: Some(tag.into()),
        }
    }

    #[test]
    fn minimal_mode_formats_one_line_per_match() {
        let path = PathBuf::from("book.xlsx");
        let parts = vec![(
            "xl/charts/chart1.xml".to_string(),
            vec![m(MatchKind::Element, "bar"), m(MatchKind::Element, "line")],
        )];

        let out = format_file(OutputMode::Minimal, false, false, &path, &parts);

        assert_eq!(
            out,
            "book.xlsx:xl/charts/chart1.xml: bar\nbook.xlsx:xl/charts/chart1.xml: line\n"
        );
    }

    #[test]
    fn minimal_mode_with_tag_in_prefix() {
        let path = PathBuf::from("book.xlsx");
        let parts = vec![(
            "xl/charts/chart1.xml".to_string(),
            vec![
                m_tag(MatchKind::Element, "bar", "<c:barChart/>"),
                m_tag(MatchKind::Element, "line", "<c:lineChart/>"),
            ],
        )];

        let out = format_file(OutputMode::Minimal, false, false, &path, &parts);

        assert_eq!(
            out,
            "book.xlsx:xl/charts/chart1.xml:<c:barChart/>: bar\nbook.xlsx:xl/charts/chart1.xml:<c:lineChart/>: line\n"
        );
    }

    #[test]
    fn minimal_mode_no_filename() {
        let path = PathBuf::from("book.xlsx");
        let parts = vec![(
            "xl/workbook.xml".to_string(),
            vec![m(MatchKind::Attribute, "Sheet1")],
        )];

        let out = format_file(OutputMode::Minimal, true, false, &path, &parts);

        assert_eq!(out, "xl/workbook.xml: Sheet1\n");
    }

    #[test]
    fn minimal_mode_no_path() {
        let path = PathBuf::from("book.xlsx");
        let parts = vec![(
            "xl/workbook.xml".to_string(),
            vec![m(MatchKind::Attribute, "Sheet1")],
        )];

        let out = format_file(OutputMode::Minimal, false, true, &path, &parts);

        assert_eq!(out, "book.xlsx: Sheet1\n");
    }

    #[test]
    fn minimal_mode_no_filename_no_path() {
        let path = PathBuf::from("book.xlsx");
        let parts = vec![(
            "xl/workbook.xml".to_string(),
            vec![m(MatchKind::Attribute, "Sheet1")],
        )];

        let out = format_file(OutputMode::Minimal, true, true, &path, &parts);

        assert_eq!(out, "Sheet1\n");
    }

    #[test]
    fn minimal_mode_no_filename_no_path_with_tag() {
        let path = PathBuf::from("book.xlsx");
        let parts = vec![(
            "xl/charts/chart1.xml".to_string(),
            vec![m_tag(MatchKind::Element, "bar", "<c:barChart/>")],
        )];

        let out = format_file(OutputMode::Minimal, true, true, &path, &parts);

        assert_eq!(out, "<c:barChart/>: bar\n");
    }

    #[test]
    fn count_mode_aggregates_matches_across_parts() {
        let path = PathBuf::from("book.xlsx");
        let parts = vec![
            (
                "xl/charts/chart1.xml".to_string(),
                vec![m(MatchKind::Element, "a"), m(MatchKind::Element, "b")],
            ),
            (
                "xl/charts/chart2.xml".to_string(),
                vec![m(MatchKind::Element, "c")],
            ),
        ];

        let out = format_file(OutputMode::Count, false, false, &path, &parts);

        assert_eq!(out, "book.xlsx:3\n");
    }

    #[test]
    fn count_mode_emits_nothing_when_the_file_has_no_matches() {
        let path = PathBuf::from("book.xlsx");
        let parts: Vec<super::PartMatches> = Vec::new();

        let out = format_file(OutputMode::Count, false, false, &path, &parts);

        assert!(out.is_empty());
    }

    #[test]
    fn only_filenames_mode_prints_path_once_when_any_match_exists() {
        let path = PathBuf::from("book.xlsx");
        let parts = vec![
            (
                "xl/workbook.xml".to_string(),
                vec![m(MatchKind::Element, "x")],
            ),
            ("xl/charts/chart1.xml".to_string(), Vec::new()),
        ];

        let out = format_file(OutputMode::OnlyFilenames, false, false, &path, &parts);

        assert_eq!(out, "book.xlsx\n");
    }

    #[test]
    fn only_filenames_mode_is_silent_without_matches() {
        let path = PathBuf::from("book.xlsx");
        let parts = vec![("xl/workbook.xml".to_string(), Vec::new())];

        let out = format_file(OutputMode::OnlyFilenames, false, false, &path, &parts);

        assert!(out.is_empty());
    }
}
