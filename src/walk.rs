//! Input resolution: take the positional path args, expand `-` from stdin and
//! directories via `walkdir`, filter to supported OOXML extensions, and skip
//! Excel's `~$`-prefixed lock files.

use std::io::{self, BufRead};
use std::path::{Path, PathBuf};

use walkdir::WalkDir;

/// OOXML spreadsheet extensions we recognise. `.xlsb` is deliberately excluded
/// since it is a binary format not amenable to XML querying.
pub const OOXML_EXTENSIONS: &[&str] = &["xlsx", "xlsm", "xltx", "xltm"];

/// Resolve positional path arguments into a concrete list of spreadsheet files.
/// Directories are walked recursively (following symlinks). A single `-` entry
/// means "read newline-separated paths from `stdin`".
///
/// Each resolved path is kept as-is — no canonicalisation — so the output the
/// user sees refers to the path they asked about.
pub fn collect<R: BufRead>(paths: &[PathBuf], stdin: R) -> io::Result<Vec<PathBuf>> {
    let mut out = Vec::new();
    let mut stdin = Some(stdin);

    for p in paths {
        if p.as_os_str() == "-" {
            // `-` is a special marker: read newline-separated paths from stdin.
            // Only honour the first `-` — subsequent ones are no-ops because
            // stdin has already been consumed.
            if let Some(reader) = stdin.take() {
                for line in reader.lines() {
                    let line = line?;
                    let trimmed = line.trim();
                    if trimmed.is_empty() {
                        continue;
                    }
                    push_from_input(&PathBuf::from(trimmed), &mut out);
                }
            }
        } else {
            push_from_input(p, &mut out);
        }
    }

    Ok(out)
}

fn push_from_input(input: &Path, out: &mut Vec<PathBuf>) {
    if input.is_dir() {
        for entry in WalkDir::new(input).follow_links(true) {
            let Ok(entry) = entry else { continue };
            if !entry.file_type().is_file() {
                continue;
            }
            let path = entry.into_path();
            if is_spreadsheet(&path) && !is_lock_file(&path) {
                out.push(path);
            }
        }
    } else {
        // Explicitly named files are passed through regardless of extension or
        // lock-file naming — the user asked for this one specifically.
        out.push(input.to_path_buf());
    }
}

fn is_spreadsheet(path: &Path) -> bool {
    let Some(ext) = path.extension().and_then(|e| e.to_str()) else {
        return false;
    };
    let lower = ext.to_ascii_lowercase();
    OOXML_EXTENSIONS.iter().any(|e| *e == lower)
}

fn is_lock_file(path: &Path) -> bool {
    path.file_name()
        .and_then(|n| n.to_str())
        .is_some_and(|n| n.starts_with("~$"))
}

#[cfg(test)]
mod tests {
    use std::fs;
    use std::fs::File;
    use std::io::{self, Write};
    use std::path::PathBuf;

    use tempfile::TempDir;

    use super::collect;

    fn touch(path: &std::path::Path) {
        if let Some(parent) = path.parent() {
            fs::create_dir_all(parent).unwrap();
        }
        File::create(path).unwrap();
    }

    fn sorted(mut v: Vec<PathBuf>) -> Vec<PathBuf> {
        v.sort();
        v
    }

    fn empty_stdin() -> io::Cursor<Vec<u8>> {
        io::Cursor::new(Vec::new())
    }

    #[test]
    fn passes_through_an_existing_xlsx_file() {
        let dir = TempDir::new().unwrap();
        let path = dir.path().join("book.xlsx");
        touch(&path);

        let result = collect(&[path.clone()], empty_stdin()).unwrap();

        assert_eq!(result, vec![path]);
    }

    #[test]
    fn recurses_into_directories_and_collects_ooxml_files() {
        let dir = TempDir::new().unwrap();
        let a = dir.path().join("a.xlsx");
        let b = dir.path().join("nested/b.xlsm");
        let c = dir.path().join("nested/deep/c.xltx");
        let d = dir.path().join("nested/d.xltm");
        let skipped_txt = dir.path().join("readme.txt");
        let skipped_csv = dir.path().join("data.csv");
        let skipped_xlsb = dir.path().join("legacy.xlsb");
        for p in [&a, &b, &c, &d, &skipped_txt, &skipped_csv, &skipped_xlsb] {
            touch(p);
        }

        let result = collect(&[dir.path().to_path_buf()], empty_stdin()).unwrap();

        assert_eq!(sorted(result), sorted(vec![a, b, c, d]));
    }

    #[test]
    fn skips_excel_lock_files_in_directories() {
        let dir = TempDir::new().unwrap();
        let real = dir.path().join("report.xlsx");
        let lock = dir.path().join("~$report.xlsx");
        touch(&real);
        touch(&lock);

        let result = collect(&[dir.path().to_path_buf()], empty_stdin()).unwrap();

        assert_eq!(result, vec![real]);
    }

    #[test]
    fn does_not_filter_explicit_file_arguments() {
        // When the user explicitly names a file, trust them — even if it has an
        // unusual extension or a lock-file name.
        let dir = TempDir::new().unwrap();
        let lock = dir.path().join("~$open.xlsx");
        let odd = dir.path().join("weird.dat");
        touch(&lock);
        touch(&odd);

        let result = collect(&[lock.clone(), odd.clone()], empty_stdin()).unwrap();

        assert_eq!(result, vec![lock, odd]);
    }

    #[test]
    fn recognises_extensions_case_insensitively() {
        let dir = TempDir::new().unwrap();
        let upper = dir.path().join("LOUD.XLSX");
        let mixed = dir.path().join("Mixed.Xlsm");
        touch(&upper);
        touch(&mixed);

        let result = collect(&[dir.path().to_path_buf()], empty_stdin()).unwrap();

        assert_eq!(sorted(result), sorted(vec![upper, mixed]));
    }

    #[test]
    fn reads_paths_from_stdin_when_dash_is_given() {
        let dir = TempDir::new().unwrap();
        let a = dir.path().join("a.xlsx");
        let b = dir.path().join("b.xlsx");
        touch(&a);
        touch(&b);

        let mut buf = Vec::new();
        writeln!(buf, "{}", a.display()).unwrap();
        writeln!(buf, "{}", b.display()).unwrap();
        // Blank line should be ignored.
        writeln!(buf).unwrap();
        let stdin = io::Cursor::new(buf);

        let result = collect(&[PathBuf::from("-")], stdin).unwrap();

        assert_eq!(result, vec![a, b]);
    }

    #[test]
    fn mixes_file_directory_and_stdin_in_a_single_call() {
        let dir = TempDir::new().unwrap();
        let explicit = dir.path().join("explicit.xlsx");
        let in_dir = dir.path().join("subdir/in_dir.xlsx");
        let from_stdin = dir.path().join("from_stdin.xlsx");
        touch(&explicit);
        touch(&in_dir);
        touch(&from_stdin);

        let mut buf = Vec::new();
        writeln!(buf, "{}", from_stdin.display()).unwrap();
        let stdin = io::Cursor::new(buf);

        let args = vec![
            explicit.clone(),
            dir.path().join("subdir"),
            PathBuf::from("-"),
        ];
        let result = collect(&args, stdin).unwrap();

        assert_eq!(sorted(result), sorted(vec![explicit, in_dir, from_stdin]));
    }
}
