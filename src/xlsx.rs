//! Opens OOXML zip archives, detects encryption, and iterates over internal XML
//! entries filtered by include/exclude globs.

use std::fs::File;
use std::io::{BufReader, Read, Seek, SeekFrom};
use std::path::Path;

use globset::{Glob, GlobSet, GlobSetBuilder};

use crate::error::SkipReason;

/// Compiled include/exclude globs, applied to zip-internal paths.
#[derive(Debug)]
pub struct EntryFilter {
    includes: Option<GlobSet>,
    excludes: Option<GlobSet>,
}

impl EntryFilter {
    /// Compile include/exclude globs. An empty `includes` list means "accept
    /// every XML-shaped entry"; anything in `excludes` is rejected regardless.
    pub fn new(includes: &[String], excludes: &[String]) -> Result<Self, globset::Error> {
        Ok(Self {
            includes: build(includes)?,
            excludes: build(excludes)?,
        })
    }

    /// Whether an entry at this zip-internal path should be processed.
    pub fn accepts(&self, entry: &str) -> bool {
        if let Some(excludes) = &self.excludes {
            if excludes.is_match(entry) {
                return false;
            }
        }
        match &self.includes {
            Some(includes) => includes.is_match(entry),
            None => true,
        }
    }
}

/// Open an OOXML workbook and stream XML-shaped entries to `on_entry`. Path
/// names passed to the callback are the zip-internal entry names (e.g.
/// `xl/workbook.xml`). The `filter` controls which entries are emitted; an
/// empty include list means "every XML-shaped entry".
///
/// Entries whose names do not end in `.xml` or `.rels` are skipped silently: we
/// assume binary parts (images, bins, ole objects) are uninteresting for XPath
/// querying.
pub fn process_entries<F>(
    path: &Path,
    filter: &EntryFilter,
    mut on_entry: F,
) -> Result<(), SkipReason>
where
    F: FnMut(&str, &[u8]),
{
    // Peek at the first bytes to spot OLE2-wrapped (i.e. encrypted) workbooks
    // before asking the zip parser to make sense of them. A single handle is
    // reused: we rewind after the peek so the zip reader starts from byte 0,
    // avoiding a second open() and the TOCTOU window between the two calls.
    let mut file = File::open(path)?;
    let mut header = [0u8; 8];
    let n = read_up_to(&mut file, &mut header)?;
    if is_ole2_header(&header[..n]) {
        return Err(SkipReason::Encrypted);
    }
    file.seek(SeekFrom::Start(0))?;
    let mut archive = zip::ZipArchive::new(BufReader::new(file))
        .map_err(|e| SkipReason::CorruptZip(e.to_string()))?;

    let mut buf = Vec::new();
    for i in 0..archive.len() {
        let mut entry = archive
            .by_index(i)
            .map_err(|e| SkipReason::CorruptZip(e.to_string()))?;
        if entry.is_dir() {
            continue;
        }
        let name = entry.name().to_string();
        if !is_xml_entry(&name) {
            continue;
        }
        if !filter.accepts(&name) {
            continue;
        }
        buf.clear();
        entry
            .read_to_end(&mut buf)
            .map_err(|e| SkipReason::CorruptZip(e.to_string()))?;
        on_entry(&name, &buf);
    }
    Ok(())
}

fn read_up_to(reader: &mut impl Read, buf: &mut [u8]) -> std::io::Result<usize> {
    let mut filled = 0;
    while filled < buf.len() {
        match reader.read(&mut buf[filled..]) {
            Ok(0) => break,
            Ok(n) => filled += n,
            Err(e) if e.kind() == std::io::ErrorKind::Interrupted => continue,
            Err(e) => return Err(e),
        }
    }
    Ok(filled)
}

fn is_xml_entry(name: &str) -> bool {
    let lower = name.to_ascii_lowercase();
    lower.ends_with(".xml") || lower.ends_with(".rels")
}

/// Eight-byte signature at the start of every OLE2 compound document. Microsoft
/// Office's password-encrypted OOXML files are wrapped in such a document, so
/// detecting this header lets us report the file as encrypted rather than
/// surfacing a confusing "not a zip archive" error.
const OLE2_MAGIC: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

fn is_ole2_header(bytes: &[u8]) -> bool {
    bytes.len() >= OLE2_MAGIC.len() && bytes[..OLE2_MAGIC.len()] == OLE2_MAGIC
}

fn build(patterns: &[String]) -> Result<Option<GlobSet>, globset::Error> {
    if patterns.is_empty() {
        return Ok(None);
    }
    let mut b = GlobSetBuilder::new();
    for p in patterns {
        b.add(Glob::new(p)?);
    }
    Ok(Some(b.build()?))
}

#[cfg(test)]
mod tests {
    use std::io::Write;
    use std::path::Path;

    use tempfile::TempDir;
    use zip::write::SimpleFileOptions;
    use zip::CompressionMethod;
    use zip::ZipWriter;

    use super::{process_entries, EntryFilter};
    use crate::error::SkipReason;

    fn write_zip(path: &Path, entries: &[(&str, &[u8])]) {
        let file = std::fs::File::create(path).unwrap();
        let mut zw = ZipWriter::new(file);
        let opts = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        for (name, data) in entries {
            zw.start_file(*name, opts).unwrap();
            zw.write_all(data).unwrap();
        }
        zw.finish().unwrap();
    }

    #[test]
    fn accepts_entries_matching_an_include_glob() {
        let filter = EntryFilter::new(&["xl/charts/*.xml".into()], &[]).unwrap();

        assert!(filter.accepts("xl/charts/chart1.xml"));
        assert!(!filter.accepts("xl/worksheets/sheet1.xml"));
    }

    #[test]
    fn empty_includes_means_accept_all_entries() {
        let filter = EntryFilter::new(&[], &[]).unwrap();

        assert!(filter.accepts("xl/workbook.xml"));
        assert!(filter.accepts("xl/charts/chart1.xml"));
        assert!(filter.accepts("_rels/.rels"));
    }

    #[test]
    fn excludes_override_includes() {
        let filter = EntryFilter::new(&["xl/**/*.xml".into()], &["xl/charts/**".into()]).unwrap();

        assert!(filter.accepts("xl/workbook.xml"));
        assert!(!filter.accepts("xl/charts/chart1.xml"));
    }

    #[test]
    fn excludes_apply_even_with_no_includes() {
        let filter = EntryFilter::new(&[], &["**/_rels/*".into()]).unwrap();

        assert!(filter.accepts("xl/workbook.xml"));
        assert!(!filter.accepts("xl/_rels/workbook.xml.rels"));
    }

    #[test]
    fn supports_double_star_globs() {
        let filter = EntryFilter::new(&["**/*.xml".into()], &[]).unwrap();

        assert!(filter.accepts("xl/workbook.xml"));
        assert!(filter.accepts("xl/charts/chart1.xml"));
        assert!(filter.accepts("top.xml"));
    }

    #[test]
    fn iterates_xml_entries_in_a_workbook() {
        let tmp = TempDir::new().unwrap();
        let path = tmp.path().join("book.xlsx");
        write_zip(
            &path,
            &[
                ("xl/workbook.xml", b"<workbook/>"),
                ("xl/media/image1.png", &[0x89, b'P', b'N', b'G']),
                ("xl/charts/chart1.xml", b"<chart/>"),
            ],
        );

        let filter = EntryFilter::new(&[], &[]).unwrap();
        let mut seen: Vec<(String, Vec<u8>)> = Vec::new();
        process_entries(&path, &filter, |name, data| {
            seen.push((name.to_string(), data.to_vec()));
        })
        .unwrap();

        seen.sort_by(|a, b| a.0.cmp(&b.0));
        assert_eq!(
            seen,
            vec![
                ("xl/charts/chart1.xml".to_string(), b"<chart/>".to_vec()),
                ("xl/workbook.xml".to_string(), b"<workbook/>".to_vec()),
            ]
        );
    }

    #[test]
    fn applies_entry_filter() {
        let tmp = TempDir::new().unwrap();
        let path = tmp.path().join("book.xlsx");
        write_zip(
            &path,
            &[
                ("xl/workbook.xml", b"<workbook/>"),
                ("xl/charts/chart1.xml", b"<chart/>"),
            ],
        );

        let filter = EntryFilter::new(&["xl/charts/*.xml".into()], &[]).unwrap();
        let mut seen: Vec<String> = Vec::new();
        process_entries(&path, &filter, |name, _| seen.push(name.to_string())).unwrap();

        assert_eq!(seen, vec!["xl/charts/chart1.xml".to_string()]);
    }

    #[test]
    fn reports_corrupt_zip() {
        let tmp = TempDir::new().unwrap();
        let path = tmp.path().join("broken.xlsx");
        std::fs::write(&path, b"this is not a zip file").unwrap();

        let filter = EntryFilter::new(&[], &[]).unwrap();
        let err = process_entries(&path, &filter, |_, _| {}).unwrap_err();

        assert!(matches!(err, SkipReason::CorruptZip(_)));
    }

    #[test]
    fn reports_encrypted_ole2_wrapped_workbook() {
        let tmp = TempDir::new().unwrap();
        let path = tmp.path().join("locked.xlsx");
        // Just the OLE2 header is enough — we detect encryption before the zip
        // parser is ever invoked.
        let mut contents = super::OLE2_MAGIC.to_vec();
        contents.extend_from_slice(&[0u8; 512]);
        std::fs::write(&path, contents).unwrap();

        let filter = EntryFilter::new(&[], &[]).unwrap();
        let err = process_entries(&path, &filter, |_, _| {}).unwrap_err();

        assert!(matches!(err, SkipReason::Encrypted));
    }

    #[test]
    fn ole2_magic_is_recognised() {
        // Password-protected OOXML workbooks are packaged as OLE2 compound
        // documents, whose files begin with this eight-byte signature.
        let ole2 = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0x00, 0x00];
        let zip = [b'P', b'K', 0x03, 0x04, 0x14, 0x00, 0x00, 0x00];

        assert!(super::is_ole2_header(&ole2));
        assert!(!super::is_ole2_header(&zip));
        assert!(!super::is_ole2_header(&[0xD0, 0xCF])); // too short
        assert!(!super::is_ole2_header(&[])); // empty
    }
}
