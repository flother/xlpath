//! Helpers shared by the integration-test binaries. Each test binary includes
//! this file via `mod common;`.

use std::io::Write;
use std::path::Path;

use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

/// Write a (fake) workbook to `path` containing the given (internal_name,
/// bytes) pairs. Minimal: we don't try to produce a valid OOXML package, only a
/// ZIP with the entries the test needs.
pub fn write_workbook(path: &Path, entries: &[(&str, &[u8])]) {
    let file = std::fs::File::create(path).unwrap();
    let mut zw = ZipWriter::new(file);
    let opts = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    for (name, data) in entries {
        zw.start_file(*name, opts).unwrap();
        zw.write_all(data).unwrap();
    }
    zw.finish().unwrap();
}

/// A trivially small but structurally real workbook containing two sheets.
pub const SIMPLE_WORKBOOK_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets>
    <sheet name="Alpha" sheetId="1"/>
    <sheet name="Beta" sheetId="2"/>
  </sheets>
</workbook>
"#;

/// A chart part with two distinct chart children — useful for testing
/// namespace-aware queries.
pub const CHART_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <c:chart>
    <c:plotArea>
      <c:barChart><c:grouping val="clustered"/></c:barChart>
    </c:plotArea>
  </c:chart>
  <c:chart>
    <c:plotArea>
      <c:lineChart/>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
"#;
