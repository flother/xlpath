mod common;

use std::io::Write;
use std::process::Command;

use assert_cmd::prelude::*;
use predicates::prelude::*;
use tempfile::TempDir;

use common::{write_workbook, CHART_XML, SIMPLE_WORKBOOK_XML};

fn xlpath() -> Command {
    Command::cargo_bin("xlpath").unwrap()
}

#[test]
fn extracts_sheet_names_via_preregistered_prefix() {
    let tmp = TempDir::new().unwrap();
    let wb = tmp.path().join("book.xlsx");
    write_workbook(&wb, &[("xl/workbook.xml", SIMPLE_WORKBOOK_XML.as_bytes())]);

    xlpath()
        .args(["//x:sheet/@name"])
        .arg(&wb)
        .assert()
        .success()
        .stdout(predicate::str::contains("xl/workbook.xml: Alpha"))
        .stdout(predicate::str::contains("xl/workbook.xml: Beta"));
}

#[test]
fn includes_filter_narrows_to_chart_parts() {
    let tmp = TempDir::new().unwrap();
    let wb = tmp.path().join("charts.xlsx");
    write_workbook(
        &wb,
        &[
            ("xl/workbook.xml", SIMPLE_WORKBOOK_XML.as_bytes()),
            ("xl/charts/chart1.xml", CHART_XML.as_bytes()),
        ],
    );

    let out = xlpath()
        .args(["//c:chart", "--include", "xl/charts/*.xml"])
        .arg(&wb)
        .output()
        .unwrap();

    assert!(
        out.status.success(),
        "stderr: {:?}",
        String::from_utf8_lossy(&out.stderr)
    );
    let stdout = String::from_utf8_lossy(&out.stdout);
    // Two <c:chart> nodes in chart1.xml, none anywhere else (workbook.xml is
    // excluded by include glob).
    assert_eq!(stdout.matches("xl/charts/chart1.xml:").count(), 2);
    assert!(!stdout.contains("xl/workbook.xml"));
}

#[test]
fn count_mode_reports_per_file_totals() {
    let tmp = TempDir::new().unwrap();
    let wb = tmp.path().join("charts.xlsx");
    write_workbook(&wb, &[("xl/charts/chart1.xml", CHART_XML.as_bytes())]);

    xlpath()
        .args(["//c:chart", "--count"])
        .arg(&wb)
        .assert()
        .success()
        .stdout(predicate::str::contains("charts.xlsx:2"));
}

#[test]
fn only_filenames_mode_prints_one_line_per_matching_workbook() {
    let tmp = TempDir::new().unwrap();
    let a = tmp.path().join("a.xlsx");
    let b = tmp.path().join("b.xlsx");
    write_workbook(&a, &[("xl/charts/chart1.xml", CHART_XML.as_bytes())]);
    write_workbook(&b, &[("xl/workbook.xml", SIMPLE_WORKBOOK_XML.as_bytes())]);

    let out = xlpath()
        .args(["//c:chart", "--only-filenames"])
        .arg(&a)
        .arg(&b)
        .output()
        .unwrap();

    assert!(out.status.success());
    let stdout = String::from_utf8_lossy(&out.stdout);
    assert!(stdout.contains("a.xlsx"));
    assert!(!stdout.contains("b.xlsx"));
}

#[test]
fn no_match_returns_exit_code_1() {
    let tmp = TempDir::new().unwrap();
    let wb = tmp.path().join("book.xlsx");
    write_workbook(&wb, &[("xl/workbook.xml", SIMPLE_WORKBOOK_XML.as_bytes())]);

    let out = xlpath().args(["//x:nothing"]).arg(&wb).output().unwrap();
    assert_eq!(out.status.code(), Some(1));
    assert!(out.stdout.is_empty());
}

#[test]
fn corrupt_zip_emits_stderr_warning_and_exit_code_2() {
    let tmp = TempDir::new().unwrap();
    let bad = tmp.path().join("broken.xlsx");
    std::fs::write(&bad, b"not a zip").unwrap();
    let good = tmp.path().join("good.xlsx");
    write_workbook(
        &good,
        &[("xl/workbook.xml", SIMPLE_WORKBOOK_XML.as_bytes())],
    );

    let out = xlpath()
        .args(["//x:sheet/@name"])
        .arg(&bad)
        .arg(&good)
        .output()
        .unwrap();

    assert_eq!(out.status.code(), Some(2));
    let stdout = String::from_utf8_lossy(&out.stdout);
    let stderr = String::from_utf8_lossy(&out.stderr);
    assert!(stdout.contains("good.xlsx:xl/workbook.xml: Alpha"));
    assert!(stderr.contains("broken.xlsx"));
}

#[test]
fn encrypted_workbook_is_reported_specifically() {
    let tmp = TempDir::new().unwrap();
    let enc = tmp.path().join("locked.xlsx");
    let mut bytes = vec![0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    bytes.extend_from_slice(&[0u8; 512]);
    std::fs::write(&enc, bytes).unwrap();

    xlpath()
        .args(["//x:sheet/@name"])
        .arg(&enc)
        .assert()
        .code(2)
        .stderr(predicate::str::contains("encrypted workbook"));
}

#[test]
fn lock_files_inside_a_directory_are_skipped() {
    let tmp = TempDir::new().unwrap();
    let real = tmp.path().join("real.xlsx");
    let lock = tmp.path().join("~$real.xlsx");
    write_workbook(
        &real,
        &[("xl/workbook.xml", SIMPLE_WORKBOOK_XML.as_bytes())],
    );
    std::fs::write(&lock, b"broken content that would error if read").unwrap();

    let out = xlpath()
        .args(["//x:sheet/@name"])
        .arg(tmp.path())
        .output()
        .unwrap();

    assert!(
        out.status.success(),
        "stderr: {:?}",
        String::from_utf8_lossy(&out.stderr)
    );
    let stderr = String::from_utf8_lossy(&out.stderr);
    assert!(!stderr.contains("~$real.xlsx"));
}

#[test]
fn dash_reads_paths_from_stdin() {
    let tmp = TempDir::new().unwrap();
    let wb = tmp.path().join("piped.xlsx");
    write_workbook(&wb, &[("xl/workbook.xml", SIMPLE_WORKBOOK_XML.as_bytes())]);

    let mut child = xlpath()
        .args(["//x:sheet/@name", "-"])
        .stdin(std::process::Stdio::piped())
        .stdout(std::process::Stdio::piped())
        .spawn()
        .unwrap();
    {
        let stdin = child.stdin.as_mut().unwrap();
        writeln!(stdin, "{}", wb.display()).unwrap();
    }
    let out = child.wait_with_output().unwrap();
    assert!(
        out.status.success(),
        "stderr: {:?}",
        String::from_utf8_lossy(&out.stderr)
    );
    let stdout = String::from_utf8_lossy(&out.stdout);
    assert!(stdout.contains("piped.xlsx:xl/workbook.xml: Alpha"));
}

#[test]
fn tag_mode_puts_element_tag_in_prefix() {
    let tmp = TempDir::new().unwrap();
    let wb = tmp.path().join("charts.xlsx");
    write_workbook(&wb, &[("xl/charts/chart1.xml", CHART_XML.as_bytes())]);

    let out = xlpath()
        .args(["//c:plotArea/*", "--include", "xl/charts/*.xml", "--tag"])
        .arg(&wb)
        .output()
        .unwrap();

    assert!(
        out.status.success(),
        "stderr: {:?}",
        String::from_utf8_lossy(&out.stderr)
    );
    let stdout = String::from_utf8_lossy(&out.stdout);
    // Tag appears in the prefix (before ': '), not as the value.
    assert!(
        stdout.contains("xl/charts/chart1.xml:<c:barChart/>:"),
        "stdout was: {stdout}"
    );
    assert!(
        stdout.contains("xl/charts/chart1.xml:<c:lineChart/>:"),
        "stdout was: {stdout}"
    );
}

#[test]
fn no_filename_flag_omits_file_path() {
    let tmp = TempDir::new().unwrap();
    let wb = tmp.path().join("book.xlsx");
    write_workbook(&wb, &[("xl/workbook.xml", SIMPLE_WORKBOOK_XML.as_bytes())]);

    let out = xlpath()
        .args(["//x:sheet/@name", "--no-filename"])
        .arg(&wb)
        .output()
        .unwrap();

    assert!(out.status.success());
    let stdout = String::from_utf8_lossy(&out.stdout);
    assert!(
        stdout.contains("xl/workbook.xml: Alpha"),
        "stdout was: {stdout}"
    );
    assert!(!stdout.contains("book.xlsx"), "stdout was: {stdout}");
}

#[test]
fn no_part_flag_omits_part_path() {
    let tmp = TempDir::new().unwrap();
    let wb = tmp.path().join("book.xlsx");
    write_workbook(&wb, &[("xl/workbook.xml", SIMPLE_WORKBOOK_XML.as_bytes())]);

    let out = xlpath()
        .args(["//x:sheet/@name", "--no-part"])
        .arg(&wb)
        .output()
        .unwrap();

    assert!(out.status.success());
    let stdout = String::from_utf8_lossy(&out.stdout);
    assert!(stdout.contains("book.xlsx: Alpha"), "stdout was: {stdout}");
    assert!(!stdout.contains("xl/workbook.xml"), "stdout was: {stdout}");
}

#[test]
fn no_filename_and_no_part_prints_bare_values() {
    let tmp = TempDir::new().unwrap();
    let wb = tmp.path().join("book.xlsx");
    write_workbook(&wb, &[("xl/workbook.xml", SIMPLE_WORKBOOK_XML.as_bytes())]);

    let out = xlpath()
        .args(["//x:sheet/@name", "--no-filename", "--no-part"])
        .arg(&wb)
        .output()
        .unwrap();

    assert!(out.status.success());
    let stdout = String::from_utf8_lossy(&out.stdout);
    let lines: Vec<&str> = stdout.lines().collect();
    assert_eq!(lines, vec!["Alpha", "Beta"], "stdout was: {stdout}");
}

#[test]
fn no_filename_no_part_with_tag_shows_parent_element_then_value() {
    let tmp = TempDir::new().unwrap();
    let wb = tmp.path().join("book.xlsx");
    write_workbook(&wb, &[("xl/workbook.xml", SIMPLE_WORKBOOK_XML.as_bytes())]);

    let out = xlpath()
        .args(["//x:sheet/@name", "--no-filename", "--no-part", "--tag"])
        .arg(&wb)
        .output()
        .unwrap();

    // --tag on an attribute match shows the parent element's synthetic tag.
    assert!(out.status.success());
    let stdout = String::from_utf8_lossy(&out.stdout);
    assert!(
        stdout.contains("<x:sheet") && stdout.contains(": Alpha"),
        "stdout was: {stdout}"
    );
}

#[test]
fn json_mode_outputs_one_ndjson_object_per_match() {
    let tmp = TempDir::new().unwrap();
    let wb = tmp.path().join("book.xlsx");
    write_workbook(&wb, &[("xl/workbook.xml", SIMPLE_WORKBOOK_XML.as_bytes())]);

    let out = xlpath()
        .args(["//x:sheet/@name", "--json"])
        .arg(&wb)
        .output()
        .unwrap();

    assert!(out.status.success());
    let stdout = String::from_utf8_lossy(&out.stdout);
    let lines: Vec<&str> = stdout.lines().collect();
    assert_eq!(
        lines.len(),
        2,
        "expected one line per match; got:\n{stdout}"
    );
    assert!(
        lines[0].contains(r#""part":"xl/workbook.xml""#),
        "stdout was: {stdout}"
    );
    assert!(
        lines[0].contains(r#""value":"Alpha""#),
        "stdout was: {stdout}"
    );
    assert!(
        lines[1].contains(r#""value":"Beta""#),
        "stdout was: {stdout}"
    );
}

#[test]
fn json_conflicts_with_count() {
    let tmp = TempDir::new().unwrap();
    let wb = tmp.path().join("book.xlsx");
    write_workbook(&wb, &[("xl/workbook.xml", SIMPLE_WORKBOOK_XML.as_bytes())]);

    xlpath()
        .args(["//x:sheet/@name", "--json", "--count"])
        .arg(&wb)
        .assert()
        .failure();
}

#[test]
fn json_conflicts_with_only_filenames() {
    let tmp = TempDir::new().unwrap();
    let wb = tmp.path().join("book.xlsx");
    write_workbook(&wb, &[("xl/workbook.xml", SIMPLE_WORKBOOK_XML.as_bytes())]);

    xlpath()
        .args(["//x:sheet/@name", "--json", "--only-filenames"])
        .arg(&wb)
        .assert()
        .failure();
}

#[test]
fn long_help_lists_preregistered_ooxml_namespaces() {
    // A representative sample of the OOXML defaults — every prefix in this list
    // must appear with its URI in `--help`, so that users can discover the
    // registered namespaces without having to read the source.
    let expected: &[(&str, &str)] = &[
        (
            "x",
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        ),
        (
            "c",
            "http://schemas.openxmlformats.org/drawingml/2006/chart",
        ),
        (
            "r",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        ),
        (
            "ct",
            "http://schemas.openxmlformats.org/package/2006/content-types",
        ),
        (
            "xr",
            "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
        ),
    ];

    let out = xlpath().arg("--help").output().unwrap();
    assert!(out.status.success());
    let stdout = String::from_utf8_lossy(&out.stdout);
    for (prefix, uri) in expected {
        assert!(
            stdout.contains(prefix) && stdout.contains(uri),
            "expected `{prefix}` → `{uri}` in --help output; got:\n{stdout}"
        );
    }
}

#[test]
fn ns_flag_registers_user_namespace() {
    let tmp = TempDir::new().unwrap();
    let wb = tmp.path().join("custom.xlsx");
    write_workbook(
        &wb,
        &[(
            "xl/custom.xml",
            br#"<r:thing xmlns:r="urn:example:custom"><r:name>ok</r:name></r:thing>"#,
        )],
    );

    xlpath()
        .args(["//my:name", "--ns", "my=urn:example:custom"])
        .arg(&wb)
        .assert()
        .success()
        .stdout(predicate::str::contains("xl/custom.xml: ok"));
}

#[test]
fn omitted_path_defaults_to_current_directory() {
    let tmp = TempDir::new().unwrap();
    let wb = tmp.path().join("book.xlsx");
    write_workbook(&wb, &[("xl/workbook.xml", SIMPLE_WORKBOOK_XML.as_bytes())]);

    xlpath()
        .args(["//x:sheet/@name"])
        .current_dir(tmp.path())
        .assert()
        .success()
        .stdout(predicate::str::contains("book.xlsx"))
        .stdout(predicate::str::contains(": Alpha"));
}
