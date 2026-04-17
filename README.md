xlpath
======

Query Excel (OOXML) XLSX files with XPath.

`xlpath` lets you play with a workbook as it really is: a bunch of XML files in a ZIP archive. You
can run a XPath 1.0 expression against that XML. The output is grep-like (`file:part: value`) by
default, so it composes with the usual shell tools for surveying feature usage across folders of
workbooks.

Install
-------

Install from the GitHub repository with `cargo`:

```text
cargo install --git https://github.com/flother/xlpath
```

Or clone and build from source:

```text
git clone https://github.com/flother/xlpath
cd xlpath
cargo install --path .
```

`xlpath` is all Rust, no system `libxml2` required.

Usage
-----

```text
xlpath <XPATH> [PATH]... [OPTIONS]
```

Each `PATH` may be a file or a directory (recursed). Pass `-` to read newline-separated paths from
`stdin`. If no `PATH` is given, `xlpath` walks the current working directory. Supported extensions:
`.xlsx`, `.xlsm`, `.xltx`, `.xltm`. Encrypted workbooks (OLE2 compound documents) are skipped with a
warning on `stderr`.

### Examples

```sh
# All sheet names in a workbook.
xlpath '//x:sheet/@name' workbook.xlsx

# All sheet names, showing the containing element for context.
xlpath '//x:sheet/@name' workbook.xlsx --tag

# All formulas in a workbook's sheets.
xlpath '/x:worksheet/x:sheetData//x:c/x:f[text()]' --include 'xl/worksheets/sheet*.xml' workbook.xlsx

# All values in a workbook's sheets.
xlpath '/x:worksheet/x:sheetData//x:c/x:v' --include 'xl/worksheets/sheet*.xml' workbook.xlsx

# Name of every theme used in a folder of workbooks.
xlpath '//a:themeElements/a:clrScheme/@name' --include 'xl/theme/*.xml' .

# Colours set in the theme.
xlpath '//a:themeElements/a:clrScheme/*/*/@val' --include 'xl/theme/*.xml' workbook.xlsx

# Filenames for workbooks in the current directory that have at least one chart.
xlpath '/c:chartSpace' --include 'xl/charts/chart*.xml' --only-filenames .

# Every chart type used across a folder of workbooks, with a count of each.
xlpath '//c:plotArea/*' . --include 'xl/charts/*.xml' --tag --no-filename --no-part \
  | sort | uniq -c

# One count line per file for the total number of relationship IDs.
xlpath '//@r:id' *.xlsx --count

# Just the filenames of workbooks that define any named ranges.
find . -name '*.xlsx' | xlpath '//x:definedName' - --only-filenames

# Output results as newline-delimited JSON
xlpath '//x:sheet/@name' workbook.xlsx --json
```

Options
-------

| Flag                      | Purpose                                                                                                                                                                                                                            |
| ------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `--include <GLOB>`        | Only evaluate matching zip-internal paths. Repeatable.                                                                                                                                                                             |
| `--exclude <GLOB>`        | Skip matching zip-internal paths. Repeatable.                                                                                                                                                                                      |
| `--ns <PREFIX=URI>`       | Register (or override) a namespace prefix. Repeatable.                                                                                                                                                                             |
| `--default-ns <PREFIX>`   | Bind the document's default `xmlns` to this prefix.                                                                                                                                                                                |
| `-c`, `--count`           | Print `file:N` per matching workbook instead of each match.                                                                                                                                                                        |
| `--only-filenames`        | Print only the names of workbooks with at least one match.                                                                                                                                                                         |
| `--tag`                   | Add the matching element's synthetic self-closing tag to the output prefix (e.g. `<x:sheet name="A" sheetId="1"/>`). For element matches, the tag is the element itself; for attribute and text matches, it is the parent element. |
| `--no-filename`           | Omit the filename from each output line.                                                                                                                                                                                           |
| `--no-part`               | Omit the zip-internal part path from each output line.                                                                                                                                                                             |
| `--json`                  | Output one newline-delimited JSON object per match                                                                                                                                                                                 |
| `-L`, `--follow`          | Follow symbolic links when walking directories. Off by default.                                                                                                                                                                    |
| `-j <N>`, `--threads <N>` | Worker threads (defaults to logical CPUs). `-j 1` forces deterministic output order.                                                                                                                                               |

`--count`, `--only-filenames`, and `--json` are mutually exclusive. `--tag`, `--no-filename`, and
`--no-part` conflict with `--count` and `--only-filenames` (which emit one line per file rather than
one per match), but are compatible with `--json`.

The synthetic tag emitted by `--tag` is a reporting artefact, not a round-trippable XML fragment: it
is always self-closing, carries no `xmlns` declarations, and uses the canonical prefix from `xlpath`
's namespace registry rather than whatever prefix the document itself declared.

### Output format

The default output format is to show one match per line à la grep:

```text
filename:part: value
```

You can use flags to control the output. The `filename` and `part` sections can be suppressed, and
you can add a section for the matching XML element's tag (or parent element for attribute and text
matches):

```text
filename:part: value          (default)
filename:part:<tag/>: value   (--tag)
part: value                   (--no-filename)
part:<tag/>: value            (--no-filename --tag)
filename: value               (--no-part)
filename:<tag/>: value        (--no-part --tag)
value                         (--no-filename --no-part)
<tag/>: value                 (--no-filename --no-part --tag)
```

Namespaces
----------

These prefixes are pre-registered for every query, so queries like `//c:chart` or `//x:sheet/@name`
work out of the box:

| Prefix      | URI                                                                   |
| ----------- | --------------------------------------------------------------------- |
| `x`, `main` | `http://schemas.openxmlformats.org/spreadsheetml/2006/main`           |
| `r`         | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` |
| `c`         | `http://schemas.openxmlformats.org/drawingml/2006/chart`              |
| `a`         | `http://schemas.openxmlformats.org/drawingml/2006/main`               |
| `xdr`       | `http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing` |
| `mc`        | `http://schemas.openxmlformats.org/markup-compatibility/2006`         |
| `rel`       | `http://schemas.openxmlformats.org/package/2006/relationships`        |
| `ct`        | `http://schemas.openxmlformats.org/package/2006/content-types`        |
| `x14`       | `http://schemas.microsoft.com/office/spreadsheetml/2009/9/main`       |
| `x15`       | `http://schemas.microsoft.com/office/spreadsheetml/2010/11/main`      |
| `xr`        | `http://schemas.microsoft.com/office/spreadsheetml/2014/revision`     |

User `--ns` pairs apply after the defaults and take precedence.

### The default-namespace quirk

OOXML files usually declare a default namespace on the root element. For example, `xl/workbook.xml`
has `xmlns=".../spreadsheetml/2006/main"`. XPath 1.0 has no way to select an unprefixed default
namespace, so `//workbook` matches nothing. `xlpath` offers two ways to make your life a little
easier:

1. Use a pre-registered prefix like `//x:workbook`. This asks for `workbook` in the `spreadsheetml`
   URI specifically, because `x` is pre-registered to that URI (see table above). If you need a
   namespace not included in the table, extend using `--ns`.
2. Pass `--default-ns x` together with `//x:workbook`. For each document, `xlpath` binds that
   document's root `xmlns` (whatever URI it happens to be) to `x` before evaluation.

Use the first option when you care which namespace you're matching; pick the second when you don't
know the URI (custom files, embedded XML) or want a single query that survives heterogeneous
documents.

Exit codes
----------

`xlpath` follows grep conventions, with errors taking precedence:

| Code | Meaning                                                                                 |
| ---- | --------------------------------------------------------------------------------------- |
| `0`  | At least one match found                                                                |
| `1`  | No matches                                                                              |
| `2`  | One or more files could not be processed (corrupt, encrypted, malformed XML, I/O error) |

Per-file errors are reported on `stderr` and do not stop the run (i.e. other inputs continue to be
processed).

Licence
-------

MIT OR Apache-2.0.
