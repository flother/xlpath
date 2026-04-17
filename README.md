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
xlpath <XPATH> <PATH>... [OPTIONS]
```

Each `PATH` may be a file or a directory (recursed). Pass `-` to read newline-separated paths from
`stdin`. Supported extensions: `.xlsx`, `.xlsm`, `.xltx`, `.xltm`. Encrypted workbooks (OLE2
compound documents) are skipped with a warning on `stderr`.

### Examples

```sh
# All sheet names in a workbook, with each match annotated by its XPath location.
xlpath '//x:sheet/@name' workbook.xlsx --with-path

# All formulas in a workbook's sheets.
xlpath '/x:worksheet/x:sheetData//x:c/x:f[text()]' --include 'xl/worksheets/sheet*.xml' workbook.xml

# All values in a workbook's sheets.
xlpath '/x:worksheet/x:sheetData//x:c/x:v' --include 'xl/worksheets/sheet*.xml' workbook.xml

# Name of every theme used in a folder of workbooks.
xlpath '//a:themeElements/a:clrScheme/@name' --include 'xl/theme/*.xml' .

# Colours set in the theme.
xlpath '//a:themeElements/a:clrScheme/*/*/@val' --include 'xl/theme/*.xml' workbook.xlsx

# Filenames for workbooks in the current directory that have at least one chart.
xlpath '/c:chartSpace' --include 'xl/charts/chart*.xml' --files-only .

# Every chart type used across a folder of workbooks.
xlpath '//c:plotArea/*' . --include 'xl/charts/*.xml' --tag --tag-only \
  | sort | uniq -c

# One count line per file for the total number of relationship IDs.
xlpath '//@r:id' *.xlsx --count

# Just the filenames of workbooks that define any named ranges.
find . -name '*.xlsx' | xlpath '//x:definedName' - --files-only
```

Options
-------

| Flag                      | Purpose                                                                                                             |
| ------------------------- | ------------------------------------------------------------------------------------------------------------------- |
| `--include <GLOB>`        | Only evaluate matching zip-internal paths. Repeatable.                                                              |
| `--exclude <GLOB>`        | Skip matching zip-internal paths. Repeatable.                                                                       |
| `--ns <PREFIX=URI>`       | Register (or override) a namespace prefix. Repeatable.                                                              |
| `--default-ns <PREFIX>`   | Bind the document's default `xmlns` to this prefix.                                                                 |
| `--count`                 | Print `file:N` per matching workbook instead of each match.                                                         |
| `--files-only`            | Print only the names of workbooks with at least one match.                                                          |
| `--with-path`             | Append an XPath-like location to each match.                                                                        |
| `--tag`                   | Render element matches as a self-closing opening tag (e.g. `<c:lineChart val="1"/>`) instead of their text content. |
| `--tag-only`              | Print only the match value on each line, with no `file:part:` prefix. Requires `--tag`.                             |
| `-j <N>`, `--threads <N>` | Worker threads (defaults to logical CPUs). `-j 1` forces deterministic output order.                                |

`--count` and `--files-only` are mutually exclusive. `--tag` only changes how element matches
render; it is silently ignored under `--count` and `--files-only` (which don't emit per-match lines)
and is a no-op for attribute, text, and atomic matches. `--tag-only` requires `--tag` and conflicts
with `--count`, `--files-only`, and `--with-path`.

The synthetic tag emitted by `--tag` is a reporting artefact, not a round-trippable XML fragment: it
is always self-closing, carries no `xmlns` declarations, and uses the canonical prefix from `xlpath`
's namespace registry rather than whatever prefix the document itself declared.

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
