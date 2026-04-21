---
icon: lucide/square-terminal
---

Usage
=====

The command takes an XPath expression as its first argument, followed by one or more paths and any
options.

``` sh
xlpath <XPATH> [PATH]... [OPTIONS]
```

Arguments
---------

### `XPATH`

An XPath 1.0 expression. See this [XPath cheatsheet](https://devhints.io/xpath) for supported
syntax.

Because XLSX files use XML namespaces heavily, most useful expressions require a namespace prefix.
See the [namespaces documentation](namespaces.md) for more information.

### `PATH`

A file or directory (recursed). Pass as many as you want. Pass `-` to read newline-separated paths
from `stdin`. If no `PATH` is given, `xlpath` walks the current working directory.

Supported extensions:

- `.xlsx`
- `.xlsm`
- `.xltx`
- `.xltm`

Encrypted workbooks are skipped with a warning on `stderr`. Binary workbooks (`.xlsb`) are not
supported.

Options
-------

### `--include <GLOB>`

Only search zip-internal XML documents whose path matches the glob. Repeatable.

An `.xlsx` file is itself a zip archive containing XML files at paths such as
`xl/worksheets/sheet1.xml`, `xl/charts/chart1.xml`, and `xl/styles.xml`. The glob is matched against
these internal paths.

For example, `--include 'xl/worksheets/*.xml'` restricts the search to sheet XML and skips
everything else (shared strings, styles, relationships, etc.).

### `--exclude <GLOB>`

Skip zip-internal XML documents whose path matches the glob. Repeatable.

### `--ns <PREFIX=URI>`

Register (or override) a namespace prefix. Repeatable. For example:

``` sh
xlpath --ns 'x14=http://schemas.microsoft.com/office/spreadsheetml/2009/9/main' '//x14:sparklineGroups'
```

The common Excel XML prefixes are pre-registered; see [namespaces](namespaces.md) for more
information.

### `--count`, `-c`

Print `file:N` per matching workbook, where `N` is the total number of matches across all XML
documents in that workbook. Produces one output line per file rather than one per match.

### `--only-filenames`

Print only the names of workbooks with at least one match. Produces one output line per file.

### `--tag`

For element matches, include a synthetic self-closing opening tag in the output (`<x:sheet
name="Sheet1" sheetId="1"/>` for example). Useful for identifying which element matched and
seeing its attributes at a glance. For attribute and text matches, the tag is the parent element. No
effect on XPath function-result matches.

The tag is a reporting aid, not a round-trippable XML fragment: it is always self-closing, carries
no `xmlns` declarations, and uses the canonical prefix from `xlpath`'s namespace registry rather
than whatever prefix the document itself declared.

### `--no-filename`

Omit the filename from each output line.

### `--no-part`

Omit the zip-internal document path from each output line.

### `--json`

Output one newline-delimited JSON object per match. Each object can have `file`, `part`, `tag`, and
`value` string fields, as controlled by the `--no-filename`, `--no-part`, and `--tag` flags.

### `--follow`, `-L`

Follow symbolic links when walking directories. Off by default.

### `--threads <N>`, `-j <N>`

Worker threads (defaults to logical CPUs). `-j 1` forces deterministic output order.

Flag compatibility
------------------

| Flag               | `--count` | `--only-filenames` | `--json` |
| ------------------ | :-------: | :----------------: | :------: |
| `--tag`            |     ✗     |         ✗          |    ✓     |
| `--no-filename`    |     ✗     |         ✗          |    ✓     |
| `--no-part`        |     ✗     |         ✗          |    ✓     |
| `--count`          |     —     |         ✗          |    ✗     |
| `--only-filenames` |     ✗     |         —          |    ✗     |

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
