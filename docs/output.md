---
icon: lucide/save
---

Output format
=============

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

JSON output
-----------

Passing `--json` switches the output to
[newline-delimited JSON](https://github.com/ndjson/ndjson-spec): one JSON object per match, one
match per line. Each object always has a `value` field, and can optionally have `file`, `part`, and
`tag` fields.

- By default `file` and `part` are included
- `tag` is only included when `--tag` is passed
- `--no-filename` drops `file`
- `--no-part` drops `part`

Default:

``` json
{"file":"workbook.xlsx","part":"xl/workbook.xml","value":"Sheet1"}
```

With `--tag`:

``` json
{"file":"workbook.xlsx","part":"xl/workbook.xml","tag":"<x:sheet name=\"Sheet1\" sheetId=\"1\"/>","value":"Sheet1"}
```

With `--no-filename` and `--no-part`:

``` json
{"value":"Sheet1"}
```

All values are strings. `--json` cannot be combined with `--count` or `--only-filenames`.

CSV output
----------

`xlpath` can't produce CSV directly, but you can combine its ND-JSON output with
[`jq`](https://jqlang.org/)'s `@csv` filter. Pick the fields you want, put them in an array, and
`jq` takes care of quoting and escaping:

``` sh
xlpath //x:sheet/@name workbooks/ --json | jq -r "[.file, .part, .value] | @csv"
```

```text
"workbooks/report.xlsx","xl/workbook.xml","Summary"
"workbooks/report.xlsx","xl/workbook.xml","Data"
"workbooks/sales.xlsx","xl/workbook.xml","Q1"
```

To include a header row, use `jq`'s slurp mode (`-s`) so the header can be emitted once before the
data rows:

``` sh
xlpath //x:sheet/@name workbooks/ --json \
    | jq -rs '(["file","part","value"] | @csv), (.[] | [.file, .part, .value] | @csv)'
```
