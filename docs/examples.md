---
icon: lucide/file-stack
---

Examples
========

These examples either assume you have a workbook named `workbook.xlsx` or a directory named
`workbooks`.

Sheet metadata
--------------

``` sh
# All sheet names in a workbook.
xlpath '//x:sheet/@name' workbook.xlsx

# All sheet names, showing the containing element for context.
xlpath '//x:sheet/@name' workbook.xlsx --tag
```

Formulas and cell values
------------------------

``` sh
# All formulas in a workbook's sheets.
xlpath '/x:worksheet/x:sheetData//x:c/x:f[text()]' --include 'xl/worksheets/sheet*.xml' workbook.xlsx

# All values in a workbook's sheets.
xlpath '/x:worksheet/x:sheetData//x:c/x:v' --include 'xl/worksheets/sheet*.xml' workbook.xlsx

# Number of formulas in each workbook in a folder, one line per file.
xlpath '/x:worksheet/x:sheetData//x:c/x:f[text()]' --include 'xl/worksheets/sheet*.xml' --count workbooks/
```

Themes and colours
------------------

``` sh
# Name of every theme used in a folder of workbooks.
xlpath '//a:themeElements/a:clrScheme/@name' --include 'xl/theme/*.xml' workbooks/

# Colours set in the theme.
xlpath '//a:themeElements/a:clrScheme/*/*/@val' --include 'xl/theme/*.xml' workbook.xlsx
```

Charts
------

``` sh
# Filenames for workbooks in workbooks/ that have at least one chart.
xlpath '/c:chartSpace' --include 'xl/charts/chart*.xml' --only-filenames workbooks/

# Every chart type used across a folder of workbooks, with a count of each.
xlpath \
    'name(//c:plotArea/*[contains(name(), "Chart")])' \
    workbooks/ \
    --include 'xl/charts/*.xml' \
    --tag \
    --no-filename \
    --no-part | sort | uniq -c | sort -rn
```

Defined names
-------------

``` sh
# Just the filenames of workbooks that define any named ranges.
xlpath //x:definedName workbooks/ --only-filenames
```

Comments
--------

``` sh
# Notes (old-school comments).
xlpath //x:comment//x:t workbooks/ --include 'xl/comments*.xml'
# Threaded comments (introduced in 2019).
xlpath //tc:threadedComment/tc:text workbooks/ --ns tc=http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments
```

Using with a database
---------------------

[DuckDB](https://duckdb.org/) can read `xlpath`'s
[newline-delimited JSON output](output.md#json-output) into a database table. Use `xlpath` to find
all the sheet names in a folder full of workbooks, and save the results to an ND-JSON file:

``` sh
xlpath '//x:sheet/@name' workbooks/ --json > sheet_names.ndjson
```

Then load that data into a table in an in-memory database:

``` sh
duckdb -cmd "CREATE TABLE sheet_names AS SELECT * FROM read_ndjson('sheet_names.ndjson')"
```

And then you can query the results using SQL. Here's an example query to find out which sheet names
are the most popular:

``` sql
SELECT
  Count(*)
    AS cnt,
  value
    AS sheet_name
FROM
  sheet_names
GROUP BY
  sheet_name
ORDER BY
  cnt DESC;
```
