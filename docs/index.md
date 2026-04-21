---
icon: lucide/rocket
---

xlpath
======

xlpath is a CLI to query the XML within XLSX files using XPath.

Excel files are just XML in a trench coat. Beneath the `.xlsx` extension lies a zip archive
containing a collection of XML files. `xlpath` lets you run XPath expressions against that XML,
querying the files within an XLSX archive directly. The [output format](output.md) follows a
grep-like pattern by default (`file:part: value`), so it's a natural fit for standard shell tools.

Here's how to get the names of all sheets in a file named `Book1.xlsx`:

``` sh
xlpath '//x:sheet/@name' Book1.xlsx
```

This would output:

``` txt
Book1.xlsx:xl/workbook.xml: Sheet1
Book1.xlsx:xl/workbook.xml: Sheet2
Book1.xlsx:xl/workbook.xml: Sheet3
```

See the [usage reference](usage.md) and [examples](examples.md) to see what else is possible.

Install
-------

You can install `xlpath` from the GitHub repository with [Cargo](https://doc.rust-lang.org/stable/cargo/), the Rust package manager:

``` sh
cargo install --git https://github.com/flother/xlpath
```

Or clone and build from source:

``` sh
git clone https://github.com/flother/xlpath
cd xlpath
cargo install --path .
```

Licence
-------

`xlpath` is licensed under either the [MIT licence](https://opensource.org/license/MIT) or [Apache licence 2.0](https://opensource.org/license/Apache-2.0).
