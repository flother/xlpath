---
icon: lucide/code-xml
---

XML namespaces
==============

Every XML element within an XLSX archive belongs to a namespace, and XPath 1.0 has no syntax for
selecting an unprefixed name in a namespace. A query like `//workbook` therefore matches nothing.

When you run `xlpath`, the most widely-used XLSX namespaces are already registered, so queries like
`//x:workbook` or `//c:chart` work out of the box.

If you need a prefix that isn't listed below, register one with `--ns` (for example,
`--ns foo=http://example.com/foo`). Your namespaces are applied after the defaults, so reusing a
built-in prefix with `--ns` overrides it.

Pre-registered prefixes
-----------------------

| Prefix | URI                                                                         |
| ------ | --------------------------------------------------------------------------- |
| `x`    | `http://schemas.openxmlformats.org/spreadsheetml/2006/main`                 |
| `r`    | `http://schemas.openxmlformats.org/officeDocument/2006/relationships`       |
| `c`    | `http://schemas.openxmlformats.org/drawingml/2006/chart`                    |
| `a`    | `http://schemas.openxmlformats.org/drawingml/2006/main`                     |
| `xdr`  | `http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing`       |
| `mc`   | `http://schemas.openxmlformats.org/markup-compatibility/2006`               |
| `rel`  | `http://schemas.openxmlformats.org/package/2006/relationships`              |
| `ct`   | `http://schemas.openxmlformats.org/package/2006/content-types`              |
| `x14`  | `http://schemas.microsoft.com/office/spreadsheetml/2009/9/main`             |
| `x15`  | `http://schemas.microsoft.com/office/spreadsheetml/2010/11/main`            |
| `xr`   | `http://schemas.microsoft.com/office/spreadsheetml/2014/revision`           |
| `xp`   | `http://schemas.openxmlformats.org/officeDocument/2006/extended-properties` |
