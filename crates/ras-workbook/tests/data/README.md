# Test data

`edit.xlsm` is the workbook the tests of `ras_workbook::patch` (`ras book edit`) edit: cells with
formulas across sheets, a formula the stand-in does not work out (a VBA function), merged cells,
a conditional format, a validation, a filter, a note, a button with a macro, a hyperlink, defined
names, a chart, a table and a VBA project.

It is made by `edit.py` with [XlsxWriter](https://github.com/jmcnamara/XlsxWriter) 3.2.9. Its VBA
project is `xlsxwriter_vbaProject.bin` of `crates/ras-ovba/tests/fixtures` (BSD-2-Clause,
Copyright (c) 2013-2026 John McNamara; the licence text is there, `LICENSE-XlsxWriter.txt`).
