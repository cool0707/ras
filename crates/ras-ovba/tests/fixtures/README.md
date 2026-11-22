# Test fixtures

Genuine VBA projects saved by Microsoft Office (with compiled p-code), used
to test `pack` / `unpack` on workbooks Excel saved. None of them was created
by this repository; all come from projects whose licences permit
redistribution. The licence texts are next to this file. The tests of the
project itself (`vbaProject.bin`) moved to umya-spreadsheet's `vba` module
together with the code, and these fixtures with two more (a project saved by
Excel for Mac and one with user forms, from oletools) to its
`tests/test_files/vba`.

| File | Origin | Licence |
|---|---|---|
| `xlsxwriter_vbaProject.bin` | `examples/vbaProject.bin` of [XlsxWriter](https://github.com/jmcnamara/XlsxWriter) 3.3.1 (PyPI sdist), unmodified. Saved by Excel for Windows: code page 1252, `_VBA_PROJECT` version 0x0085, modules `ThisWorkbook`, `Sheet1`, `Module1`, `ThisWorkbook1`, `Sheet2`. | BSD-2-Clause, Copyright (c) 2013-2026 John McNamara — `LICENSE-XlsxWriter.txt` |
| `rust_xlsxwriter_macro04_signed.xlsm` | `tests/input/macro04.xlsm` of the rust_xlsxwriter 0.79.0 crate, unmodified. A workbook whose VBA project is digitally signed (`xl/vbaProjectSignature.bin`). | MIT OR Apache-2.0, Copyright 2022-2024 John McNamara — `LICENSE-rust_xlsxwriter-MIT.txt` |

SHA-256:

```
0ced1464b3677e98f5e3a8c5d80135e18dc98dca39299f1a8cfd2a00999fbf9f  xlsxwriter_vbaProject.bin
d49e30414b59e66446689d77119fa5b9f4f99fa86dbcade003f4d3fe1b8e225d  rust_xlsxwriter_macro04_signed.xlsm
```
