# FormSpec fixtures and xlflow golden files

Each `NN-*.form.json` is a FormSpec; the `NN-*.golden.json` next to it is
what [xlflow](https://github.com/harumiWeb/xlflow) (MIT License, Copyright
(c) 2026 harumiWeb) produces for the same file with its pure-Go compiler, at
commit `2085f49` of its `main` branch. `tests/golden.rs` checks that
`ras_form::compile` writes exactly the same bytes.

A golden file holds either

* `streams`: every stream of the designer storage, keyed by its path below
  the form storage, as hex, and `storages`: every storage (`""` = the form
  storage) with its CLSID in wire byte order, or
* `error`: the message xlflow rejected the file with (fixtures `9x-*`); RAS
  must reject these files too.

The golden files were produced with a throwaway Go program placed inside an
xlflow checkout (so that its `internal/` packages can be imported); it is not
part of this repository. For every argument it does, in essence:

```go
body = strings.Replace(body, `"ras.userform"`, `"xlflow.userform"`, 1)
s, err := spec.ParseFormSpec(spec.SpecInput{Path: path, DisplayPath: path, Format: "json"}, body)
form, err := compiler.CompileNew(s, 932)
ser, err := oforms.SerializeForm(form, 932)
// write ser.Streams / ser.Storages (paths relative to the form storage, hex)
// or err.Error() to <name>.golden.json
```

`../data/excel-mutationform.bin` is the `xl/vbaProject.bin` of a workbook
saved by Excel 16 with a hand-made user form, taken from xlflow's test data
(`internal/vba/userforms/compiler/testdata/excel-authored/00-baseline.bin`).
