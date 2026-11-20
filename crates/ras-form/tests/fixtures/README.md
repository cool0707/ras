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

Picture paths in the fixtures are relative to this directory (the test
compiles them with this directory as the project root). `assets/logo.bmp`
and `assets/logo.jpg` are xlflow's
`internal/vba/userforms/compiler/testdata/pictures-excel-authored/logo.*`.
`assets/logo.png` holds the pixels of `logo.bmp` with a transparent frame;
it is used by `tests/snapshot.rs` (xlflow has no conversion to compare
with: the test checks that a PNG gives the same form as the bitmap it is
converted to).

Fixtures only cover what xlflow accepts as well. What RAS accepts beyond
that — `tag` / `controlTipText` on every control, `list` / `selectedIndex`
of ComboBox and ListBox — is tested in `tests/snapshot.rs` by comparison
with the equivalent xlflow-compatible spec.

### Where RAS writes what Excel writes instead

xlflow writes a ComboBox and a ListBox without `ShowDropButtonWhen` and
`MatchEntry`, so they take the MS-OFORMS defaults: no drop button and no
auto-completion. Excel stores `ShowDropButtonWhen = 2` and `MatchEntry = 1`
for a new ComboBox and `MatchEntry = 0` and `ScrollBars = 3` for a new
ListBox (as the forms in `../data` show; MS-OFORMS requires the scroll bars
of a ListBox), and RAS writes the same. The `o` streams of
`03-all-controls.golden.json` and `12-bag-appearance.golden.json` differ
from xlflow's output by these fields only.

## Projects saved by Excel (`../data`)

`../data/<label>.bin` are `xl/vbaProject.bin` files of workbooks saved by
Excel 16, taken from xlflow's test data
(`internal/vba/userforms/compiler/testdata`):

| File | xlflow test data |
|---|---|
| `mutation.bin` | `excel-authored/00-baseline.bin` |
| `disabled.bin` | `excel-authored/04-all-disabled.bin` |
| `frame.bin` | `frame-excel-authored/baseline.bin` |
| `generation.bin` | `generation-excel-authored/baseline.bin` |
| `multipage.bin`, `multipage-disabled.bin`, `multipage-empty.bin` | `multipage-excel-authored/{baseline,disabled,empty}.bin` |
| `pictures.bin` | `pictures-excel-authored/baseline.bin` |
| `typed.bin` | `typed-properties-excel/vbaProject.bin` |

`../data/<label>.<Form>.projected.json` holds what xlflow makes of the form
`<Form>` of `<label>.bin`: `spec` is the result of `oforms.ReadForm` +
`projection.Project` (code page 932), `pictures` describes what
`picture.Decode` finds in the Image controls. `tests/snapshot.rs` checks
that `ras_form::project` gives the same FormSpec (apart from `kind` and the
language of the warning messages). RAS also projects the plain properties
of the controls (`src/scalars.rs`: `style`, `multiSelect`, `specialEffect`,
`orientation`, the bits of `VariousPropertyBits` ...) where they differ
from what Excel stores for a new control; xlflow lists them as
`unsupported`. The goldens have been changed to match: the names RAS
models are gone from `unsupported` (and so are the warnings of controls
left with none), and a ComboBox of `typed.bin`, written by xlflow, shows the
`matchEntry` and `showDropButtonWhen` it really has.

`../data/triple-state/` holds the streams (`f`, `o`, `CompObj` for
`\x01CompObj`, `VBFrame` for `\x03VBFrame`) of a form made in the VBA
editor of Excel 16 and saved: two OptionButtons and two CheckBoxes, the first
of each with TripleState True. It shows where Excel keeps TripleState, which
MS-OFORMS does not document: in the `MultiSelect` field of the MorphData
record (1 for True). Only the form's streams are kept, not the project.

`../data/excel-export.frx` is `example/maze-chase/src/forms/MazeChaseForm.frx`
of xlflow, a file written by `VBComponent.Export` of Excel 16.

## How the golden files were made

With a throwaway Go program placed inside an xlflow checkout at commit
`2085f49` (so that its `internal/` packages can be imported); it is not part
of this repository. In essence:

```go
// <name>.golden.json
body = strings.Replace(body, `"ras.userform"`, `"xlflow.userform"`, 1)
s, err := spec.ParseFormSpec(spec.SpecInput{Path: path, DisplayPath: path, Format: "json"}, body)
for each control with picture.path: asset, err := picture.Load(dir, path); control.Picture.Data = asset.Data
form, err := compiler.CompileNew(s, 932)
ser, err := oforms.SerializeForm(form, 932)
// write ser.Streams / ser.Storages (paths relative to the form storage, hex)
// or err.Error()

// <label>.<Form>.projected.json
container, err := cfb.Open(bin)
for _, name := range oforms.DiscoverForms(container) {
    form, err := oforms.ReadForm(container, name, 932)
    s, err := projection.Project(form)
    // write s as "spec"; for every Image control with a picture,
    // picture.Decode(control.Record.Pictures["Picture"]) as "pictures"
}
```
