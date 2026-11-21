# Test data of `ras-form`

## FormSpec fixtures (this directory)

Each `NN-*.form.json` is a FormSpec. `tests/fixtures.rs` checks that

* a valid one (`01`–`14`) is written, read back with `workspace::pull` and
  written again with the same designer storage (nothing is lost or changed
  by reading), and
* an invalid one (`9x-error-*`) is rejected with the diagnostic listed in
  `ERRORS` there.

Picture paths in the fixtures are relative to this directory (the test
copies `assets/` next to the form). `assets/logo.bmp` and `assets/logo.jpg`
are xlflow's
`internal/vba/userforms/compiler/testdata/pictures-excel-authored/logo.*`
([xlflow](https://github.com/harumiWeb/xlflow), MIT License, Copyright (c)
2026 harumiWeb). `assets/logo.png` holds the pixels of `logo.bmp` with a
transparent frame; `tests/snapshot.rs` checks that a PNG gives the same form
as the bitmap it is converted to.

What every single property writes and reads is tested in `src/scalars.rs`;
the check against the real Excel is `tools/excel-check` at the root of the
repository.

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

`../data/<label>.<Form>.expected.json` holds what `ras_form::snapshot` reads
from the form `<Form>` of `<label>.bin` (code page 932): `spec`, the FormSpec
with its warnings, and `pictures`, the pictures it extracts.
`tests/snapshot.rs` checks the reading against them, and that the spec read
is written back with the same design. After a deliberate change to the
reading, write them again with

```sh
RAS_UPDATE_EXPECTED=1 cargo test -p ras-form --test snapshot forms_saved_by_excel
```

and check the change with `git diff`.

These forms show what Excel stores, which MS-OFORMS does not always say and
RAS writes the same way:

* A new ComboBox has `ShowDropButtonWhen = 2` and `MatchEntry = 1`; a new
  ListBox `MatchEntry = 0` and `ScrollBars = 3` (MS-OFORMS requires the scroll
  bars of a ListBox).
* The text of a disabled Label, TextBox, CommandButton, ComboBox, CheckBox or
  OptionButton is grayed in its font: the TextProps hold `FontEffects`
  `0x40002000` (`CFE_DISABLED | CFE_AUTOCOLOR`, `disabled.bin`; not for a
  ListBox or TabStrip). RAS does the same for a ToggleButton, which the data
  does not have.

`../data/triple-state/` holds the streams (`f`, `o`, `CompObj` for
`\x01CompObj`, `VBFrame` for `\x03VBFrame`) of a form made in the VBA
editor of Excel 16 and saved: two OptionButtons and two CheckBoxes, the first
of each with TripleState True. It shows where Excel keeps TripleState, which
MS-OFORMS does not document: in the `MultiSelect` field of the MorphData
record (1 for True). Only the form's streams are kept, not the project.

`../data/excel-export.frx` is `example/maze-chase/src/forms/MazeChaseForm.frx`
of xlflow, a file written by `VBComponent.Export` of Excel 16.
