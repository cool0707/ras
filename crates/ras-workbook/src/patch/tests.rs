use std::collections::{BTreeMap, HashSet};
use std::io::Read as _;

use serde_json::json;

use super::*;
use crate::book::Read;

/// A workbook made by XlsxWriter (`tests/data/edit.py`).
const BOOK: &[u8] = include_bytes!("../../tests/data/edit.xlsm");

fn ops(json: &str) -> Vec<Op> {
    parse_ops(&serde_json::from_str(json).unwrap(), Some("Data")).unwrap()
}

fn edit(json: &str) -> Vec<u8> {
    apply(BOOK, &ops(json)).unwrap()
}

fn refused(json: &str) -> String {
    apply(BOOK, &ops(json)).unwrap_err()
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut out = BTreeMap::new();
    for i in 0..zip.len() {
        let mut file = zip.by_index(i).unwrap();
        let mut data = Vec::new();
        file.read_to_end(&mut data).unwrap();
        out.insert(file.name().to_string(), data);
    }
    out
}

fn part(bytes: &[u8], name: &str) -> String {
    String::from_utf8(
        parts(bytes)
            .remove(name)
            .unwrap_or_else(|| panic!("no {name}")),
    )
    .unwrap()
}

/// The parts that differ between two files.
fn changed(a: &[u8], b: &[u8]) -> Vec<String> {
    let (a, b) = (parts(a), parts(b));
    let names: HashSet<&String> = a.keys().chain(b.keys()).collect();
    let mut out: Vec<String> = names
        .into_iter()
        .filter(|n| a.get(*n) != b.get(*n))
        .cloned()
        .collect();
    out.sort();
    out
}

/// The cells of a range as text: formulas as they are typed, values as `Val`s print.
fn cells(bytes: &[u8], sheet: &str, range: &str, what: Read) -> Vec<Vec<String>> {
    let mut book = Book::open_bytes(bytes).unwrap();
    let si = book.find_sheet(sheet).unwrap();
    book.get(si, RangeRef::parse(range).unwrap(), what)
        .unwrap()
        .into_iter()
        .map(|row| {
            row.into_iter()
                .map(|v| match v {
                    Val::Empty => String::new(),
                    Val::Num(n) | Val::Date(n) => format!("{n}"),
                    Val::Text(s) => s,
                    Val::Bool(b) => b.to_string(),
                    Val::Err(e) => format!("#{e}"),
                })
                .collect()
        })
        .collect()
}

fn cell(bytes: &[u8], sheet: &str, at: &str, what: Read) -> String {
    cells(bytes, sheet, at, what).remove(0).remove(0)
}

const SHEET1: &str = "xl/worksheets/sheet1.xml";
const SHEET2: &str = "xl/worksheets/sheet2.xml";

#[test]
fn edits_are_read_from_json() {
    assert_eq!(
        ops(
            r#"[{"op":"set","range":"B2:B3","text":"=A1"},{"op":"clear","sheet":"S","range":"$A$1"}]"#
        ),
        [
            Op::Set {
                sheet: "Data".into(),
                range: RangeRef::parse("B2:B3").unwrap(),
                text: "=A1".into()
            },
            Op::Clear {
                sheet: "S".into(),
                range: RangeRef::parse("A1").unwrap()
            },
        ]
    );
    // Whole rows move down and whole columns right, whatever is asked; cells as asked.
    let shifts = ops(
        r#"[{"op":"insert","range":"3:4","shift":"right"},{"op":"delete","range":"B:C"},{"op":"insert","range":"B2:C3","shift":"right"},{"op":"delete","range":"B2"}]"#,
    );
    assert!(matches!(shifts[0], Op::Insert { right: false, .. }));
    assert!(matches!(shifts[1], Op::Delete { left: true, .. }));
    assert!(matches!(shifts[2], Op::Insert { right: true, .. }));
    assert!(matches!(shifts[3], Op::Delete { left: false, .. }));
    let bad = |json: &str| parse_ops(&serde_json::from_str(json).unwrap(), None).unwrap_err();
    assert!(bad(r#"{"op":"set","range":"A1","text":"x"}"#).contains("シート"));
    assert!(bad(r#"{"op":"nope","sheet":"S"}"#).contains("nope"));
    assert!(bad(r#"{"op":"clear","sheet":"S","range":"A0"}"#).contains("A0"));
}

#[test]
fn a_value_typed_in_changes_its_cell_and_the_formulas_that_read_it() {
    let out = edit(r#"{"op":"set","range":"C2","text":"7"}"#);
    assert_eq!(cell(&out, "Data", "C2", Read::Value), "7");
    assert_eq!(cell(&out, "Data", "D2", Read::Value), "840");
    assert_eq!(cell(&out, "Data", "D6", Read::Value), "3840");
    // On the other sheet too, through a reference and through a name.
    assert_eq!(cell(&out, "Summary", "A1", Read::Value), "7680");
    assert_eq!(cell(&out, "Summary", "A3", Read::Value), "3841");
    // Nothing else of the file is written again: the VBA project, the drawings, the styles,
    // the shared strings, the sheet nothing changed on.
    assert_eq!(changed(BOOK, &out), [SHEET1, SHEET2]);
    let sheet = part(&out, SHEET1);
    // The cells that did not change are as they were written (their shared strings, formats).
    assert!(sheet.contains(r#"<c r="A2" t="s"><v>4</v></c><c r="B2" s="1"><v>120</v></c><c r="C2"><v>7</v></c><c r="D2" s="1"><f>B2*C2</f><v>840</v></c>"#), "{sheet}");
    // What is outside the cells is byte for byte the same.
    let before = part(BOOK, SHEET1);
    assert_eq!(
        sheet.split("</sheetData>").nth(1),
        before.split("</sheetData>").nth(1)
    );
    // The same value again: the file is not written.
    assert_eq!(edit(r#"{"op":"set","range":"C2","text":"3"}"#), BOOK);
}

#[test]
fn a_formula_the_stand_in_does_not_work_out_is_kept() {
    // E2 is `=MyUdf(D2)`, a function of the VBA project, saved with the value 42.
    let out = edit(r#"{"op":"set","range":"C2","text":"7"}"#);
    assert!(part(&out, SHEET1).contains(r#"<c r="E2"><f>MyUdf(D2)</f><v>42</v></c>"#));
    // It moves with its cell, and so do its references.
    let out = edit(r#"{"op":"insert","range":"2:2"}"#);
    let sheet = part(&out, SHEET1);
    assert!(
        sheet.contains(r#"<c r="E3"><f>MYUDF(D3)</f><v>42</v></c>"#),
        "{sheet}"
    );
    // Typed over and cleared, it is gone.
    let out = edit(r#"{"op":"set","range":"E2","text":"42"}"#);
    assert!(part(&out, SHEET1).contains(r#"<c r="E2"><v>42</v></c>"#));
    let out = edit(r#"{"op":"clear","range":"E2:E3"}"#);
    assert!(!part(&out, SHEET1).contains("MyUdf"));
}

#[test]
fn text_dates_and_formulas_are_read_as_excel_reads_them() {
    let out = edit(
        r#"[{"op":"set","range":"A12","text":"a<b & \"c\""},{"op":"set","range":"B12","text":"2026/4/1"},{"op":"set","range":"C12:C13","text":"=B2*2"},{"op":"set","range":"D12","text":"50%"},{"op":"set","range":"E12","text":"TRUE"}]"#,
    );
    let sheet = part(&out, SHEET1);
    assert!(sheet.contains(r#"<c r="A12" t="inlineStr"><is><t xml:space="preserve">a&lt;b &amp; &quot;c&quot;</t></is></c>"#), "{sheet}");
    assert_eq!(cell(&out, "Data", "A12", Read::Value), "a<b & \"c\"");
    // Line breaks, and a carriage return, are the cell's as they were typed.
    let lines = edit(r#"{"op":"set","range":"A12","text":" a\r\nb\n c "}"#);
    assert_eq!(cell(&lines, "Data", "A12", Read::Value), " a\r\nb\n c ");
    // A date gets the date format: a format added to the styles, made from the cell's.
    assert_eq!(cell(&out, "Data", "B12", Read::NumberFormat), "m/d/yyyy");
    assert_eq!(cell(&out, "Data", "B12", Read::Value2), "46113");
    let styles = part(&out, "xl/styles.xml");
    assert!(styles.contains(r#"<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>"#), "{styles}");
    // The same entry in every cell of the range, its relative references moved.
    assert_eq!(
        cells(&out, "Data", "C12:C13", Read::Formula),
        [["=B2*2"], ["=B3*2"]]
    );
    assert_eq!(
        cells(&out, "Data", "C12:C13", Read::Value),
        [["240"], ["160"]]
    );
    assert_eq!(cell(&out, "Data", "D12", Read::Value), "0.5");
    assert!(
        sheet.contains(r#"<c r="E12" t="b"><v>1</v></c>"#),
        "{sheet}"
    );
    assert!(changed(BOOK, &out).contains(&"xl/styles.xml".to_string()));
}

#[test]
fn rows_inserted_move_everything_below_them() {
    let out = edit(r#"{"op":"insert","range":"3:4"}"#);
    assert_eq!(
        cells(&out, "Data", "A2:D7", Read::Formula),
        [
            ["りんご", "120", "3", "=B2*C2"],
            ["", "", "", ""],
            ["", "", "", ""],
            ["みかん", "80", "10", "=B5*C5"],
            ["ぶどう", "500", "2", "=B6*C6"],
            ["もも", "300", "4", "=B7*C7"],
        ]
    );
    assert_eq!(cell(&out, "Data", "D8", Read::Formula), "=SUM(D2:D7)");
    assert_eq!(cell(&out, "Summary", "A1", Read::Formula), "=Data!D8*2");
    assert_eq!(
        cell(&out, "Summary", "A2", Read::Formula),
        "=SUM(Data!D2:D7)"
    );
    let sheet = part(&out, SHEET1);
    for expected in [
        r#"<dimension ref="A1:E12"/>"#,
        // The rows inserted are formatted as the row above them.
        r#"<row r="3"><c r="B3" s="1"/><c r="D3" s="1"/></row>"#,
        // A row keeps its height where it goes.
        r#"<row r="9" ht="30" customHeight="1"/>"#,
        r#"<autoFilter ref="A1:D7"/>"#,
        r#"<mergeCell ref="A10:C10"/>"#,
        r#"<conditionalFormatting sqref="D2:D7">"#,
        r#"sqref="C2:C7""#,
        r#"<hyperlink ref="A12" r:id="rId1"/>"#,
    ] {
        assert!(sheet.contains(expected), "{expected}: {sheet}");
    }
    let workbook = part(&out, "xl/workbook.xml");
    assert!(
        workbook.contains(">Data!$A$2:$A$7</definedName>"),
        "{workbook}"
    );
    assert!(workbook.contains(">Data!$D$8</definedName>"), "{workbook}");
    // The filter's own name too.
    assert!(
        workbook.contains(">Data!$A$1:$D$7</definedName>"),
        "{workbook}"
    );
    // The note goes with its cell; the chart's bottom edge and its references follow.
    assert!(part(&out, "xl/comments1.xml").contains(r#"<comment ref="B5""#));
    let vml = part(&out, "xl/drawings/vmlDrawing1.vml");
    assert!(vml.contains("<x:Row>4</x:Row>"), "{vml}");
    let chart = part(&out, "xl/charts/chart1.xml");
    assert!(chart.contains("<c:f>Data!$D$2:$D$7</c:f>"), "{chart}");
    assert!(part(&out, "xl/drawings/drawing1.xml").contains("<xdr:row>16</xdr:row>"));
    // The VBA project and the other sheet's table are not touched.
    let touched = changed(BOOK, &out);
    assert!(
        !touched
            .iter()
            .any(|p| p.contains("vbaProject") || p.contains("tables/")),
        "{touched:?}"
    );
}

#[test]
fn rows_deleted_take_what_belongs_to_them() {
    // Row 3 has the note; row 10 the hyperlink.
    let out = edit(r#"[{"op":"delete","range":"10:10"},{"op":"delete","range":"3:3"}]"#);
    assert_eq!(
        cells(&out, "Data", "A2:D5", Read::Formula),
        [
            ["りんご", "120", "3", "=B2*C2"],
            ["ぶどう", "500", "2", "=B3*C3"],
            ["もも", "300", "4", "=B4*C4"],
            ["合計", "", "", "=SUM(D2:D4)"],
        ]
    );
    assert_eq!(cell(&out, "Data", "D5", Read::Value), "2560");
    assert_eq!(cell(&out, "Summary", "A3", Read::Value), "2561");
    let sheet = part(&out, SHEET1);
    assert!(!sheet.contains("<hyperlink"), "{sheet}");
    assert!(sheet.contains(r#"<mergeCell ref="A7:C7"/>"#), "{sheet}");
    assert!(!part(&out, "xl/comments1.xml").contains("<comment "));
    let vml = part(&out, "xl/drawings/vmlDrawing1.vml");
    assert!(!vml.contains("ObjectType=\"Note\""), "{vml}");
    // The button stays.
    assert!(vml.contains("ObjectType=\"Button\""), "{vml}");
}

#[test]
fn columns_inserted_and_deleted() {
    let out = edit(r#"{"op":"insert","range":"B:B"}"#);
    assert_eq!(
        cells(&out, "Data", "A2:E2", Read::Formula),
        [["りんご", "", "120", "3", "=C2*D2"]]
    );
    let sheet = part(&out, SHEET1);
    // The column inserted is as wide as the one before it; the others keep their widths.
    assert!(sheet.contains(r#"<cols><col min="1" max="2" width="14.7109375" customWidth="1"/><col min="5" max="5" width="12.7109375" style="1" customWidth="1"/></cols>"#), "{sheet}");
    assert!(sheet.contains(r#"<autoFilter ref="A1:E5"/>"#), "{sheet}");
    assert!(sheet.contains(r#"<mergeCell ref="A8:D8"/>"#), "{sheet}");
    assert!(part(&out, "xl/workbook.xml").contains(">Data!$E$6</definedName>"));

    let out = edit(r#"{"op":"delete","range":"C:C"}"#);
    assert_eq!(cell(&out, "Data", "C2", Read::Formula), "=B2*#REF!");
    let sheet = part(&out, SHEET1);
    // The validation was of the column deleted.
    assert!(!sheet.contains("<dataValidation"), "{sheet}");
    assert!(
        sheet.contains(r#"<conditionalFormatting sqref="C2:C5">"#),
        "{sheet}"
    );
    assert!(
        sheet.contains(r#"<col min="3" max="3" width="12.7109375" style="1" customWidth="1"/>"#),
        "{sheet}"
    );
}

#[test]
fn cells_inserted_and_deleted_move_their_row_or_column_only() {
    let out = edit(r#"{"op":"delete","range":"A2:A3","shift":"left"}"#);
    assert_eq!(
        cells(&out, "Data", "A2:D4", Read::Formula),
        [
            ["120", "3", "=A2*B2", "42"],
            ["80", "10", "=A3*B3", ""],
            ["ぶどう", "500", "2", "=B4*C4"],
        ]
    );
    assert!(part(&out, SHEET1).contains(r#"<c r="D2"><f>MYUDF(C2)</f><v>42</v></c>"#));
    let out = edit(r#"{"op":"insert","range":"E2:E3","shift":"down"}"#);
    assert_eq!(cell(&out, "Data", "D2", Read::Formula), "=B2*C2");
    assert!(part(&out, SHEET1).contains(r#"<c r="E4"><f>MyUdf(D2)</f><v>42</v></c>"#));
    // What Excel refuses is refused, and says why.
    assert!(refused(r#"{"op":"insert","range":"B2:C3","shift":"down"}"#).contains("結合"));
}

#[test]
fn cells_copied_moved_and_filled() {
    let out = edit(
        r#"[{"op":"copy","from":"A1:D2","to":"A12"},{"op":"fill","range":"D5:D7"},{"op":"values","at":"G10","rows":[["a","1"],["=G10&\"x\"",""]]},{"op":"clear","range":"A10"}]"#,
    );
    assert_eq!(
        cells(&out, "Data", "A12:D13", Read::Formula),
        [
            ["品名", "単価", "数量", "金額"],
            ["りんご", "120", "3", "=B13*C13"]
        ]
    );
    // The formats are pasted with the cells: those of the cells copied.
    let sheet = part(&out, SHEET1);
    assert!(
        sheet.contains(r#"<c r="A12" s="2" t="inlineStr">"#),
        "{sheet}"
    );
    assert!(
        sheet.contains(r#"<c r="D13" s="1"><f>B13*C13</f><v>360</v></c>"#),
        "{sheet}"
    );
    assert_eq!(
        cells(&out, "Data", "D6:D7", Read::Formula),
        [["=B6*C6"], ["=B7*C7"]]
    );
    assert_eq!(
        cells(&out, "Data", "G10:H11", Read::Value),
        [["a", "1"], ["ax", ""]]
    );
    assert_eq!(cell(&out, "Data", "A10", Read::Value), "");

    // A range pasted over a multiple of itself is repeated.
    let out = edit(r#"{"op":"copy","from":"C2:C3","to":"G2:H5"}"#);
    assert_eq!(
        cells(&out, "Data", "G2:H5", Read::Value),
        [["3", "3"], ["10", "10"], ["3", "3"], ["10", "10"]]
    );

    // Cut and paste: the references to the cells moved follow them, and so do the names.
    let out = edit(r#"{"op":"move","from":"A6:D6","to":"A20"}"#);
    assert_eq!(cell(&out, "Data", "D20", Read::Formula), "=SUM(D2:D5)");
    assert_eq!(cell(&out, "Data", "D6", Read::Formula), "");
    assert_eq!(cell(&out, "Summary", "A1", Read::Formula), "=Data!D20*2");
    assert!(part(&out, "xl/workbook.xml").contains(">Data!$D$20</definedName>"));
    assert_eq!(cell(&out, "Summary", "A3", Read::Value), "3361");

    // Pasted on another sheet.
    let out =
        edit(r#"{"op":"copy","sheet":"Summary","fromSheet":"Data","from":"A2:B3","to":"D1"}"#);
    assert_eq!(
        cells(&out, "Summary", "D1:E2", Read::Value),
        [["りんご", "120"], ["みかん", "80"]]
    );
}

#[test]
fn what_the_file_cannot_follow_is_refused() {
    // A table's columns are listed in its part: rows may come and go, columns not.
    let table = |json: &str| {
        apply(
            BOOK,
            &parse_ops(&serde_json::from_str(json).unwrap(), Some("Tbl")).unwrap(),
        )
    };
    let out = table(r#"{"op":"insert","range":"4:4"}"#).unwrap();
    assert!(part(&out, "xl/tables/table1.xml").contains(r#"ref="B2:D6""#));
    let out = table(r#"{"op":"insert","range":"A:A"}"#).unwrap();
    assert!(part(&out, "xl/tables/table1.xml").contains(r#"ref="C2:E5""#));
    assert!(
        table(r#"{"op":"insert","range":"C:C"}"#)
            .unwrap_err()
            .contains("テーブル")
    );
    assert!(
        table(r#"{"op":"delete","range":"2:2"}"#)
            .unwrap_err()
            .contains("テーブル")
    );
    assert!(refused(r#"{"op":"set","sheet":"Nope","range":"A1","text":"x"}"#).contains("Nope"));
    assert!(refused(r#"{"op":"set","range":"A:C","text":"x"}"#).contains("大きすぎ"));
}

#[test]
fn shared_formulas_are_written_out_when_their_sheet_is() {
    // The second sheet with a formula shared by three cells, as Excel writes filled formulas.
    let sheet = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B3"/><sheetData><row r="1" spans="1:2"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2</f><v>2</v></c></row><row r="2" spans="1:2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row><row r="3" spans="1:2"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>6</v></c></row></sheetData></worksheet>"#;
    let mut pack = Pack::open(BOOK).unwrap();
    let book = pack
        .rewrite(
            &BTreeMap::from([(SHEET2.to_string(), sheet.as_bytes().to_vec())]),
            &HashSet::new(),
        )
        .unwrap();
    let edit = |json: &str| {
        apply(
            &book,
            &parse_ops(&serde_json::from_str(json).unwrap(), Some("Summary")).unwrap(),
        )
        .unwrap()
    };
    let out = edit(r#"{"op":"set","range":"A2","text":"5"}"#);
    let new = part(&out, SHEET2);
    assert!(!new.contains("shared"), "{new}");
    assert!(
        new.contains(r#"<c r="B1"><f>A1*2</f><v>2</v></c>"#),
        "{new}"
    );
    assert!(
        new.contains(r#"<c r="A2"><v>5</v></c><c r="B2"><f>A2*2</f><v>10</v></c>"#),
        "{new}"
    );
    let out = edit(r#"{"op":"insert","range":"2:2"}"#);
    assert_eq!(
        cells(&out, "Summary", "B1:B4", Read::Formula),
        [["=A1*2"], [""], ["=A3*2"], ["=A4*2"]]
    );
}

#[test]
fn positions_after_an_insertion_or_a_deletion() {
    let at = |a: &str| CellRef::parse(a).unwrap();
    let range = |a: &str| RangeRef::parse(a).unwrap();
    let rows = edit_of(range("3:4"), false, true);
    assert_eq!(map_cell(at("B2"), &rows), Some(at("B2")));
    assert_eq!(map_cell(at("B3"), &rows), Some(at("B5")));
    let gone = edit_of(range("B:C"), true, false);
    assert_eq!(map_cell(at("A9"), &gone), Some(at("A9")));
    assert_eq!(map_cell(at("C9"), &gone), None);
    assert_eq!(map_cell(at("E9"), &gone), Some(at("C9")));
    // Cells: only within the columns (rows) of the range.
    let down = edit_of(range("B2:C3"), false, true);
    assert_eq!(map_cell(at("B2"), &down), Some(at("B4")));
    assert_eq!(map_cell(at("D2"), &down), Some(at("D2")));
    assert_eq!(map_cell(at("C1"), &down), Some(at("C1")));
    let left = edit_of(range("B2:C3"), true, false);
    assert_eq!(map_cell(at("C3"), &left), None);
    assert_eq!(map_cell(at("E3"), &left), Some(at("C3")));
    assert_eq!(map_cell(at("E4"), &left), Some(at("E4")));
}

#[test]
fn formats_are_added_to_the_styles_and_cells_keep_what_they_hold() {
    let out = edit(
        r##"[{"op":"format","range":"A2:B3","set":{"bold":true,"fill":"#FFFF00","color":"#FF0000","border":"all","align":"center","fontSize":14}},{"op":"format","range":"C2","set":{"numberFormat":"0.00"}},{"op":"format","range":"E:E","set":{"italic":true}}]"##,
    );
    assert_eq!(changed(BOOK, &out), ["xl/styles.xml", SHEET1]);
    let sheet = part(&out, SHEET1);
    // Only the format of the cells changes: their shared strings and formulas are as written.
    assert!(sheet.contains(r#"<c r="A2" t="s" s="4"><v>4</v></c><c r="B2" s="5"><v>120</v></c><c r="C2" s="6"><v>3</v></c>"#), "{sheet}");
    assert!(
        sheet.contains(r#"<c r="E2" s="7"><f>MyUdf(D2)</f><v>42</v></c>"#),
        "{sheet}"
    );
    // A whole column: its format in `<cols>`, without a width of its own.
    assert!(
        sheet.contains(r#"<col min="5" max="5" width="9.140625" style="7"/>"#),
        "{sheet}"
    );
    let styles = part(&out, "xl/styles.xml");
    for expected in [
        r#"<font><b/><sz val="14"/><color rgb="FFFF0000"/>"#,
        r#"<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>"#,
        r#"<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>"#,
        r#"<cellXfs count="8">"#,
        // The number format alone: made from the format the cell had.
        r#"<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>"#,
        r#"applyAlignment="1"><alignment horizontal="center"/></xf>"#,
    ] {
        assert!(styles.contains(expected), "{expected}: {styles}");
    }
    // What the workbook reads back.
    let mut book = Book::open_bytes(&out).unwrap();
    let a2 = RangeRef::parse("A2:B3").unwrap();
    assert_eq!(
        book.format_of(0, a2, "font.bold").unwrap(),
        Some(Val::Bool(true))
    );
    assert_eq!(
        book.format_of(0, a2, "font.size").unwrap(),
        Some(Val::Num(14.0))
    );
    assert_eq!(
        book.format_of(0, a2, "interior.color").unwrap(),
        Some(Val::Num(f64::from(0x00FFFFu32)))
    );
    assert_eq!(cell(&out, "Data", "C2", Read::NumberFormat), "0.00");
    // The same formats again use the formats added; cleared, the cells have the first one.
    let again = apply(
        &out,
        &ops(r#"[{"op":"format","range":"A4","set":{"italic":true}},{"op":"clearFormats","range":"A2:B2"}]"#),
    )
    .unwrap();
    let sheet = part(&again, SHEET1);
    assert!(
        sheet.contains(r#"<c r="A2" t="s" s="0"><v>4</v></c><c r="B2" s="0"><v>120</v></c>"#),
        "{sheet}"
    );
    assert!(
        sheet.contains(r#"<c r="A4" t="s" s="7"><v>6</v></c>"#),
        "{sheet}"
    );
    assert!(part(&again, "xl/styles.xml").contains(r#"<cellXfs count="8">"#));
}

#[test]
fn a_size_shows_hidden_lines_and_zero_hides_them() {
    let out =
        edit(r#"[{"op":"hide","range":"5:5","hidden":true},{"op":"size","range":"B:B","px":0}]"#);
    let sheet = part(&out, SHEET1);
    assert!(sheet.contains(r#"<row r="5" hidden="1">"#), "{sheet}");
    assert!(
        sheet.contains(r#"hidden="1""#) && sheet.contains(r#"<col min="2" max="2""#),
        "{sheet}"
    );
    // As Excel: sizing hidden lines shows them again.
    let back = apply(
        &out,
        &ops(r#"[{"op":"size","range":"4:6","px":30},{"op":"size","range":"B:B","px":80}]"#),
    )
    .unwrap();
    let sheet = part(&back, SHEET1);
    assert!(
        sheet.contains(r#"<row r="5" ht="22.5" customHeight="1">"#),
        "{sheet}"
    );
    assert!(!sheet.contains(r#"hidden="1""#), "{sheet}");
}

#[test]
fn rows_fit_to_their_fonts() {
    // Row 3 is given a height first: a fit leaves it as it is.
    let out = edit(
        r#"[{"op":"size","range":"3:3","px":40},{"op":"format","range":"B4","set":{"fontSize":20}},{"op":"fit","rows":[[2,35],[3,35]]}]"#,
    );
    let sheet = part(&out, SHEET1);
    assert!(
        sheet.contains(r#"<row r="3" ht="30" customHeight="1">"#),
        "{sheet}"
    );
    // As Excel keeps a row it fits: its height without customHeight.
    assert!(sheet.contains(r#"<row r="4" ht="26.25">"#), "{sheet}");
    // A smaller font again: back to its height.
    let back = apply(&out, &ops(r#"[{"op":"fit","rows":[[3,20]]}]"#)).unwrap();
    assert!(part(&back, SHEET1).contains(r#"<row r="4" ht="15">"#));
    let bad = serde_json::from_str(r#"[{"op":"fit","rows":[[3]]}]"#).unwrap();
    assert!(parse_ops(&bad, Some("Data")).is_err());
}

#[test]
fn sizes_hidden_lines_and_merged_cells() {
    let out = edit(
        r#"[{"op":"size","range":"B:C","px":120},{"op":"size","range":"3:3","px":40},{"op":"hide","range":"5:5","hidden":true},{"op":"hide","range":"A:A","hidden":true},{"op":"merge","range":"A2:B3"},{"op":"unmerge","range":"A8"}]"#,
    );
    assert_eq!(changed(BOOK, &out), [SHEET1, SHEET2]);
    let sheet = part(&out, SHEET1);
    assert!(sheet.contains(r#"<cols><col min="1" max="1" width="14.7109375" customWidth="1" hidden="1"/><col min="2" max="3" width="15.00390625" customWidth="1"/><col min="4" max="4" width="12.7109375" style="1" customWidth="1"/></cols>"#), "{sheet}");
    assert!(
        sheet.contains(r#"<row r="3" ht="30" customHeight="1">"#),
        "{sheet}"
    );
    assert!(sheet.contains(r#"<row r="5" hidden="1">"#), "{sheet}");
    assert!(
        sheet.contains(r#"<mergeCells count="1"><mergeCell ref="A2:B3"/></mergeCells>"#),
        "{sheet}"
    );
    // A merged cell keeps what its first cell holds.
    assert_eq!(
        cells(&out, "Data", "A2:B3", Read::Value),
        [["りんご", ""], ["", ""]]
    );
    assert_eq!(cell(&out, "Data", "D2", Read::Value), "0");
    // Shown again, as high as the others again.
    let back = apply(
        &out,
        &ops(r#"[{"op":"hide","range":"1:1048576","hidden":false},{"op":"hide","range":"A:XFD","hidden":false}]"#),
    )
    .unwrap();
    let sheet = part(&back, SHEET1);
    assert!(!sheet.contains("hidden"), "{sheet}");
    let bad = serde_json::from_str(r#"{"op":"size","range":"B2:C3","px":10}"#).unwrap();
    assert!(
        parse_ops(&bad, Some("Data"))
            .unwrap_err()
            .contains("行（3:5）か列（B:D）")
    );
}

#[test]
fn rows_sorted_and_filtered() {
    let out = edit(r#"{"op":"sort","range":"B3","by":"B","desc":true}"#);
    // The data around the cell (the filter's range), under its headings.
    assert_eq!(
        cells(&out, "Data", "A1:D5", Read::Formula),
        [
            ["品名", "単価", "数量", "金額"],
            ["ぶどう", "500", "2", "=B2*C2"],
            ["もも", "300", "4", "=B3*C3"],
            ["りんご", "120", "3", "=B4*C4"],
            ["みかん", "80", "10", "=B5*C5"],
        ]
    );
    // A formula beside the rows sorted stays where it is.
    assert!(part(&out, SHEET1).contains(r#"<c r="E2"><f>MyUdf(D2)</f><v>42</v></c>"#));

    let out = edit(r#"{"op":"filterBy","column":"A","values":["りんご","もも"]}"#);
    let sheet = part(&out, SHEET1);
    assert!(sheet.contains(r#"<autoFilter ref="A1:D5"><filterColumn colId="0"><filters><filter val="りんご"/><filter val="もも"/></filters></filterColumn></autoFilter>"#), "{sheet}");
    assert!(
        sheet.contains(r#"<row r="3" hidden="1">"#) && sheet.contains(r#"<row r="4" hidden="1">"#),
        "{sheet}"
    );
    assert!(!sheet.contains(r#"<row r="2" hidden"#), "{sheet}");
    // No longer filtered by the column: the rows show; without the filter, its name goes too.
    let shown = apply(&out, &ops(r#"{"op":"filterBy","column":"A","clear":true}"#)).unwrap();
    let sheet = part(&shown, SHEET1);
    assert!(
        sheet.contains(r#"<autoFilter ref="A1:D5"/>"#) && !sheet.contains("hidden"),
        "{sheet}"
    );
    let off = apply(&out, &ops(r#"{"op":"filter","off":true}"#)).unwrap();
    assert!(!part(&off, SHEET1).contains("autoFilter"));
    assert!(!part(&off, "xl/workbook.xml").contains("_FilterDatabase"));
    // Put on the data around a cell of another sheet: the element in its place, and the name.
    let on = apply(
        BOOK,
        &parse_ops(
            &serde_json::from_str(r#"[{"op":"set","range":"D1:E1","text":"h"},{"op":"set","range":"D2:E4","text":"1"},{"op":"filter","range":"D2"}]"#).unwrap(),
            Some("Summary"),
        )
        .unwrap(),
    )
    .unwrap();
    assert!(part(&on, SHEET2).contains(r#"</sheetData><autoFilter ref="D1:E4"/><pageMargins"#));
    assert!(part(&on, "xl/workbook.xml").contains(r#"<definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">Summary!$D$1:$E$4</definedName>"#));
    assert!(refused(r#"{"op":"filterBy","column":"H","values":[]}"#).contains("範囲の外"));
}

#[test]
fn the_fill_handle_continues_what_it_is_dragged_from() {
    let out = edit(
        r#"[{"op":"values","at":"G1","rows":[["1","第1回","2026/4/1","a","=C1*2","No.008"],["3","第2回","2026/4/3","b","","No.009"]]},{"op":"autofill","from":"G1:L2","to":"G1:L6"}]"#,
    );
    assert_eq!(
        cells(&out, "Data", "G3:I6", Read::Value2),
        [
            ["5", "第3回", "46117"],
            ["7", "第4回", "46119"],
            ["9", "第5回", "46121"],
            ["11", "第6回", "46123"],
        ]
    );
    // Anything else is repeated; formulas with their references moved.
    assert_eq!(
        cells(&out, "Data", "J3:L6", Read::Formula),
        [
            ["a", "=C3*2", "No.010"],
            ["b", "", "No.011"],
            ["a", "=C5*2", "No.012"],
            ["b", "", "No.013"],
        ]
    );
    assert_eq!(cell(&out, "Data", "I6", Read::NumberFormat), "m/d/yyyy");
    // One number is repeated, one date goes on by days; to the right and upwards too.
    let out = edit(
        r#"[{"op":"values","at":"G5","rows":[["7","2026/4/1"]]},{"op":"autofill","from":"G5:H5","to":"G3:H5"},{"op":"autofill","from":"G5","to":"G5:G6"},{"op":"autofill","from":"C2:C3","to":"C2:E3"}]"#,
    );
    assert_eq!(
        cells(&out, "Data", "G3:H6", Read::Value2),
        [["7", "46111"], ["7", "46112"], ["7", "46113"], ["7", ""]]
    );
    assert_eq!(
        cells(&out, "Data", "C2:E3", Read::Value),
        [["3", "3", "3"], ["10", "10", "10"]]
    );
    assert!(refused(r#"{"op":"autofill","from":"A1:B2","to":"A1:C3"}"#).contains("どちらかにだけ"));
}

#[test]
fn a_macro_is_assigned_to_a_button_and_a_chart() {
    let out = edit(
        r#"[{"op":"macro","shape":1025,"macro":"Sheet1.runIt"},{"op":"macro","object":"Chart 1","macro":"Sheet1.chartClick"}]"#,
    );
    assert_eq!(
        changed(BOOK, &out),
        ["xl/drawings/drawing1.xml", "xl/drawings/vmlDrawing1.vml"]
    );
    let vml = part(&out, "xl/drawings/vmlDrawing1.vml");
    assert!(vml.contains("<x:FmlaMacro>[0]!Sheet1.runIt</x:FmlaMacro>"));
    assert!(!vml.contains("say_hello"));
    assert!(
        part(&out, "xl/drawings/drawing1.xml")
            .contains("<xdr:graphicFrame macro=\"[0]!Sheet1.chartClick\">")
    );
    // The cells are those they were.
    assert_eq!(
        cell(&out, "Data", "A1", Read::Formula),
        cell(BOOK, "Data", "A1", Read::Formula)
    );
    assert!(refused(r#"{"op":"macro","object":"なし","macro":"Sheet1.a"}"#).contains("なし"));
    assert!(
        parse_ops(
            &serde_json::json!({"op":"macro","object":"Chart 1","macro":"a b"}),
            Some("Data")
        )
        .is_err()
    );
}

/// What the grid draws of a workbook file (`ras book view`).
fn view(bytes: &[u8]) -> Json {
    let mut book = Book::open_bytes(bytes).unwrap();
    crate::view::data(&mut book, "edit.xlsm", Some(bytes), None, None).unwrap()
}

#[test]
fn panes_frozen_and_no_longer() {
    // `<sheetView .../>` gets the pane and its selections, as Excel writes them.
    let out = edit(r#"{"op":"freeze","at":"B2"}"#);
    assert_eq!(changed(BOOK, &out), ["xl/worksheets/sheet1.xml"]);
    let doc = part(&out, "xl/worksheets/sheet1.xml");
    assert!(doc.contains(
        "<sheetView tabSelected=\"1\" workbookViewId=\"0\"><pane xSplit=\"1\" ySplit=\"1\" topLeftCell=\"B2\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"topRight\"/><selection pane=\"bottomLeft\"/><selection pane=\"bottomRight\" activeCell=\"B2\" sqref=\"B2\"/></sheetView>"
    ));
    assert_eq!(view(&out)["sheets"][0]["frozen"], json!([1, 1]));
    // Frozen again elsewhere: one pane, in place of the other.
    let rows = apply(&out, &ops(r#"{"op":"freeze","at":"A4"}"#)).unwrap();
    let doc = part(&rows, "xl/worksheets/sheet1.xml");
    assert_eq!(doc.matches("<pane ").count(), 1);
    assert!(doc.contains("<pane ySplit=\"3\" topLeftCell=\"A4\" activePane=\"bottomLeft\" state=\"frozen\"/><selection pane=\"bottomLeft\" activeCell=\"A4\" sqref=\"A4\"/></sheetView>"));
    assert_eq!(view(&rows)["sheets"][0]["frozen"], json!([3, 0]));
    // No longer frozen: the pane and its selections go.
    let off = apply(&rows, &ops(r#"{"op":"freeze","off":true}"#)).unwrap();
    let doc = part(&off, "xl/worksheets/sheet1.xml");
    assert!(!doc.contains("<pane") && !doc.contains("<selection"));
    assert_eq!(view(&off)["sheets"][0]["frozen"], Json::Null);
    // The cells are those they were; A1 is no pane; a range is refused.
    assert_eq!(cell(&out, "Data", "D6", Read::Formula), "=SUM(D2:D5)");
    assert_eq!(
        apply(BOOK, &ops(r#"{"op":"freeze","at":"A1"}"#)).unwrap(),
        BOOK
    );
    assert!(
        parse_ops(
            &serde_json::json!({"op":"freeze","at":"A1:B2"}),
            Some("Data")
        )
        .is_err()
    );
}

#[test]
fn notes_lists_and_names_for_the_grid() {
    // The input rule of C2:C5 made a list of its own, and B2:B5 one of the name `Items`.
    let sheet = part(BOOK, "xl/worksheets/sheet1.xml").replace(
        "<dataValidations count=\"1\"><dataValidation type=\"whole\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" sqref=\"C2:C5\"><formula1>1</formula1><formula2>100</formula2></dataValidation></dataValidations>",
        "<dataValidations count=\"2\"><dataValidation type=\"list\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" promptTitle=\"数量\" prompt=\"選んでください\" sqref=\"C2:C5\"><formula1>\"1,2, 3\"</formula1></dataValidation><dataValidation type=\"list\" errorStyle=\"warning\" showErrorMessage=\"1\" sqref=\"B2:B5 E7\"><formula1>Items</formula1></dataValidation></dataValidations>",
    );
    let mut all = parts(BOOK);
    all.insert("xl/worksheets/sheet1.xml".into(), sheet.into_bytes());
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, data) in all {
        zip.start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(&data).unwrap();
    }
    let bytes = zip.finish().unwrap().into_inner();
    let data = view(&bytes);
    let sheet = &data["sheets"][0];
    assert_eq!(
        sheet["notes"],
        json!([{ "r": 2, "c": 1, "text": "メモです", "author": "" }])
    );
    assert_eq!(
        sheet["lists"],
        json!([
            { "ranges": [[1, 2, 4, 2]], "values": ["1", "2", "3"], "all": ["1", "2", "3"], "stop": "", "title": "数量", "prompt": "選んでください" },
            { "ranges": [[1, 1, 4, 1], [6, 4, 6, 4]], "values": ["りんご", "みかん", "ぶどう", "もも"] }
        ])
    );
    assert_eq!(
        data["names"],
        json!([{ "name": "Items", "ref": "Data!$A$2:$A$5" }, { "name": "Total", "ref": "Data!$D$6" }])
    );
    // Numbers keep their value, formatted or not (the status bar sums them).
    let b2 = sheet["cells"]
        .as_array()
        .unwrap()
        .iter()
        .find(|c| c["r"] == json!(1) && c["c"] == json!(1))
        .unwrap();
    assert_eq!(
        (b2["x"].clone(), b2["n"].clone()),
        (json!(120.0), json!("#,##0"))
    );
}

#[test]
fn pasted_special_transposed_and_multiplied() {
    // B2:C3 (120 3 / 80 10) as values, rows made columns.
    let out = edit(
        r#"{"op":"pasteSpecial","from":"B2:C3","to":"F10","paste":"values","transpose":true}"#,
    );
    assert_eq!(
        cells(&out, "Data", "F10:G11", Read::Value),
        [["120", "80"], ["3", "10"]]
    );
    // The quantities multiplied into the prices; the formulas of D follow them.
    let out = edit(
        r#"{"op":"pasteSpecial","from":"C2:C3","to":"B2:B3","paste":"values","operation":"multiply"}"#,
    );
    assert_eq!(
        cells(&out, "Data", "B2:B3", Read::Value),
        [["360"], ["800"]]
    );
    assert_eq!(cell(&out, "Data", "D2", Read::Value), "1080");
    let wrong =
        |json: &str| parse_ops(&serde_json::from_str(json).unwrap(), Some("Data")).unwrap_err();
    assert!(
        wrong(r#"{"op":"pasteSpecial","from":"A1","to":"B1","paste":"notes"}"#).contains("notes")
    );
}

#[test]
fn sorted_by_several_keys_and_filtered_by_comparisons() {
    let out = edit(
        r#"{"op":"sort","range":"A1:D5","keys":[{"by":"C","desc":true},{"by":"A"}],"header":true}"#,
    );
    assert_eq!(
        cells(&out, "Data", "A2:A5", Read::Value),
        [["みかん"], ["もも"], ["りんご"], ["ぶどう"]]
    );
    let wrong =
        |json: &str| parse_ops(&serde_json::from_str(json).unwrap(), Some("Data")).unwrap_err();
    assert!(wrong(r#"{"op":"sort","range":"A1:D5","keys":[]}"#).contains("keys"));
    // Prices from 100 up to 400: the rows of 80 and 500 are hidden.
    let out = edit(r#"{"op":"filterBy","column":"B","criteria":[[">=","100"],["<","400"]]}"#);
    let sheet = part(&out, "xl/worksheets/sheet1.xml");
    assert!(sheet.contains(r#"<filterColumn colId="1"><customFilters and="1"><customFilter operator="greaterThanOrEqual" val="100"/><customFilter operator="lessThan" val="400"/></customFilters></filterColumn>"#), "{sheet}");
    let hidden = |doc: &str, r: u32| {
        doc.split("<row ").any(|row| {
            row.starts_with(&format!("r=\"{r}\""))
                && row
                    .split('>')
                    .next()
                    .is_some_and(|t| t.contains("hidden=\"1\""))
        })
    };
    assert!(hidden(&sheet, 3) && hidden(&sheet, 4) && !hidden(&sheet, 2) && !hidden(&sheet, 5));
    // The largest amount only.
    let out = edit(r#"{"op":"filterBy","column":"D","top":1}"#);
    let sheet = part(&out, "xl/worksheets/sheet1.xml");
    assert!(sheet.contains(r#"<top10 val="1"/>"#));
    assert!(hidden(&sheet, 2) && hidden(&sheet, 3) && hidden(&sheet, 4) && !hidden(&sheet, 5));
    assert!(wrong(r#"{"op":"filterBy","column":"B","criteria":[["~","1"]]}"#).contains("~"));
}

#[test]
fn links_conditional_formats_tables_and_functions_for_the_grid() {
    let data = view(BOOK);
    let sheet = &data["sheets"][0];
    assert_eq!(
        sheet["links"],
        json!([{ "r": 9, "c": 0, "url": "https://example.com" }])
    );
    // D2:D5 over 500 (800, 1000, 1200) in the rule's format: bold, filled, with borders.
    let css = |sheet: &Json, r: u64, c: u64| -> String {
        let cell = sheet["cells"]
            .as_array()
            .unwrap()
            .iter()
            .find(|x| x["r"] == json!(r) && x["c"] == json!(c))
            .unwrap();
        sheet["styles"][cell["s"].as_u64().unwrap_or(0) as usize]
            .as_str()
            .unwrap()
            .to_string()
    };
    assert!(!css(sheet, 1, 3).contains("font-weight:bold"));
    for r in 2..=4 {
        let s = css(sheet, r, 3);
        assert!(
            s.contains("font-weight:bold") && s.contains("background:#DDEEFF"),
            "{s}"
        );
    }
    // The table of Tbl (XlsxWriter's Table Style Medium 9): its heading filled, its rows in
    // stripes of two tints of the accent.
    let tbl = &data["sheets"][2];
    let (head, first, second) = (css(tbl, 1, 1), css(tbl, 2, 1), css(tbl, 3, 1));
    assert!(
        head.contains("background:#4F81BD") && head.contains("font-weight:bold"),
        "{head}"
    );
    assert!(
        first.contains("background:#B8CCE4") && second.contains("background:#DCE6F1"),
        "{first} / {second}"
    );
    // The functions to complete formulas with: the evaluator's own too.
    let functions = crate::view::functions();
    let find = |n: &str| {
        functions
            .as_array()
            .unwrap()
            .iter()
            .find(|f| f[0] == json!(n))
            .cloned()
    };
    assert_eq!(
        find("VLOOKUP").unwrap()[1],
        json!("VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])")
    );
    assert_eq!(
        find("SUBTOTAL").unwrap()[1],
        json!("SUBTOTAL(function_num, ref1, [ref2], ...)")
    );
}

#[test]
fn hyperlinks_inserted_and_taken_away() {
    // A web page: the sheet's hyperlinks get one, its relationships the page.
    let out =
        edit(r#"{"op":"link","range":"C2","url":"https://example.org/a?b=1&c=2","tip":"説明"}"#);
    assert_eq!(
        changed(BOOK, &out),
        [
            "xl/worksheets/_rels/sheet1.xml.rels",
            "xl/worksheets/sheet1.xml"
        ]
    );
    let doc = part(&out, SHEET1);
    assert!(doc.contains(r#"<hyperlinks><hyperlink ref="A10" r:id="rId1"/><hyperlink ref="C2" tooltip="説明" r:id="rId5"/></hyperlinks>"#), "{doc}");
    assert!(part(&out, "xl/worksheets/_rels/sheet1.xml.rels").contains(r#"<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/a?b=1&amp;c=2" TargetMode="External"/>"#));
    assert_eq!(
        view(&out)["sheets"][0]["links"],
        json!([
            { "r": 1, "c": 2, "url": "https://example.org/a?b=1&c=2", "tip": "説明" },
            { "r": 9, "c": 0, "url": "https://example.com" }
        ])
    );
    // A place in the workbook, in place of the cell's link: the page's relationship goes.
    let to = apply(
        &out,
        &ops(r#"{"op":"link","range":"C2","to":"Summary!A1","display":"集計へ"}"#),
    )
    .unwrap();
    let doc = part(&to, SHEET1);
    assert!(doc.contains(r#"<hyperlink ref="A10" r:id="rId1"/><hyperlink ref="C2" location="Summary!A1" display="集計へ"/></hyperlinks>"#), "{doc}");
    assert!(!part(&to, "xl/worksheets/_rels/sheet1.xml.rels").contains("rId5"));
    assert_eq!(
        view(&to)["sheets"][0]["links"][0],
        json!({ "r": 1, "c": 2, "to": "Summary!A1" })
    );
    // Taken away: the last one takes `<hyperlinks>` with it.
    let off = apply(&to, &ops(r#"[{"op":"link","range":"A1:C10","off":true}]"#)).unwrap();
    let doc = part(&off, SHEET1);
    assert!(!doc.contains("hyperlink"), "{doc}");
    assert!(!part(&off, "xl/worksheets/_rels/sheet1.xml.rels").contains("/hyperlink\""));
    assert_eq!(view(&off)["sheets"][0]["links"], json!([]));
    // A sheet without relationships gets them; nothing to take away changes nothing.
    let other =
        edit(r#"{"op":"link","sheet":"Summary","range":"B2","url":"mailto:a@example.com"}"#);
    let doc = part(&other, SHEET2);
    assert!(
        doc.contains(r#"<hyperlinks><hyperlink ref="B2" r:id="rId1"/></hyperlinks><pageMargins"#),
        "{doc}"
    );
    assert!(
        part(&other, "xl/worksheets/_rels/sheet2.xml.rels")
            .contains(r#"Target="mailto:a@example.com" TargetMode="External""#)
    );
    assert_eq!(
        view(&other)["sheets"][1]["links"],
        json!([{ "r": 1, "c": 1, "url": "mailto:a@example.com" }])
    );
    assert_eq!(edit(r#"{"op":"link","range":"B2","off":true}"#), BOOK);
    assert!(parse_ops(&json!({"op":"link","range":"A1"}), Some("Data")).is_err());
    assert!(
        parse_ops(
            &json!({"op":"link","range":"A1","url":"x","to":"y"}),
            Some("Data")
        )
        .is_err()
    );
}

#[test]
fn notes_written_edited_and_deleted() {
    // A new note: the comments part and the VML drawing the sheet has get it.
    let out = edit(r#"{"op":"note","at":"C2","text":"新しい\nメモ & <注>","author":"ras"}"#);
    assert_eq!(
        changed(BOOK, &out),
        ["xl/comments1.xml", "xl/drawings/vmlDrawing1.vml"]
    );
    let comments = part(&out, "xl/comments1.xml");
    assert!(
        comments.contains("<authors><author></author><author>ras</author></authors>"),
        "{comments}"
    );
    assert!(comments.contains(r#"<comment ref="C2" authorId="1"><text><r>"#));
    let vml = part(&out, "xl/drawings/vmlDrawing1.vml");
    assert!(
        vml.contains(r##"<v:shape id="_x0000_s1027" type="#_x0000_t202""##),
        "{vml}"
    );
    assert!(vml.contains("<x:Row>1</x:Row>\n   <x:Column>2</x:Column>"));
    let notes = |b: &[u8], i: usize| view(b)["sheets"][i]["notes"].clone();
    assert_eq!(
        notes(&out, 0),
        json!([
            { "r": 2, "c": 1, "text": "メモです", "author": "" },
            { "r": 1, "c": 2, "text": "新しい\nメモ & <注>", "author": "ras" }
        ])
    );
    // Edited: its text only.
    let edited = apply(&out, &ops(r#"{"op":"note","at":"B3","text":"書き換えた"}"#)).unwrap();
    assert_eq!(changed(&out, &edited), ["xl/comments1.xml"]);
    assert_eq!(notes(&edited, 0)[0]["text"], json!("書き換えた"));
    // Deleted: the note and its box.
    let gone = apply(&edited, &ops(r#"{"op":"note","at":"B3","off":true}"#)).unwrap();
    assert!(!part(&gone, "xl/comments1.xml").contains("ref=\"B3\""));
    let vml = part(&gone, "xl/drawings/vmlDrawing1.vml");
    assert!(
        !vml.contains("<x:Row>2</x:Row>") && vml.contains("_x0000_s1025"),
        "{vml}"
    );
    assert_eq!(notes(&gone, 0).as_array().unwrap().len(), 1);
    assert_eq!(edit(r#"{"op":"note","at":"D4","off":true}"#), BOOK);
    // A sheet without notes: a comments part and a VML drawing of its own.
    let other = edit(r#"{"op":"note","sheet":"Summary","at":"A1","text":"合計の2倍"}"#);
    assert_eq!(
        changed(BOOK, &other),
        [
            "[Content_Types].xml",
            "xl/comments2.xml",
            "xl/drawings/vmlDrawing2.vml",
            "xl/worksheets/_rels/sheet2.xml.rels",
            "xl/worksheets/sheet2.xml"
        ]
    );
    let doc = part(&other, SHEET2);
    assert!(doc.contains(r#"<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><legacyDrawing r:id="rId2"/></worksheet>"#), "{doc}");
    assert!(part(&other, "[Content_Types].xml").contains(r#"<Override PartName="/xl/comments2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>"#));
    let vml = part(&other, "xl/drawings/vmlDrawing2.vml");
    assert!(
        vml.contains(r#"<o:idmap v:ext="edit" data="2"/>"#) && vml.contains("_x0000_s2049"),
        "{vml}"
    );
    assert_eq!(
        notes(&other, 1),
        json!([{ "r": 0, "c": 0, "text": "合計の2倍", "author": "" }])
    );
    // The cells are as they were; a row inserted moves the new note too.
    assert_eq!(cell(&other, "Summary", "A1", Read::Formula), "=Data!D6*2");
    let moved = apply(
        &other,
        &ops(r#"{"op":"insert","sheet":"Summary","range":"1:1"}"#),
    )
    .unwrap();
    assert_eq!(notes(&moved, 1)[0]["r"], json!(1));
}

#[test]
fn tables_made_of_the_data() {
    // A cell of the filter's range: the table is the filter's range, and takes its place.
    let out = edit(r#"{"op":"table","range":"B3"}"#);
    let doc = part(&out, SHEET1);
    assert!(!doc.contains("<autoFilter"), "{doc}");
    assert!(doc.contains(r#"<legacyDrawing r:id="rId3"/><tableParts count="1"><tablePart r:id="rId5"/></tableParts></worksheet>"#), "{doc}");
    let table = part(&out, "xl/tables/table2.xml");
    assert!(table.contains(r#"id="2" name="テーブル1" displayName="テーブル1" ref="A1:D5" totalsRowShown="0"><autoFilter ref="A1:D5"/><tableColumns count="4"><tableColumn id="1" name="品名"/><tableColumn id="2" name="単価"/><tableColumn id="3" name="数量"/><tableColumn id="4" name="金額"/></tableColumns><tableStyleInfo name="TableStyleMedium2""#), "{table}");
    assert!(
        part(&out, "[Content_Types].xml").contains(r#"<Override PartName="/xl/tables/table2.xml""#)
    );
    assert!(
        part(&out, "xl/worksheets/_rels/sheet1.xml.rels")
            .contains(r#"Target="../tables/table2.xml""#)
    );
    assert_eq!(cell(&out, "Data", "D6", Read::Formula), "=SUM(D2:D5)");
    let mut book = Book::open_bytes(&out).unwrap();
    let info = book.sheet_summary(0).unwrap();
    assert_eq!(
        (info.tables[0].name.as_str(), info.tables[0].range.as_str()),
        ("テーブル1", "A1:D5")
    );
    assert_eq!(info.filter, None);
    // A row inserted in it makes it longer.
    let longer = apply(&out, &ops(r#"{"op":"insert","range":"3:3"}"#)).unwrap();
    assert!(part(&longer, "xl/tables/table2.xml").contains(r#"ref="A1:D6""#));

    // Headings that are not texts of their own are made so: a formula, a number, one again.
    let sum = edit(
        r#"[{"op":"set","sheet":"Summary","range":"B1","text":"6720"},{"op":"set","sheet":"Summary","range":"B2","text":"x"}]"#,
    );
    let out = apply(
        &sum,
        &ops(r#"{"op":"table","sheet":"Summary","range":"A1:B3","name":"集計"}"#),
    )
    .unwrap();
    assert_eq!(
        cells(&out, "Summary", "A1:B1", Read::Formula),
        [["6720", "67202"]]
    );
    assert!(
        part(&out, "xl/tables/table2.xml")
            .contains(r#"name="集計" displayName="集計" ref="A1:B3""#)
    );
    assert!(
        part(&out, "xl/tables/table2.xml")
            .contains(r#"<tableColumn id="1" name="6720"/><tableColumn id="2" name="67202"/>"#)
    );
    // Empty headings, in the middle and last, are 列1, 列2, ... as Excel names them.
    let gaps = edit(
        r#"[{"op":"clear","sheet":"Summary","range":"A1:D3"},{"op":"set","sheet":"Summary","range":"A1","text":"a"},{"op":"set","sheet":"Summary","range":"C1","text":"c"},{"op":"set","sheet":"Summary","range":"A2:D2","text":"1"}]"#,
    );
    let out = apply(
        &gaps,
        &ops(r#"{"op":"table","sheet":"Summary","range":"A1:D2"}"#),
    )
    .unwrap();
    assert_eq!(
        cells(&out, "Summary", "A1:D1", Read::Formula),
        [["a", "列1", "c", "列2"]]
    );
    assert!(
        part(&out, "xl/tables/table2.xml").contains(r#"<tableColumns count="4"><tableColumn id="1" name="a"/><tableColumn id="2" name="列1"/><tableColumn id="3" name="c"/><tableColumn id="4" name="列2"/></tableColumns>"#),
        "{}",
        part(&out, "xl/tables/table2.xml")
    );
    // No headings: a row of them is inserted above the data.
    let out = edit(r#"{"op":"table","sheet":"Summary","range":"A1:A3","header":false}"#);
    assert_eq!(
        cells(&out, "Summary", "A1:A4", Read::Formula),
        [["列1"], ["=Data!D6*2"], ["=SUM(Data!D2:D5)"], ["=Total+1"]]
    );
    // (Only the cells of the table move down.)
    assert_eq!(
        cell(&out, "Summary", "B1", Read::Formula),
        cell(BOOK, "Summary", "B1", Read::Formula)
    );
    assert!(part(&out, "xl/tables/table2.xml").contains(r#"ref="A1:A4""#));

    // Refused: over a table, over merged cells, a name taken.
    assert!(refused(r#"{"op":"table","sheet":"Tbl","range":"A1:C3"}"#).contains("重ね"));
    assert!(refused(r#"{"op":"table","range":"A8:C9"}"#).contains("結合"));
    assert!(refused(r#"{"op":"table","range":"A1:D5","name":"Items"}"#).contains("既に"));
    assert!(refused(r#"{"op":"table","range":"A1:D5","name":"A1"}"#).contains("使えません"));
}

/// A PNG of 3 × 2 pixels.
const PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x03, 0x00, 0x00, 0x00, 0x02, 0x08, 0x02, 0x00, 0x00, 0x00, 0x12, 0x16, 0xF1,
    0x4D, 0x00, 0x00, 0x00, 0x12, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0xF0,
    0x1F, 0x08, 0x19, 0x18, 0xFE, 0x03, 0x00, 0x2F, 0xF6, 0x05, 0xFB, 0x66, 0x29, 0xC3, 0x1E, 0x00,
    0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
];

/// The objects the grid draws on a sheet: (kind, name, from, to).
fn objects(bytes: &[u8], sheet: usize) -> Vec<(String, String, Json, Json)> {
    view(bytes)["sheets"][sheet]["objects"]
        .as_array()
        .cloned()
        .unwrap_or_default()
        .into_iter()
        .map(|o| {
            (
                o["kind"].as_str().unwrap_or("").to_string(),
                o["name"].as_str().unwrap_or("").to_string(),
                o["from"].clone(),
                o["to"].clone(),
            )
        })
        .collect()
}

#[test]
fn shapes_and_pictures_put_on_a_sheet() {
    // A sheet without a drawing gets one: its part, relationship, content type, `<drawing>`.
    let out = edit(
        r#"{"op":"shape","sheet":"Summary","shape":"ellipse","at":"B2","dx":10,"dy":5,"width":100,"height":50,"text":"こんにちは\n2 行目"}"#,
    );
    assert_eq!(
        changed(BOOK, &out),
        [
            "[Content_Types].xml",
            "xl/drawings/drawing2.xml",
            "xl/worksheets/_rels/sheet2.xml.rels",
            "xl/worksheets/sheet2.xml"
        ]
    );
    assert!(part(&out, SHEET2).ends_with(r#"<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><drawing r:id="rId1"/></worksheet>"#));
    assert!(part(&out, "xl/worksheets/_rels/sheet2.xml.rels").contains(r#"Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml""#));
    assert!(part(&out, "[Content_Types].xml").contains(r#"<Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>"#));
    let drawing = part(&out, "xl/drawings/drawing2.xml");
    // (Columns 72 px wide, rows 20 px high: B2 is at 72, 20.)
    assert!(drawing.contains(r#"<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>95250</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>47625</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>361950</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>142875</xdr:rowOff></xdr:to><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="2" name="楕円 1"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="781050" y="238125"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="ellipse">"#), "{drawing}");
    assert!(drawing.contains(r#"<a:p><a:pPr algn="ctr"/><a:r><a:rPr kumimoji="1" lang="ja-JP" altLang="en-US" sz="1100"/><a:t>こんにちは</a:t></a:r></a:p><a:p><a:pPr algn="ctr"/><a:r><a:rPr kumimoji="1" lang="ja-JP" altLang="en-US" sz="1100"/><a:t>2 行目</a:t></a:r></a:p>"#), "{drawing}");
    let shapes = objects(&out, 1);
    assert_eq!(
        shapes,
        [(
            "shape".to_string(),
            "楕円 1".to_string(),
            json!([1, 10.0, 1, 5.0]),
            json!([2, 38.0, 3, 15.0])
        )]
    );
    assert!(
        view(&out)["sheets"][1]["objects"][0]["html"]
            .as_str()
            .unwrap()
            .contains("こんにちは")
    );
    // The cells are as they were.
    assert_eq!(cell(&out, "Summary", "A1", Read::Formula), "=Data!D6*2");

    // Over a range; a text box, a line drawn up to the right, an arrow: on the sheet's drawing,
    // after its chart.
    let out = edit(
        r#"[{"op":"shape","shape":"textBox","range":"A12:B13","text":"メモ"},{"op":"shape","shape":"line","at":"D12","width":64,"height":40,"flipV":true},{"op":"shape","shape":"arrow","at":"F12","width":30}]"#,
    );
    assert_eq!(changed(BOOK, &out), ["xl/drawings/drawing1.xml"]);
    let drawing = part(&out, "xl/drawings/drawing1.xml");
    assert!(
        drawing
            .contains(r#"<xdr:cNvPr id="3" name="テキスト ボックス 2"/><xdr:cNvSpPr txBox="1"/>"#),
        "{drawing}"
    );
    assert!(drawing.contains(r#"<xdr:cNvPr id="4" name="直線コネクタ 3"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr><a:xfrm flipV="1">"#), "{drawing}");
    assert!(
        drawing.contains(r#"name="直線矢印コネクタ 4"/>"#),
        "{drawing}"
    );
    assert!(drawing.contains(r#"<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom><a:ln><a:tailEnd type="triangle"/></a:ln>"#), "{drawing}");
    let names: Vec<(String, String)> = objects(&out, 0).into_iter().map(|o| (o.0, o.1)).collect();
    assert_eq!(
        names,
        [
            ("chart", "Chart 1"),
            ("shape", "テキスト ボックス 2"),
            ("shape", "直線コネクタ 3"),
            ("shape", "直線矢印コネクタ 4"),
            ("control", "")
        ]
        .map(|(a, b)| (a.to_string(), b.to_string()))
    );
    // The text box covers its range.
    assert_eq!(objects(&out, 0)[1].2, json!([0, 0.0, 11, 0.0]));
    assert_eq!(objects(&out, 0)[1].3, json!([2, 0.0, 13, 0.0]));

    // A picture: its bytes in xl/media, the relationship of the drawing, the content type of
    // its extension; its own size.
    let out = edit(&format!(
        r#"{{"op":"picture","at":"A12","data":"data:image/png;base64,{}"}}"#,
        crate::view::drawing::base64(PNG)
    ));
    assert_eq!(
        changed(BOOK, &out),
        [
            "[Content_Types].xml",
            "xl/drawings/_rels/drawing1.xml.rels",
            "xl/drawings/drawing1.xml",
            "xl/media/image1.png"
        ]
    );
    assert_eq!(parts(&out)["xl/media/image1.png"], PNG);
    assert!(
        part(&out, "[Content_Types].xml")
            .contains(r#"<Default Extension="png" ContentType="image/png"/>"#)
    );
    assert!(part(&out, "xl/drawings/_rels/drawing1.xml.rels").contains(r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>"#));
    let drawing = part(&out, "xl/drawings/drawing1.xml");
    assert!(drawing.contains(r#"<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>11</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>0</xdr:col><xdr:colOff>28575</xdr:colOff><xdr:row>11</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="3" name="図 2"/>"#), "{drawing}");
    assert!(drawing.contains(r#"<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId2"/>"#));
    let pic = &view(&out)["sheets"][0]["objects"][1];
    assert_eq!(pic["kind"], json!("picture"));
    assert!(pic["html"].as_str().unwrap().contains("data-img"));
    // Sized by its width only: its proportion kept.
    let out2 = edit(&format!(
        r#"{{"op":"picture","sheet":"Summary","at":"B2","width":30,"data":"{}"}}"#,
        crate::view::drawing::base64(PNG)
    ));
    assert_eq!(objects(&out2, 1)[0].3, json!([1, 30.0, 2, 0.0]));

    // Refused: a shape there is not, a picture that is not one.
    assert!(
        parse_ops(
            &json!({"op":"shape","shape":"cloud","at":"A1"}),
            Some("Data")
        )
        .unwrap_err()
        .contains("cloud")
    );
    assert!(refused(r#"{"op":"picture","at":"A1","data":"aGVsbG8="}"#).contains("形式"));
    assert!(
        parse_ops(&json!({"op":"shape","at":"A1:B2"}), Some("Data"))
            .unwrap_err()
            .contains("range")
    );
}

#[test]
fn objects_moved_sized_given_a_text_and_deleted() {
    let png = crate::view::drawing::base64(PNG);
    let out = edit(&format!(
        r#"[{{"op":"shape","at":"A12","width":50,"height":30,"text":"前"}},{{"op":"picture","at":"D12","data":"{png}"}}]"#
    ));
    // Moved (its size kept), sized, both.
    let moved = apply(
        &out,
        &ops(r#"{"op":"object","name":"正方形/長方形 2","at":"B15","dx":4}"#),
    )
    .unwrap();
    assert_eq!(changed(&out, &moved), ["xl/drawings/drawing1.xml"]);
    let o = &objects(&moved, 0)[1];
    assert_eq!(
        (o.2.clone(), o.3.clone()),
        (json!([1, 4.0, 14, 0.0]), json!([1, 54.0, 15, 10.0]))
    );
    let drawing = part(&moved, "xl/drawings/drawing1.xml");
    // (A is 118 px wide; row 7 is 40 px high.)
    assert!(
        drawing.contains(r#"<a:off x="1162050" y="2857500"/><a:ext cx="476250" cy="285750"/>"#),
        "{drawing}"
    );
    let sized = apply(
        &moved,
        &ops(r#"{"op":"object","name":"図 3","range":"D12:E13"}"#),
    )
    .unwrap();
    let o = &objects(&sized, 0)[2];
    assert_eq!(
        (o.2.clone(), o.3.clone()),
        (json!([3, 0.0, 11, 0.0]), json!([5, 0.0, 13, 0.0]))
    );

    // A text: the first paragraph's properties kept.
    let texted = apply(
        &sized,
        &ops(r#"{"op":"object","name":"正方形/長方形 2","text":"後\n2"}"#),
    )
    .unwrap();
    let drawing = part(&texted, "xl/drawings/drawing1.xml");
    assert!(drawing.contains(r#"<a:p><a:pPr algn="ctr"/><a:r><a:rPr kumimoji="1" lang="ja-JP" altLang="en-US" sz="1100"/><a:t>後</a:t></a:r></a:p><a:p><a:pPr algn="ctr"/><a:r><a:rPr kumimoji="1" lang="ja-JP" altLang="en-US" sz="1100"/><a:t>2</a:t></a:r></a:p></xdr:txBody>"#), "{drawing}");
    assert!(!drawing.contains("前"));
    assert!(
        apply(&texted, &ops(r#"{"op":"object","name":"図 3","text":"x"}"#))
            .unwrap_err()
            .contains("文字")
    );

    // Deleted: a picture takes its bytes, its relationship and nothing else with it.
    let gone = apply(
        &texted,
        &ops(r#"{"op":"object","name":"図 3","delete":true}"#),
    )
    .unwrap();
    assert_eq!(
        changed(&texted, &gone),
        [
            "xl/drawings/_rels/drawing1.xml.rels",
            "xl/drawings/drawing1.xml",
            "xl/media/image1.png"
        ]
    );
    assert!(!parts(&gone).contains_key("xl/media/image1.png"));
    assert!(!part(&gone, "xl/drawings/_rels/drawing1.xml.rels").contains("image"));
    // A chart: its part and content type.
    let gone = apply(
        &gone,
        &ops(r#"{"op":"object","name":"Chart 1","delete":true}"#),
    )
    .unwrap();
    assert!(!parts(&gone).contains_key("xl/charts/chart1.xml"));
    assert!(!part(&gone, "[Content_Types].xml").contains("chart1"));
    assert!(!part(&gone, "xl/drawings/_rels/drawing1.xml.rels").contains("chart"));
    let names: Vec<String> = objects(&gone, 0).into_iter().map(|o| o.1).collect();
    assert_eq!(names, ["正方形/長方形 2", ""]);
    assert_eq!(cell(&gone, "Data", "D6", Read::Formula), "=SUM(D2:D5)");

    assert!(
        apply(
            &gone,
            &ops(r#"{"op":"object","name":"なし","delete":true}"#)
        )
        .unwrap_err()
        .contains("なし")
    );
    assert!(
        parse_ops(&json!({"op":"object","name":"x"}), Some("Data"))
            .unwrap_err()
            .contains("delete")
    );
}

#[test]
fn objects_filled_lined_turned_and_reordered() {
    let png = crate::view::drawing::base64(PNG);
    let out = edit(&format!(
        r#"[{{"op":"shape","at":"A12","width":50,"height":30}},{{"op":"shape","shape":"line","at":"D12","width":40}},{{"op":"picture","at":"F12","data":"{png}"}}]"#
    ));
    let shape = |bytes: &[u8], name: &str| -> String {
        let doc = part(bytes, "xl/drawings/drawing1.xml");
        let at = doc.find(&format!("name=\"{name}\"")).unwrap();
        let start = doc[..at].rfind("<xdr:twoCellAnchor").unwrap();
        let end = at + doc[at..].find("</xdr:twoCellAnchor>").unwrap();
        doc[start..end].to_string()
    };
    // Fill and line: in the shape's properties, in the schema's order (after the geometry; the
    // line after the fill), over what the shape's style gives.
    let styled = apply(
        &out,
        &ops(r##"{"op":"object","name":"正方形/長方形 2","fill":"#ff0000","line":"#00ff00","lineWidth":2.25}"##),
    )
    .unwrap();
    assert_eq!(changed(&out, &styled), ["xl/drawings/drawing1.xml"]);
    assert!(shape(&styled, "正方形/長方形 2").contains(r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:ln w="28575"><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></xdr:spPr>"#), "{}", shape(&styled, "正方形/長方形 2"));
    // Again: replaced, not added; no fill, no line.
    let none = apply(
        &styled,
        &ops(r#"{"op":"object","name":"正方形/長方形 2","fill":null,"line":null}"#),
    )
    .unwrap();
    assert!(
        shape(&none, "正方形/長方形 2")
            .contains(r#"</a:prstGeom><a:noFill/><a:ln w="28575"><a:noFill/></a:ln></xdr:spPr>"#),
        "{}",
        shape(&none, "正方形/長方形 2")
    );
    // The width alone keeps the line's color; the arrow's head stays.
    let arrow = edit(r#"{"op":"shape","shape":"arrow","at":"A20","width":40}"#);
    let wide = apply(
        &arrow,
        &ops(r#"{"op":"object","name":"直線矢印コネクタ 2","lineWidth":3}"#),
    )
    .unwrap();
    assert!(
        shape(&wide, "直線矢印コネクタ 2")
            .contains(r#"<a:ln w="38100"><a:tailEnd type="triangle"/></a:ln>"#)
    );
    let red = apply(
        &wide,
        &ops(r##"{"op":"object","name":"直線矢印コネクタ 2","line":"#C00000"}"##),
    )
    .unwrap();
    assert!(shape(&red, "直線矢印コネクタ 2").contains(r#"<a:ln w="38100"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:tailEnd type="triangle"/></a:ln>"#));
    // A picture gets a frame.
    let framed = apply(
        &out,
        &ops(r##"{"op":"object","name":"図 4","line":"#000000"}"##),
    )
    .unwrap();
    assert!(shape(&framed, "図 4").contains(r#"</a:prstGeom><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></xdr:spPr>"#));

    // Rotation and flips: the transform's attributes; 0 and false take them away.
    let turned = apply(
        &styled,
        &ops(r#"{"op":"object","name":"正方形/長方形 2","rotation":-30,"flipH":true}"#),
    )
    .unwrap();
    assert!(
        shape(&turned, "正方形/長方形 2").contains(r#"<a:xfrm rot="19800000" flipH="1">"#),
        "{}",
        shape(&turned, "正方形/長方形 2")
    );
    let shown = &view(&turned)["sheets"][0]["objects"][1];
    assert_eq!(
        (
            shown["rot"].clone(),
            shown["flipH"].clone(),
            shown["el"].clone()
        ),
        (json!(330.0), json!(true), json!("sp"))
    );
    assert!(shown["html"].as_str().unwrap().contains("rotate(330deg)"));
    let back = apply(
        &turned,
        &ops(r#"{"op":"object","name":"正方形/長方形 2","rotation":360,"flipH":false}"#),
    )
    .unwrap();
    assert!(shape(&back, "正方形/長方形 2").contains("<a:xfrm><a:off"));
    // Turned sideways, its anchor is its box turned a quarter around its center (Excel's); its
    // transform keeps the box. Moved so, the anchor stays turned; turned back, it is the box.
    let side = apply(
        &out,
        &ops(r#"{"op":"object","name":"正方形/長方形 2","rotation":90}"#),
    )
    .unwrap();
    let doc = shape(&side, "正方形/長方形 2");
    // (The box 50 × 30 at 0, 240: the anchor 30 × 50 at 10, 230.)
    assert!(doc.contains(r#"<xdr:from><xdr:col>0</xdr:col><xdr:colOff>95250</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>95250</xdr:rowOff></xdr:from><xdr:to><xdr:col>0</xdr:col><xdr:colOff>381000</xdr:colOff><xdr:row>13</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>"#), "{doc}");
    assert!(
        doc.contains(
            r#"<a:xfrm rot="5400000"><a:off x="0" y="2286000"/><a:ext cx="476250" cy="285750"/>"#
        ),
        "{doc}"
    );
    let moved = apply(
        &side,
        &ops(r#"{"op":"object","name":"正方形/長方形 2","at":"B15"}"#),
    )
    .unwrap();
    let doc = shape(&moved, "正方形/長方形 2");
    assert!(doc.contains(r#"<a:ext cx="476250" cy="285750"/>"#), "{doc}");
    let o = objects(&moved, 0)
        .into_iter()
        .find(|o| o.1 == "正方形/長方形 2")
        .unwrap();
    // (B15 is at 118, 300: the anchor from 128, 290 to 158, 340.)
    assert_eq!(
        (o.2, o.3),
        (json!([1, 10.0, 13, 10.0]), json!([1, 40.0, 16, 0.0]))
    );
    let back = apply(
        &side,
        &ops(r#"{"op":"object","name":"正方形/長方形 2","rotation":180}"#),
    )
    .unwrap();
    assert_eq!(
        shape(&back, "正方形/長方形 2").split("<xdr:sp ").next(),
        shape(&out, "正方形/長方形 2").split("<xdr:sp ").next()
    );

    // Moved and turned in one edit.
    let both = apply(
        &out,
        &ops(r#"{"op":"object","name":"正方形/長方形 2","at":"B15","rotation":90}"#),
    )
    .unwrap();
    let doc = shape(&both, "正方形/長方形 2");
    assert!(
        doc.contains(r#"<xdr:from><xdr:col>1</xdr:col>"#)
            && doc.contains(r#"<a:xfrm rot="5400000">"#),
        "{doc}"
    );

    // The order: the drawing's order is the order they lie over one another.
    let names = |bytes: &[u8]| -> Vec<String> {
        objects(bytes, 0)
            .into_iter()
            .map(|o| o.1)
            .filter(|n| !n.is_empty())
            .collect()
    };
    assert_eq!(
        names(&out),
        ["Chart 1", "正方形/長方形 2", "直線コネクタ 3", "図 4"]
    );
    let order = |name: &str, how: &str| {
        names(
            &apply(
                &out,
                &ops(&format!(
                    r#"{{"op":"object","name":"{name}","order":"{how}"}}"#
                )),
            )
            .unwrap(),
        )
    };
    assert_eq!(
        order("正方形/長方形 2", "front"),
        ["Chart 1", "直線コネクタ 3", "図 4", "正方形/長方形 2"]
    );
    assert_eq!(
        order("正方形/長方形 2", "forward"),
        ["Chart 1", "直線コネクタ 3", "正方形/長方形 2", "図 4"]
    );
    assert_eq!(
        order("正方形/長方形 2", "backward"),
        ["正方形/長方形 2", "Chart 1", "直線コネクタ 3", "図 4"]
    );
    assert_eq!(
        order("図 4", "back"),
        ["図 4", "Chart 1", "正方形/長方形 2", "直線コネクタ 3"]
    );
    // Already in front: nothing to write.
    assert_eq!(
        apply(
            &out,
            &ops(r#"{"op":"object","name":"図 4","order":"front"}"#)
        )
        .unwrap(),
        out
    );
    // Styled and put in front at once.
    let front = apply(
        &out,
        &ops(r##"{"op":"object","name":"正方形/長方形 2","fill":"#123456","order":"front"}"##),
    )
    .unwrap();
    assert_eq!(names(&front).last().unwrap(), "正方形/長方形 2");
    assert!(shape(&front, "正方形/長方形 2").contains("123456"));

    // Refused: a fill of a line, the format of a chart; colors and orders not so written.
    let no = |json: &str| apply(&out, &ops(json)).unwrap_err();
    assert!(no(r##"{"op":"object","name":"直線コネクタ 3","fill":"#FF0000"}"##).contains("線"));
    assert!(no(r#"{"op":"object","name":"Chart 1","rotation":10}"#).contains("グラフ"));
    // (A chart is moved and put elsewhere in the order.)
    assert_eq!(order("Chart 1", "front").last().unwrap(), "Chart 1");
    for bad in [
        json!({"op":"object","name":"x","fill":"red"}),
        json!({"op":"object","name":"x","order":"top"}),
        json!({"op":"object","name":"x","lineWidth":-1}),
    ] {
        assert!(parse_ops(&bad, Some("Data")).is_err(), "{bad}");
    }
    // The file reads back.
    let mut book = Book::open_bytes(&front).unwrap();
    assert!(book.sheet_summary(0).is_ok());
}

/// The controls the grid draws on a sheet: (name, shape number, from, to).
fn controls_of(bytes: &[u8], sheet: usize) -> Vec<(String, u64, Json, Json)> {
    view(bytes)["sheets"][sheet]["objects"]
        .as_array()
        .cloned()
        .unwrap_or_default()
        .into_iter()
        .filter(|o| o["kind"] == json!("control"))
        .map(|o| {
            (
                o["name"].as_str().unwrap_or("").to_string(),
                o["shapeId"].as_u64().unwrap_or(0),
                o["from"].clone(),
                o["to"].clone(),
            )
        })
        .collect()
}

#[test]
fn form_controls_in_vml_put_moved_and_deleted() {
    // A sheet whose controls are in VML alone (XlsxWriter's button, as Excel 2007 writes them):
    // a new one is too, after its shapes.
    let out = edit(
        r#"{"op":"control","control":"checkBox","at":"H18","width":100,"height":22,"link":"$I$18","checked":true,"text":"済み"}"#,
    );
    // (The linked cell gets the state, as Excel writes it.)
    assert_eq!(
        changed(BOOK, &out),
        ["xl/drawings/vmlDrawing1.vml", "xl/worksheets/sheet1.xml"]
    );
    assert_eq!(cell(&out, "Data", "I18", Read::Value), "true");
    let vml = part(&out, "xl/drawings/vmlDrawing1.vml");
    assert!(vml.ends_with(r#"<x:ClientData ObjectType="Checkbox"><x:Anchor>7, 0, 17, 0, 8, 28, 18, 2</x:Anchor><x:AutoFill>False</x:AutoFill><x:AutoLine>False</x:AutoLine><x:TextVAlign>Center</x:TextVAlign><x:FmlaLink>$I$18</x:FmlaLink><x:Checked>1</x:Checked><x:NoThreeD/></x:ClientData></v:shape></xml>"#), "{vml}");
    assert!(vml.contains(r##"<v:shape id="_x0000_s1027" type="#_x0000_t201" style="position:absolute;z-index:3;mso-wrap-style:tight;margin-left:435pt;margin-top:270pt;width:75pt;height:16.5pt""##), "{vml}");
    assert!(vml.contains(">済み</font>"));
    assert_eq!(vml.matches("id=\"_x0000_t201\"").count(), 1);
    let drawn = controls_of(&out, 0);
    assert_eq!(drawn[1].1, 1027);
    assert_eq!(
        (drawn[1].2.clone(), drawn[1].3.clone()),
        (json!([7, 0.0, 17, 0.0]), json!([8, 28.0, 18, 2.0]))
    );

    // A sheet without VML gets a drawing of its own, numbered in a block no other has.
    let out = edit(
        r#"{"op":"control","sheet":"Summary","control":"button","at":"C2","macro":"Module1.run"}"#,
    );
    assert_eq!(
        changed(BOOK, &out),
        [
            "xl/drawings/vmlDrawing2.vml",
            "xl/worksheets/_rels/sheet2.xml.rels",
            "xl/worksheets/sheet2.xml"
        ]
    );
    let vml = part(&out, "xl/drawings/vmlDrawing2.vml");
    assert!(
        vml.contains(r#"<o:idmap v:ext="edit" data="2"/>"#) && vml.contains(r#"id="_x0000_s2049""#),
        "{vml}"
    );
    assert!(
        vml.contains("<x:FmlaMacro>[0]!Module1.run</x:FmlaMacro>")
            && vml.contains(">ボタン 1</font>")
    );
    assert!(part(&out, SHEET2).contains(r#"<legacyDrawing r:id="rId1"/></worksheet>"#));
    assert_eq!(controls_of(&out, 1)[0].1, 2049);

    // Moved: its VML anchor and style; deleted: its shape (the note's box stays).
    let moved = edit(r#"{"op":"object","shape":1025,"at":"B12","dx":4}"#);
    assert_eq!(changed(BOOK, &moved), ["xl/drawings/vmlDrawing1.vml"]);
    let vml = part(&moved, "xl/drawings/vmlDrawing1.vml");
    assert!(
        vml.contains("<x:Anchor>1, 4, 11, 0, 2, 20, 12, 10</x:Anchor>"),
        "{vml}"
    );
    assert!(
        vml.contains("margin-left:91.5pt;margin-top:180pt;width:66pt;height:22.5pt"),
        "{vml}"
    );
    let gone = edit(r#"{"op":"object","shape":1025,"delete":true}"#);
    let vml = part(&gone, "xl/drawings/vmlDrawing1.vml");
    assert!(!vml.contains("_x0000_s1025") && vml.contains("ObjectType=\"Note\""));
    assert!(controls_of(&gone, 0).is_empty());
    assert_eq!(cell(&gone, "Data", "D6", Read::Formula), "=SUM(D2:D5)");

    // Refused: a note's box, what a control does not have, kinds there are not.
    assert!(
        refused(r#"{"op":"object","shape":1026,"delete":true}"#)
            .contains("コントロール `_x0000_s1026` がありません")
    );
    assert!(
        refused(r##"{"op":"object","shape":1025,"fill":"#FF0000"}"##)
            .contains("塗りつぶし・線・回転・重なり順は変えられません")
    );
    assert!(refused(r#"{"op":"object","shape":1099,"delete":true}"#).contains("1099"));
    for bad in [
        json!({"op":"control","control":"slider","at":"A1"}),
        json!({"op":"control","control":"checkBox","at":"A1","link":"どこか"}),
        json!({"op":"control","control":"spinner","at":"A1","min":10,"max":5}),
        json!({"op":"control","control":"button","at":"A1","macro":"a b"}),
    ] {
        assert!(parse_ops(&bad, Some("Data")).is_err(), "{bad}");
    }
}

#[test]
fn form_controls_as_excel_2010_keeps_them() {
    // A sheet saved by Excel 2010 or later: `<controls>`, the properties, the mirrors.
    let sample = crate::view::tests::sample();
    let sheet1 = "xl/worksheets/sheet1.xml";
    let ed = |bytes: &[u8], json: &str| {
        apply(
            bytes,
            &parse_ops(&serde_json::from_str(json).unwrap(), Some("Sheet1")).unwrap(),
        )
    };
    let out = ed(&sample, r#"{"op":"control","control":"listBox","at":"B2","width":80,"height":60,"items":"$K$18:$K$20","link":"$L$1"}"#).unwrap();
    assert_eq!(
        changed(&sample, &out),
        [
            "[Content_Types].xml",
            "xl/ctrlProps/ctrlProp4.xml",
            "xl/drawings/drawing1.xml",
            "xl/drawings/vmlDrawing1.vml",
            "xl/worksheets/_rels/sheet1.xml.rels",
            sheet1
        ]
    );
    assert_eq!(
        part(&out, "xl/ctrlProps/ctrlProp4.xml"),
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<formControlPr xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" objectType=\"List\" dx=\"16\" fmlaRange=\"$K$18:$K$20\" fmlaLink=\"$L$1\" noThreeD=\"1\" sel=\"0\"/>"
    );
    assert!(part(&out, "[Content_Types].xml").contains(r#"<Override PartName="/xl/ctrlProps/ctrlProp4.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>"#));
    let doc = part(&out, sheet1);
    assert!(doc.contains(r#"<control shapeId="1031" r:id="rId5" name="リスト 7"><controlPr defaultSize="0" autoFill="0" autoLine="0" autoPict="0"><anchor moveWithCells="1"><from><xdr:col>1</xdr:col>"#), "{doc}");
    assert!(doc.contains(r#"<mc:Fallback><control shapeId="1031" r:id="rId5" name="リスト 7"/></mc:Fallback></mc:AlternateContent></controls>"#));
    assert!(part(&out, "xl/drawings/drawing1.xml").contains(r#"<xdr:cNvPr id="1031" name="リスト 7" hidden="1"><a:extLst><a:ext uri="{63B3BB69-23CF-44E3-9099-C40C66FF867C}"><a14:compatExt spid="_x0000_s1031"/>"#));
    assert!(
        part(&out, "xl/drawings/vmlDrawing1.vml").contains(r#"<x:ClientData ObjectType="List">"#)
    );
    let drawn = controls_of(&out, 0);
    let list = drawn.iter().find(|c| c.0 == "リスト 7").unwrap();
    assert_eq!((list.1, list.2.clone()), (1031, json!([1, 0.0, 1, 0.0])));

    // Moved by its name: its entry, its VML shape and its mirror.
    let moved = ed(&sample, r#"{"op":"object","name":"Button 1","at":"B3"}"#).unwrap();
    assert_eq!(
        changed(&sample, &moved),
        [
            "xl/drawings/drawing1.xml",
            "xl/drawings/vmlDrawing1.vml",
            sheet1
        ]
    );
    let doc = part(&moved, sheet1);
    assert!(doc.contains(r#"name="Button 1"><controlPr defaultSize="0" autoLine="0"><anchor moveWithCells="1"><from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></from><to><xdr:col>3</xdr:col>"#), "{doc}");
    assert!(
        part(&moved, "xl/drawings/vmlDrawing1.vml")
            .contains("<x:Anchor>1, 0, 2, 0, 3, 0, 4, 0</x:Anchor>")
    );
    let button = controls_of(&moved, 0)
        .into_iter()
        .find(|c| c.0 == "Button 1")
        .unwrap();
    assert_eq!(
        (button.2, button.3),
        (json!([1, 0.0, 2, 0.0]), json!([3, 0.0, 4, 0.0]))
    );

    // Deleted: its entry, its properties (the part, its relationship, its content type), its
    // VML shape; the others stay.
    let gone = ed(
        &sample,
        r#"{"op":"object","name":"Check Box 3","delete":true}"#,
    )
    .unwrap();
    assert!(!parts(&gone).contains_key("xl/ctrlProps/ctrlProp2.xml"));
    assert!(!part(&gone, "[Content_Types].xml").contains("ctrlProp2"));
    assert!(!part(&gone, "xl/worksheets/_rels/sheet1.xml.rels").contains("ctrlProp2"));
    assert!(!part(&gone, sheet1).contains("Check Box 3"));
    assert!(!part(&gone, "xl/drawings/vmlDrawing1.vml").contains("_x0000_s1027"));
    let names: Vec<String> = controls_of(&gone, 0).into_iter().map(|c| c.0).collect();
    assert!(
        !names.contains(&"Check Box 3".to_string()) && names.contains(&"List Box 4".to_string()),
        "{names:?}"
    );
    // With its mirror.
    let gone = ed(
        &sample,
        r#"{"op":"object","name":"Button 1","delete":true}"#,
    )
    .unwrap();
    assert!(!part(&gone, "xl/drawings/drawing1.xml").contains("_x0000_s1025"));
}

#[test]
fn form_controls_given_texts_links_and_values() {
    // In VML alone: the text of XlsxWriter's button (its font kept).
    let out = edit(r#"{"op":"object","shape":1025,"text":"集計\n実行"}"#);
    assert_eq!(changed(BOOK, &out), ["xl/drawings/vmlDrawing1.vml"]);
    let vml = part(&out, "xl/drawings/vmlDrawing1.vml");
    assert!(vml.contains(r##"<div style="text-align:center"><font face="Calibri" size="220" color="#000000">集計</font></div><div style="text-align:center"><font face="Calibri" size="220" color="#000000">実行</font></div></v:textbox>"##), "{vml}");
    let drawn = &view(&out)["sheets"][0]["objects"];
    let button = drawn
        .as_array()
        .unwrap()
        .iter()
        .find(|o| o["shapeId"] == json!(1025))
        .unwrap();
    assert_eq!(button["text"], json!("集計\n実行"));
    assert_eq!(button["form"]["type"], json!("Button"));

    // A check box linked and checked: its VML, and the cell it is linked to.
    let added = edit(r#"{"op":"control","control":"checkBox","at":"H18"}"#);
    let checked = apply(
        &added,
        &ops(r#"{"op":"object","shape":1027,"link":"$I$18","checked":true}"#),
    )
    .unwrap();
    let vml = part(&checked, "xl/drawings/vmlDrawing1.vml");
    assert!(
        vml.contains("<x:FmlaLink>$I$18</x:FmlaLink><x:Checked>1</x:Checked>"),
        "{vml}"
    );
    assert_eq!(cell(&checked, "Data", "I18", Read::Value), "true");
    let form = &view(&checked)["sheets"][0]["objects"][2]["form"];
    assert_eq!(
        (form["link"].clone(), form["checked"].clone()),
        (json!("$I$18"), json!(true))
    );
    // Unchecked and no longer linked.
    let off = apply(
        &checked,
        &ops(r#"{"op":"object","shape":1027,"checked":false,"link":null}"#),
    )
    .unwrap();
    let vml = part(&off, "xl/drawings/vmlDrawing1.vml");
    assert!(
        !vml.contains("<x:Checked>") && !vml.contains("$I$18"),
        "{vml}"
    );

    // A spin button: its range and value, the value in the cell.
    let spin = edit(r#"{"op":"control","control":"spinner","at":"F15","link":"Data!$G$15"}"#);
    let set = apply(
        &spin,
        &ops(r#"{"op":"object","shape":1027,"min":5,"max":50,"step":5,"value":20}"#),
    )
    .unwrap();
    let vml = part(&set, "xl/drawings/vmlDrawing1.vml");
    assert!(
        vml.contains("<x:Val>20</x:Val><x:Min>5</x:Min><x:Max>50</x:Max><x:Inc>5</x:Inc>"),
        "{vml}"
    );
    assert_eq!(cell(&set, "Data", "G15", Read::Value), "20");
    // A range that leaves the value out: the value comes into it.
    let narrowed = apply(&set, &ops(r#"{"op":"object","shape":1027,"max":10}"#)).unwrap();
    assert!(part(&narrowed, "xl/drawings/vmlDrawing1.vml").contains("<x:Val>10</x:Val>"));
    assert_eq!(cell(&narrowed, "Data", "G15", Read::Value), "10");

    // Refused: what a kind does not have, values out of their range, settings of a shape.
    let no = |bytes: &[u8], json: &str| apply(bytes, &ops(json)).unwrap_err();
    assert!(
        no(BOOK, r#"{"op":"object","shape":1025,"link":"$A$1"}"#)
            .contains("ボタンにはリンクするセルがありません")
    );
    assert!(no(&spin, r#"{"op":"object","shape":1027,"page":3}"#).contains("ページ"));
    assert!(
        no(
            &spin,
            r#"{"op":"object","shape":1027,"value":200,"max":100}"#
        )
        .contains("最大値")
    );
    assert!(no(&spin, r#"{"op":"object","shape":1027,"min":60,"max":50}"#).contains("最小値"));
    assert!(no(&spin, r#"{"op":"object","shape":1027,"text":"x"}"#).contains("文字"));
    assert!(
        no(BOOK, r#"{"op":"object","name":"Chart 1","link":"$A$1"}"#)
            .contains("フォーム コントロールではない")
    );
    assert!(
        parse_ops(
            &json!({"op":"object","shape":1025,"link":"どこ"}),
            Some("Data")
        )
        .is_err()
    );

    // As Excel 2010 keeps them: the properties too, and the text of the mirror.
    let sample = crate::view::tests::sample();
    let ed = |bytes: &[u8], json: &str| {
        apply(
            bytes,
            &parse_ops(&serde_json::from_str(json).unwrap(), Some("Sheet1")).unwrap(),
        )
    };
    let out = ed(&sample, r#"[{"op":"object","name":"Check Box 3","checked":true},{"op":"object","name":"List Box 4","items":"$K$18:$K$19","link":"$L$1"},{"op":"object","name":"Button 1","text":"実行する"}]"#).unwrap();
    assert_eq!(
        part(&out, "xl/ctrlProps/ctrlProp2.xml"),
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<formControlPr xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" objectType=\"CheckBox\" fmlaLink=\"$K$17\" lockText=\"1\" noThreeD=\"1\" checked=\"Checked\"/>"
    );
    assert!(
        part(&out, "xl/ctrlProps/ctrlProp3.xml")
            .contains(r#"fmlaRange="$K$18:$K$19" noThreeD="1" sel="2" fmlaLink="$L$1""#)
    );
    assert_eq!(cell(&out, "Sheet1", "K17", Read::Value), "true");
    assert!(part(&out, "xl/drawings/drawing1.xml").contains("<a:t>実行する</a:t>"));
    assert!(part(&out, "xl/drawings/vmlDrawing1.vml").contains(">実行する</font>"));
    let drawn = view(&out)["sheets"][0]["objects"]
        .as_array()
        .unwrap()
        .clone();
    let list = drawn
        .iter()
        .find(|o| o["name"] == json!("List Box 4"))
        .unwrap();
    assert_eq!(
        (list["form"]["items"].clone(), list["form"]["link"].clone()),
        (json!("$K$18:$K$19"), json!("$L$1"))
    );
}

/// A workbook file with some of its parts replaced (None: taken away).
fn with_parts(bytes: &[u8], change: &[(&str, Option<Vec<u8>>)]) -> Vec<u8> {
    let mut all = parts(bytes);
    for (name, data) in change {
        match data {
            Some(d) => {
                all.insert(name.to_string(), d.clone());
            }
            None => {
                all.remove(*name);
            }
        }
    }
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, data) in all {
        zip.start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
        std::io::Write::write_all(&mut zip, &data).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

#[test]
fn list_items_selected() {
    let sample = crate::view::tests::sample();
    let ed = |bytes: &[u8], json: &str| {
        apply(
            bytes,
            &parse_ops(&serde_json::from_str(json).unwrap(), Some("Sheet1")).unwrap(),
        )
    };
    let out = ed(
        &sample,
        r#"{"op":"object","name":"List Box 4","selected":3,"link":"$L$2"}"#,
    )
    .unwrap();
    assert!(part(&out, "xl/ctrlProps/ctrlProp3.xml").contains(r#"sel="3""#));
    assert!(part(&out, "xl/drawings/vmlDrawing1.vml").contains("<x:Sel>3</x:Sel>"));
    assert_eq!(cell(&out, "Sheet1", "L2", Read::Value), "3");
    let list = view(&out)["sheets"][0]["objects"]
        .as_array()
        .unwrap()
        .iter()
        .find(|o| o["name"] == json!("List Box 4"))
        .unwrap()
        .clone();
    assert_eq!(list["form"]["sel"], json!(3));
    // Three items: no fourth; a check box has none.
    assert!(
        ed(
            &sample,
            r#"{"op":"object","name":"List Box 4","selected":4}"#
        )
        .unwrap_err()
        .contains("3 まで")
    );
    assert!(
        ed(
            &sample,
            r#"{"op":"object","name":"Check Box 3","selected":1}"#
        )
        .unwrap_err()
        .contains("選択している項目")
    );
}

#[test]
fn activex_controls_changed_moved_and_deleted() {
    let sample = crate::view::tests::sample();
    let ed = |bytes: &[u8], json: &str| {
        apply(
            bytes,
            &parse_ops(&serde_json::from_str(json).unwrap(), Some("Sheet1")).unwrap(),
        )
    };
    let html = |bytes: &[u8], name: &str| -> Json {
        view(bytes)["sheets"][0]["objects"]
            .as_array()
            .unwrap()
            .iter()
            .find(|o| o["name"] == json!(name))
            .cloned()
            .unwrap_or(Json::Null)
    };
    // A command button's caption: in its data (its font and the rest kept).
    let out = ed(
        &sample,
        r#"{"op":"object","name":"CommandButton1","text":"OK"}"#,
    )
    .unwrap();
    assert_eq!(changed(&sample, &out), ["xl/activeX/activeX1.bin"]);
    let button = html(&out, "CommandButton1");
    assert!(
        button["html"].as_str().unwrap().contains(">OK<"),
        "{button}"
    );
    assert_eq!(
        (button["text"].clone(), button["form"]["type"].clone()),
        (json!("OK"), json!("CommandButton"))
    );
    let before =
        crate::activex::Stream::read("CommandButton", &parts(&sample)["xl/activeX/activeX1.bin"])
            .unwrap();
    let after =
        crate::activex::Stream::read("CommandButton", &parts(&out)["xl/activeX/activeX1.bin"])
            .unwrap();
    assert_eq!(after.number("ForeColor"), before.number("ForeColor"));
    assert!(
        parts(&out)["xl/activeX/activeX1.bin"]
            .starts_with(&parts(&sample)["xl/activeX/activeX1.bin"][..16])
    );

    // Moved and sized: its entry, its VML shape, its mirror, its own size.
    let moved = ed(
        &sample,
        r#"{"op":"object","name":"CommandButton1","at":"B2","width":96,"height":48}"#,
    )
    .unwrap();
    assert_eq!(
        changed(&sample, &moved),
        [
            "xl/activeX/activeX1.bin",
            "xl/drawings/drawing1.xml",
            "xl/drawings/vmlDrawing1.vml",
            "xl/worksheets/sheet1.xml"
        ]
    );
    let b = html(&moved, "CommandButton1");
    assert_eq!(b["from"], json!([1, 0.0, 1, 0.0]));
    let st =
        crate::activex::Stream::read("CommandButton", &parts(&moved)["xl/activeX/activeX1.bin"])
            .unwrap();
    assert_eq!(
        st.write().len(),
        parts(&moved)["xl/activeX/activeX1.bin"].len()
    );
    let vml = part(&moved, "xl/drawings/vmlDrawing1.vml");
    assert!(
        vml.contains("<x:Anchor>1, 0, 1, 0, 2, 24, 2, 23</x:Anchor>"),
        "{vml}"
    );

    // Deleted: its part, its data, its VML shape, its mirror.
    let gone = ed(
        &sample,
        r#"{"op":"object","name":"CommandButton1","delete":true}"#,
    )
    .unwrap();
    let names = parts(&gone);
    assert!(
        !names.contains_key("xl/activeX/activeX1.xml")
            && !names.contains_key("xl/activeX/activeX1.bin")
    );
    assert!(!names.contains_key("xl/activeX/_rels/activeX1.xml.rels"));
    assert!(!part(&gone, "[Content_Types].xml").contains("activeX1.xml"));
    assert!(!part(&gone, "xl/worksheets/sheet1.xml").contains("CommandButton1"));
    assert!(!part(&gone, "xl/drawings/vmlDrawing1.vml").contains("CommandButton1"));
    assert!(!part(&gone, "xl/drawings/drawing1.xml").contains("_x0000_s1026"));
    assert_eq!(html(&gone, "CommandButton1"), Json::Null);

    // What a command button does not have.
    assert!(
        ed(
            &sample,
            r#"{"op":"object","name":"CommandButton1","link":"$A$1"}"#
        )
        .unwrap_err()
        .contains("リンクするセル")
    );

    // A check box: its caption and state in its data, its linked cell in the sheet, the cell.
    let check = with_parts(
        &sample,
        &[
            (
                "xl/activeX/activeX1.xml",
                Some(
                    part(&sample, "xl/activeX/activeX1.xml")
                        .replace(
                            "D7053240-CE69-11CD-A777-00DD01143C57",
                            "8BD21D40-EC42-11CE-9E0D-00AA006002F3",
                        )
                        .into_bytes(),
                ),
            ),
            (
                "xl/activeX/activeX1.bin",
                Some(
                    [
                        &parts(&sample)["xl/activeX/activeX1.bin"][..16],
                        &crate::activex::tests::check_box("確認", "0")[..],
                    ]
                    .concat(),
                ),
            ),
        ],
    );
    let on = ed(
        &check,
        r#"{"op":"object","name":"CommandButton1","text":"同意する","checked":true,"link":"$M$1"}"#,
    )
    .unwrap();
    let st =
        crate::activex::Stream::read("MorphData", &parts(&on)["xl/activeX/activeX1.bin"]).unwrap();
    assert_eq!(
        (
            st.string("Caption").as_deref(),
            st.string("Value").as_deref()
        ),
        (Some("同意する"), Some("1"))
    );
    assert!(
        part(&on, "xl/worksheets/sheet1.xml")
            .contains(r#"<controlPr defaultSize="0" autoLine="0" linkedCell="$M$1">"#)
    );
    assert_eq!(cell(&on, "Sheet1", "M1", Read::Value), "true");
    let shown = html(&on, "CommandButton1");
    assert_eq!(
        (
            shown["form"]["type"].clone(),
            shown["form"]["link"].clone(),
            shown["form"]["checked"].clone()
        ),
        (json!("CheckBox"), json!("$M$1"), json!(true))
    );
    let off = ed(
        &on,
        r#"{"op":"object","name":"CommandButton1","link":null}"#,
    )
    .unwrap();
    assert!(!part(&off, "xl/worksheets/sheet1.xml").contains("linkedCell"));

    // Properties kept in the part (persistPropertyBag).
    let bag = with_parts(
        &sample,
        &[("xl/activeX/activeX1.xml", Some(part(&sample, "xl/activeX/activeX1.xml").replace("persistStream\" r:id=\"rId1\"/>", "persistPropertyBag\"><ax:ocxPr ax:name=\"Caption\" ax:value=\"前\"/></ax:ocx>").into_bytes()))],
    );
    let out = ed(
        &bag,
        r#"{"op":"object","name":"CommandButton1","text":"後"}"#,
    )
    .unwrap();
    assert!(
        part(&out, "xl/activeX/activeX1.xml")
            .contains(r#"<ax:ocxPr ax:name="Caption" ax:value="後"/>"#)
    );
}

/// The ActiveX controls the grid draws on a sheet: (name, Microsoft Forms type).
fn activex_of(bytes: &[u8], sheet: usize) -> Vec<(String, String)> {
    view(bytes)["sheets"][sheet]["objects"]
        .as_array()
        .cloned()
        .unwrap_or_default()
        .into_iter()
        .filter(|o| o["kind"] == json!("control") && o.get("activex").is_some())
        .map(|o| {
            (
                o["name"].as_str().unwrap_or("").to_string(),
                o["activex"].as_str().unwrap_or("").to_string(),
            )
        })
        .collect()
}

#[test]
fn activex_controls_put_on_a_sheet() {
    // Each kind, on a sheet with VML (a button, a note) and no `<controls>`.
    let kinds = [
        (
            "commandButton",
            "CommandButton",
            "CommandButton",
            Some("CommandButton1"),
        ),
        ("textBox", "TextBox", "MorphData", None),
        ("checkBox", "CheckBox", "MorphData", Some("CheckBox1")),
        (
            "optionButton",
            "OptionButton",
            "MorphData",
            Some("OptionButton1"),
        ),
        (
            "toggleButton",
            "ToggleButton",
            "MorphData",
            Some("ToggleButton1"),
        ),
        ("comboBox", "ComboBox", "MorphData", None),
        ("listBox", "ListBox", "MorphData", None),
        ("label", "Label", "Label", Some("Label1")),
        ("spinButton", "SpinButton", "SpinButton", None),
        ("scrollBar", "ScrollBar", "ScrollBar", None),
    ];
    for (key, forms, class, caption) in kinds {
        let out = edit(&format!(
            r#"{{"op":"activeX","control":"{key}","at":"H18","width":96,"height":32}}"#
        ));
        let name = format!("{forms}1");
        assert_eq!(
            changed(BOOK, &out),
            [
                "[Content_Types].xml",
                "xl/activeX/_rels/activeX1.xml.rels",
                "xl/activeX/activeX1.bin",
                "xl/activeX/activeX1.xml",
                "xl/drawings/vmlDrawing1.vml",
                "xl/worksheets/_rels/sheet1.xml.rels",
                "xl/worksheets/sheet1.xml"
            ],
            "{key}"
        );
        // Its part, its data's relationship, their content types (the file's `bin` is its VBA
        // project's: the data's own).
        let clsid = crate::activex::clsid_of(forms).unwrap();
        assert_eq!(
            part(&out, "xl/activeX/activeX1.xml"),
            format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>\n<ax:ocx ax:classid=\"{{{clsid}}}\" ax:persistence=\"persistStreamInit\" r:id=\"rId1\" xmlns:ax=\"http://schemas.microsoft.com/office/2006/activeX\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>"
            )
        );
        assert!(part(&out, "xl/activeX/_rels/activeX1.xml.rels").contains(r#"<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary" Target="activeX1.bin"/>"#));
        let types = part(&out, "[Content_Types].xml");
        assert!(types.contains(r#"<Override PartName="/xl/activeX/activeX1.xml" ContentType="application/vnd.ms-office.activeX+xml"/>"#));
        assert!(types.contains(r#"<Override PartName="/xl/activeX/activeX1.bin" ContentType="application/vnd.ms-office.activeX"/>"#));
        assert!(types.contains(
            r#"<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>"#
        ));
        // The sheet: its relationship, `<controls>` after `<legacyDrawing>`.
        assert!(part(&out, "xl/worksheets/_rels/sheet1.xml.rels").contains(r#"Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/control" Target="../activeX/activeX1.xml""#));
        let doc = part(&out, SHEET1);
        assert!(doc.contains(&format!(r#"<legacyDrawing r:id="rId3"/><controls><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><mc:Choice Requires="x14"><control shapeId="1027" r:id="rId5" name="{name}"><controlPr defaultSize="0" autoLine="0" autoPict="0"><anchor moveWithCells="1" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><from><xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>17</xdr:row><xdr:rowOff>0</xdr:rowOff></from>"#)), "{doc}");
        assert!(doc.ends_with(&format!(r#"</controlPr></control></mc:Choice><mc:Fallback><control shapeId="1027" r:id="rId5" name="{name}"/></mc:Fallback></mc:AlternateContent></controls></worksheet>"#)), "{doc}");
        // Its VML shape, named after it.
        let vml = part(&out, "xl/drawings/vmlDrawing1.vml");
        assert!(vml.contains(&format!(r##"<v:shape id="{name}" o:spid="_x0000_s1027" type="#_x0000_t201" style="position:absolute;z-index:3;margin-left:435pt;margin-top:270pt;width:72pt;height:24pt""##)), "{vml}");
        assert!(vml.ends_with(r#"<x:ClientData ObjectType="Pict"><x:SizeWithCells/><x:Anchor>7, 0, 17, 0, 8, 24, 18, 12</x:Anchor><x:CF>Pict</x:CF><x:AutoPict/></x:ClientData></v:shape></xml>"#), "{vml}");

        // Its data: its CLSID and a record read back whole, its caption, its size.
        let bin = &parts(&out)["xl/activeX/activeX1.bin"];
        let st = crate::activex::Stream::read(class, bin).unwrap_or_else(|| panic!("{key}"));
        assert_eq!(st.write(), *bin, "{key}");
        assert_eq!(st.string("Caption").as_deref(), caption, "{key}");
        assert_eq!(st.size(), Some((2540, 847)), "{key}");
        let mut ax = umya_spreadsheet::drawingml::controls::ActiveX {
            class,
            ..Default::default()
        };
        umya_spreadsheet::drawingml::controls::read_activex_stream(&mut ax, &bin[16..]);
        assert_eq!(ax.caption.as_deref(), caption, "{key}");
        if matches!(class, "CommandButton" | "Label" | "MorphData") {
            assert_eq!(
                (ax.font.as_deref(), ax.size),
                (Some("MS UI Gothic"), Some(9.0)),
                "{key}"
            );
        }
        // The grid draws it as an ActiveX control of its type.
        assert!(
            activex_of(&out, 0).contains(&(name.clone(), forms.to_string())),
            "{key}: {:?}",
            activex_of(&out, 0)
        );
    }

    // Names as Excel gives them; a text; a name given.
    let two = edit(
        r#"[{"op":"activeX","at":"B12"},{"op":"activeX","control":"commandButton","at":"B15","text":"実行"},{"op":"activeX","control":"textBox","at":"D15","name":"txtName","text":"既定"}]"#,
    );
    let names: Vec<String> = activex_of(&two, 0).into_iter().map(|c| c.0).collect();
    assert_eq!(names, ["CommandButton1", "CommandButton2", "txtName"]);
    assert!(parts(&two).contains_key("xl/activeX/activeX3.bin"));
    let st = crate::activex::Stream::read("CommandButton", &parts(&two)["xl/activeX/activeX2.bin"])
        .unwrap();
    assert_eq!(st.string("Caption").as_deref(), Some("実行"));
    let st =
        crate::activex::Stream::read("MorphData", &parts(&two)["xl/activeX/activeX3.bin"]).unwrap();
    assert_eq!(st.string("Value").as_deref(), Some("既定"));
    let doc = part(&two, SHEET1);
    assert_eq!(doc.matches("<controls>").count(), 1);
    assert!(
        doc.contains(r#"<control shapeId="1028" r:id="rId6" name="CommandButton2">"#),
        "{doc}"
    );
    // Refused: a name the sheet has, names VBA does not take, kinds there are not, a text a
    // kind does not show.
    assert!(
        apply(
            &two,
            &ops(r#"{"op":"activeX","at":"A1","name":"commandbutton1"}"#)
        )
        .unwrap_err()
        .contains("commandbutton1")
    );
    for bad in [
        json!({"op":"activeX","control":"slider","at":"A1"}),
        json!({"op":"activeX","at":"A1","name":"1st"}),
        json!({"op":"activeX","at":"A1","name":"a b"}),
        json!({"op":"activeX","control":"spinButton","at":"A1","text":"x"}),
        json!({"op":"activeX","control":"commandButton"}),
    ] {
        assert!(parse_ops(&bad, Some("Data")).is_err(), "{bad}");
    }

    // A sheet without VML: its drawing made, numbered in a block no other has; `<controls>`
    // in the schema's place.
    let out = edit(r#"{"op":"activeX","sheet":"Summary","control":"checkBox","at":"C2"}"#);
    let vml = part(&out, "xl/drawings/vmlDrawing2.vml");
    assert!(
        vml.contains(r#"<o:idmap v:ext="edit" data="2"/>"#)
            && vml.contains(r#"<v:shape id="CheckBox1" o:spid="_x0000_s2049""#),
        "{vml}"
    );
    let doc = part(&out, SHEET2);
    assert!(
        doc.contains(r#"<legacyDrawing r:id="rId1"/><controls><mc:AlternateContent"#),
        "{doc}"
    );
    assert_eq!(
        activex_of(&out, 1),
        [("CheckBox1".to_string(), "CheckBox".to_string())]
    );
    // Its caption, moved, deleted.
    let named = apply(
        &out,
        &ops(
            r#"{"op":"object","sheet":"Summary","name":"CheckBox1","text":"確認","checked":true}"#,
        ),
    )
    .unwrap();
    let st = crate::activex::Stream::read("MorphData", &parts(&named)["xl/activeX/activeX1.bin"])
        .unwrap();
    assert_eq!(
        (
            st.string("Caption").as_deref(),
            st.string("Value").as_deref()
        ),
        (Some("確認"), Some("1"))
    );
    let moved = apply(
        &out,
        &ops(r#"{"op":"object","sheet":"Summary","name":"CheckBox1","at":"E4","width":120,"height":40}"#),
    )
    .unwrap();
    assert!(
        part(&moved, "xl/drawings/vmlDrawing2.vml")
            .contains("<x:Anchor>4, 0, 3, 0, 5, 48, 5, 0</x:Anchor>"),
        "{}",
        part(&moved, "xl/drawings/vmlDrawing2.vml")
    );
    assert!(part(&moved, SHEET2).contains("<from><xdr:col>4</xdr:col>"));
    let st = crate::activex::Stream::read("MorphData", &parts(&moved)["xl/activeX/activeX1.bin"])
        .unwrap();
    assert_eq!(st.size(), Some((3175, 1058)));
    let gone = apply(
        &out,
        &ops(r#"{"op":"object","sheet":"Summary","name":"CheckBox1","delete":true}"#),
    )
    .unwrap();
    let names = parts(&gone);
    assert!(
        !names.contains_key("xl/activeX/activeX1.xml")
            && !names.contains_key("xl/activeX/activeX1.bin")
    );
    assert!(!part(&gone, "[Content_Types].xml").contains("activeX1"));
    assert!(!part(&gone, SHEET2).contains("<controls>"));
    assert!(activex_of(&gone, 1).is_empty());

    // What the compiler reads of the sheet's controls (ras-ovba): their names and types.
    let file = std::env::temp_dir().join(format!("ras-activex-{}.xlsm", std::process::id()));
    std::fs::write(&file, &two).unwrap();
    let listed = ras_ovba::workbook_sheet_controls(&file).unwrap().unwrap();
    std::fs::remove_file(&file).ok();
    let data = listed.iter().find(|s| s.code_name == "Sheet1").unwrap();
    let got: Vec<(&str, &str)> = data
        .controls
        .iter()
        .map(|c| (c.name.as_str(), c.class.as_str()))
        .collect();
    assert_eq!(
        got,
        [
            ("CommandButton1", "CommandButton"),
            ("CommandButton2", "CommandButton"),
            ("txtName", "TextBox")
        ]
    );
}

#[test]
fn activex_controls_put_where_excel_2010_keeps_controls() {
    // A sheet with `<controls>` (form and ActiveX controls): the next name, the entry with the
    // others.
    let sample = crate::view::tests::sample();
    let ed = |bytes: &[u8], json: &str| {
        apply(
            bytes,
            &parse_ops(&serde_json::from_str(json).unwrap(), Some("Sheet1")).unwrap(),
        )
    };
    let out = ed(&sample, r#"{"op":"activeX","at":"B20"}"#).unwrap();
    let doc = part(&out, "xl/worksheets/sheet1.xml");
    assert_eq!(doc.matches("<controls>").count(), 1);
    assert!(doc.contains(r#"name="CommandButton2"><controlPr"#), "{doc}");
    assert!(parts(&out).contains_key("xl/activeX/activeX2.bin"));
    // (The file's `bin` is the controls' already.)
    assert!(!part(&out, "[Content_Types].xml").contains("activeX2.bin"));
    let names: Vec<(String, String)> = activex_of(&out, 0);
    assert!(
        names.contains(&("CommandButton1".into(), "CommandButton".into()))
            && names.contains(&("CommandButton2".into(), "CommandButton".into())),
        "{names:?}"
    );

    // `<controls>` in an `mc:AlternateContent` with a fallback: both get it.
    let sheet = part(&sample, "xl/worksheets/sheet1.xml");
    let (head, rest) = sheet.split_once("<controls>").unwrap();
    let (inner, tail) = rest.split_once("</controls>").unwrap();
    let wrapped = format!(
        "{head}<mc:AlternateContent><mc:Choice Requires=\"x14\"><controls>{inner}</controls></mc:Choice><mc:Fallback><controls><control shapeId=\"1026\" r:id=\"rId4\" name=\"CommandButton1\"/></controls></mc:Fallback></mc:AlternateContent>{tail}"
    );
    let book = with_parts(
        &sample,
        &[("xl/worksheets/sheet1.xml", Some(wrapped.into_bytes()))],
    );
    let out = ed(&book, r#"{"op":"activeX","control":"label","at":"B20"}"#).unwrap();
    let doc = part(&out, "xl/worksheets/sheet1.xml");
    assert_eq!(doc.matches(r#"name="Label1""#).count(), 3, "{doc}");
    assert!(
        doc.contains(r#"name="Label1"/></controls></mc:Fallback>"#),
        "{doc}"
    );
}

/// The `xf` (in `cellXfs`) of the cell `at` of the first sheet.
fn xf_of(bytes: &[u8], at: &str) -> String {
    let sheet = part(bytes, SHEET1);
    let tag = &sheet[sheet.find(&format!("<c r=\"{at}\"")).unwrap()..];
    let tag = &tag[..tag.find('>').unwrap()];
    let s: usize = tag
        .split(" s=\"")
        .nth(1)
        .map_or(0, |s| s[..s.find('"').unwrap()].parse().unwrap());
    let styles = part(bytes, "xl/styles.xml");
    let xfs = &styles[styles.find("<cellXfs").unwrap()..styles.find("</cellXfs>").unwrap()];
    xfs.split("<xf ").nth(s + 1).unwrap().to_string()
}

#[test]
fn links_get_the_cell_style_hyperlink() {
    // As the grid gives a new link: Excel's cell style ハイパーリンク (built-in 8).
    let link = r#"[{"op":"format","range":"C2","set":{"style":"Hyperlink"}},{"op":"link","range":"C2","url":"https://example.org/"}]"#;
    let out = edit(link);
    let styles = part(&out, "xl/styles.xml");
    let named = &styles[styles.find("<cellStyles").unwrap()..];
    let id = named
        .split("<cellStyle ")
        .find(|s| s.contains("builtinId=\"8\""))
        .and_then(|s| s.split("xfId=\"").nth(1))
        .map(|s| s[..s.find('"').unwrap()].to_string())
        .expect("ハイパーリンク");
    // (The workbook has it already, as English Excel names it: it is taken, not added.)
    assert_eq!(styles.matches("builtinId=\"8\"").count(), 1);
    assert!(named.contains("name=\"Hyperlink\""), "{named}");
    let xf = xf_of(&out, "C2");
    assert!(xf.contains(&format!("xfId=\"{id}\"")), "{xf}");
    // Its font: the standard one underlined, in the theme's colour of links.
    let font_id: usize = xf
        .split("fontId=\"")
        .nth(1)
        .unwrap()
        .split('"')
        .next()
        .unwrap()
        .parse()
        .unwrap();
    let fonts = &styles[styles.find("<fonts").unwrap()..styles.find("</fonts>").unwrap()];
    let font = fonts.split("<font>").nth(font_id + 1).unwrap();
    assert!(
        font.starts_with("<u/>") && font.contains("<color theme=\"10\"/>"),
        "{font}"
    );
    // Another link takes the style there.
    let again = apply(
        &out,
        &ops(r#"[{"op":"format","range":"D2","set":{"style":"Hyperlink"}},{"op":"link","range":"D2","url":"https://example.org/"}]"#),
    )
    .unwrap();
    assert_eq!(
        part(&again, "xl/styles.xml")
            .matches("builtinId=\"8\"")
            .count(),
        1
    );
    assert!(xf_of(&again, "D2").contains(&format!("xfId=\"{id}\"")));
    // Taken away: the standard style again.
    let off = apply(
        &again,
        &ops(r#"[{"op":"link","range":"C2","off":true},{"op":"format","range":"C2","set":{"style":"Normal"}}]"#),
    )
    .unwrap();
    let xf = xf_of(&off, "C2");
    assert!(xf.contains("xfId=\"0\""), "{xf}");
    assert!(
        apply(
            &out,
            &ops(r#"{"op":"format","range":"C2","set":{"style":"Good"}}"#)
        )
        .unwrap_err()
        .contains("Hyperlink")
    );
}

#[test]
fn the_color_of_a_tab() {
    // `<sheetPr codeName="Sheet1"/>` gets a tabColor, as Excel writes it.
    let out = edit(r##"{"op":"tabColor","color":"#ff0000"}"##);
    assert_eq!(changed(BOOK, &out), ["xl/worksheets/sheet1.xml"]);
    let doc = part(&out, "xl/worksheets/sheet1.xml");
    assert!(doc.contains("<sheetPr codeName=\"Sheet1\"><tabColor rgb=\"FFFF0000\"/></sheetPr>"));
    assert_eq!(view(&out)["sheets"][0]["tabColor"], json!("#ff0000"));
    // Another color takes its place; none takes it away.
    let blue = apply(&out, &ops(r##"{"op":"tabColor","color":"0070C0"}"##)).unwrap();
    let doc = part(&blue, "xl/worksheets/sheet1.xml");
    assert_eq!(doc.matches("<tabColor").count(), 1);
    assert!(doc.contains("<tabColor rgb=\"FF0070C0\"/>"));
    let none = apply(&blue, &ops(r#"{"op":"tabColor","color":null}"#)).unwrap();
    assert!(!part(&none, "xl/worksheets/sheet1.xml").contains("<tabColor"));
    assert_eq!(view(&none)["sheets"][0]["tabColor"], Json::Null);
    assert!(parse_ops(&json!({"op":"tabColor","color":"red"}), Some("Data")).is_err());
}

#[test]
fn sheets_hidden_shown_and_moved() {
    let names = |bytes: &[u8]| -> Vec<String> {
        view(bytes)["sheets"]
            .as_array()
            .unwrap()
            .iter()
            .map(|s| s["name"].as_str().unwrap().to_string())
            .collect()
    };
    // シートを表示しない: the active sheet gives its place to the next visible one.
    let hidden = edit(r#"{"op":"sheetVisible","visible":false}"#);
    assert_eq!(changed(BOOK, &hidden), ["xl/workbook.xml"]);
    assert_eq!(view(&hidden)["sheets"][0]["visible"], json!("hidden"));
    assert!(part(&hidden, "xl/workbook.xml").contains("activeTab=\"1\""));
    // シートの再表示.
    let shown = apply(&hidden, &ops(r#"{"op":"sheetVisible","visible":true}"#)).unwrap();
    assert_eq!(view(&shown)["sheets"][0]["visible"], json!("visible"));
    assert!(!part(&shown, "xl/workbook.xml").contains("state="));
    // Not every sheet.
    let all = r#"[{"op":"sheetVisible","sheet":"Data","visible":false},{"op":"sheetVisible","sheet":"Summary","visible":false},{"op":"sheetVisible","sheet":"Tbl","visible":false}]"#;
    assert!(
        apply(BOOK, &ops(all))
            .unwrap_err()
            .contains("すべてのシート")
    );
    // Moved to the end, and before another; the names of a sheet follow it.
    let end = edit(r#"{"op":"moveSheet"}"#);
    assert_eq!(names(&end), ["Summary", "Tbl", "Data"]);
    let back = apply(&end, &ops(r#"{"op":"moveSheet","before":"Summary"}"#)).unwrap();
    assert_eq!(names(&back), ["Data", "Summary", "Tbl"]);
    assert_eq!(
        part(&back, "xl/workbook.xml"),
        part(BOOK, "xl/workbook.xml")
    );
    // Copied before a sheet.
    let copied = edit(r#"{"op":"sheet","sheet":"Summary","action":"copy","before":"Data"}"#);
    assert_eq!(names(&copied), ["Summary (2)", "Data", "Summary", "Tbl"]);
    let last = edit(r#"{"op":"sheet","sheet":"Summary","action":"copy","after":"Tbl"}"#);
    assert_eq!(names(&last), ["Data", "Summary", "Tbl", "Summary (2)"]);
}

#[test]
fn sheets_protected_and_cells_unlocked() {
    const SHEET1: &str = "xl/worksheets/sheet1.xml";
    assert!(view(BOOK)["sheets"][0]["protection"].is_null());
    // Without a password: what it allows besides selecting is written.
    let on = edit(
        r#"{"op":"protect","allow":["selectLockedCells","selectUnlockedCells","formatCells"]}"#,
    );
    assert_eq!(changed(BOOK, &on), [SHEET1]);
    assert!(part(&on, SHEET1).contains("</sheetData><sheetProtection sheet=\"1\" formatCells=\"0\" objects=\"1\" scenarios=\"1\"/><autoFilter"));
    let p = &view(&on)["sheets"][0]["protection"];
    assert_eq!(p["password"], json!(false));
    assert_eq!(p["allow"]["formatCells"], json!(true));
    assert_eq!(p["allow"]["formatColumns"], json!(false));
    assert_eq!(p["allow"]["selectLockedCells"], json!(true));
    let off = apply(&on, &ops(r#"{"op":"protect","on":false}"#)).unwrap();
    assert_eq!(part(&off, SHEET1), part(BOOK, SHEET1));
    // With one: hashed as Excel hashes it, and asked for again.
    let locked = edit(r#"{"op":"protect","password":"秘密"}"#);
    let doc = part(&locked, SHEET1);
    assert!(
        doc.contains("algorithmName=\"SHA-512\"")
            && doc.contains("spinCount=\"100000\"")
            && !doc.contains("秘密")
    );
    assert_eq!(
        view(&locked)["sheets"][0]["protection"]["password"],
        json!(true)
    );
    let wrong = apply(
        &locked,
        &ops(r#"{"op":"protect","on":false,"password":"himitsu"}"#),
    )
    .unwrap_err();
    assert!(wrong.contains("パスワードが間違っています"));
    let open = apply(
        &locked,
        &ops(r#"{"op":"protect","on":false,"password":"秘密"}"#),
    )
    .unwrap();
    assert!(!part(&open, SHEET1).contains("sheetProtection"));
    // セルのロック off: the cell's format says so (and the page sees it).
    let unlocked = edit(r#"{"op":"format","range":"B2","set":{"locked":false}}"#);
    let mut book = Book::open_bytes(&unlocked).unwrap();
    assert_eq!(
        book.format_of(0, RangeRef::parse("B2").unwrap(), "locked")
            .unwrap(),
        Some(Val::Bool(false))
    );
    let v = view(&unlocked);
    let cell = v["sheets"][0]["cells"]
        .as_array()
        .unwrap()
        .iter()
        .find(|c| c["r"] == json!(1) && c["c"] == json!(1))
        .unwrap();
    let css = v["sheets"][0]["styles"][cell["s"].as_u64().unwrap() as usize]
        .as_str()
        .unwrap();
    assert!(css.contains("--unlocked:1"));
}

#[test]
fn the_standard_width_of_a_sheet() {
    let out = edit(r#"{"op":"standardWidth","width":12}"#);
    assert!(part(&out, "xl/worksheets/sheet1.xml").contains("defaultColWidth=\"12.625\""));
    assert_eq!(view(&out)["sheets"][0]["defaultWidth"], json!(101.0));
}

#[test]
fn sheets_added_copied_renamed_and_removed_from_their_tabs() {
    let names = |bytes: &[u8]| -> Vec<String> {
        view(bytes)["sheets"]
            .as_array()
            .unwrap()
            .iter()
            .map(|s| s["name"].as_str().unwrap().to_string())
            .collect()
    };
    // Added after the sheet, named and colored in one edit.
    let out = edit(
        r##"[{"op":"sheet","action":"add","name":"売上","codeName":"Sales"},{"op":"tabColor","sheet":"売上","color":"#00B050"}]"##,
    );
    assert_eq!(names(&out), ["Data", "売上", "Summary", "Tbl"]);
    assert_eq!(view(&out)["sheets"][1]["tabColor"], json!("#00b050"));
    assert_eq!(view(&out)["sheets"][1]["codeName"], json!("Sales"));
    // Renamed: the formulas that refer to it follow.
    let renamed = apply(
        &out,
        &ops(r#"{"op":"sheet","action":"rename","name":"Input"}"#),
    )
    .unwrap();
    assert_eq!(names(&renamed), ["Input", "売上", "Summary", "Tbl"]);
    assert!(part(&renamed, "xl/workbook.xml").contains("Input!$A$2:$A$5"));
    // Copied, as Excel names a copy.
    let copied = apply(
        &out,
        &ops(r#"{"op":"sheet","sheet":"売上","action":"copy"}"#),
    )
    .unwrap();
    assert_eq!(
        names(&copied),
        ["Data", "売上", "売上 (2)", "Summary", "Tbl"]
    );
    // Removed.
    let removed = apply(
        &out,
        &ops(r#"{"op":"sheet","sheet":"売上","action":"remove"}"#),
    )
    .unwrap();
    assert_eq!(names(&removed), ["Data", "Summary", "Tbl"]);
    // What cannot be done is said in Japanese.
    assert!(refused(r#"{"op":"sheet","action":"add","name":"Summary"}"#).contains("既にあります"));
    assert!(
        refused(r#"{"op":"sheet","action":"rename","name":"a/b"}"#).contains("名前にできません")
    );
}

#[test]
fn conditional_formats_put_on_cells_and_cleared() {
    const SHEET1: &str = "xl/worksheets/sheet1.xml";
    // 指定の値より小さい: the format in the styles' dxfs, the rule first in priority.
    let out = edit(
        r##"{"op":"conditional","range":"D2:D5","type":"cellIs","operator":"lessThan","values":["1000"],"format":{"color":"#9C0006","fill":"#FFC7CE"}}"##,
    );
    assert_eq!(changed(BOOK, &out), ["xl/styles.xml", SHEET1]);
    let styles = part(&out, "xl/styles.xml");
    assert!(
        styles.contains(r#"<dxfs count="2">"#)
            && styles.contains(r#"<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>"#),
        "{styles}"
    );
    let doc = part(&out, SHEET1);
    assert!(
        doc.contains(r#"<conditionalFormatting sqref="D2:D5"><cfRule type="cellIs" dxfId="0" priority="2" operator="greaterThan"><formula>500</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="D2:D5"><cfRule type="cellIs" dxfId="1" priority="1" operator="lessThan"><formula>1000</formula></cfRule></conditionalFormatting>"#),
        "{doc}"
    );
    let data = view(&out);
    let sheet = &data["sheets"][0];
    let css = |r: u64, c: u64| -> String {
        let cell = sheet["cells"]
            .as_array()
            .unwrap()
            .iter()
            .find(|x| x["r"] == json!(r) && x["c"] == json!(c))
            .unwrap();
        sheet["styles"][cell["s"].as_u64().unwrap_or(0) as usize]
            .as_str()
            .unwrap()
            .to_string()
    };
    // 800 is under 1000 (and over 500); 1200 is not.
    let s = css(2, 3);
    assert!(
        s.contains("background:#FFC7CE") && s.contains("font-weight:bold"),
        "{s}"
    );
    assert!(!css(4, 3).contains("#FFC7CE"));

    // Texts, top and bottom, dates: the formulas Excel writes, of the first cell.
    let text = edit(
        r##"{"op":"conditional","range":"A2:A7 C2","type":"containsText","text":"a\"b","format":{"bold":true,"border":"FF0000"}}"##,
    );
    let doc = part(&text, SHEET1);
    assert!(
        doc.contains(r#"<conditionalFormatting sqref="A2:A7 C2"><cfRule type="containsText" dxfId="1" priority="1" operator="containsText" text="a&quot;b"><formula>NOT(ISERROR(SEARCH(&quot;a&quot;&quot;b&quot;,A2)))</formula></cfRule></conditionalFormatting>"#),
        "{doc}"
    );
    let top = edit(
        r##"{"op":"conditional","range":"B2:B7","type":"top10","rank":3,"bottom":true,"format":{"fill":"#C6EFCE"}}"##,
    );
    assert!(
        part(&top, SHEET1)
            .contains(r#"<cfRule type="top10" dxfId="1" priority="1" bottom="1" rank="3"/>"#)
    );
    // 標準偏差 1 下: below the average by a standard deviation, as Excel writes it.
    let sd = edit(
        r##"{"op":"conditional","range":"B2:B7","type":"aboveAverage","below":true,"stdDev":1,"format":{"fill":"#C6EFCE"}}"##,
    );
    assert!(part(&sd, SHEET1).contains(
        r#"<cfRule type="aboveAverage" dxfId="1" priority="1" aboveAverage="0" stdDev="1"/>"#
    ));
    let rule = &view(&sd)["sheets"][0]["condRules"][0];
    assert_eq!(
        (&rule["below"], &rule["stdDev"], &rule["equal"]),
        (&json!(true), &json!(1), &json!(false))
    );
    let week = edit(
        r##"{"op":"conditional","range":"B2","type":"timePeriod","period":"today","format":{"fill":"#C6EFCE"}}"##,
    );
    assert!(
        part(&week, SHEET1)
            .contains(r#"timePeriod="today"><formula>FLOOR(B2,1)=TODAY()</formula>"#)
    );
    // A text value in quotes, a formula as it is.
    let between = edit(
        r##"{"op":"conditional","range":"A2:A7","type":"cellIs","operator":"between","values":["a","=$B$1"],"format":{"italic":true}}"##,
    );
    assert!(part(&between, SHEET1).contains(
        r#"operator="between"><formula>&quot;a&quot;</formula><formula>$B$1</formula></cfRule>"#
    ));

    // Data bars, color scales and icon sets: no format of the styles.
    let bars =
        edit(r##"{"op":"conditional","range":"B2:B5","type":"dataBar","colors":["#638EC6"]}"##);
    assert_eq!(changed(BOOK, &bars), [SHEET1]);
    // (As Excel 2010 writes it: its look in the extLst, joined by its id; 自動 for the least and
    // the greatest, a border of its color.)
    let doc = part(&bars, SHEET1);
    assert!(doc.contains(r#"<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>{6F1C3A2E-8B4D-4C7A-9E5F-000000000001}</x14:id></ext></extLst></cfRule>"#), "{doc}");
    assert!(doc.ends_with(r#"<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="{6F1C3A2E-8B4D-4C7A-9E5F-000000000001}"><x14:dataBar minLength="0" maxLength="100" border="1" negativeBarBorderColorSameAsPositive="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:borderColor rgb="FF638EC6"/><x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>B2:B5</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>"#), "{doc}");
    let data = view(&bars);
    assert!(
        data["sheets"][0]["cells"]
            .as_array()
            .unwrap()
            .iter()
            .any(|x| x["r"] == json!(1) && x["c"] == json!(1) && x.get("bar").is_some())
    );
    let scale = edit(
        r##"{"op":"conditional","range":"B2:B5","type":"colorScale","colors":["F8696B","FFEB84","63BE7B"]}"##,
    );
    assert!(part(&scale, SHEET1).contains(r#"<colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale>"#));
    let icons =
        edit(r##"{"op":"conditional","range":"B2:B5","type":"iconSet","icons":"4Rating"}"##);
    assert!(part(&icons, SHEET1).contains(r#"<iconSet iconSet="4Rating"><cfvo type="percent" val="0"/><cfvo type="percent" val="25"/><cfvo type="percent" val="50"/><cfvo type="percent" val="75"/></iconSet>"#));

    // Cleared from some cells: the rules keep the others; from all of theirs: they go.
    let some = apply(
        &out,
        &ops(r#"{"op":"conditional","range":"D3","clear":true}"#),
    )
    .unwrap();
    let doc = part(&some, SHEET1);
    assert_eq!(
        doc.matches(r#"<conditionalFormatting sqref="D2 D4:D5">"#)
            .count(),
        2,
        "{doc}"
    );
    let gone = apply(
        &out,
        &ops(r#"{"op":"conditional","range":"C1:E9","clear":true}"#),
    )
    .unwrap();
    assert!(!part(&gone, SHEET1).contains("conditionalFormatting"));
    let sheet = apply(&out, &ops(r#"{"op":"conditional","clear":true}"#)).unwrap();
    assert!(!part(&sheet, SHEET1).contains("conditionalFormatting"));
    let none = apply(
        BOOK,
        &ops(r#"{"op":"conditional","range":"A1","clear":true}"#),
    )
    .unwrap();
    assert!(changed(BOOK, &none).is_empty());

    let wrong =
        |json: &str| parse_ops(&serde_json::from_str(json).unwrap(), Some("Data")).unwrap_err();
    assert!(
        wrong(r#"{"op":"conditional","range":"A1","type":"nope","format":{"bold":true}}"#)
            .contains("nope")
    );
    assert!(wrong(r#"{"op":"conditional","range":"A1","type":"cellIs","operator":"between","values":["1"],"format":{"bold":true}}"#).contains("2"));
    assert!(
        wrong(
            r#"{"op":"conditional","range":"A1","type":"cellIs","operator":"equal","values":["1"]}"#
        )
        .contains("format")
    );
    assert!(
        wrong(r#"{"op":"conditional","range":"A1","type":"colorScale","colors":["red","blue"]}"#)
            .contains("RRGGBB")
    );
    assert!(
        wrong(r#"{"op":"conditional","range":"A1","type":"aboveAverage","stdDev":4,"format":{"bold":true}}"#)
            .contains("stdDev")
    );
}

#[test]
fn conditional_formats_of_excel_2010_cleared() {
    // A data bar written in both places, joined by its id: the cells cleared from both.
    let x14 = r#"<conditionalFormatting sqref="B2:B5"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>{00000000-0000-0000-0000-000000000001}</x14:id></ext></extLst></cfRule></conditionalFormatting>"#;
    let ext = r#"<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="{00000000-0000-0000-0000-000000000001}"><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule><xm:sqref>B2:B5</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>"#;
    let mut parts = parts(BOOK);
    let doc = String::from_utf8(parts["xl/worksheets/sheet1.xml"].clone()).unwrap();
    let doc = doc
        .replace(
            "</conditionalFormatting>",
            &format!("</conditionalFormatting>{x14}"),
        )
        .replace("</worksheet>", &format!("{ext}</worksheet>"));
    parts.insert("xl/worksheets/sheet1.xml".into(), doc.into_bytes());
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, data) in &parts {
        zip.start_file(name.as_str(), zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(data).unwrap();
    }
    let book = zip.finish().unwrap().into_inner();
    let some = apply(
        &book,
        &ops(r#"{"op":"conditional","range":"B2","clear":true}"#),
    )
    .unwrap();
    let doc = part(&some, "xl/worksheets/sheet1.xml");
    assert!(
        doc.contains(r#"<conditionalFormatting sqref="B3:B5"><cfRule type="dataBar""#),
        "{doc}"
    );
    assert!(doc.contains("<xm:sqref>B3:B5</xm:sqref>"), "{doc}");
    // A new rule comes first: the others, the Excel 2010 one too, after it.
    let first = apply(&book, &ops(r##"{"op":"conditional","range":"A1","type":"duplicateValues","format":{"fill":"#FFC7CE"}}"##)).unwrap();
    let doc = part(&first, "xl/worksheets/sheet1.xml");
    assert!(
        doc.contains(r#"<cfRule type="duplicateValues" dxfId="1" priority="1"/>"#)
            && doc.contains(r#"<cfRule type="dataBar" priority="2">"#),
        "{doc}"
    );
    let all = apply(
        &book,
        &ops(r#"{"op":"conditional","range":"A:E","clear":true}"#),
    )
    .unwrap();
    let doc = part(&all, "xl/worksheets/sheet1.xml");
    assert!(
        !doc.contains("conditionalFormatting") && !doc.contains("<extLst>"),
        "{doc}"
    );
}

#[test]
fn conditional_rules_listed_and_managed() {
    const SHEET1: &str = "xl/worksheets/sheet1.xml";
    let out = edit(
        r##"{"op":"conditional","range":"D2:D5","type":"cellIs","operator":"lessThan","values":["1000"],"format":{"color":"#9C0006","fill":"#FFC7CE"}}"##,
    );
    // Listed in priority order, with what the dialog edits them with.
    let data = view(&out);
    let rules = &data["sheets"][0]["condRules"];
    assert_eq!(rules[0]["id"], json!("m1"));
    assert_eq!(rules[0]["values"], json!(["1000"]));
    assert_eq!(
        rules[0]["format"],
        json!({ "color": "#9C0006", "fill": "#FFC7CE" })
    );
    assert_eq!(rules[0]["editable"], json!(true));
    assert_eq!(rules[1]["id"], json!("m0"));
    assert_eq!(rules[1]["operator"], json!("greaterThan"));
    assert_eq!(rules[1]["format"]["bold"], json!(true));
    assert_eq!(rules[1]["range"], json!("D2:D5"));
    assert!(data["sheets"][1].get("condRules").is_none());

    // One deleted, the other first, its range and 停止 changed.
    let one = apply(
        &out,
        &ops(r#"{"op":"conditionalRules","rules":[{"id":"m0","range":"D2:D9","stop":true}]}"#),
    )
    .unwrap();
    let doc = part(&one, SHEET1);
    assert_eq!(doc.matches("<conditionalFormatting").count(), 1, "{doc}");
    assert!(
        doc.contains(r#"<conditionalFormatting sqref="D2:D9"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan" stopIfTrue="1"><formula>500</formula></cfRule></conditionalFormatting>"#),
        "{doc}"
    );
    // Edited, kept, new: in that order of priority.
    let edited = apply(
        &out,
        &ops(r##"{"op":"conditionalRules","rules":[{"id":"m1","type":"cellIs","operator":"equal","values":["=$B$1"],"format":{"bold":true}},{"id":"m0"},{"type":"dataBar","colors":["#638EC6"],"range":"B2:B5"}]}"##),
    )
    .unwrap();
    let doc = part(&edited, SHEET1);
    assert!(
        doc.contains(r#"<conditionalFormatting sqref="D2:D5"><cfRule type="cellIs" dxfId="2" priority="1" operator="equal"><formula>$B$1</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="D2:D5"><cfRule type="cellIs" dxfId="0" priority="2" operator="greaterThan"><formula>500</formula></cfRule></conditionalFormatting><conditionalFormatting sqref="B2:B5"><cfRule type="dataBar" priority="3">"#),
        "{doc}"
    );
    let rules = &view(&edited)["sheets"][0]["condRules"];
    assert_eq!(rules[0]["values"], json!(["=$B$1"]));
    assert_eq!(rules[2]["colors"], json!(["#638EC6"]));
    // All deleted.
    let none = apply(&out, &ops(r#"{"op":"conditionalRules","rules":[]}"#)).unwrap();
    assert!(!part(&none, SHEET1).contains("conditionalFormatting"));

    let wrong = |json: &str| apply(&out, &ops(json)).unwrap_err();
    assert!(wrong(r#"{"op":"conditionalRules","rules":[{"id":"m7"}]}"#).contains("m7"));
    assert!(
        wrong(r#"{"op":"conditionalRules","rules":[{"id":"m0"},{"id":"m0"}]}"#).contains("2 回")
    );
    assert!(
        wrong(r##"{"op":"conditionalRules","rules":[{"type":"dataBar","colors":["#638EC6"]}]}"##)
            .contains("range")
    );
}

#[test]
fn conditional_rules_of_excel_2010_managed() {
    // A data bar of the sheet with its half in the extLst, and an icon set of the extLst alone.
    let x14 = r#"<conditionalFormatting sqref="B2:B5"><cfRule type="dataBar" priority="2"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>{00000000-0000-0000-0000-000000000001}</x14:id></ext></extLst></cfRule></conditionalFormatting>"#;
    let ext = r#"<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="{00000000-0000-0000-0000-000000000001}"><x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/></x14:dataBar></x14:cfRule><xm:sqref>B2:B5</xm:sqref></x14:conditionalFormatting><x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="iconSet" priority="3" id="{00000000-0000-0000-0000-000000000002}"><x14:iconSet iconSet="3Stars"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule><xm:sqref>C2:C5</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>"#;
    let mut parts = parts(BOOK);
    let doc = String::from_utf8(parts["xl/worksheets/sheet1.xml"].clone()).unwrap();
    let doc = doc
        .replace(
            "</conditionalFormatting>",
            &format!("</conditionalFormatting>{x14}"),
        )
        .replace("</worksheet>", &format!("{ext}</worksheet>"));
    parts.insert("xl/worksheets/sheet1.xml".into(), doc.into_bytes());
    let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for (name, data) in &parts {
        zip.start_file(name.as_str(), zip::write::SimpleFileOptions::default())
            .unwrap();
        zip.write_all(data).unwrap();
    }
    let book = zip.finish().unwrap().into_inner();
    let rules = &view(&book)["sheets"][0]["condRules"];
    let ids: Vec<&str> = rules
        .as_array()
        .unwrap()
        .iter()
        .map(|r| r["id"].as_str().unwrap())
        .collect();
    assert_eq!(ids, ["m0", "m1", "x1"]);
    // (The data bar's look and 自動 from its half; the stars editable.)
    assert_eq!(
        (
            &rules[1]["solid"],
            &rules[1]["barBorder"],
            &rules[1]["thresholds"][0]["type"]
        ),
        (&json!(false), &json!(false), &json!("auto"))
    );
    assert_eq!(rules[2]["editable"], json!(true));
    assert_eq!(rules[2]["icons"], json!("3Stars"));
    // A new one of Excel 2010: with the others of the extLst.
    let more = apply(
        &book,
        &ops(r##"{"op":"conditional","range":"D2:D5","type":"dataBar","colors":["#63C384"],"solid":true}"##),
    )
    .unwrap();
    let doc = part(&more, "xl/worksheets/sheet1.xml");
    assert_eq!(
        doc.matches("<x14:conditionalFormattings>").count(),
        1,
        "{doc}"
    );
    assert_eq!(
        doc.matches("<x14:conditionalFormatting ").count(),
        3,
        "{doc}"
    );
    // The stars edited: boxes, of the extLst still.
    let boxes = apply(
        &book,
        &ops(r#"{"op":"conditionalRules","rules":[{"id":"m0"},{"id":"m1"},{"id":"x1","type":"iconSet","icons":"5Boxes","reverse":true}]}"#),
    )
    .unwrap();
    let doc = part(&boxes, "xl/worksheets/sheet1.xml");
    assert!(
        doc.contains(r#"<x14:cfRule type="iconSet" priority="3" id="{6F1C3A2E-8B4D-4C7A-9E5F-000000000001}"><x14:iconSet iconSet="5Boxes" reverse="1"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo>"#)
            && !doc.contains("3Stars")
            && doc.contains(r#"<x14:cfRule type="dataBar" id="{00000000-0000-0000-0000-000000000001}">"#),
        "{doc}"
    );
    // The icon set first, the data bar's range changed (its half's too), the cell rule deleted.
    let out = apply(
        &book,
        &ops(r#"{"op":"conditionalRules","rules":[{"id":"x1"},{"id":"m1","range":"B2:B9"}]}"#),
    )
    .unwrap();
    let doc = part(&out, "xl/worksheets/sheet1.xml");
    assert!(!doc.contains(r#"type="cellIs""#), "{doc}");
    assert!(
        doc.contains(
            r#"<conditionalFormatting sqref="B2:B9"><cfRule type="dataBar" priority="2">"#
        ),
        "{doc}"
    );
    assert!(
        doc.contains(r#"<x14:cfRule type="dataBar" id="{00000000-0000-0000-0000-000000000001}">"#)
            && doc.contains("<xm:sqref>B2:B9</xm:sqref>"),
        "{doc}"
    );
    assert!(
        doc.contains(r#"<x14:cfRule type="iconSet" priority="1""#)
            && doc.contains("<xm:sqref>C2:C5</xm:sqref>"),
        "{doc}"
    );
    // The data bar deleted: its half goes with it; then nothing of the extLst is left.
    let out = apply(
        &book,
        &ops(r#"{"op":"conditionalRules","rules":[{"id":"m0"}]}"#),
    )
    .unwrap();
    let doc = part(&out, "xl/worksheets/sheet1.xml");
    assert!(
        !doc.contains("<extLst>") && !doc.contains("dataBar"),
        "{doc}"
    );
}

#[test]
fn conditional_thresholds_of_bars_scales_and_icons() {
    const SHEET1: &str = "xl/worksheets/sheet1.xml";
    // A bar from 0 to the 90th percentile, the bar alone.
    let bar = edit(
        r##"{"op":"conditional","range":"B2:B5","type":"dataBar","colors":["#638EC6"],"barOnly":true,"thresholds":[{"type":"num","value":0},{"type":"percentile","value":"90"}]}"##,
    );
    let doc = part(&bar, SHEET1);
    assert!(
        doc.contains(r#"<dataBar showValue="0"><cfvo type="num" val="0"/><cfvo type="percentile" val="90"/><color rgb="FF638EC6"/></dataBar>"#),
        "{doc}"
    );
    let rule = &view(&bar)["sheets"][0]["condRules"][0];
    assert_eq!(
        rule["thresholds"],
        json!([{ "type": "num", "value": "0", "gte": true }, { "type": "percentile", "value": "90", "gte": true }])
    );
    assert_eq!(rule["barOnly"], json!(true));
    // A scale from a cell's value (a formula) through a number to the greatest.
    let scale = edit(
        r##"{"op":"conditional","range":"B2:B5","type":"colorScale","colors":["F8696B","FFEB84","63BE7B"],"thresholds":[{"type":"formula","value":"=$B$1*2"},{"type":"num","value":"=$C$1"},{"type":"max"}]}"##,
    );
    let doc = part(&scale, SHEET1);
    assert!(doc.contains(r#"<colorScale><cfvo type="formula" val="$B$1*2"/><cfvo type="num" val="$C$1"/><cfvo type="max"/>"#), "{doc}");
    let rule = &view(&scale)["sheets"][0]["condRules"][0];
    assert_eq!(
        rule["thresholds"][0],
        json!({ "type": "formula", "value": "=$B$1*2", "gte": true })
    );
    assert_eq!(rule["thresholds"][1]["value"], json!("=$C$1"));
    // Icons by numbers, the greatest above 300 (not from it).
    let icons = edit(
        r##"{"op":"conditional","range":"B2:B5","type":"iconSet","icons":"3Arrows","thresholds":[{"type":"percent","value":0},{"type":"num","value":100},{"type":"num","value":300,"gte":false}]}"##,
    );
    let doc = part(&icons, SHEET1);
    assert!(doc.contains(r#"<iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="num" val="100"/><cfvo type="num" val="300" gte="0"/></iconSet>"#), "{doc}");
    assert_eq!(
        view(&icons)["sheets"][0]["condRules"][0]["thresholds"][2]["gte"],
        json!(false)
    );
    // Kept as they are when the rule is moved; written again when it is edited.
    let moved = apply(
        &icons,
        &ops(r#"{"op":"conditionalRules","rules":[{"id":"m0"},{"id":"m1","range":"B2:B9"}]}"#),
    )
    .unwrap();
    assert!(part(&moved, SHEET1).contains(r#"<cfvo type="num" val="300" gte="0"/>"#));

    let wrong =
        |json: &str| parse_ops(&serde_json::from_str(json).unwrap(), Some("Data")).unwrap_err();
    assert!(wrong(r##"{"op":"conditional","range":"A1","type":"dataBar","colors":["#638EC6"],"thresholds":[{"type":"max"},{"type":"max"}]}"##).contains("max"));
    assert!(wrong(r##"{"op":"conditional","range":"A1","type":"dataBar","colors":["#638EC6"],"thresholds":[{"type":"percent","value":"150"},{"type":"max"}]}"##).contains("0〜100"));
    assert!(wrong(r##"{"op":"conditional","range":"A1","type":"dataBar","colors":["#638EC6"],"thresholds":[{"type":"num","value":"abc"},{"type":"max"}]}"##).contains("abc"));
    assert!(wrong(r#"{"op":"conditional","range":"A1","type":"iconSet","icons":"4Rating","thresholds":[{"type":"num","value":1}]}"#).contains("4 つ"));
    assert!(wrong(r#"{"op":"conditional","range":"A1","type":"colorScale","colors":["000000","FFFFFF"],"thresholds":[{"type":"min"},{"type":"formula","value":"="}]}"#).contains("数式"));
}

#[test]
fn icon_sets_drawn_for_the_dialogs() {
    // Every icon set a rule can be written with has its icons, as many as the set has.
    let sets = crate::view::icon_sets();
    for (set, n) in conditional::ICON_SETS {
        let icons = sets[*set].as_array().unwrap_or_else(|| panic!("{set}"));
        assert_eq!(icons.len(), *n, "{set}");
        assert!(
            icons
                .iter()
                .all(|s| s.as_str().unwrap().starts_with("<svg")),
            "{set}"
        );
    }
}

#[test]
fn conditional_formats_of_excel_2010_written() {
    const SHEET1: &str = "xl/worksheets/sheet1.xml";
    let rules = |out: &[u8]| view(out)["sheets"][0]["condRules"].clone();
    let bar_of = |out: &[u8]| {
        view(out)["sheets"][0]["cells"]
            .as_array()
            .unwrap()
            .iter()
            .find(|x| x["r"] == json!(4) && x["c"] == json!(1))
            .map(|x| x["bar"].clone())
            .unwrap()
    };
    // 塗りつぶし (単色): of one color, without a border.
    let solid = edit(
        r##"{"op":"conditional","range":"B2:B5","type":"dataBar","colors":["#63C384"],"solid":true}"##,
    );
    let doc = part(&solid, SHEET1);
    assert!(
        doc.contains(r#"<x14:dataBar minLength="0" maxLength="100" gradient="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar>"#),
        "{doc}"
    );
    let bar = bar_of(&solid);
    assert!(
        !bar["bg"].as_str().unwrap().contains("gradient") && bar.get("bd").is_none(),
        "{bar}"
    );
    let r = &rules(&solid)[0];
    assert_eq!(
        (&r["solid"], &r["barBorder"], &r["thresholds"]),
        (
            &json!(true),
            &json!(false),
            &json!([{ "type": "auto", "value": "", "gte": true }, { "type": "auto", "value": "", "gte": true }])
        )
    );
    // A gradient with a border of another color; none.
    let border = edit(
        r##"{"op":"conditional","range":"B2:B5","type":"dataBar","colors":["#638EC6"],"barBorder":"#000000"}"##,
    );
    assert!(part(&border, SHEET1).contains(r#"<x14:borderColor rgb="FF000000"/>"#));
    let bar = bar_of(&border);
    assert!(
        bar["bg"].as_str().unwrap().contains("gradient") && bar.get("bd").is_some(),
        "{bar}"
    );
    assert_eq!(rules(&border)[0]["barBorder"], json!("#000000"));
    let plain = edit(
        r##"{"op":"conditional","range":"B2:B5","type":"dataBar","colors":["#638EC6"],"barBorder":false}"##,
    );
    assert!(!part(&plain, SHEET1).contains("border="));

    // 3 つの星: in the extLst alone, first in priority (the rule there moves down).
    let stars = edit(r##"{"op":"conditional","range":"B2:B5","type":"iconSet","icons":"3Stars"}"##);
    let doc = part(&stars, SHEET1);
    assert!(
        doc.contains(r#"<cfRule type="cellIs" dxfId="0" priority="2""#)
            && doc.contains(r#"<x14:cfRule type="iconSet" priority="1" id="{6F1C3A2E-8B4D-4C7A-9E5F-000000000001}"><x14:iconSet iconSet="3Stars"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule><xm:sqref>B2:B5</xm:sqref>"#),
        "{doc}"
    );
    let listed = rules(&stars);
    assert_eq!(
        (
            &listed[0]["icons"],
            &listed[0]["editable"],
            &listed[0]["id"]
        ),
        (&json!("3Stars"), &json!(true), &json!("x0"))
    );
    let data = view(&stars);
    assert!(
        data["sheets"][0]["cells"]
            .as_array()
            .unwrap()
            .iter()
            .any(|x| x["ic"].as_str().is_some_and(|k| k.starts_with("3Stars:"))),
        "{data}"
    );
    // Triangles and boxes too; an icon set of Excel 2007 stays in the sheet's rules.
    for set in ["3Triangles", "5Boxes"] {
        let out = edit(&format!(
            r#"{{"op":"conditional","range":"B2:B5","type":"iconSet","icons":"{set}","reverse":true,"iconOnly":true}}"#
        ));
        assert!(
            part(&out, SHEET1).contains(&format!(
                r#"<x14:iconSet iconSet="{set}" showValue="0" reverse="1">"#
            )),
            "{set}"
        );
    }
    let arrows =
        edit(r##"{"op":"conditional","range":"B2:B5","type":"iconSet","icons":"3Arrows"}"##);
    assert!(!part(&arrows, SHEET1).contains("extLst"));

    // ルールの管理: a new solid bar and stars, the rule there kept; then all deleted.
    let managed = apply(
        BOOK,
        &ops(r##"{"op":"conditionalRules","rules":[{"type":"iconSet","icons":"3Stars","range":"C2:C5"},{"id":"m0"},{"type":"dataBar","colors":["#FF555A"],"solid":true,"range":"B2:B5"}]}"##),
    )
    .unwrap();
    let doc = part(&managed, SHEET1);
    assert!(
        doc.contains(r#"<x14:cfRule type="iconSet" priority="1" id="{6F1C3A2E-8B4D-4C7A-9E5F-000000000001}">"#)
            && doc.contains(r#"<cfRule type="cellIs" dxfId="0" priority="2""#)
            && doc.contains(r#"<cfRule type="dataBar" priority="3"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF555A"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:id>{6F1C3A2E-8B4D-4C7A-9E5F-000000000002}</x14:id>"#)
            && doc.contains(r#"<x14:cfRule type="dataBar" id="{6F1C3A2E-8B4D-4C7A-9E5F-000000000002}"><x14:dataBar minLength="0" maxLength="100" gradient="0">"#),
        "{doc}"
    );
    let ids: Vec<serde_json::Value> = rules(&managed)
        .as_array()
        .unwrap()
        .iter()
        .map(|r| r["id"].clone())
        .collect();
    assert_eq!(ids, [json!("x0"), json!("m0"), json!("m1")]);
    let none = apply(&managed, &ops(r#"{"op":"conditionalRules","rules":[]}"#)).unwrap();
    let doc = part(&none, SHEET1);
    assert!(
        !doc.contains("extLst") && !doc.contains("conditionalFormatting"),
        "{doc}"
    );

    let wrong =
        |json: &str| parse_ops(&serde_json::from_str(json).unwrap(), Some("Data")).unwrap_err();
    assert!(
        wrong(r##"{"op":"conditional","range":"A1","type":"dataBar","colors":["#638EC6"],"barBorder":"blue"}"##)
            .contains("barBorder")
    );
    assert!(
        wrong(r##"{"op":"conditional","range":"A1","type":"colorScale","colors":["#638EC6","#FFFFFF"],"thresholds":[{"type":"auto"},{"type":"max"}]}"##)
            .contains("auto")
    );
}

#[test]
fn conditional_bars_negative_values_axis_and_direction() {
    const SHEET1: &str = "xl/worksheets/sheet1.xml";
    let cell = |out: &[u8], r: u64| {
        view(out)["sheets"][0]["cells"]
            .as_array()
            .unwrap()
            .iter()
            .find(|x| x["r"] == json!(r) && x["c"] == json!(1))
            .cloned()
            .unwrap()
    };
    // 右から左, the axis in the middle in gray, negative values in the positive's fill and in a
    // blue border.
    let out = edit(
        r##"{"op":"conditional","range":"B2:B5","type":"dataBar","colors":["#638EC6"],"direction":"rightToLeft","negativeFillSame":true,"negativeBorder":"#0000FF","axis":"middle","axisColor":"#808080"}"##,
    );
    let doc = part(&out, SHEET1);
    assert!(
        doc.contains(r#"<x14:dataBar minLength="0" maxLength="100" border="1" direction="rightToLeft" negativeBarColorSameAsPositive="1" negativeBarBorderColorSameAsPositive="0" axisPosition="middle"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:borderColor rgb="FF638EC6"/><x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FF0000FF"/><x14:axisColor rgb="FF808080"/></x14:dataBar>"#),
        "{doc}"
    );
    let r = &view(&out)["sheets"][0]["condRules"][0];
    for (key, want) in [
        ("direction", json!("rightToLeft")),
        ("negativeFillSame", json!(true)),
        ("negativeFill", json!("#FF0000")),
        ("negativeBorderSame", json!(false)),
        ("negativeBorder", json!("#0000FF")),
        ("axis", json!("middle")),
        ("axisColor", json!("#808080")),
    ] {
        assert_eq!(r[key], want, "{key}");
    }
    // 300 of 0〜500 from the middle, to the left: 0.2〜0.5 of the cell, the axis at 0.5.
    let bar = &cell(&out, 4)["bar"];
    let near = |v: &serde_json::Value, x: f64| (v.as_f64().unwrap() - x).abs() < 1e-9;
    assert!(
        near(&bar["x"], 0.2) && near(&bar["w"], 0.3) && near(&bar["axis"], 0.5),
        "{bar}"
    );
    assert_eq!(bar["ac"], json!("#808080"));

    // A negative value: its own color from the axis (automatic: where 0 is); without an axis,
    // the same way as the others, in its color.
    let neg = |extra: &str| {
        let json = format!(
            r##"[{{"op":"set","range":"B3","text":"-100"}},{{"op":"conditional","range":"B2:B5","type":"dataBar","colors":["#638EC6"],"solid":true,"negativeFill":"#00FF00"{extra}}}]"##
        );
        apply(BOOK, &ops(&json)).unwrap()
    };
    let auto = neg("");
    let bar = &cell(&auto, 2)["bar"];
    assert_eq!(bar["bg"], json!("#00FF00"));
    assert!(near(&bar["axis"], 100.0 / 600.0), "{bar}");
    let none = neg(r#","axis":"none""#);
    assert!(part(&none, SHEET1).contains(r#"axisPosition="none""#));
    let bar = &cell(&none, 2)["bar"];
    assert!(
        bar["bg"] == json!("#00FF00") && bar["axis"].is_null() && near(&bar["x"], 0.0),
        "{bar}"
    );

    let wrong =
        |json: &str| parse_ops(&serde_json::from_str(json).unwrap(), Some("Data")).unwrap_err();
    assert!(
        wrong(r##"{"op":"conditional","range":"A1","type":"dataBar","colors":["#638EC6"],"direction":"up"}"##)
            .contains("direction")
    );
    assert!(
        wrong(r##"{"op":"conditional","range":"A1","type":"dataBar","colors":["#638EC6"],"axis":"left"}"##)
            .contains("axis")
    );
}

#[test]
fn conditional_scales_bars_and_icons_together() {
    // As Excel draws them on the same cells: the scale's fill, the bar, the icon.
    let out = apply(
        BOOK,
        &ops(
            r##"[{"op":"conditional","range":"B2:B5","type":"colorScale","colors":["#F8696B","#63BE7B"]},{"op":"conditional","range":"B2:B5","type":"dataBar","colors":["#D6007B"],"solid":true},{"op":"conditional","range":"B2:B5","type":"iconSet","icons":"3TrafficLights2"}]"##,
        ),
    )
    .unwrap();
    let data = view(&out);
    let sheet = &data["sheets"][0];
    for r in 1..=4u64 {
        let cell = sheet["cells"]
            .as_array()
            .unwrap()
            .iter()
            .find(|x| x["r"] == json!(r) && x["c"] == json!(1))
            .unwrap();
        let css = sheet["styles"][cell["s"].as_u64().unwrap_or(0) as usize]
            .as_str()
            .unwrap();
        assert!(
            css.contains("background:#") && cell.get("bar").is_some() && cell.get("ic").is_some(),
            "{r}: {cell} {css}"
        );
    }
}

#[test]
fn borders_take_the_line_and_the_color_of_the_ribbon() {
    let out = edit(
        r##"[{"op":"format","range":"A2:B3","set":{"border":{"preset":"topDoubleBottom","line":"dashed","color":"#FF0000"}}},{"op":"format","range":"C2","set":{"border":"thickBottom"}}]"##,
    );
    let styles = part(&out, "xl/styles.xml");
    for expected in [
        // The top of the range (the 線のスタイル), its bottom (double), in the 線の色.
        r#"<left/><right/><top style="dashed"><color rgb="FFFF0000"/></top><bottom/>"#,
        r#"<left/><right/><top/><bottom style="double"><color rgb="FFFF0000"/></bottom>"#,
        r#"<left/><right/><top/><bottom style="thick"/>"#,
    ] {
        assert!(styles.contains(expected), "{expected}: {styles}");
    }
    // Automatic again: the line without its color.
    let again = apply(
        &out,
        &ops(r#"[{"op":"format","range":"A2:B2","set":{"border":"top"}}]"#),
    )
    .unwrap();
    let styles = part(&again, "xl/styles.xml");
    assert!(
        styles.contains(r#"<left/><right/><top style="thin"/><bottom/>"#),
        "{styles}"
    );
}

#[test]
fn adjust_values_of_a_shape() {
    // A shape put on a sheet: the page gets its adjust handles, with no values of its own.
    let out = edit(r#"{"op":"shape","shape":"roundRect","at":"A20","width":100,"height":50}"#);
    let objects = view(&out)["sheets"][0]["objects"].clone();
    let shown = objects
        .as_array()
        .unwrap()
        .iter()
        .find(|o| o["adjust"]["prst"] == json!("roundRect"))
        .unwrap()
        .clone();
    assert_eq!(shown["adjust"]["adj"], json!({}));
    assert_eq!(shown["adjust"]["ah"][0]["x"], json!(["adj", "0", "50000"]));
    let name = shown["name"].as_str().unwrap().to_string();
    // Its handle dragged: the value in the avLst, as Excel writes it; the others stay.
    let op = |adjust: &str| format!(r#"{{"op":"object","name":"{name}","adjust":{adjust}}}"#);
    let set = apply(&out, &ops(&op(r#"{"adj":30000}"#))).unwrap();
    let doc = part(&set, "xl/drawings/drawing1.xml");
    assert!(
        doc.contains(r#"<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 30000"/></a:avLst></a:prstGeom>"#),
        "{doc}"
    );
    let shown = view(&set)["sheets"][0]["objects"]
        .as_array()
        .unwrap()
        .iter()
        .find(|o| o["name"] == json!(name))
        .unwrap()["adjust"]["adj"]
        .clone();
    assert_eq!(shown, json!({ "adj": 30000.0 }));
    let again = apply(&set, &ops(&op(r#"{"adj":5000,"adj2":1}"#))).unwrap();
    let doc = part(&again, "xl/drawings/drawing1.xml");
    assert!(doc.contains(r#"<a:avLst><a:gd name="adj" fmla="val 5000"/><a:gd name="adj2" fmla="val 1"/></a:avLst>"#), "{doc}");
    // Not a name of an adjust value, not a number, a picture: refused.
    assert!(
        parse_ops(
            &json!({"op":"object","name":"x","adjust":{"a b":1}}),
            Some("Data")
        )
        .is_err()
    );
    assert!(
        parse_ops(
            &json!({"op":"object","name":"x","adjust":{"adj":"1"}}),
            Some("Data")
        )
        .is_err()
    );
    assert!(
        parse_ops(
            &json!({"op":"object","name":"x","adjust":[1]}),
            Some("Data")
        )
        .is_err()
    );
}

#[test]
fn diagonal_borders_are_added_and_none_clears_them() {
    let out = edit(
        r##"[{"op":"format","range":"A2","set":{"border":"diagonalDown"}},{"op":"format","range":"A2","set":{"border":"diagonalUp"}},{"op":"format","range":"B2","set":{"border":"diagonalUp"}}]"##,
    );
    let styles = part(&out, "xl/styles.xml");
    for expected in [
        r#"<border diagonalUp="1" diagonalDown="1"><left/><right/><top/><bottom/><diagonal style="thin"/></border>"#,
        r#"<border diagonalUp="1"><left/><right/><top/><bottom/><diagonal style="thin"/></border>"#,
    ] {
        assert!(styles.contains(expected), "{expected}: {styles}");
    }
    let mut book = Book::open_bytes(&out).unwrap();
    let a2 = RangeRef::parse("A2").unwrap();
    assert_eq!(
        book.format_of(0, a2, "borders.5.linestyle").unwrap(),
        Some(Val::Num(1.0))
    );
    // 枠なし: the diagonals go too.
    let none = apply(
        &out,
        &ops(r#"[{"op":"format","range":"A2:B2","set":{"border":"none"}}]"#),
    )
    .unwrap();
    let mut book = Book::open_bytes(&none).unwrap();
    for index in ["5", "6"] {
        assert_eq!(
            book.format_of(
                0,
                RangeRef::parse("A2:B2").unwrap(),
                &format!("borders.{index}.linestyle")
            )
            .unwrap(),
            Some(Val::Num(-4142.0))
        );
    }
}
