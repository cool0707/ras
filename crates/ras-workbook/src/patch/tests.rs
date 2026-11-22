use std::collections::{BTreeMap, HashSet};
use std::io::Read as _;

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
fn sizes_hidden_lines_and_merged_cells() {
    let out = edit(
        r#"[{"op":"size","range":"B:C","px":120},{"op":"size","range":"3:3","px":40},{"op":"hide","range":"5:5","hidden":true},{"op":"hide","range":"A:A","hidden":true},{"op":"merge","range":"A2:B3"},{"op":"unmerge","range":"A8"}]"#,
    );
    assert_eq!(changed(BOOK, &out), [SHEET1, SHEET2]);
    let sheet = part(&out, SHEET1);
    assert!(sheet.contains(r#"<cols><col min="1" max="1" width="14.7109375" customWidth="1" hidden="1"/><col min="2" max="3" width="15.625" customWidth="1"/><col min="4" max="4" width="12.7109375" style="1" customWidth="1"/></cols>"#), "{sheet}");
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
