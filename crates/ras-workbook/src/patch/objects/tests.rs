//! The objects of a sheet edited in the file: read back as the grid reads them, and the parts
//! of the file that changed.

use std::collections::BTreeMap;
use std::io::{Cursor, Read as _};

use serde_json::{Value as Json, json};

use crate::book::Book;
use crate::patch::{apply, parse_ops};

/// A workbook made by XlsxWriter: a button with a macro, a note and a chart on `Data`.
const BOOK: &[u8] = include_bytes!("../../../tests/data/edit.xlsm");

fn edit_of(bytes: &[u8], sheet: &str, ops: Json) -> Result<Vec<u8>, String> {
    apply(bytes, &parse_ops(&ops, Some(sheet))?)
}

fn edit(bytes: &[u8], sheet: &str, ops: Json) -> Vec<u8> {
    edit_of(bytes, sheet, ops).unwrap()
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

fn changed(a: &[u8], b: &[u8]) -> Vec<String> {
    let (a, b) = (parts(a), parts(b));
    let mut out: Vec<String> = a
        .keys()
        .chain(b.keys())
        .filter(|n| a.get(*n) != b.get(*n))
        .cloned()
        .collect();
    out.sort();
    out.dedup();
    out
}

/// The objects of a sheet as the grid draws them.
fn objects(bytes: &[u8], sheet: usize) -> Vec<Json> {
    let mut book = Book::open_bytes(bytes).unwrap();
    let data = crate::view::data(&mut book, "t.xlsx", Some(bytes), None, None).unwrap();
    data["sheets"][sheet]["objects"].as_array().unwrap().clone()
}

fn object(bytes: &[u8], sheet: usize, name: &str) -> Json {
    objects(bytes, sheet)
        .into_iter()
        .find(|o| o["name"] == name)
        .unwrap_or_else(|| panic!("no object {name}"))
}

#[test]
fn edits_of_objects_are_read_from_json() {
    use crate::patch::{Op, Order, Place, What};
    let ops = parse_ops(
        &json!([
            { "op": "addShape", "shape": "rect", "range": "B2:C3", "set": { "text": "a" } },
            { "op": "placeObject", "object": "x", "from": [1, 2.5, 3, 4], "to": [5, 0, 6, 0] },
            { "op": "addControl", "control": "button", "at": "$B$2", "width": 80, "height": 24 },
            { "op": "orderObject", "object": "x", "to": "front" },
        ]),
        Some("S"),
    )
    .unwrap();
    assert!(
        matches!(&ops[0], Op::Object { sheet, what: What::AddShape { shape, place: Place::Range(_), set } } if sheet == "S" && shape == "rect" && set.len() == 1)
    );
    assert!(
        matches!(&ops[1], Op::Object { what: What::Place { place: Place::Markers((1, x, 3, _), (5, ..)), .. }, .. } if *x == 2.5)
    );
    assert!(
        matches!(&ops[2], Op::Object { what: What::AddControl { place: Place::At(_, w, _), .. }, .. } if *w == 80.0)
    );
    assert!(matches!(
        &ops[3],
        Op::Object {
            what: What::Order {
                to: Order::Front,
                ..
            },
            ..
        }
    ));
    let bad = |v: Json| parse_ops(&v, Some("S")).unwrap_err();
    assert!(bad(json!({ "op": "addShape", "shape": "rect" })).contains("位置"));
    assert!(bad(json!({ "op": "setObject", "object": "x" })).contains("set"));
    assert!(bad(json!({ "op": "orderObject", "object": "x", "to": "up" })).contains("up"));
}

#[test]
fn shapes_are_added_to_a_sheet_without_a_drawing() {
    let out = edit(
        BOOK,
        "Summary",
        json!([
            { "op": "addShape", "shape": "rect", "range": "C2:D4", "set": { "text": "一行目\n二行目", "fill": "#FFC000", "macro": "Main.run" } },
            { "op": "addShape", "shape": "arrow", "at": "C6", "width": 100, "height": 40, "set": { "line": "#0070C0", "lineWidth": 2 } },
            { "op": "addShape", "shape": "textBox", "range": "F2:G3", "set": { "text": "注", "bold": true, "fontSize": 14, "color": "#C00000", "align": "right", "valign": "middle" } },
        ]),
    );
    // The drawing is new: the sheet names it, its type is said; nothing else changed.
    assert_eq!(
        changed(BOOK, &out),
        [
            "[Content_Types].xml",
            "xl/drawings/drawing2.xml",
            "xl/worksheets/_rels/sheet2.xml.rels",
            "xl/worksheets/sheet2.xml",
        ]
    );
    assert!(part(&out, "[Content_Types].xml").contains(r#"<Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>"#));
    assert!(
        part(&out, "xl/worksheets/sheet2.xml").ends_with(r#"<drawing r:id="rId1"/></worksheet>"#)
    );
    let list = objects(&out, 1);
    assert_eq!(list.len(), 3);
    let rect = &list[0];
    assert_eq!(rect["name"], "正方形/長方形 1");
    assert_eq!(
        (&rect["from"], &rect["to"]),
        (&json!([2, 0.0, 1, 0.0]), &json!([4, 0.0, 4, 0.0]))
    );
    assert_eq!(rect["macro"], "Main.run");
    assert_eq!(
        rect["props"],
        json!({ "type": "shape", "sub": "rect", "placement": "moveAndSize", "text": "一行目\n二行目", "fill": "#FFC000" })
    );
    assert_eq!(list[1]["props"]["sub"], "straightConnector1");
    assert_eq!(list[1]["props"]["line"], "#0070C0");
    assert_eq!(list[1]["props"]["lineWidth"], 2.0);
    assert_eq!(list[2]["props"]["sub"], "textBox");
    let drawing = part(&out, "xl/drawings/drawing2.xml");
    assert!(drawing.contains(r#"<a:pPr algn="r"/><a:r><a:rPr kumimoji="1" lang="ja-JP" altLang="en-US" sz="1400" b="1"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:rPr><a:t>注</a:t></a:r>"#), "{drawing}");
    assert!(drawing.contains(r#"anchor="ctr""#));
    let refused = |ops: Json| edit_of(BOOK, "Summary", ops).unwrap_err();
    assert!(refused(json!({ "op": "addShape", "shape": "nope", "range": "A1" })).contains("nope"));
    assert!(
        refused(json!({ "op": "setObject", "object": "nope", "set": { "text": "x" } }))
            .contains("オブジェクト `nope` がありません")
    );
}

#[test]
fn objects_of_the_drawing_are_placed_set_ordered_and_deleted() {
    let out = edit(
        BOOK,
        "Data",
        json!([
            { "op": "addShape", "shape": "ellipse", "range": "A12:B14" },
            { "op": "orderObject", "object": "楕円 2", "to": "back" },
            { "op": "placeObject", "object": "Chart 1", "at": "J20", "width": 300, "height": 200 },
            { "op": "setObject", "object": "Chart 1", "set": { "name": "売上グラフ", "placement": "free" } },
        ]),
    );
    assert_eq!(changed(BOOK, &out), ["xl/drawings/drawing1.xml"]);
    let list = objects(&out, 0);
    let names: Vec<&str> = list.iter().filter_map(|o| o["name"].as_str()).collect();
    assert_eq!(names, ["楕円 2", "売上グラフ", "ボタン 1"]);
    let chart = &list[1];
    assert_eq!(chart["from"], json!([9, 0.0, 19, 0.0]));
    assert_eq!(chart["props"]["placement"], "free");
    // The chart keeps its size: 300 × 200 px, in the box its frame keeps too.
    assert!(
        part(&out, "xl/drawings/drawing1.xml").contains(r#"<a:ext cx="2857500" cy="1905000"/>"#)
    );
    let out = edit(
        &out,
        "Data",
        json!({ "op": "deleteObject", "object": "s3" }),
    );
    let names: Vec<String> = objects(&out, 0)
        .iter()
        .map(|o| o["name"].as_str().unwrap().to_string())
        .collect();
    assert_eq!(names, ["売上グラフ", "ボタン 1"]);
}

#[test]
fn form_controls_are_added_as_excel_2007_saves_them() {
    let out = edit(
        BOOK,
        "Summary",
        json!([
            { "op": "addControl", "control": "button", "range": "C2:D3", "set": { "text": "実行", "macro": "Main.run" } },
            { "op": "addControl", "control": "checkBox", "range": "C5:D5", "set": { "linkedCell": "b1", "checked": true } },
            { "op": "addControl", "control": "comboBox", "range": "C7:D7", "set": { "inputRange": "Data!A2:A5", "linkedCell": "B2", "value": 2, "dropLines": 4 } },
            { "op": "addControl", "control": "scrollBar", "range": "F2:F8", "set": { "min": 0, "max": 50, "value": 10, "step": 5 } },
        ]),
    );
    // (The workbook already says the type of `.vml` parts.)
    assert_eq!(
        changed(BOOK, &out),
        [
            "xl/drawings/vmlDrawing2.vml",
            "xl/worksheets/_rels/sheet2.xml.rels",
            "xl/worksheets/sheet2.xml",
        ]
    );
    let vml = part(&out, "xl/drawings/vmlDrawing2.vml");
    // Shape ids of a block no other sheet's drawing uses.
    assert!(vml.contains(r#"<o:idmap v:ext="edit" data="2"/>"#), "{vml}");
    assert!(
        vml.contains(r##"<v:shape id="ボタン_x0020_1" o:spid="_x0000_s2049" type="#_x0000_t201""##),
        "{vml}"
    );
    assert!(vml.contains("<x:FmlaMacro>[0]!Main.run</x:FmlaMacro>"));
    assert!(vml.contains("<x:Checked>1</x:Checked><x:FmlaLink>$B$1</x:FmlaLink>"));
    assert!(vml.contains("<x:FmlaRange>Data!$A$2:$A$5</x:FmlaRange>"));
    assert!(
        part(&out, "xl/worksheets/sheet2.xml")
            .ends_with(r#"<legacyDrawing r:id="rId1"/></worksheet>"#)
    );
    let list = objects(&out, 1);
    assert_eq!(list.len(), 4);
    assert_eq!(list[0]["name"], "ボタン 1");
    assert_eq!(list[0]["id"], "c2049");
    assert_eq!(list[0]["macro"], "Main.run");
    assert_eq!(list[0]["props"]["text"], "実行");
    assert_eq!(
        list[1]["props"],
        json!({ "type": "form", "sub": "CheckBox", "placement": "move", "linkedCell": "$B$1", "text": "チェック 2", "fontSize": 11.0, "checked": true, "threeD": false })
    );
    // The combo box shows the item of the list that is selected.
    assert!(
        list[2]["html"].as_str().unwrap().contains("みかん"),
        "{}",
        list[2]["html"]
    );
    assert_eq!(list[2]["props"]["dropLines"], 4);
    assert_eq!(list[3]["props"]["max"], 50);
    assert_eq!(list[3]["props"]["step"], 5);
    // Named by its name or its id; renamed, moved and deleted.
    let out = edit(
        &out,
        "Summary",
        json!([
            { "op": "setObject", "object": "c2049", "set": { "name": "Run", "placement": "free" } },
            { "op": "placeObject", "object": "run", "from": [5, 10, 9, 5], "to": [6, 20, 11, 0] },
            { "op": "deleteObject", "object": "チェック 2" },
        ]),
    );
    let list = objects(&out, 1);
    assert_eq!(list.len(), 3);
    let run = object(&out, 1, "Run");
    assert_eq!(
        (&run["from"], &run["to"]),
        (&json!([5, 10.0, 9, 5.0]), &json!([6, 20.0, 11, 0.0]))
    );
    assert_eq!(run["props"]["placement"], "free");
    let refused = edit_of(
        &out,
        "Summary",
        json!({ "op": "setObject", "object": "Run", "set": { "checked": true } }),
    )
    .unwrap_err();
    assert!(refused.contains("checked"), "{refused}");
}

#[test]
fn a_control_of_the_sheets_vml_drawing_is_set() {
    // The button XlsxWriter wrote, next to the box of a note.
    let out = edit(
        BOOK,
        "Data",
        json!([
            { "op": "setObject", "object": "ボタン 1", "set": { "text": "走る", "macro": "other", "fontSize": 9 } },
            { "op": "addControl", "control": "label", "range": "A12:B12" },
        ]),
    );
    assert_eq!(changed(BOOK, &out), ["xl/drawings/vmlDrawing1.vml"]);
    let vml = part(&out, "xl/drawings/vmlDrawing1.vml");
    assert!(
        vml.contains(r##"<font face="Calibri" size="180" color="#000000">走る</font>"##),
        "{vml}"
    );
    assert!(vml.contains("<x:FmlaMacro>[0]!other</x:FmlaMacro>"));
    // The new shape has the next id after the note's.
    assert!(vml.contains(r#"o:spid="_x0000_s1027""#));
    assert!(vml.contains(r#"ObjectType="Note""#));
    assert_eq!(object(&out, 0, "ボタン 1")["macro"], "other");
    assert_eq!(object(&out, 0, "ラベル 3")["props"]["sub"], "Label");
}

#[test]
fn activex_controls_are_added_set_and_deleted() {
    let out = edit(
        BOOK,
        "Summary",
        json!([
            { "op": "addActiveX", "control": "CommandButton", "range": "C2:D3", "set": { "text": "押す", "bold": true, "fontSize": 12, "color": "#FF0000" } },
            { "op": "addActiveX", "control": "CheckBox", "range": "C5:D5", "set": { "value": true, "linkedCell": "B5" } },
            { "op": "addActiveX", "control": "ScrollBar", "range": "C7:E7", "set": { "max": 10, "value": 3, "page": 2 } },
            { "op": "addActiveX", "control": "ComboBox", "range": "C9:D9", "set": { "inputRange": "Data!A2:A5", "text": "もも" } },
        ]),
    );
    let all = parts(&out);
    for n in 1..=4 {
        for name in [
            format!("xl/activeX/activeX{n}.xml"),
            format!("xl/activeX/activeX{n}.bin"),
            format!("xl/activeX/_rels/activeX{n}.xml.rels"),
        ] {
            assert!(all.contains_key(&name), "{name}");
        }
    }
    let types = part(&out, "[Content_Types].xml");
    assert!(types.contains(r#"<Override PartName="/xl/activeX/activeX1.xml" ContentType="application/vnd.ms-office.activeX+xml"/><Override PartName="/xl/activeX/activeX1.bin" ContentType="application/vnd.ms-office.activeX"/>"#), "{types}");
    assert!(part(&out, "xl/activeX/activeX1.xml").contains(r#"ax:classid="{D7053240-CE69-11CD-A777-00DD01143C57}" ax:persistence="persistStreamInit" r:id="rId1""#));
    // The stream starts with the class id.
    assert_eq!(
        all["xl/activeX/activeX1.bin"][..18],
        [
            0x40, 0x32, 0x05, 0xD7, 0x69, 0xCE, 0xCD, 0x11, 0xA7, 0x77, 0x00, 0xDD, 0x01, 0x14,
            0x3C, 0x57, 0, 2
        ]
    );
    let sheet = part(&out, "xl/worksheets/sheet2.xml");
    assert!(sheet.ends_with(r#"<legacyDrawing r:id="rId1"/><controls><control shapeId="2049" r:id="rId2" name="CommandButton1"/><control shapeId="2050" r:id="rId3" name="CheckBox1"/><control shapeId="2051" r:id="rId4" name="ScrollBar1"/><control shapeId="2052" r:id="rId5" name="ComboBox1"/></controls></worksheet>"#), "{sheet}");
    let button = object(&out, 1, "CommandButton1");
    assert_eq!(
        button["props"],
        json!({ "type": "activex", "sub": "CommandButton", "placement": "move", "text": "押す", "color": "#ff0000", "fontName": "Calibri", "fontSize": 12.0, "bold": true, "italic": false })
    );
    let check = object(&out, 1, "CheckBox1");
    assert_eq!(check["props"]["value"], true);
    assert_eq!(check["props"]["linkedCell"], "$B$5");
    assert_eq!(check["props"]["text"], "CheckBox1");
    let scroll = object(&out, 1, "ScrollBar1");
    assert_eq!(
        (
            &scroll["props"]["value"],
            &scroll["props"]["max"],
            &scroll["props"]["page"]
        ),
        (&json!(3), &json!(10), &json!(2))
    );
    assert_eq!(object(&out, 1, "ComboBox1")["props"]["text"], "もも");
    // The compiler reads the controls of the sheet from the same parts (`ras-ovba`).
    // Resized: the control keeps its size itself (HIMETRIC).
    let moved = edit(
        &out,
        "Summary",
        json!({ "op": "placeObject", "object": "CommandButton1", "at": "C2", "width": 96, "height": 48 }),
    );
    let bin = &parts(&moved)["xl/activeX/activeX1.bin"];
    let record = crate::oforms::Record::parse("CommandButton", &bin[16..]).unwrap();
    assert_eq!(record.size(), (2540, 1270));
    assert_eq!(record.string("Caption"), Some("押す"));
    // Deleted: its parts, its relationship and its entry go; the last one takes `<controls>`.
    let mut rest = moved;
    for name in ["CommandButton1", "CheckBox1", "ScrollBar1"] {
        rest = edit(
            &rest,
            "Summary",
            json!({ "op": "deleteObject", "object": name }),
        );
    }
    let names: Vec<String> = parts(&rest)
        .keys()
        .filter(|n| n.contains("activeX"))
        .cloned()
        .collect();
    assert_eq!(
        names,
        [
            "xl/activeX/_rels/activeX4.xml.rels",
            "xl/activeX/activeX4.bin",
            "xl/activeX/activeX4.xml"
        ]
    );
    assert!(part(&rest, "xl/worksheets/sheet2.xml").ends_with(
        r#"<controls><control shapeId="2052" r:id="rId5" name="ComboBox1"/></controls></worksheet>"#
    ));
    let rest = edit(
        &rest,
        "Summary",
        json!({ "op": "deleteObject", "object": "ComboBox1" }),
    );
    assert!(
        part(&rest, "xl/worksheets/sheet2.xml")
            .ends_with(r#"<legacyDrawing r:id="rId1"/></worksheet>"#)
    );
    assert!(!part(&rest, "[Content_Types].xml").contains("activeX"));
    assert!(objects(&rest, 1).is_empty());
    let refused = edit_of(
        BOOK,
        "Summary",
        json!({ "op": "addActiveX", "control": "Calendar", "range": "A1" }),
    )
    .unwrap_err();
    assert!(refused.contains("Calendar"), "{refused}");
}

/// A workbook as Excel 2010 and later save it: the controls also in the sheet's `<controls>`
/// (with their anchors), their properties, and their mirrors in the drawing.
#[test]
fn controls_with_the_parts_of_newer_excel_are_changed_in_all_of_them() {
    let book = crate::view::tests::sample();
    let before = object(&book, 0, "Button 1");
    assert_eq!(before["id"], "c1025");
    let out = edit(
        &book,
        "Sheet1",
        json!([
            { "op": "setObject", "object": "Button 1", "set": { "text": "集計する", "macro": "Report.run", "name": "RunButton" } },
            { "op": "placeObject", "object": "RunButton", "range": "J5:K6" },
            { "op": "setObject", "object": "CommandButton1", "set": { "text": "更新", "linkedCell": "A1" } },
        ]),
    );
    let sheet = part(&out, "xl/worksheets/sheet1.xml");
    // Both entries (for newer Excel and for older) have the name; the anchor moved.
    assert_eq!(sheet.matches(r#"name="RunButton""#).count(), 2, "{sheet}");
    assert!(sheet.contains("<from><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></from><to><xdr:col>11</xdr:col>"), "{sheet}");
    let drawing = part(&out, "xl/drawings/drawing1.xml");
    assert!(drawing.contains(r#"<xdr:sp macro="[0]!Report.run" textlink=""><xdr:nvSpPr><xdr:cNvPr id="1025" name="RunButton" hidden="1">"#), "{drawing}");
    assert!(drawing.contains("<a:t>集計する</a:t>"));
    let vml = part(&out, "xl/drawings/vmlDrawing1.vml");
    assert!(
        vml.contains("<x:Anchor>9, 0, 4, 0, 11, 0, 6, 0</x:Anchor>"),
        "{vml}"
    );
    assert!(vml.contains("<x:FmlaMacro>[0]!Report.run</x:FmlaMacro>"));
    let button = object(&out, 0, "RunButton");
    assert_eq!(button["macro"], "Report.run");
    assert_eq!(button["props"]["text"], "集計する");
    assert_eq!(button["from"], json!([9, 0.0, 4, 0.0]));
    let ax = object(&out, 0, "CommandButton1");
    assert_eq!(ax["props"]["text"], "更新");
    assert_eq!(ax["props"]["linkedCell"], "$A$1");
    // The font the control had is kept.
    assert_eq!(ax["props"]["fontName"], "ＭＳ Ｐゴシック");
    // Deleted: the mirror, the entries, the properties.
    let count = objects(&out, 0).len();
    let out = edit(
        &out,
        "Sheet1",
        json!([{ "op": "deleteObject", "object": "RunButton" }, { "op": "deleteObject", "object": "CommandButton1" }]),
    );
    assert_eq!(objects(&out, 0).len(), count - 2);
    let all = parts(&out);
    assert!(!all.contains_key("xl/ctrlProps/ctrlProp1.xml"));
    assert!(!all.keys().any(|n| n.starts_with("xl/activeX/")));
    assert!(!part(&out, "xl/drawings/drawing1.xml").contains("compatExt"));
    assert!(!part(&out, "xl/worksheets/sheet1.xml").contains("shapeId=\"1025\""));
}
