//! The drawing of a sheet as the grid shows it, from a workbook made as Excel saves one with
//! shapes, a picture, form controls (with and without their mirror in the drawing) and an
//! ActiveX control.

use std::io::Write as _;

use serde_json::Value as Json;

use crate::book::Book;

/// A small PNG (8 × 8, a gradient).
const PNG: [u8; 179] = [
    137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 8, 0, 0, 0, 8, 8, 2, 0, 0, 0, 75, 109, 41, 220, 0, 0, 0, 122, 73, 68, 65, 84, 120, 156, 13, 201, 33, 1, 4, 33, 16, 0, 192, 79,
    114, 53, 48, 132, 216, 16, 24, 26, 140, 70, 211, 0, 67, 136, 13, 129, 161, 211, 51, 118, 126, 159, 83, 156, 112, 186, 51, 156, 229, 164, 115, 157, 223, 39, 139, 12, 217, 229, 144, 75, 166, 188,
    242, 197, 46, 118, 216, 221, 30, 246, 178, 211, 190, 246, 139, 89, 204, 48, 187, 57, 204, 101, 166, 121, 205, 23, 10, 65, 103, 176, 72, 46, 47, 90, 209, 66, 235, 218, 208, 150, 150, 218, 213, 94,
    68, 17, 33, 186, 24, 98, 137, 20, 87, 188, 168, 69, 13, 181, 171, 67, 93, 106, 170, 87, 253, 3, 17, 50, 91, 225, 95, 130, 247, 25, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
];

const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn rels(list: &[(&str, &str, &str)]) -> String {
    let mut s = String::from("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
    for (id, kind, target) in list {
        s.push_str(&format!("<Relationship Id=\"{id}\" Type=\"{kind}\" Target=\"{target}\"/>"));
    }
    s.push_str("</Relationships>");
    s
}

fn marker(tag: &str, col: u32, col_off: i64, row: u32, row_off: i64) -> String {
    format!("<xdr:{tag}><xdr:col>{col}</xdr:col><xdr:colOff>{col_off}</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>{row_off}</xdr:rowOff></xdr:{tag}>")
}

/// The mirror of a form or ActiveX control in the drawing (Excel 2010 and later).
fn mirror(id: u32, name: &str, from: (u32, u32), to: (u32, u32), text: &str) -> String {
    let body = if text.is_empty() {
        String::new()
    } else {
        format!("<xdr:txBody><a:bodyPr vertOverflow=\"clip\" wrap=\"square\" lIns=\"27432\" tIns=\"27432\" rIns=\"27432\" bIns=\"27432\" anchor=\"ctr\" upright=\"1\"/><a:lstStyle/><a:p><a:pPr algn=\"ctr\" rtl=\"0\"><a:defRPr sz=\"1000\"/></a:pPr><a:r><a:rPr lang=\"ja-JP\" altLang=\"en-US\" sz=\"1100\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" baseline=\"0\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill><a:latin typeface=\"游ゴシック\"/><a:ea typeface=\"游ゴシック\"/></a:rPr><a:t>{text}</a:t></a:r></a:p></xdr:txBody>")
    };
    format!(
        "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" Requires=\"a14\"><xdr:twoCellAnchor editAs=\"oneCell\">{}{}<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"{id}\" name=\"{name}\" hidden=\"1\"><a:extLst><a:ext uri=\"{{63B3BB69-23CF-44E3-9099-C40C66FF867C}}\"><a14:compatExt spid=\"_x0000_s{id}\"/></a:ext></a:extLst></xdr:cNvPr><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr bwMode=\"auto\"><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr>{body}</xdr:sp><xdr:clientData fPrintsWithSheet=\"0\"/></xdr:twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent>",
        marker("from", from.0, 0, from.1, 0),
        marker("to", to.0, 0, to.1, 0)
    )
}

/// A control of the sheet's `<controls>` with its anchor.
fn control(id: u32, rid: &str, name: &str, from: (u32, u32), to: (u32, u32)) -> String {
    format!(
        "<mc:AlternateContent><mc:Choice Requires=\"x14\"><control shapeId=\"{id}\" r:id=\"{rid}\" name=\"{name}\"><controlPr defaultSize=\"0\" autoLine=\"0\"><anchor moveWithCells=\"1\">{}{}</anchor></controlPr></control></mc:Choice><mc:Fallback><control shapeId=\"{id}\" r:id=\"{rid}\" name=\"{name}\"/></mc:Fallback></mc:AlternateContent>",
        marker("from", from.0, 0, from.1, 0).replace("xdr:from", "from"),
        marker("to", to.0, 0, to.1, 0).replace("xdr:to", "to")
    )
}

fn vml_shape(id: u32, object: &str, anchor: &str, text: &str, extra: &str) -> String {
    format!(
        "<v:shape id=\"_x0000_s{id}\" type=\"#_x0000_t201\" style='position:absolute;z-index:1' o:button=\"t\" fillcolor=\"buttonFace [67]\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\"><v:textbox style='mso-direction-alt:auto' o:singleclick=\"f\"><div style='text-align:left'><font face=\"ＭＳ Ｐゴシック\" size=\"220\" color=\"#000000\">{text}</font></div></v:textbox><x:ClientData ObjectType=\"{object}\"><x:Anchor>{anchor}</x:Anchor>{extra}</x:ClientData></v:shape>"
    )
}

/// A workbook with what a sheet's drawing holds, as Excel saves it.
pub(crate) fn sample() -> Vec<u8> {
    let mut parts: Vec<(String, Vec<u8>)> = Vec::new();
    let mut add = |name: &str, data: Vec<u8>| parts.push((name.to_string(), data));
    add(
        "[Content_Types].xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Default Extension="png" ContentType="image/png"/><Default Extension="bin" ContentType="application/vnd.ms-office.activeX"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/><Override PartName="/xl/ctrlProps/ctrlProp2.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/><Override PartName="/xl/ctrlProps/ctrlProp3.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/><Override PartName="/xl/activeX/activeX1.xml" ContentType="application/vnd.ms-office.activeX+xml"/></Types>"#
            .to_vec(),
    );
    add("_rels/.rels", rels(&[("rId1", &format!("{R}/officeDocument"), "xl/workbook.xml")]).into_bytes());
    add(
        "xl/workbook.xml",
        format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"{R}\"><bookViews><workbookView/></bookViews><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>").into_bytes(),
    );
    add("xl/_rels/workbook.xml.rels", rels(&[("rId1", &format!("{R}/worksheet"), "worksheets/sheet1.xml"), ("rId2", &format!("{R}/styles"), "styles.xml")]).into_bytes());
    add(
        "xl/styles.xml",
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="游ゴシック"/><family val="2"/><charset val="128"/><scheme val="minor"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#
            .as_bytes()
            .to_vec(),
    );
    let cell = |a: &str, t: &str| format!("<c r=\"{a}\" t=\"inlineStr\"><is><t>{t}</t></is></c>");
    let rows = [
        (5, vec![cell("C5", "r")]),
        (7, vec![cell("D7", "qwer"), cell("E7", "wq")]),
        (9, vec![cell("B9", "wewer"), cell("D9", "qre")]),
        (10, vec![cell("E10", "wq")]),
        (12, vec![cell("B12", "rwetwert")]),
        (14, vec![cell("B14", "長い文字列は右の空いたセルの上に続きます"), "<c r=\"H14\" s=\"1\"/>".to_string()]),
        (17, vec!["<c r=\"K17\" t=\"b\"><v>1</v></c>".to_string()]),
        (18, vec![cell("K18", "東京")]),
        (19, vec![cell("K19", "大阪")]),
        (20, vec![cell("K20", "名古屋")]),
    ];
    let data: String = rows.iter().map(|(r, cells)| format!("<row r=\"{r}\">{}</row>", cells.concat())).collect();
    let controls = [
        control(1025, "rId3", "Button 1", (5, 4), (7, 6)),
        control(1026, "rId4", "CommandButton1", (6, 11), (7, 13)),
        control(1027, "rId6", "Check Box 3", (1, 16), (3, 17)),
        control(1028, "rId7", "List Box 4", (4, 16), (6, 20)),
    ]
    .concat();
    add(
        "xl/worksheets/sheet1.xml",
        format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"{R}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"><dimension ref=\"B5:K20\"/><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetFormatPr defaultRowHeight=\"18.75\"/><sheetData>{data}</sheetData><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><drawing r:id=\"rId1\"/><legacyDrawing r:id=\"rId2\"/><controls>{controls}</controls></worksheet>"
        )
        .into_bytes(),
    );
    add(
        "xl/worksheets/_rels/sheet1.xml.rels",
        rels(&[
            ("rId1", &format!("{R}/drawing"), "../drawings/drawing1.xml"),
            ("rId2", &format!("{R}/vmlDrawing"), "../drawings/vmlDrawing1.vml"),
            ("rId3", &format!("{R}/ctrlProp"), "../ctrlProps/ctrlProp1.xml"),
            ("rId4", &format!("{R}/control"), "../activeX/activeX1.xml"),
            ("rId6", &format!("{R}/ctrlProp"), "../ctrlProps/ctrlProp2.xml"),
            ("rId7", &format!("{R}/ctrlProp"), "../ctrlProps/ctrlProp3.xml"),
        ])
        .into_bytes(),
    );
    // The drawing: an oval, a picture, the mirrors of the button and the ActiveX control, an
    // arrow and WordArt.
    let oval = format!(
        "<xdr:twoCellAnchor>{}{}<xdr:sp macro=\"[0]!Oval_Click\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"2\" name=\"楕円 1\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1000000\" cy=\"700000\"/></a:xfrm><a:prstGeom prst=\"ellipse\"><a:avLst/></a:prstGeom></xdr:spPr><xdr:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"50000\"/></a:schemeClr></a:lnRef><a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef></xdr:style></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>",
        marker("from", 2, 300000, 2, 150000),
        marker("to", 4, 100000, 6, 50000)
    );
    let picture = "<xdr:oneCellAnchor><xdr:from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"952500\" cy=\"952500\"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"3\" name=\"図 2\"/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed=\"rId1\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>";
    let arrow = format!(
        "<xdr:twoCellAnchor>{}{}<xdr:cxnSp macro=\"\"><xdr:nvCxnSpPr><xdr:cNvPr id=\"5\" name=\"直線矢印コネクタ 4\"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr><a:prstGeom prst=\"straightConnector1\"><a:avLst/></a:prstGeom><a:ln w=\"19050\"><a:solidFill><a:srgbClr val=\"C00000\"/></a:solidFill><a:tailEnd type=\"triangle\"/></a:ln></xdr:spPr></xdr:cxnSp><xdr:clientData/></xdr:twoCellAnchor>",
        marker("from", 8, 0, 8, 0),
        marker("to", 10, 0, 10, 0)
    );
    let wordart = format!(
        "<xdr:twoCellAnchor>{}{}<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"6\" name=\"テキスト ボックス 5\"/><xdr:cNvSpPr txBox=\"1\"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr><xdr:txBody><a:bodyPr wrap=\"none\"><a:prstTxWarp prst=\"textArchUp\"><a:avLst/></a:prstTxWarp></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang=\"ja-JP\" sz=\"1800\" b=\"1\"><a:solidFill><a:srgbClr val=\"2F5597\"/></a:solidFill></a:rPr><a:t>ワードアート</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>",
        marker("from", 0, 0, 21, 0),
        marker("to", 3, 0, 26, 0)
    );
    add(
        "xl/drawings/drawing1.xml",
        format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"{R}\">{oval}{picture}{}{}{arrow}{wordart}</xdr:wsDr>",
            mirror(1025, "Button 1", (5, 4), (7, 6), "ボタン 1"),
            mirror(1026, "CommandButton1", (6, 11), (7, 13), "")
        )
        .into_bytes(),
    );
    add("xl/drawings/_rels/drawing1.xml.rels", rels(&[("rId1", &format!("{R}/image"), "../media/image1.png")]).into_bytes());
    add("xl/media/image1.png", PNG.to_vec());
    // VML: the button, the ActiveX control's picture, a check box and a list box without mirrors,
    // a comment.
    let vml = format!(
        "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout><v:shapetype id=\"_x0000_t201\" coordsize=\"21600,21600\" o:spt=\"201\" path=\"m,l,21600r21600,l21600,xe\"/>{}<v:shape id=\"CommandButton1\" o:spid=\"_x0000_s1026\" type=\"#_x0000_t201\" style='position:absolute;z-index:2' filled=\"f\" stroked=\"f\"><v:imagedata o:relid=\"rId1\" o:title=\"\"/><x:ClientData ObjectType=\"Pict\"><x:Anchor> 6, 0, 11, 0, 7, 0, 13, 0</x:Anchor><x:CF>Pict</x:CF></x:ClientData></v:shape>{}{}<v:shape id=\"_x0000_s1030\" type=\"#_x0000_t202\" style='position:absolute;visibility:hidden'><v:textbox><div>メモ</div></v:textbox><x:ClientData ObjectType=\"Note\"><x:Anchor> 1, 15, 0, 2, 3, 15, 4, 2</x:Anchor><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>",
        vml_shape(1025, "Button", " 5, 0, 4, 0, 7, 0, 6, 0", "ボタン 1", "<x:PrintObject>False</x:PrintObject><x:FmlaMacro>[0]!Button1_Click</x:FmlaMacro><x:TextHAlign>Center</x:TextHAlign><x:TextVAlign>Center</x:TextVAlign>"),
        vml_shape(1027, "Checkbox", " 1, 0, 16, 0, 3, 0, 17, 0", "チェック 3", "<x:Checked>0</x:Checked><x:FmlaLink>$K$17</x:FmlaLink><x:TextVAlign>Center</x:TextVAlign>"),
        vml_shape(1028, "List", " 4, 0, 16, 0, 6, 0, 20, 0", "", "<x:FmlaRange>$K$18:$K$20</x:FmlaRange><x:Sel>2</x:Sel>")
    );
    add("xl/drawings/vmlDrawing1.vml", vml.into_bytes());
    let props = |p: &str| format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<formControlPr xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" {p}/>").into_bytes();
    add("xl/ctrlProps/ctrlProp1.xml", props("objectType=\"Button\" lockText=\"1\""));
    add("xl/ctrlProps/ctrlProp2.xml", props("objectType=\"CheckBox\" fmlaLink=\"$K$17\" lockText=\"1\" noThreeD=\"1\""));
    add("xl/ctrlProps/ctrlProp3.xml", props("objectType=\"List\" dx=\"22\" fmlaRange=\"$K$18:$K$20\" noThreeD=\"1\" sel=\"2\""));
    add(
        "xl/activeX/activeX1.xml",
        format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>\n<ax:ocx xmlns:ax=\"http://schemas.microsoft.com/office/2006/activeX\" xmlns:r=\"{R}\" ax:classid=\"{{D7053240-CE69-11CD-A777-00DD01143C57}}\" ax:persistence=\"persistStream\" r:id=\"rId1\"/>").into_bytes(),
    );
    add("xl/activeX/_rels/activeX1.xml.rels", rels(&[("rId1", "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary", "activeX1.bin")]).into_bytes());
    let mut bin = vec![0x40, 0x32, 0x05, 0xD7, 0x69, 0xCE, 0xCD, 0x11, 0xA7, 0x77, 0x00, 0xDD, 0x01, 0x14, 0x3C, 0x57];
    bin.extend(super::controls::command_button_for_tests("実行", "ＭＳ Ｐゴシック", 220));
    add("xl/activeX/activeX1.bin", bin);

    zip_of(parts)
}

fn zip_of(parts: Vec<(String, Vec<u8>)>) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    for (name, data) in parts {
        zip.start_file(name, zip::write::SimpleFileOptions::default()).unwrap();
        zip.write_all(&data).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn view_of(bytes: &[u8]) -> Json {
    let mut book = Book::open_bytes(bytes).unwrap();
    super::data(&mut book, "Drawing.xlsx", Some(bytes), None, None).unwrap()
}

#[test]
fn the_drawing_over_the_grid() {
    let bytes = sample();
    // The workbook, for looking at it in a browser (`ras book view`).
    let _ = std::fs::write(std::env::temp_dir().join("ras-view-drawing.xlsx"), &bytes);
    let data = view_of(&bytes);
    let sheet = &data["sheets"][0];
    let objects = sheet["objects"].as_array().unwrap();
    let list: Vec<(&str, &str)> = objects.iter().map(|o| (o["kind"].as_str().unwrap(), o["name"].as_str().unwrap_or(""))).collect();
    // In the order of the drawing, the controls with their mirrors in place; then those without.
    assert_eq!(
        list,
        [
            ("shape", "楕円 1"),
            ("picture", "図 2"),
            ("control", "Button 1"),
            ("control", "CommandButton1"),
            ("shape", "直線矢印コネクタ 4"),
            ("shape", "テキスト ボックス 5"),
            ("control", "Check Box 3"),
            ("control", "List Box 4"),
        ]
    );
    let html = |i: usize| objects[i]["html"].as_str().unwrap();
    // The oval: the theme's accent color, from C3 (with its offsets in pixels).
    assert!(html(0).contains("fill=\"#4472C4\""), "{}", html(0));
    assert_eq!(objects[0]["from"], serde_json::json!([2, 31.5, 2, 15.75]));
    assert_eq!(objects[0]["macro"], "Oval_Click");
    // The picture, once in the page's pictures.
    assert!(html(1).contains("data-img=\"0\""));
    assert!(data["images"][0].as_str().unwrap().starts_with("data:image/png;base64,iVBORw0KGgo"));
    assert_eq!(objects[1]["ext"], serde_json::json!([100.0, 100.0]));
    // The form button: its text and its macro.
    assert!(html(2).contains("class=\"ctl button\"") && html(2).contains("ボタン 1"), "{}", html(2));
    assert_eq!(objects[2]["macro"], "Button1_Click");
    // The ActiveX button: its caption and font from its persisted data.
    assert!(html(3).contains("ctl button") && html(3).contains(">実行<") && html(3).contains("ＭＳ Ｐゴシック"), "{}", html(3));
    // The arrow with its head; WordArt laid out by the page.
    assert!(html(4).contains("marker-end") && html(4).contains("#C00000"), "{}", html(4));
    assert!(html(5).contains("class=\"warp\""), "{}", html(5));
    // The check box follows its linked cell (TRUE), not what was saved; the list shows its
    // items with the second one chosen.
    assert!(html(6).contains("class=\"box\"><svg"), "{}", html(6));
    assert!(html(7).contains("<span>東京</span><span class=\"sel\">大阪</span><span>名古屋</span>"), "{}", html(7));
    // The comment is not drawn; the grid reaches to the objects (the WordArt to row 27).
    assert!(!objects.iter().any(|o| o["html"].as_str().unwrap().contains("メモ")));
    assert_eq!(sheet["rows"], 27);
    // The grid goes on past the cells.
    assert_eq!((sheet["maxRows"].as_u64(), sheet["maxCols"].as_u64()), (Some(1_048_576), Some(16_384)));
    assert_eq!((sheet["defaultWidth"].as_f64(), sheet["defaultHeight"].as_f64()), (Some(72.0), Some(25.0)));
    assert_eq!(sheet["gridLines"], true);
}

/// Without the file, the cells alone.
#[test]
fn without_the_drawing() {
    let mut book = Book::open_bytes(&sample()).unwrap();
    let data = super::data(&mut book, "x.xlsx", None, None, None).unwrap();
    assert_eq!(data["sheets"][0]["objects"], serde_json::json!([]));
    assert_eq!(data["images"], serde_json::json!([]));
    assert_eq!(data["sheets"][0]["gridLines"], true);
}

/// A sheet whose gridlines are off says so.
#[test]
fn gridlines_off() {
    let parts = super::Parts::read(&sample());
    assert_eq!(parts.grid_lines("Sheet1"), Some(true));
    let mut files = parts_of(&sample());
    let sheet = files.iter_mut().find(|(n, _)| n == "xl/worksheets/sheet1.xml").unwrap();
    sheet.1 = String::from_utf8_lossy(&sheet.1).replace("<sheetView workbookViewId=\"0\"/>", "<sheetView showGridLines=\"0\" workbookViewId=\"0\"/>").into_bytes();
    let bytes = zip_of(files);
    assert_eq!(view_of(&bytes)["sheets"][0]["gridLines"], false);
}

fn parts_of(bytes: &[u8]) -> Vec<(String, Vec<u8>)> {
    use std::io::Read as _;
    let mut zr = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    (0..zr.len())
        .map(|i| {
            let mut f = zr.by_index(i).unwrap();
            let mut data = Vec::new();
            f.read_to_end(&mut data).unwrap();
            (f.name().to_string(), data)
        })
        .collect()
}


/// A workbook the stand-in saved keeps its cells, not its controls: drawn with the drawings of
/// the workbook it was read from, the controls follow the cells as they are now.
#[test]
fn a_saved_workbook_with_the_drawings_it_was_read_from() {
    let bytes = sample();
    let mut book = Book::open_bytes(&bytes).unwrap();
    book.set_values(0, crate::RangeRef::parse("K17").unwrap(), &[vec![crate::book::Val::Bool(false)]]).unwrap();
    let path = std::env::temp_dir().join(format!("ras-view-saved-{}.xlsx", std::process::id()));
    book.save(&path).unwrap();
    let saved = std::fs::read(&path).unwrap();
    let _ = std::fs::remove_file(&path);
    // Its own drawing has lost the ActiveX control.
    let own = view_of(&saved);
    assert!(!own["sheets"][0]["objects"].as_array().unwrap().iter().any(|o| o["name"] == "CommandButton1"));
    let mut book = Book::open_bytes(&saved).unwrap();
    let data = super::data(&mut book, "Drawing.xlsx", Some(&bytes), None, None).unwrap();
    let objects = data["sheets"][0]["objects"].as_array().unwrap();
    let html = |name: &str| objects.iter().find(|o| o["name"] == name).and_then(|o| o["html"].as_str()).unwrap_or_default().to_string();
    assert!(html("CommandButton1").contains(">実行<"));
    // The check box: its linked cell is FALSE now.
    assert!(html("Check Box 3").contains("<span class=\"box\"></span>"), "{}", html("Check Box 3"));
}
