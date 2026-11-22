//! The drawing of a sheet as the grid shows it, from a workbook made as Excel saves one with
//! shapes, a picture, form controls (with and without their mirror in the drawing) and an
//! ActiveX control.

use std::io::Write as _;

use serde_json::{Value as Json, json};

use crate::book::Book;

/// A small PNG (8 × 8, a gradient).
const PNG: [u8; 179] = [
    137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 8, 0, 0, 0, 8, 8, 2, 0,
    0, 0, 75, 109, 41, 220, 0, 0, 0, 122, 73, 68, 65, 84, 120, 156, 13, 201, 33, 1, 4, 33, 16, 0,
    192, 79, 114, 53, 48, 132, 216, 16, 24, 26, 140, 70, 211, 0, 67, 136, 13, 129, 161, 211, 51,
    118, 126, 159, 83, 156, 112, 186, 51, 156, 229, 164, 115, 157, 223, 39, 139, 12, 217, 229, 144,
    75, 166, 188, 242, 197, 46, 118, 216, 221, 30, 246, 178, 211, 190, 246, 139, 89, 204, 48, 187,
    57, 204, 101, 166, 121, 205, 23, 10, 65, 103, 176, 72, 46, 47, 90, 209, 66, 235, 218, 208, 150,
    150, 218, 213, 94, 68, 17, 33, 186, 24, 98, 137, 20, 87, 188, 168, 69, 13, 181, 171, 67, 93,
    106, 170, 87, 253, 3, 17, 50, 91, 225, 95, 130, 247, 25, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66,
    96, 130,
];

const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn rels(list: &[(&str, &str, &str)]) -> String {
    let mut s = String::from(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    );
    for (id, kind, target) in list {
        s.push_str(&format!(
            "<Relationship Id=\"{id}\" Type=\"{kind}\" Target=\"{target}\"/>"
        ));
    }
    s.push_str("</Relationships>");
    s
}

fn marker(tag: &str, col: u32, col_off: i64, row: u32, row_off: i64) -> String {
    format!(
        "<xdr:{tag}><xdr:col>{col}</xdr:col><xdr:colOff>{col_off}</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>{row_off}</xdr:rowOff></xdr:{tag}>"
    )
}

/// The mirror of a form or ActiveX control in the drawing (Excel 2010 and later).
fn mirror(id: u32, name: &str, from: (u32, u32), to: (u32, u32), text: &str) -> String {
    let body = if text.is_empty() {
        String::new()
    } else {
        format!(
            "<xdr:txBody><a:bodyPr vertOverflow=\"clip\" wrap=\"square\" lIns=\"27432\" tIns=\"27432\" rIns=\"27432\" bIns=\"27432\" anchor=\"ctr\" upright=\"1\"/><a:lstStyle/><a:p><a:pPr algn=\"ctr\" rtl=\"0\"><a:defRPr sz=\"1000\"/></a:pPr><a:r><a:rPr lang=\"ja-JP\" altLang=\"en-US\" sz=\"1100\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" baseline=\"0\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill><a:latin typeface=\"游ゴシック\"/><a:ea typeface=\"游ゴシック\"/></a:rPr><a:t>{text}</a:t></a:r></a:p></xdr:txBody>"
        )
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
    book(false)
}

/// The same with an embedded object, without the mirrors of the controls in the drawing
/// (umya-spreadsheet takes every mirror for the mirror of an embedded object).
fn with_object() -> Vec<u8> {
    book(true)
}

fn book(object: bool) -> Vec<u8> {
    let mut parts: Vec<(String, Vec<u8>)> = Vec::new();
    let mut add = |name: &str, data: Vec<u8>| parts.push((name.to_string(), data));
    add(
        "[Content_Types].xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Default Extension="png" ContentType="image/png"/><Default Extension="bin" ContentType="application/vnd.ms-office.activeX"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/ctrlProps/ctrlProp1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/><Override PartName="/xl/ctrlProps/ctrlProp2.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/><Override PartName="/xl/ctrlProps/ctrlProp3.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/><Override PartName="/xl/activeX/activeX1.xml" ContentType="application/vnd.ms-office.activeX+xml"/></Types>"#
            .to_vec(),
    );
    add(
        "_rels/.rels",
        rels(&[("rId1", &format!("{R}/officeDocument"), "xl/workbook.xml")]).into_bytes(),
    );
    add(
        "xl/workbook.xml",
        format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"{R}\"><bookViews><workbookView/></bookViews><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>").into_bytes(),
    );
    add(
        "xl/_rels/workbook.xml.rels",
        rels(&[
            ("rId1", &format!("{R}/worksheet"), "worksheets/sheet1.xml"),
            ("rId2", &format!("{R}/styles"), "styles.xml"),
        ])
        .into_bytes(),
    );
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
        (
            14,
            vec![
                cell("B14", "長い文字列は右の空いたセルの上に続きます"),
                "<c r=\"H14\" s=\"1\"/>".to_string(),
            ],
        ),
        (17, vec!["<c r=\"K17\" t=\"b\"><v>1</v></c>".to_string()]),
        (18, vec![cell("K18", "東京")]),
        (19, vec![cell("K19", "大阪")]),
        (20, vec![cell("K20", "名古屋")]),
        (30, vec![cell("L30", "月"), cell("M30", "売上")]),
        (
            31,
            vec![
                cell("L31", "4月"),
                "<c r=\"M31\"><v>120</v></c>".to_string(),
            ],
        ),
        (
            32,
            vec![
                cell("L32", "5月"),
                "<c r=\"M32\"><v>150</v></c>".to_string(),
            ],
        ),
        (
            33,
            vec![cell("L33", "6月"), "<c r=\"M33\"><v>90</v></c>".to_string()],
        ),
    ];
    let data: String = rows
        .iter()
        .map(|(r, cells)| format!("<row r=\"{r}\">{}</row>", cells.concat()))
        .collect();
    let controls = [
        control(1025, "rId3", "Button 1", (5, 4), (7, 6)),
        control(1026, "rId4", "CommandButton1", (6, 11), (7, 13)),
        control(1027, "rId6", "Check Box 3", (1, 16), (3, 17)),
        control(1028, "rId7", "List Box 4", (4, 16), (6, 20)),
    ]
    .concat();
    let objects = if object {
        "<oleObjects><mc:AlternateContent><mc:Choice Requires=\"x14\"><oleObject progId=\"Word.Document.12\" shapeId=\"1029\" r:id=\"rId8\"><objectPr defaultSize=\"0\" r:id=\"rId9\"><anchor moveWithCells=\"1\"><from><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></from><to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>19</xdr:row><xdr:rowOff>0</xdr:rowOff></to></anchor></objectPr></oleObject></mc:Choice><mc:Fallback><oleObject progId=\"Word.Document.12\" shapeId=\"1029\" r:id=\"rId8\"/></mc:Fallback></mc:AlternateContent></oleObjects>"
    } else {
        ""
    };
    add(
        "xl/worksheets/sheet1.xml",
        format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"{R}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"><dimension ref=\"B5:K20\"/><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetFormatPr defaultRowHeight=\"18.75\"/><sheetData>{data}</sheetData><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/><drawing r:id=\"rId1\"/><legacyDrawing r:id=\"rId2\"/>{objects}<controls>{controls}</controls></worksheet>"
        )
        .into_bytes(),
    );
    add(
        "xl/worksheets/_rels/sheet1.xml.rels",
        rels(&[
            ("rId1", &format!("{R}/drawing"), "../drawings/drawing1.xml"),
            (
                "rId2",
                &format!("{R}/vmlDrawing"),
                "../drawings/vmlDrawing1.vml",
            ),
            (
                "rId3",
                &format!("{R}/ctrlProp"),
                "../ctrlProps/ctrlProp1.xml",
            ),
            ("rId4", &format!("{R}/control"), "../activeX/activeX1.xml"),
            (
                "rId6",
                &format!("{R}/ctrlProp"),
                "../ctrlProps/ctrlProp2.xml",
            ),
            (
                "rId7",
                &format!("{R}/ctrlProp"),
                "../ctrlProps/ctrlProp3.xml",
            ),
            (
                "rId8",
                &format!("{R}/package"),
                "../embeddings/Document.docx",
            ),
            ("rId9", &format!("{R}/image"), "../media/image2.emf"),
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
    let chart = format!(
        "<xdr:twoCellAnchor>{}{}<xdr:graphicFrame macro=\"[0]!Chart_Click\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"7\" name=\"グラフ 6\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" r:id=\"rId2\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>",
        marker("from", 13, 0, 28, 0),
        marker("to", 19, 0, 40, 0)
    );
    let smartart = format!(
        "<xdr:twoCellAnchor>{}{}<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"8\" name=\"図表 7\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"3048000\" cy=\"952500\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"><dgm:relIds xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" r:dm=\"rId3\" r:lo=\"rId4\" r:qs=\"rId5\" r:cs=\"rId6\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>",
        marker("from", 0, 0, 28, 0),
        marker("to", 4, 0, 32, 0)
    );
    let emf = "<xdr:oneCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>28</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"1905000\" cy=\"952500\"/><xdr:pic macro=\"[0]!Picture_Click\"><xdr:nvPicPr><xdr:cNvPr id=\"9\" name=\"図 8\"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed=\"rId7\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>";
    let wordart365 = format!(
        "<xdr:twoCellAnchor>{}{}<xdr:sp macro=\"[0]!WordArt_Click\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"10\" name=\"テキスト ボックス 9\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr><xdr:txBody><a:bodyPr wrap=\"none\" lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang=\"ja-JP\" sz=\"3600\" b=\"1\"><a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"C55A11\"/></a:solidFill></a:ln><a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FFC000\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"ED7D31\"/></a:gs></a:gsLst><a:lin ang=\"5400000\"/></a:gradFill><a:effectLst><a:outerShdw blurRad=\"38100\" dist=\"19050\" dir=\"2700000\" algn=\"tl\" rotWithShape=\"0\"><a:srgbClr val=\"000000\"><a:alpha val=\"40000\"/></a:srgbClr></a:outerShdw></a:effectLst></a:rPr><a:t>ワードアート</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>",
        marker("from", 4, 0, 21, 0),
        marker("to", 9, 0, 25, 0)
    );
    let wordart = format!(
        "<xdr:twoCellAnchor>{}{}<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"6\" name=\"テキスト ボックス 5\"/><xdr:cNvSpPr txBox=\"1\"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></xdr:spPr><xdr:txBody><a:bodyPr wrap=\"none\"><a:prstTxWarp prst=\"textArchUp\"><a:avLst/></a:prstTxWarp></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr lang=\"ja-JP\" sz=\"1800\" b=\"1\"><a:solidFill><a:srgbClr val=\"2F5597\"/></a:solidFill></a:rPr><a:t>ワードアート</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>",
        marker("from", 0, 0, 21, 0),
        marker("to", 3, 0, 26, 0)
    );
    add(
        "xl/drawings/drawing1.xml",
        format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"{R}\">{oval}{picture}{}{}{arrow}{wordart}{chart}{smartart}{emf}{wordart365}</xdr:wsDr>",
            if object { String::new() } else { mirror(1025, "Button 1", (5, 4), (7, 6), "ボタン 1") },
            if object { String::new() } else { mirror(1026, "CommandButton1", (6, 11), (7, 13), "") }
        )
        .into_bytes(),
    );
    add(
        "xl/drawings/_rels/drawing1.xml.rels",
        rels(&[
            ("rId1", &format!("{R}/image"), "../media/image1.png"),
            ("rId2", &format!("{R}/chart"), "../charts/chart1.xml"),
            ("rId3", &format!("{R}/diagramData"), "../diagrams/data1.xml"),
            (
                "rId4",
                &format!("{R}/diagramLayout"),
                "../diagrams/layout1.xml",
            ),
            (
                "rId5",
                &format!("{R}/diagramQuickStyle"),
                "../diagrams/quickStyle1.xml",
            ),
            (
                "rId6",
                &format!("{R}/diagramColors"),
                "../diagrams/colors1.xml",
            ),
            ("rId7", &format!("{R}/image"), "../media/image3.emf"),
            (
                "rId8",
                "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
                "../diagrams/drawing1.xml",
            ),
        ])
        .into_bytes(),
    );
    add("xl/media/image1.png", PNG.to_vec());
    add("xl/media/image2.emf", super::emf::tests::sample());
    add("xl/media/image3.emf", super::emf::tests::sample());
    add("xl/embeddings/Document.docx", b"PK".to_vec());
    // A column chart of L30:M33, with its cache (the cells win).
    add(
        "xl/charts/chart1.xml",
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>月別売上</a:t></a:r></a:p></c:rich></c:tx></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Sheet1!$M$30</c:f></c:strRef></c:tx><c:cat><c:strRef><c:f>Sheet1!$L$31:$L$33</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$M$31:$M$33</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:gapWidth val="150"/><c:axId val="10"/><c:axId val="20"/></c:barChart><c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="20"/></c:catAx><c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:crossAx val="10"/></c:valAx></c:plotArea><c:legend><c:legendPos val="b"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>"#
            .as_bytes()
            .to_vec(),
    );
    // SmartArt: its data refers to the drawing Excel keeps of it.
    add(
        "xl/diagrams/data1.xml",
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst/><dgm:cxnLst/><dgm:bg/><dgm:whole/><dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rId8" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst></dgm:dataModel>"#
            .to_vec(),
    );
    let step = |i: i64, text: &str| {
        format!(
            "<dsp:sp modelId=\"{{0000000{i}-0000-0000-0000-000000000000}}\"><dsp:nvSpPr><dsp:cNvPr id=\"0\" name=\"\"/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x=\"{}\" y=\"0\"/><a:ext cx=\"914400\" cy=\"952500\"/></a:xfrm><a:prstGeom prst=\"chevron\"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val=\"4472C4\"/></a:solidFill></dsp:spPr><dsp:txBody><a:bodyPr anchor=\"ctr\"/><a:p><a:pPr algn=\"ctr\"/><a:r><a:rPr sz=\"1400\"><a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill></a:rPr><a:t>{text}</a:t></a:r></a:p></dsp:txBody><dsp:txXfrm><a:off x=\"{}\" y=\"0\"/><a:ext cx=\"457200\" cy=\"952500\"/></dsp:txXfrm></dsp:sp>",
            i * 1066800,
            i * 1066800 + 228600
        )
    };
    add(
        "xl/diagrams/drawing1.xml",
        format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<dsp:drawing xmlns:dsp=\"http://schemas.microsoft.com/office/drawing/2008/diagram\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id=\"0\" name=\"\"/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>{}{}{}</dsp:spTree></dsp:drawing>",
            step(0, "企画"),
            step(1, "開発"),
            step(2, "公開")
        )
        .into_bytes(),
    );
    // VML: the button, the ActiveX control's picture, a check box and a list box without mirrors,
    // a comment.
    let vml = format!(
        "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"1\"/></o:shapelayout><v:shapetype id=\"_x0000_t201\" coordsize=\"21600,21600\" o:spt=\"201\" path=\"m,l,21600r21600,l21600,xe\"/>{}<v:shape id=\"CommandButton1\" o:spid=\"_x0000_s1026\" type=\"#_x0000_t201\" style='position:absolute;z-index:2' filled=\"f\" stroked=\"f\"><v:imagedata o:relid=\"rId1\" o:title=\"\"/><x:ClientData ObjectType=\"Pict\"><x:Anchor> 6, 0, 11, 0, 7, 0, 13, 0</x:Anchor><x:CF>Pict</x:CF></x:ClientData></v:shape>{}{}<v:shape id=\"_x0000_s1030\" type=\"#_x0000_t202\" style='position:absolute;visibility:hidden'><v:textbox><div>メモ</div></v:textbox><x:ClientData ObjectType=\"Note\"><x:Anchor> 1, 15, 0, 2, 3, 15, 4, 2</x:Anchor><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>",
        vml_shape(
            1025,
            "Button",
            " 5, 0, 4, 0, 7, 0, 6, 0",
            "ボタン 1",
            "<x:PrintObject>False</x:PrintObject><x:FmlaMacro>[0]!Button1_Click</x:FmlaMacro><x:TextHAlign>Center</x:TextHAlign><x:TextVAlign>Center</x:TextVAlign>"
        ),
        vml_shape(
            1027,
            "Checkbox",
            " 1, 0, 16, 0, 3, 0, 17, 0",
            "チェック 3",
            "<x:Checked>0</x:Checked><x:FmlaLink>$K$17</x:FmlaLink><x:TextVAlign>Center</x:TextVAlign>"
        ),
        vml_shape(
            1028,
            "List",
            " 4, 0, 16, 0, 6, 0, 20, 0",
            "",
            "<x:FmlaRange>$K$18:$K$20</x:FmlaRange><x:Sel>2</x:Sel>"
        )
    );
    // The embedded object's picture.
    let vml = if !object {
        vml
    } else {
        vml.replace(
        "</xml>",
        "<v:shape id=\"_x0000_s1029\" type=\"#_x0000_t75\" style='position:absolute;z-index:5'><v:imagedata o:relid=\"rId1\" o:title=\"\"/><x:ClientData ObjectType=\"Pict\"><x:Anchor> 8, 0, 15, 0, 10, 0, 19, 0</x:Anchor><x:CF>Pict</x:CF></x:ClientData></v:shape></xml>",
        )
    };
    add("xl/drawings/vmlDrawing1.vml", vml.into_bytes());
    add(
        "xl/drawings/_rels/vmlDrawing1.vml.rels",
        rels(&[("rId1", &format!("{R}/image"), "../media/image2.emf")]).into_bytes(),
    );
    let props = |p: &str| {
        format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<formControlPr xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" {p}/>").into_bytes()
    };
    add(
        "xl/ctrlProps/ctrlProp1.xml",
        props("objectType=\"Button\" lockText=\"1\""),
    );
    add(
        "xl/ctrlProps/ctrlProp2.xml",
        props("objectType=\"CheckBox\" fmlaLink=\"$K$17\" lockText=\"1\" noThreeD=\"1\""),
    );
    add(
        "xl/ctrlProps/ctrlProp3.xml",
        props("objectType=\"List\" dx=\"22\" fmlaRange=\"$K$18:$K$20\" noThreeD=\"1\" sel=\"2\""),
    );
    add(
        "xl/activeX/activeX1.xml",
        format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>\n<ax:ocx xmlns:ax=\"http://schemas.microsoft.com/office/2006/activeX\" xmlns:r=\"{R}\" ax:classid=\"{{D7053240-CE69-11CD-A777-00DD01143C57}}\" ax:persistence=\"persistStream\" r:id=\"rId1\"/>").into_bytes(),
    );
    add(
        "xl/activeX/_rels/activeX1.xml.rels",
        rels(&[(
            "rId1",
            "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary",
            "activeX1.bin",
        )])
        .into_bytes(),
    );
    let mut bin = vec![
        0x40, 0x32, 0x05, 0xD7, 0x69, 0xCE, 0xCD, 0x11, 0xA7, 0x77, 0x00, 0xDD, 0x01, 0x14, 0x3C,
        0x57,
    ];
    bin.extend(super::controls::command_button_for_tests(
        "実行",
        "ＭＳ Ｐゴシック",
        220,
    ));
    add("xl/activeX/activeX1.bin", bin);

    zip_of(parts)
}

fn zip_of(parts: Vec<(String, Vec<u8>)>) -> Vec<u8> {
    let mut zip = zip::ZipWriter::new(std::io::Cursor::new(Vec::new()));
    for (name, data) in parts {
        zip.start_file(name, zip::write::SimpleFileOptions::default())
            .unwrap();
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
    let list: Vec<(&str, &str)> = objects
        .iter()
        .map(|o| {
            (
                o["kind"].as_str().unwrap(),
                o["name"].as_str().unwrap_or(""),
            )
        })
        .collect();
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
            ("chart", "グラフ 6"),
            ("group", "図表 7"),
            ("picture", "図 8"),
            ("shape", "テキスト ボックス 9"),
            ("control", "Check Box 3"),
            ("control", "List Box 4"),
        ]
    );
    let html = |i: usize| objects[i]["html"].as_str().unwrap();
    // The oval: the theme's accent color, from C3 (with its offsets in pixels).
    assert!(html(0).contains("fill=\"#4472C4\""), "{}", html(0));
    assert_eq!(objects[0]["from"], json!([2, 31.5, 2, 15.75]));
    assert_eq!(objects[0]["macro"], "Oval_Click");
    // The picture, once in the page's pictures.
    assert!(html(1).contains("data-img=\"0\""));
    assert!(
        data["images"][0]
            .as_str()
            .unwrap()
            .starts_with("data:image/png;base64,iVBORw0KGgo")
    );
    assert_eq!(objects[1]["ext"], json!([100.0, 100.0]));
    // The form button: its text and its macro.
    assert!(
        html(2).contains("class=\"ctl button\"") && html(2).contains("ボタン 1"),
        "{}",
        html(2)
    );
    assert_eq!(objects[2]["macro"], "Button1_Click");
    // The ActiveX button: its caption and font from its persisted data.
    assert!(
        html(3).contains("ctl button")
            && html(3).contains(">実行<")
            && html(3).contains("ＭＳ Ｐゴシック"),
        "{}",
        html(3)
    );
    // The arrow with its head; WordArt laid out by the page.
    assert!(
        html(4).contains("marker-end") && html(4).contains("#C00000"),
        "{}",
        html(4)
    );
    assert!(html(5).contains("class=\"warp\""), "{}", html(5));
    // The check box follows its linked cell (TRUE), not what was saved; the list shows its
    // items with the second one chosen.
    assert!(html(10).contains("class=\"box\"><svg"), "{}", html(10));
    assert!(
        html(11).contains("<span>東京</span><span class=\"sel\">大阪</span><span>名古屋</span>"),
        "{}",
        html(11)
    );
    // The chart: its series from the cells (not its cache), for the page to draw.
    let chart = &objects[6]["chart"];
    assert_eq!(chart["title"], "月別売上");
    assert_eq!(chart["legend"], "b");
    let series = &chart["groups"][0]["series"][0];
    assert_eq!(
        (&series["name"], &series["cat"], &series["val"]),
        (
            &json!("売上"),
            &json!(["4月", "5月", "6月"]),
            &json!([120.0, 150.0, 90.0])
        )
    );
    assert_eq!(chart["groups"][0]["kind"], "bar");
    assert_eq!(objects[6]["macro"], "Chart_Click");
    // SmartArt: the shapes Excel keeps drawn of it.
    assert_eq!(html(7).matches("<path").count(), 3, "{}", html(7));
    assert_eq!(
        html(7).matches("class=\"shape\"").count(),
        1 + 6,
        "{}",
        html(7)
    );
    assert!(
        html(7).contains(">企画</span>") && html(7).contains(">公開</span>"),
        "{}",
        html(7)
    );
    // The EMF picture as SVG.
    assert!(html(8).contains("data-img=\"1\""), "{}", html(8));
    assert!(
        data["images"][1]
            .as_str()
            .unwrap()
            .starts_with("data:image/svg+xml;base64,")
    );
    assert_eq!(objects[8]["macro"], "Picture_Click");
    // WordArt as Office 365 makes it: an outline, a gradient, a shadow.
    assert!(
        html(9).contains("-webkit-text-stroke:1pt #C55A11")
            && html(9).contains("background-clip:text")
            && html(9).contains("drop-shadow"),
        "{}",
        html(9)
    );
    // The comment is not drawn; the grid reaches to the objects (the chart to row 41).
    assert!(
        !objects
            .iter()
            .any(|o| o["html"].as_str().unwrap().contains("メモ"))
    );
    assert_eq!(sheet["rows"], 41);
    // The grid goes on past the cells.
    assert_eq!(
        (sheet["maxRows"].as_u64(), sheet["maxCols"].as_u64()),
        (Some(1_048_576), Some(16_384))
    );
    assert_eq!(
        (
            sheet["defaultWidth"].as_f64(),
            sheet["defaultHeight"].as_f64()
        ),
        (Some(72.0), Some(25.0))
    );
    assert_eq!(sheet["gridLines"], true);
}

/// What a run changed in the cells, the chart shows.
#[test]
fn charts_follow_the_cells() {
    let bytes = sample();
    let mut book = Book::open_bytes(&bytes).unwrap();
    book.set_values(
        0,
        crate::RangeRef::parse("M31").unwrap(),
        &[vec![crate::book::Val::Num(500.0)]],
    )
    .unwrap();
    let data = super::data(&mut book, "x.xlsx", Some(&bytes), None, None).unwrap();
    let chart = data["sheets"][0]["objects"]
        .as_array()
        .unwrap()
        .iter()
        .find(|o| o["kind"] == "chart")
        .unwrap();
    assert_eq!(
        chart["chart"]["groups"][0]["series"][0]["val"],
        json!([500.0, 150.0, 90.0])
    );
}

/// An embedded object: its picture (EMF) where the sheet puts it.
#[test]
fn an_embedded_object() {
    let data = view_of(&with_object());
    let objects = data["sheets"][0]["objects"].as_array().unwrap();
    let ole = objects
        .iter()
        .find(|o| o["name"] == "Word.Document.12")
        .expect("the object");
    assert_eq!(ole["from"], json!([8, 0.0, 15, 0.0]));
    let i: usize = ole["html"]
        .as_str()
        .unwrap()
        .split("data-img=\"")
        .nth(1)
        .unwrap()
        .split('"')
        .next()
        .unwrap()
        .parse()
        .unwrap();
    assert!(
        data["images"][i]
            .as_str()
            .unwrap()
            .starts_with("data:image/svg+xml;base64,")
    );
}

/// Without the file, the cells alone.
#[test]
fn without_the_drawing() {
    let mut book = Book::open_bytes(&sample()).unwrap();
    let data = super::data(&mut book, "x.xlsx", None, None, None).unwrap();
    assert_eq!(data["sheets"][0]["objects"], json!([]));
    assert_eq!(data["images"], json!([]));
    assert_eq!(data["sheets"][0]["gridLines"], true);
}

/// A sheet whose gridlines are off says so.
#[test]
fn gridlines_off() {
    let parts = super::Parts::read(&sample());
    assert_eq!(parts.grid_lines("Sheet1"), Some(true));
    let mut files = parts_of(&sample());
    let sheet = files
        .iter_mut()
        .find(|(n, _)| n == "xl/worksheets/sheet1.xml")
        .unwrap();
    sheet.1 = String::from_utf8_lossy(&sheet.1)
        .replace(
            "<sheetView workbookViewId=\"0\"/>",
            "<sheetView showGridLines=\"0\" workbookViewId=\"0\"/>",
        )
        .into_bytes();
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
    book.set_values(
        0,
        crate::RangeRef::parse("K17").unwrap(),
        &[vec![crate::book::Val::Bool(false)]],
    )
    .unwrap();
    let path = std::env::temp_dir().join(format!("ras-view-saved-{}.xlsx", std::process::id()));
    book.save(&path).unwrap();
    let saved = std::fs::read(&path).unwrap();
    let _ = std::fs::remove_file(&path);
    // Its own drawing has lost the ActiveX control.
    let own = view_of(&saved);
    assert!(
        !own["sheets"][0]["objects"]
            .as_array()
            .unwrap()
            .iter()
            .any(|o| o["name"] == "CommandButton1")
    );
    let mut book = Book::open_bytes(&saved).unwrap();
    let data = super::data(&mut book, "Drawing.xlsx", Some(&bytes), None, None).unwrap();
    let objects = data["sheets"][0]["objects"].as_array().unwrap();
    let html = |name: &str| {
        objects
            .iter()
            .find(|o| o["name"] == name)
            .and_then(|o| o["html"].as_str())
            .unwrap_or_default()
            .to_string()
    };
    assert!(html("CommandButton1").contains(">実行<"));
    // The check box: its linked cell is FALSE now.
    assert!(
        html("Check Box 3").contains("<span class=\"box\"></span>"),
        "{}",
        html("Check Box 3")
    );
}

/// A workbook as Japanese Excel saves one with formats the grid draws as Excel does: column
/// widths, the Normal style centred vertically, a built-in format whose code is Japanese Excel's,
/// a small number in General, a pattern fill, turned and shrunk text, a color scale and a data bar.
fn formatted() -> Vec<u8> {
    let parts = vec![
        (
            "[Content_Types].xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#
                .to_vec(),
        ),
        (
            "_rels/.rels".to_string(),
            rels(&[("rId1", &format!("{R}/officeDocument"), "xl/workbook.xml")]).into_bytes(),
        ),
        (
            "xl/workbook.xml".to_string(),
            format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"{R}\"><sheets><sheet name=\"Grid\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>").into_bytes(),
        ),
        (
            "xl/_rels/workbook.xml.rels".to_string(),
            rels(&[
                ("rId1", &format!("{R}/worksheet"), "worksheets/sheet1.xml"),
                ("rId2", &format!("{R}/styles"), "styles.xml"),
            ])
            .into_bytes(),
        ),
        (
            "xl/styles.xml".to_string(),
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="游ゴシック"/><family val="2"/><charset val="128"/><scheme val="minor"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="darkGrid"><fgColor rgb="FF00AA00"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf><xf numFmtId="38" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"><alignment vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"><alignment vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" textRotation="45"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" shrinkToFit="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#
                .as_bytes()
                .to_vec(),
        ),
        (
            "xl/worksheets/sheet1.xml".to_string(),
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetFormatPr defaultRowHeight="18.75"/><cols><col min="1" max="1" width="14.625" customWidth="1"/><col min="2" max="2" width="6.625" customWidth="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>plain</t></is></c><c r="B1" s="1"><v>-1500</v></c><c r="C1"><v>1.2345E-5</v></c><c r="D1" s="2" t="inlineStr"><is><t>pattern</t></is></c><c r="E1" s="3" t="inlineStr"><is><t>turned</t></is></c><c r="F1" s="4" t="inlineStr"><is><t>shrunk</t></is></c></row><row r="2"><c r="A2"><v>10</v></c><c r="B2"><v>1</v></c></row><row r="3"><c r="A3"><v>50</v></c><c r="B3"><v>5</v></c></row><row r="4"><c r="A4"><v>90</v></c><c r="B4"><v>9</v></c></row></sheetData><conditionalFormatting sqref="A2:A4"><cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar></cfRule></conditionalFormatting><conditionalFormatting sqref="B2:B4"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting></worksheet>"#
                .as_bytes()
                .to_vec(),
        ),
    ];
    zip_of(parts)
}

#[test]
fn formats_as_excel_draws_them() {
    let data = view_of(&formatted());
    let sheet = &data["sheets"][0];
    // Columns 14 and 6 characters wide (the file's widths have the padding in them): 8 px a
    // digit and 5 px of padding, as Excel shows them at 100%.
    assert_eq!(sheet["colWidths"][0], json!(117.0));
    assert_eq!(sheet["colWidths"][1], json!(53.0));
    let cell = |a: &str| -> Json {
        let at = gridcraft_core::addr::CellRef::parse(a).unwrap();
        sheet["cells"]
            .as_array()
            .unwrap()
            .iter()
            .find(|c| c["r"] == json!(at.row) && c["c"] == json!(at.col))
            .cloned()
            .unwrap_or(Json::Null)
    };
    let css = |a: &str| -> String {
        let s = cell(a)["s"].as_u64().unwrap_or(0) as usize;
        sheet["styles"][s].as_str().unwrap_or("").to_string()
    };
    // A cell without a format of its own has the Normal style's alignment.
    assert!(css("A1").contains("vertical-align:middle"), "{}", css("A1"));
    // Built-in format 38 as Japanese Excel has it, in red.
    assert_eq!(cell("B1")["t"], json!("-1,500"));
    assert!(css("B1").contains("color:#FF0000"), "{}", css("B1"));
    // General: all the digits of a small number when they fit in 11 characters.
    assert_eq!(cell("C1")["t"], json!("0.000012345"));
    assert_eq!(cell("C1")["x"], json!(1.2345e-5));
    // A pattern over white; turned and shrunk text for the page.
    assert!(
        css("D1").contains("background:#FFFFFF url(\"data:image/svg+xml,"),
        "{}",
        css("D1")
    );
    assert!(css("D1").contains("fill='%2300AA00'"), "{}", css("D1"));
    assert!(css("E1").contains("--rot:45;"), "{}", css("E1"));
    assert!(css("F1").contains("--shrink:1;"), "{}", css("F1"));
    // The data bar: 10% to 90% of the cell from the lowest to the highest value.
    let bar = |a: &str| cell(a)["bar"].clone();
    assert_eq!(bar("A2")["w"], json!(0.1));
    assert_eq!(bar("A3")["w"], json!(0.5));
    assert_eq!(bar("A4")["w"], json!(0.9));
    assert!(
        bar("A2")["bg"]
            .as_str()
            .unwrap()
            .starts_with("linear-gradient(to right,#638EC6,")
    );
    // The color scale: the lowest, the middle and the highest colors.
    assert!(css("B2").contains("background:#F8696B;"), "{}", css("B2"));
    assert!(css("B3").contains("background:#FFEB84;"), "{}", css("B3"));
    assert!(css("B4").contains("background:#63BE7B;"), "{}", css("B4"));
}
