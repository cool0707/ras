//! Shapes, text boxes, lines and pictures put on a sheet as Excel's 挿入 > 図 puts them (the
//! sheet's drawing part, made when it has none; a picture's bytes in `xl/media/`), and the
//! objects of the drawing moved, resized, given a text or deleted as they are with the mouse in
//! Excel. Nothing else of the file changes.

use std::collections::HashSet;

use super::insert::{Sheet, place, relative, rid_attr, tree};
use super::scan::{self, Edits, El};
use super::{Place, Result};
use crate::book::Book;
use crate::book::view::Sizes;

const XDR_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const A_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const DRAWING_TYPE: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";
/// EMU in a pixel at 96 dpi.
const EMU_PX: f64 = 9525.0;
/// The size of a shape put without one (px): Excel's when the sheet is clicked, not dragged.
const SHAPE_SIZE: (f64, f64) = (96.0, 96.0);
/// The largest box of an object (px).
pub(super) const MAX_SIZE: f64 = 100_000.0;

/// The shapes that can be put on a sheet: the preset geometry, and the name Excel gives a new
/// one (followed by its number). `textBox` is a rectangle with `txBox`; `line` and `arrow` are
/// connectors.
pub(super) const SHAPES: [(&str, &str); 24] = [
    ("rect", "正方形/長方形"),
    ("roundRect", "四角形: 角を丸くする"),
    ("ellipse", "楕円"),
    ("triangle", "二等辺三角形"),
    ("rtTriangle", "直角三角形"),
    ("diamond", "ひし形"),
    ("parallelogram", "平行四辺形"),
    ("trapezoid", "台形"),
    ("pentagon", "五角形"),
    ("hexagon", "六角形"),
    ("octagon", "八角形"),
    ("rightArrow", "矢印: 右"),
    ("leftArrow", "矢印: 左"),
    ("upArrow", "矢印: 上"),
    ("downArrow", "矢印: 下"),
    ("star5", "星: 5 pt"),
    ("heart", "ハート"),
    ("smileyFace", "スマイル"),
    ("wedgeRectCallout", "吹き出し: 四角形"),
    ("wedgeRoundRectCallout", "吹き出し: 角を丸めた四角形"),
    ("wedgeEllipseCallout", "吹き出し: 円形"),
    ("textBox", "テキスト ボックス"),
    ("line", "直線コネクタ"),
    ("arrow", "直線矢印コネクタ"),
];

// ---------------------------------------------------------------------------- where it goes

/// The box of a place on the sheet in pixels: (left, top, right, bottom). `own`: the size of
/// the object when the place gives none (one side given: the other keeps the proportion).
fn box_of(sizes: &Sizes, place: &Place, own: (f64, f64)) -> (f64, f64, f64, f64) {
    match place {
        Place::Cells(range) => (
            sizes.x(range.start.col),
            sizes.y(range.start.row),
            sizes.x(range.end.col + 1),
            sizes.y(range.end.row + 1),
        ),
        Place::At {
            at,
            dx,
            dy,
            width,
            height,
        } => {
            let x = (sizes.x(at.col) + dx).max(0.0);
            let y = (sizes.y(at.row) + dy).max(0.0);
            let (w, h) = match (width, height) {
                (Some(w), Some(h)) => (*w, *h),
                (Some(w), None) if own.0 > 0.0 => (*w, own.1 * w / own.0),
                (None, Some(h)) if own.1 > 0.0 => (own.0 * h / own.1, *h),
                (Some(w), None) => (*w, own.1),
                (None, Some(h)) => (own.0, *h),
                (None, None) => own,
            };
            (x, y, x + w.max(0.0), y + h.max(0.0))
        }
    }
}

fn emu(px: f64) -> i64 {
    (px * EMU_PX).round() as i64
}

/// A marker of an anchor (`from`, `to`) at a point of the sheet.
fn marker_xml(x: &str, tag: &str, sizes: &Sizes, px: f64, py: f64) -> String {
    let (col, dx, row, dy) = sizes.marker(px, py);
    format!(
        "<{x}{tag}><{x}col>{col}</{x}col><{x}colOff>{}</{x}colOff><{x}row>{row}</{x}row><{x}rowOff>{}</{x}rowOff></{x}{tag}>",
        emu(dx),
        emu(dy)
    )
}

fn sizes_of(bytes: &[u8], sheet: &str) -> Result<Sizes> {
    let book = Book::open_bytes(bytes)?;
    let si = book
        .find_sheet(sheet)
        .ok_or_else(|| format!("シート `{sheet}` がありません"))?;
    book.sizes(si)
        .ok_or_else(|| format!("シート `{sheet}` がありません"))
}

// ---------------------------------------------------------------------------- the drawing

/// The drawing of a sheet: its part and XML; a new one when the sheet has none.
struct Drawing {
    part: String,
    doc: String,
    /// The sheet's XML with the `<drawing>` that refers to a new drawing.
    sheet_doc: Option<String>,
}

/// The sheet's drawing (made, with its relationship and content type, when it has none).
fn drawing_of(s: &mut Sheet, make: bool) -> Result<Option<Drawing>> {
    let sdoc = s.doc.clone();
    let root = tree(&sdoc, &s.part)?;
    let rels = s.rels()?;
    let found = root
        .child("drawing")
        .and_then(|d| super::insert::prefixed_attr(&sdoc, d, "id"))
        .and_then(|id| rels.iter().find(|r| r.0 == id))
        .map(|r| r.2.clone());
    if let Some(part) = found {
        let doc = s
            .text(&part)?
            .ok_or_else(|| format!("{part} がありません"))?;
        return Ok(Some(Drawing {
            part,
            doc,
            sheet_doc: None,
        }));
    }
    if !make {
        return Ok(None);
    }
    let part = s.free("xl/drawings/drawing", ".xml");
    let doc = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<xdr:wsDr xmlns:xdr=\"{XDR_NS}\" xmlns:a=\"{A_NS}\"></xdr:wsDr>"
    );
    let id = s.add_rel("drawing", &relative(&s.part.clone(), &part), false)?;
    s.content_type(Some(&part), None, DRAWING_TYPE)?;
    let p = scan::prefix(root.name);
    let mut edits = Edits::default();
    edits.insert(
        place(&root, "drawing"),
        format!("<{p}drawing {}/>", rid_attr(&sdoc, &root, &id)),
    );
    Ok(Some(Drawing {
        part,
        doc,
        sheet_doc: Some(edits.apply(&sdoc)),
    }))
}

/// The prefix (with its colon) the root of a document gives a namespace; None when it does
/// not declare it.
fn ns_prefix(doc: &str, root: &El, ns: &str) -> Option<String> {
    scan::attrs(doc, root.open.clone())
        .into_iter()
        .find(|(name, span)| {
            (*name == "xmlns" || name.starts_with("xmlns:"))
                && scan::unescape(&doc[span.clone()]) == ns
        })
        .map(|(name, _)| match name.strip_prefix("xmlns:") {
            Some(p) => format!("{p}:"),
            None => String::new(),
        })
}

/// The drawing with a new anchor (`make` writes it with the drawing's prefix for the
/// spreadsheet drawing and `a:` for DrawingML, and the id its object gets) at its end.
fn append(doc: &str, make: impl FnOnce(&str, &str, u32) -> String) -> Result<String> {
    let root = scan::tree(doc, &[]).ok_or("描画のパートを読み取れません")?;
    let x = ns_prefix(doc, &root, XDR_NS).unwrap_or_else(|| scan::prefix(root.name).to_string());
    // `a:` declared on the anchor where the root does not have it so.
    let a_decl = if ns_prefix(doc, &root, A_NS).as_deref() == Some("a:") {
        String::new()
    } else {
        format!(" xmlns:a=\"{A_NS}\"")
    };
    let id = next_id(doc);
    let anchor = make(&x, &a_decl, id);
    let mut edits = Edits::default();
    if root.open.end == root.end {
        let head = doc[root.open.clone()].trim_end_matches("/>");
        edits.replace(
            root.open.clone(),
            format!("{head}>{anchor}</{}>", root.name),
        );
    } else {
        edits.insert(root.inner.end, anchor);
    }
    Ok(edits.apply(doc))
}

/// The id a new object of the drawing gets: one more than the largest (2 for the first, as
/// Excel numbers them).
fn next_id(doc: &str) -> u32 {
    scan::tags(doc, 0)
        .filter(|t| !t.is_close() && t.local() == "cNvPr")
        .filter_map(|t| scan::attr(doc, t.start..t.end, "id")?.parse::<u32>().ok())
        .max()
        .unwrap_or(1)
        + 1
}

/// The paragraphs of a text, one a line, as Excel writes those typed into a shape.
fn paragraphs(text: &str, algn: &str, ppr: Option<&str>, rpr: Option<&str>) -> String {
    let end_rpr = "<a:endParaRPr kumimoji=\"1\" lang=\"ja-JP\" altLang=\"en-US\" sz=\"1100\"/>";
    let rpr = rpr.map_or_else(
        || "<a:rPr kumimoji=\"1\" lang=\"ja-JP\" altLang=\"en-US\" sz=\"1100\"/>".to_string(),
        str::to_string,
    );
    let ppr = ppr.map_or_else(
        || {
            if algn.is_empty() {
                String::new()
            } else {
                format!("<a:pPr algn=\"{algn}\"/>")
            }
        },
        str::to_string,
    );
    let text = text.replace("\r\n", "\n").replace('\r', "\n");
    text.split('\n')
        .map(|line| {
            if line.is_empty() {
                format!("<a:p>{ppr}{end_rpr}</a:p>")
            } else {
                format!(
                    "<a:p>{ppr}<a:r>{rpr}<a:t>{}</a:t></a:r></a:p>",
                    scan::escape(line)
                )
            }
        })
        .collect()
}

// ---------------------------------------------------------------------------- shapes

/// The file with a shape (`kind`: one of `SHAPES`) put at `place` on the sheet `sheet`, saying
/// `text`. `flip`: a line drawn right to left, bottom to top (flipped as Excel writes it).
pub(super) fn shape(
    bytes: &[u8],
    sheet: &str,
    kind: &str,
    place: &Place,
    text: &str,
    flip: (bool, bool),
) -> Result<Vec<u8>> {
    let base = SHAPES
        .iter()
        .find(|(k, _)| *k == kind)
        .map(|(_, n)| *n)
        .ok_or_else(|| format!("図形 `{kind}` はありません"))?;
    let sizes = sizes_of(bytes, sheet)?;
    let line = matches!(kind, "line" | "arrow");
    let own = match kind {
        "line" | "arrow" => (SHAPE_SIZE.0, 0.0),
        "textBox" => (144.0, 36.0),
        _ => SHAPE_SIZE,
    };
    let (x0, y0, x1, y1) = box_of(&sizes, place, own);
    let mut s = Sheet::open(bytes, sheet)?;
    let d = drawing_of(&mut s, true)?.ok_or("描画のパートを作れません")?;
    let doc = append(&d.doc, |x, a_decl, id| {
        let name = scan::escape(&format!("{base} {}", id - 1));
        let flips = format!(
            "{}{}",
            if flip.0 { " flipH=\"1\"" } else { "" },
            if flip.1 { " flipV=\"1\"" } else { "" }
        );
        let xfrm = format!(
            "<a:xfrm{flips}><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/></a:xfrm>",
            emu(x0),
            emu(y0),
            emu(x1 - x0),
            emu(y1 - y0)
        );
        let object = if line {
            let (prst, ln) = if kind == "arrow" {
                (
                    "straightConnector1",
                    "<a:ln><a:tailEnd type=\"triangle\"/></a:ln>",
                )
            } else {
                ("line", "")
            };
            format!(
                "<{x}cxnSp macro=\"\"><{x}nvCxnSpPr><{x}cNvPr id=\"{id}\" name=\"{name}\"/><{x}cNvCxnSpPr/></{x}nvCxnSpPr><{x}spPr>{xfrm}<a:prstGeom prst=\"{prst}\"><a:avLst/></a:prstGeom>{ln}</{x}spPr><{x}style><a:lnRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:lnRef><a:fillRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></a:fontRef></{x}style></{x}cxnSp>"
            )
        } else if kind == "textBox" {
            format!(
                "<{x}sp macro=\"\" textlink=\"\"><{x}nvSpPr><{x}cNvPr id=\"{id}\" name=\"{name}\"/><{x}cNvSpPr txBox=\"1\"/></{x}nvSpPr><{x}spPr>{xfrm}<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill><a:ln w=\"9525\" cmpd=\"sng\"><a:solidFill><a:schemeClr val=\"lt1\"><a:shade val=\"50000\"/></a:schemeClr></a:solidFill></a:ln></{x}spPr><{x}style><a:lnRef idx=\"0\"><a:scrgbClr r=\"0\" g=\"0\" b=\"0\"/></a:lnRef><a:fillRef idx=\"0\"><a:scrgbClr r=\"0\" g=\"0\" b=\"0\"/></a:fillRef><a:effectRef idx=\"0\"><a:scrgbClr r=\"0\" g=\"0\" b=\"0\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"dk1\"/></a:fontRef></{x}style><{x}txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\" wrap=\"square\" rtlCol=\"0\" anchor=\"t\"/><a:lstStyle/>{}</{x}txBody></{x}sp>",
                paragraphs(text, "l", None, None)
            )
        } else {
            format!(
                "<{x}sp macro=\"\" textlink=\"\"><{x}nvSpPr><{x}cNvPr id=\"{id}\" name=\"{name}\"/><{x}cNvSpPr/></{x}nvSpPr><{x}spPr>{xfrm}<a:prstGeom prst=\"{kind}\"><a:avLst/></a:prstGeom></{x}spPr><{x}style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"50000\"/></a:schemeClr></a:lnRef><a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef></{x}style><{x}txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\" rtlCol=\"0\" anchor=\"ctr\"/><a:lstStyle/>{}</{x}txBody></{x}sp>",
                paragraphs(text, "ctr", None, None)
            )
        };
        format!(
            "<{x}twoCellAnchor{a_decl}>{}{}{object}<{x}clientData/></{x}twoCellAnchor>",
            marker_xml(x, "from", &sizes, x0, y0),
            marker_xml(x, "to", &sizes, x1, y1)
        )
    })?;
    s.put(&d.part, doc);
    s.finish(d.sheet_doc)
}

// ---------------------------------------------------------------------------- pictures

/// A picture's format (its extension and content type) and size in pixels, from its first
/// bytes. None for what Excel does not take as a picture of its own.
pub(super) fn picture_info(data: &[u8]) -> Option<(&'static str, &'static str, (f64, f64))> {
    let be16 = |i: usize| Some(u16::from_be_bytes([*data.get(i)?, *data.get(i + 1)?]) as f64);
    let le16 = |i: usize| Some(u16::from_le_bytes([*data.get(i)?, *data.get(i + 1)?]) as f64);
    let be32 = |i: usize| Some(u32::from_be_bytes(data.get(i..i + 4)?.try_into().ok()?) as f64);
    let le32 = |i: usize| {
        Some(i32::from_le_bytes(data.get(i..i + 4)?.try_into().ok()?).unsigned_abs() as f64)
    };
    if data.starts_with(b"\x89PNG\r\n\x1a\n") {
        return Some(("png", "image/png", (be32(16)?, be32(20)?)));
    }
    if data.starts_with(b"GIF87a") || data.starts_with(b"GIF89a") {
        return Some(("gif", "image/gif", (le16(6)?, le16(8)?)));
    }
    if data.starts_with(b"BM") && data.len() > 26 {
        return Some(("bmp", "image/bmp", (le32(18)?, le32(22)?)));
    }
    if data.starts_with(&[0xFF, 0xD8]) {
        // The frame's header (SOFn) has the size.
        let mut i = 2;
        while i + 9 < data.len() {
            if data[i] != 0xFF {
                i += 1;
                continue;
            }
            let m = data[i + 1];
            if m == 0xFF {
                i += 1;
                continue;
            }
            if matches!(m, 0xD8 | 0x01 | 0xD0..=0xD7) {
                i += 2;
                continue;
            }
            let len = be16(i + 2)? as usize;
            if (0xC0..=0xCF).contains(&m) && !matches!(m, 0xC4 | 0xC8 | 0xCC) {
                return Some(("jpeg", "image/jpeg", (be16(i + 7)?, be16(i + 5)?)));
            }
            i += 2 + len;
        }
        return None;
    }
    None
}

/// The file with the picture `data` (PNG, JPEG, GIF, BMP) put at `place` on the sheet `sheet`
/// (its own size where the place gives none), as Excel's 挿入 > 画像 puts one: moving with its
/// cells, not sized with them.
pub(super) fn picture(bytes: &[u8], sheet: &str, place: &Place, data: &[u8]) -> Result<Vec<u8>> {
    let (ext, mime, own) = picture_info(data)
        .ok_or("画像の形式が読み取れません（PNG・JPEG・GIF・BMP の画像を指定してください）")?;
    if own.0 <= 0.0 || own.1 <= 0.0 {
        return Err("画像の大きさが読み取れません".into());
    }
    let sizes = sizes_of(bytes, sheet)?;
    let (x0, y0, x1, y1) = box_of(&sizes, place, own);
    let mut s = Sheet::open(bytes, sheet)?;
    let d = drawing_of(&mut s, true)?.ok_or("描画のパートを作れません")?;
    let media = s.free("xl/media/image", &format!(".{ext}"));
    s.put_bytes(&media, data.to_vec());
    s.content_type(None, Some(ext), mime)?;
    let rid = s.add_rel_of(&d.part, "image", &relative(&d.part, &media), false)?;
    let doc = append(&d.doc, |x, a_decl, id| {
        let name = scan::escape(&format!("図 {}", id - 1));
        format!(
            "<{x}twoCellAnchor editAs=\"oneCell\"{a_decl}>{}{}<{x}pic><{x}nvPicPr><{x}cNvPr id=\"{id}\" name=\"{name}\"/><{x}cNvPicPr><a:picLocks noChangeAspect=\"1\"/></{x}cNvPicPr></{x}nvPicPr><{x}blipFill><a:blip xmlns:r=\"{REL_NS}\" r:embed=\"{rid}\"/><a:stretch><a:fillRect/></a:stretch></{x}blipFill><{x}spPr><a:xfrm><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></{x}spPr></{x}pic><{x}clientData/></{x}twoCellAnchor>",
            marker_xml(x, "from", &sizes, x0, y0),
            marker_xml(x, "to", &sizes, x1, y1),
            emu(x0),
            emu(y0),
            emu(x1 - x0),
            emu(y1 - y0)
        )
    })?;
    s.put(&d.part, doc);
    s.finish(d.sheet_doc)
}

/// The bytes of a picture given as base64 (a `data:` URL too).
pub(super) fn base64_decode(text: &str) -> Result<Vec<u8>> {
    let text = match text.split_once(";base64,") {
        Some((head, rest)) if head.starts_with("data:") => rest,
        _ => text,
    };
    let mut out = Vec::with_capacity(text.len() / 4 * 3);
    let (mut acc, mut bits) = (0u32, 0);
    for c in text.bytes() {
        let v = match c {
            b'A'..=b'Z' => c - b'A',
            b'a'..=b'z' => c - b'a' + 26,
            b'0'..=b'9' => c - b'0' + 52,
            b'+' | b'-' => 62,
            b'/' | b'_' => 63,
            b'=' => break,
            b' ' | b'\n' | b'\r' | b'\t' => continue,
            _ => return Err("`data` は base64 で指定してください".into()),
        };
        acc = acc << 6 | v as u32;
        bits += 6;
        if bits >= 8 {
            bits -= 8;
            out.push((acc >> bits) as u8);
            acc &= (1 << bits) - 1;
        }
    }
    Ok(out)
}

// ---------------------------------------------------------------------------- the objects

/// An object of the drawing as the grid lists it: its anchors (the alternatives of one in
/// `mc:AlternateContent`) and what is deleted with it.
struct Unit<'a, 'd> {
    span: std::ops::Range<usize>,
    anchors: Vec<&'a El<'d>>,
}

const OBJECTS: [&str; 5] = ["sp", "cxnSp", "grpSp", "pic", "graphicFrame"];

fn object_of<'a, 'd>(anchor: &'a El<'d>) -> Option<&'a El<'d>> {
    anchor
        .children
        .iter()
        .find(|c| OBJECTS.contains(&c.local()))
}

fn find_all<'a, 'd>(el: &'a El<'d>, name: &str, out: &mut Vec<&'a El<'d>>) {
    for c in &el.children {
        if c.local() == name {
            out.push(c);
        }
        find_all(c, name, out);
    }
}

fn first<'a, 'd>(el: &'a El<'d>, name: &str) -> Option<&'a El<'d>> {
    let mut all = Vec::new();
    find_all(el, name, &mut all);
    all.into_iter().next()
}

/// The objects of the drawing, in order.
fn units<'a, 'd>(root: &'a El<'d>) -> Vec<Unit<'a, 'd>> {
    let mut out = Vec::new();
    for c in &root.children {
        if c.local().ends_with("Anchor") {
            out.push(Unit {
                span: c.span(),
                anchors: vec![c],
            });
        } else if c.local() == "AlternateContent" {
            let mut anchors = Vec::new();
            for alt in &c.children {
                anchors.extend(
                    alt.children
                        .iter()
                        .filter(|a| a.local().ends_with("Anchor")),
                );
            }
            if !anchors.is_empty() {
                out.push(Unit {
                    span: c.span(),
                    anchors,
                });
            }
        }
    }
    out
}

fn name_of(doc: &str, anchor: &El) -> Option<String> {
    let obj = object_of(anchor)?;
    let nv = obj.children.iter().find(|c| c.local().starts_with("nv"))?;
    nv.child("cNvPr")?.attr(doc, "name")
}

/// The point of a marker (`from`, `to`) of an anchor in pixels.
fn marker_point(doc: &str, sizes: &Sizes, m: &El) -> Option<(f64, f64)> {
    let n = |name: &str| -> Option<f64> { m.child(name)?.text(doc).trim().parse().ok() };
    Some((
        sizes.x(n("col")? as u32) + n("colOff")? / EMU_PX,
        sizes.y(n("row")? as u32) + n("rowOff")? / EMU_PX,
    ))
}

/// The box of an anchor in pixels.
fn anchor_box(doc: &str, sizes: &Sizes, a: &El) -> Option<(f64, f64, f64, f64)> {
    let ext = |el: &El| -> Option<(f64, f64)> {
        Some((
            el.attr(doc, "cx")?.parse::<f64>().ok()? / EMU_PX,
            el.attr(doc, "cy")?.parse::<f64>().ok()? / EMU_PX,
        ))
    };
    let (x, y) = match (a.child("from"), a.child("pos")) {
        (Some(f), _) => marker_point(doc, sizes, f)?,
        (None, Some(p)) => (
            p.attr(doc, "x")?.parse::<f64>().ok()? / EMU_PX,
            p.attr(doc, "y")?.parse::<f64>().ok()? / EMU_PX,
        ),
        _ => return None,
    };
    let (x2, y2) = match (a.child("to"), a.child("ext")) {
        (Some(t), _) => marker_point(doc, sizes, t)?,
        (None, Some(e)) => {
            let (w, h) = ext(e)?;
            (x + w, y + h)
        }
        _ => return None,
    };
    Some((x, y, x2, y2))
}

/// The file with the object named `name` (as the name box shows it) of the sheet `sheet`
/// moved and sized to `place` (the size it has where the place gives none), saying `text`, or
/// (`delete`) taken away with its picture or chart.
pub(super) fn object(
    bytes: &[u8],
    sheet: &str,
    name: &str,
    place: Option<&Place>,
    text: Option<&str>,
    delete: bool,
) -> Result<Vec<u8>> {
    let sizes = sizes_of(bytes, sheet)?;
    let mut s = Sheet::open(bytes, sheet)?;
    let missing = || format!("シート `{sheet}` に図形・画像 `{name}` がありません");
    let d = drawing_of(&mut s, false)?.ok_or_else(missing)?;
    let doc = d.doc.clone();
    let root = scan::tree(&doc, &[]).ok_or("描画のパートを読み取れません")?;
    let all = units(&root);
    let unit = all
        .iter()
        .find(|u| {
            u.anchors
                .iter()
                .any(|a| name_of(&doc, a).as_deref() == Some(name))
        })
        .ok_or_else(missing)?;
    // A form or ActiveX control is drawn from the sheet's controls and its VML shape too.
    if unit.anchors.iter().any(|a| first(a, "compatExt").is_some()) {
        return Err(format!(
            "`{name}` はコントロールのため、ここでは移動・削除できません（Excel で行ってください）"
        ));
    }
    let mut edits = Edits::default();
    if delete {
        edits.replace(unit.span.clone(), "");
        let removed = &doc[unit.span.clone()];
        let after = edits.apply(&doc);
        let after = drop_unused(&mut s, &d.part, removed, after)?;
        s.put(&d.part, after);
        return s.finish(None);
    }
    if let Some(place) = place {
        for a in &unit.anchors {
            let now = anchor_box(&doc, &sizes, a)
                .ok_or_else(|| format!("`{name}` の位置を読み取れません"))?;
            let own = (now.2 - now.0, now.3 - now.1);
            let (x0, y0, x1, y1) = box_of(&sizes, place, own);
            let x = scan::prefix(a.name).to_string();
            for (tag, px, py) in [("from", x0, y0), ("to", x1, y1)] {
                if let Some(m) = a.child(tag) {
                    edits.replace(m.span(), marker_xml(&x, tag, &sizes, px, py));
                }
            }
            if let Some(p) = a.child("pos") {
                scan::set_attr(&doc, p.open.clone(), "x", &emu(x0).to_string(), &mut edits);
                scan::set_attr(&doc, p.open.clone(), "y", &emu(y0).to_string(), &mut edits);
            }
            if let Some(e) = a.child("ext") {
                scan::set_attr(
                    &doc,
                    e.open.clone(),
                    "cx",
                    &emu(x1 - x0).to_string(),
                    &mut edits,
                );
                scan::set_attr(
                    &doc,
                    e.open.clone(),
                    "cy",
                    &emu(y1 - y0).to_string(),
                    &mut edits,
                );
            }
            // The transform of the object says the same in EMU (a chart's frame keeps its own).
            let obj = object_of(a).filter(|o| o.local() != "graphicFrame");
            let props = obj.and_then(|o| {
                o.children
                    .iter()
                    .find(|c| c.local() == "spPr" || c.local() == "grpSpPr")
            });
            if let Some(xfrm) = props.and_then(|p| p.child("xfrm")) {
                if let Some(off) = xfrm.child("off") {
                    scan::set_attr(
                        &doc,
                        off.open.clone(),
                        "x",
                        &emu(x0).to_string(),
                        &mut edits,
                    );
                    scan::set_attr(
                        &doc,
                        off.open.clone(),
                        "y",
                        &emu(y0).to_string(),
                        &mut edits,
                    );
                }
                if let Some(ext) = xfrm.child("ext") {
                    scan::set_attr(
                        &doc,
                        ext.open.clone(),
                        "cx",
                        &emu(x1 - x0).to_string(),
                        &mut edits,
                    );
                    scan::set_attr(
                        &doc,
                        ext.open.clone(),
                        "cy",
                        &emu(y1 - y0).to_string(),
                        &mut edits,
                    );
                }
            }
        }
    }
    if let Some(text) = text {
        for a in &unit.anchors {
            let obj = object_of(a).filter(|o| o.local() == "sp").ok_or_else(|| {
                format!(
                    "`{name}` には文字を入れられません（図形・テキスト ボックスに入れられます）"
                )
            })?;
            set_text(&doc, obj, text, &mut edits);
        }
    }
    if edits.is_empty() {
        return Ok(bytes.to_vec());
    }
    s.put(&d.part, edits.apply(&doc));
    s.finish(None)
}

/// The text of a shape replaced: its paragraphs, one a line, keep the first one's paragraph
/// and run properties (the font, the alignment); the body's properties stay.
fn set_text(doc: &str, sp: &El, text: &str, edits: &mut Edits) {
    let x = scan::prefix(sp.name);
    let Some(body) = sp.child("txBody") else {
        // A shape without a text body gets one, centered as a new shape's.
        let at = sp.children.last().map_or(sp.inner.end, |c| c.end);
        edits.insert(
            at,
            format!(
                "<{x}txBody><a:bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\" rtlCol=\"0\" anchor=\"ctr\"/><a:lstStyle/>{}</{x}txBody>",
                paragraphs(text, "ctr", None, None)
            ),
        );
        return;
    };
    let paras: Vec<&El> = body.children.iter().filter(|c| c.local() == "p").collect();
    let a = paras.first().map_or("a:", |p| scan::prefix(p.name));
    let ppr = paras
        .first()
        .and_then(|p| p.child("pPr"))
        .map(|e| &doc[e.span()]);
    let rpr = paras.iter().find_map(|p| {
        p.children
            .iter()
            .find(|r| r.local() == "r")
            .and_then(|r| r.child("rPr"))
            .map(|e| &doc[e.span()])
    });
    let mut new = paragraphs(text, "", ppr, rpr);
    if a != "a:" {
        new = new
            .replace("<a:", &format!("<{a}"))
            .replace("</a:", &format!("</{a}"));
    }
    match (paras.first(), paras.last()) {
        (Some(f), Some(l)) => edits.replace(f.open.start..l.end, new),
        _ => {
            let at = body.children.last().map_or(body.inner.end, |c| c.end);
            edits.insert(at, new);
        }
    }
}

/// The drawing `doc` (with `removed` taken out of it) without the relationships only the
/// removed XML used, and the parts they pointed to when nothing else points to them.
fn drop_unused(s: &mut Sheet, drawing: &str, removed: &str, doc: String) -> Result<String> {
    let rels = s.rels_of(drawing)?;
    let used = |text: &str, id: &str| {
        text.contains(&format!("=\"{id}\"")) || text.contains(&format!("='{id}'"))
    };
    let ids: HashSet<String> = rels
        .iter()
        .filter(|r| used(removed, &r.0) && !used(&doc, &r.0))
        .map(|r| r.0.clone())
        .collect();
    s.remove_rels_of(drawing, &ids)?;
    for (_, _, target, external) in rels.iter().filter(|r| ids.contains(&r.0)) {
        if !external {
            drop_part(s, target)?;
        }
    }
    Ok(doc)
}

/// A part taken out of the file when no relationship points to it any more, with its own
/// relationships and the parts only it pointed to (a chart's style and colors), and its content
/// type.
fn drop_part(s: &mut Sheet, part: &str) -> Result<()> {
    if !s.has(part) || referenced(s, part)? {
        return Ok(());
    }
    let rels = s.rels_of(part)?;
    s.remove(part);
    s.remove(&super::insert::rels_part_of(part));
    remove_override(s, part)?;
    for (_, _, target, external) in rels {
        if !external {
            drop_part(s, &target)?;
        }
    }
    Ok(())
}

/// Whether a relationship of any part of the file points to `part`.
fn referenced(s: &mut Sheet, part: &str) -> Result<bool> {
    let owners: Vec<String> = s
        .pack
        .names()
        .iter()
        .filter(|n| n.ends_with(".rels"))
        .filter_map(|n| {
            let (dir, file) = n
                .rsplit_once("/_rels/")
                .or_else(|| n.strip_prefix("_rels/").map(|f| ("", f)))?;
            let file = file.strip_suffix(".rels")?;
            Some(if dir.is_empty() {
                file.to_string()
            } else {
                format!("{dir}/{file}")
            })
        })
        .collect();
    for owner in owners {
        if s.has(&super::insert::rels_part_of(&owner))
            && s.rels_of(&owner)?
                .iter()
                .any(|r| !r.3 && r.2.eq_ignore_ascii_case(part))
        {
            return Ok(true);
        }
    }
    Ok(false)
}

fn remove_override(s: &mut Sheet, part: &str) -> Result<()> {
    let ct = "[Content_Types].xml";
    let Some(doc) = s.text(ct)? else {
        return Ok(());
    };
    let name = format!("/{part}");
    let Some(root) = scan::tree(&doc, &[]) else {
        return Ok(());
    };
    let mut edits = Edits::default();
    for c in root.children.iter().filter(|c| c.local() == "Override") {
        if c.attr(&doc, "PartName")
            .is_some_and(|n| n.eq_ignore_ascii_case(&name))
        {
            edits.replace(c.span(), "");
        }
    }
    if !edits.is_empty() {
        let text = edits.apply(&doc);
        s.put(ct, text);
    }
    Ok(())
}
