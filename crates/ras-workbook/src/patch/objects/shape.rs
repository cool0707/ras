//! Shapes of a sheet's drawing (DrawingML): added, and given a name, a text, a macro, colours.

use serde_json::Value as Json;

use super::{
    Pkg, REL, Rect, Result, Settings, Sheet, colour_of, emu, flag_of, items, macro_ref, nv_pr,
    object_of, placement_of, rel_prefix, sheet_insert, text_of, wrong,
};
use crate::patch::scan::{self, Edits, El};

const XDR: &str = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const DRAWING_TYPE: &str = "application/vnd.openxmlformats-officedocument.drawing+xml";

/// The shapes with a name of their own: (what the edit calls it, the preset geometry, the
/// name Excel gives a new one).
const SHAPES: &[(&str, &str, &str)] = &[
    ("rect", "rect", "正方形/長方形"),
    ("roundRect", "roundRect", "四角形: 角を丸くする"),
    ("ellipse", "ellipse", "楕円"),
    ("triangle", "triangle", "二等辺三角形"),
    ("rtTriangle", "rtTriangle", "直角三角形"),
    ("diamond", "diamond", "ひし形"),
    ("parallelogram", "parallelogram", "平行四辺形"),
    ("trapezoid", "trapezoid", "台形"),
    ("pentagon", "pentagon", "五角形"),
    ("hexagon", "hexagon", "六角形"),
    ("rightArrow", "rightArrow", "矢印: 右"),
    ("leftArrow", "leftArrow", "矢印: 左"),
    ("upArrow", "upArrow", "矢印: 上"),
    ("downArrow", "downArrow", "矢印: 下"),
    ("star5", "star5", "星: 5 pt"),
    ("wedgeRectCallout", "wedgeRectCallout", "吹き出し: 四角形"),
    (
        "wedgeRoundRectCallout",
        "wedgeRoundRectCallout",
        "吹き出し: 角を丸めた四角形",
    ),
    ("textBox", "rect", "テキスト ボックス"),
    ("line", "line", "直線コネクタ"),
    ("arrow", "straightConnector1", "直線矢印コネクタ"),
];

/// The prefix of DrawingML's main namespace in the drawing (`a:`), and the declaration to write
/// on an element when the root has none.
fn dml_prefix(doc: &str, root: &El) -> (String, String) {
    for (name, value) in scan::attrs(doc, root.open.clone()) {
        if let Some(prefix) = name.strip_prefix("xmlns:")
            && &doc[value] == DML
        {
            return (format!("{prefix}:"), String::new());
        }
    }
    ("a:".into(), format!(" xmlns:a=\"{DML}\""))
}

/// The drawing of the sheet; made when the sheet has none.
fn ensure_drawing(pkg: &mut Pkg, sheet: &Sheet) -> Result<String> {
    if let Some(part) = super::drawing_part(pkg, sheet)? {
        return Ok(part);
    }
    let part = pkg.free("xl/drawings/drawing", ".xml");
    pkg.put(
        &part,
        format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<xdr:wsDr xmlns:xdr=\"{XDR}\" xmlns:a=\"{DML}\"></xdr:wsDr>").into_bytes(),
    );
    pkg.content_type(Some(&part), "xml", DRAWING_TYPE)?;
    let id = pkg.relate(&sheet.part, &format!("{REL}/drawing"), &part)?;
    pkg.change(&sheet.part, &["sheetData"], |doc, root, edits| {
        let (r, declare) = rel_prefix(doc, root);
        let p = scan::prefix(root.name);
        match root.child("drawing") {
            Some(el) => scan::set_attr(doc, el.open.clone(), &format!("{r}id"), &id, edits),
            None => sheet_insert(
                root,
                "drawing",
                format!("<{p}drawing{declare} {r}id=\"{id}\"/>"),
                edits,
            ),
        }
        Ok(())
    })?;
    Ok(part)
}

fn marker(p: &str, tag: &str, m: super::Marker) -> String {
    format!(
        "<{p}{tag}><{p}col>{}</{p}col><{p}colOff>{}</{p}colOff><{p}row>{}</{p}row><{p}rowOff>{}</{p}rowOff></{p}{tag}>",
        m.0,
        emu(m.1),
        m.2,
        emu(m.3)
    )
}

/// Adds a shape; its name.
pub(super) fn add(pkg: &mut Pkg, sheet: &Sheet, shape: &str, rect: &Rect) -> Result<String> {
    let known = SHAPES.iter().find(|s| s.0.eq_ignore_ascii_case(shape));
    let (kind, prst, base) = match known {
        Some(&(kind, prst, base)) => (kind, prst, base),
        None if crate::view::geom::preset(shape, None, 100.0, 100.0).is_some() => {
            (shape, shape, shape)
        }
        None => {
            return Err(format!(
                "図形 `{shape}` はありません（{}、または既定の図形の名前）",
                SHAPES.iter().map(|s| s.0).collect::<Vec<_>>().join(" / ")
            ));
        }
    };
    let taken = super::names(pkg, sheet)?;
    let part = ensure_drawing(pkg, sheet)?;
    let mut name = String::new();
    pkg.change(&part, &[], |doc, root, edits| {
        let p = scan::prefix(root.name);
        let (a, declare) = dml_prefix(doc, root);
        let id = scan::tags(doc, 0)
            .filter(|t| !t.is_close() && t.local() == "cNvPr")
            .filter_map(|t| scan::attr(doc, t.start..t.end, "id")?.parse::<u32>().ok())
            .max()
            .unwrap_or(1)
            + 1;
        name = super::fresh_name(&taken, base, id - 1, true);
        let xfrm = format!(
            "<{a}xfrm><{a}off x=\"{}\" y=\"{}\"/><{a}ext cx=\"{}\" cy=\"{}\"/></{a}xfrm>",
            emu(rect.x),
            emu(rect.y),
            emu(rect.w),
            emu(rect.h)
        );
        let geom = format!("<{a}prstGeom prst=\"{prst}\"><{a}avLst/></{a}prstGeom>");
        let style = |line: (&str, &str), fill: &str, font: &str| {
            format!(
                "<{p}style><{a}lnRef idx=\"{}\"><{a}schemeClr val=\"accent1\">{}</{a}schemeClr></{a}lnRef><{a}fillRef idx=\"{fill}\"><{a}schemeClr val=\"accent1\"/></{a}fillRef><{a}effectRef idx=\"0\"><{a}schemeClr val=\"accent1\"/></{a}effectRef><{a}fontRef idx=\"minor\"><{a}schemeClr val=\"{font}\"/></{a}fontRef></{p}style>",
                line.0,
                if line.1.is_empty() {
                    String::new()
                } else {
                    format!("<{a}shade val=\"{}\"/>", line.1)
                }
            )
        };
        let text = |wrap: &str, align: &str| {
            format!(
                "<{p}txBody><{a}bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\"{wrap} rtlCol=\"0\" anchor=\"t\"/><{a}lstStyle/><{a}p><{a}pPr algn=\"{align}\"/><{a}endParaRPr kumimoji=\"1\" lang=\"ja-JP\" altLang=\"en-US\" sz=\"1100\"/></{a}p></{p}txBody>"
            )
        };
        let pr = format!("<{p}cNvPr id=\"{id}\" name=\"{}\"/>", scan::escape(&name));
        let object = match kind {
            "line" | "arrow" => format!(
                "<{p}cxnSp macro=\"\"><{p}nvCxnSpPr>{pr}<{p}cNvCxnSpPr/></{p}nvCxnSpPr><{p}spPr>{xfrm}{geom}{}</{p}spPr>{}</{p}cxnSp>",
                if kind == "arrow" {
                    format!("<{a}ln><{a}tailEnd type=\"triangle\"/></{a}ln>")
                } else {
                    String::new()
                },
                style(("1", ""), "0", "tx1")
            ),
            "textBox" => format!(
                "<{p}sp macro=\"\" textlink=\"\"><{p}nvSpPr>{pr}<{p}cNvSpPr txBox=\"1\"/></{p}nvSpPr><{p}spPr>{xfrm}{geom}<{a}solidFill><{a}schemeClr val=\"lt1\"/></{a}solidFill><{a}ln w=\"9525\" cmpd=\"sng\"><{a}solidFill><{a}schemeClr val=\"lt1\"><{a}shade val=\"50000\"/></{a}schemeClr></{a}solidFill></{a}ln></{p}spPr>{}{}</{p}sp>",
                style(("0", ""), "0", "dk1"),
                text(" wrap=\"square\"", "l")
            ),
            _ => format!(
                "<{p}sp macro=\"\" textlink=\"\"><{p}nvSpPr>{pr}<{p}cNvSpPr/></{p}nvSpPr><{p}spPr>{xfrm}{geom}</{p}spPr>{}{}</{p}sp>",
                style(("2", "50000"), "1", "lt1"),
                text("", "ctr")
            ),
        };
        let anchor = format!(
            "<{p}twoCellAnchor{declare}>{}{}{object}<{p}clientData/></{p}twoCellAnchor>",
            marker(p, "from", rect.from),
            marker(p, "to", rect.to)
        );
        append(doc, root, anchor, edits);
        Ok(())
    })?;
    Ok(name)
}

/// Puts `xml` at the end of what an element holds.
pub(super) fn append(doc: &str, el: &El, xml: String, edits: &mut Edits) {
    if el.end == el.open.end {
        let open = doc[el.open.clone()].trim_end_matches("/>").trim_end();
        edits.replace(el.span(), format!("{open}>{xml}</{}>", el.name));
    } else {
        edits.insert(el.inner.end, xml);
    }
}

/// Puts `xml` into an element before the first of its children named in `before`, or at its
/// end.
fn insert_before(doc: &str, el: &El, before: &[&str], xml: String, edits: &mut Edits) {
    match el.children.iter().find(|c| before.contains(&c.local())) {
        Some(next) => edits.insert(next.open.start, xml),
        None => append(doc, el, xml, edits),
    }
}

const FILLS: &[&str] = &[
    "noFill",
    "solidFill",
    "gradFill",
    "blipFill",
    "pattFill",
    "grpFill",
];

fn fill_xml(a: &str, colour: Option<&str>) -> String {
    match colour {
        Some(rgb) => format!("<{a}solidFill><{a}srgbClr val=\"{rgb}\"/></{a}solidFill>"),
        None => format!("<{a}noFill/>"),
    }
}

/// The fill of an element that takes one among its children (a shape's properties, a line,
/// the properties of a run of text).
fn set_fill(doc: &str, el: &El, a: &str, colour: Option<&str>, after: &[&str], edits: &mut Edits) {
    let xml = fill_xml(a, colour);
    if let Some(old) = el.children.iter().find(|c| FILLS.contains(&c.local())) {
        edits.replace(old.span(), xml);
        return;
    }
    // After what comes before a fill, where there is any of it.
    match el
        .children
        .iter()
        .rev()
        .find(|c| after.contains(&c.local()))
    {
        Some(last) => edits.insert(last.end, xml),
        None if el.end == el.open.end => append(doc, el, xml, edits),
        None => edits.insert(el.inner.start, xml),
    }
}

/// The paragraphs of a text, written anew with the format of the first one.
fn set_text(doc: &str, obj: &El, a: &str, text: &str, edits: &mut Edits) -> Result<()> {
    if !matches!(obj.local(), "sp") {
        return Err(
            "このオブジェクトには文字を設定できません（図形・テキスト ボックスに設定できます）"
                .into(),
        );
    }
    let p = scan::prefix(obj.name);
    const RUN: &str = " kumimoji=\"1\" lang=\"ja-JP\" altLang=\"en-US\" sz=\"1100\"";
    let body = obj.child("txBody");
    let first = body.and_then(|b| b.child("p"));
    let para = first
        .and_then(|f| f.child("pPr"))
        .map_or(String::new(), |pr| doc[pr.span()].to_string());
    // The format of the first run; of the end of the paragraph when it has no run.
    let run = first.and_then(|f| {
        f.children
            .iter()
            .find(|c| c.local() == "r")
            .and_then(|r| r.child("rPr"))
            .or_else(|| f.child("endParaRPr"))
    });
    let props = |tag: &str| match run {
        Some(r) => {
            let attrs = super::super::attrs_without(doc, r.open.clone(), &[]);
            if r.end == r.open.end {
                format!("<{a}{tag}{attrs}/>")
            } else {
                format!("<{a}{tag}{attrs}>{}</{a}{tag}>", &doc[r.inner.clone()])
            }
        }
        None => format!("<{a}{tag}{RUN}/>"),
    };
    let mut paragraphs = String::new();
    for line in text.replace("\r\n", "\n").split('\n') {
        if line.is_empty() {
            paragraphs.push_str(&format!("<{a}p>{para}{}</{a}p>", props("endParaRPr")));
        } else {
            paragraphs.push_str(&format!(
                "<{a}p>{para}<{a}r>{}<{a}t>{}</{a}t></{a}r></{a}p>",
                props("rPr"),
                scan::escape(line)
            ));
        }
    }
    match body {
        Some(b) => {
            let all: Vec<&El> = b.children.iter().filter(|c| c.local() == "p").collect();
            match (all.first(), all.last()) {
                (Some(first), Some(last)) => {
                    edits.replace(first.open.start..last.end, paragraphs);
                }
                _ => append(doc, b, paragraphs, edits),
            }
        }
        None => append(
            doc,
            obj,
            format!(
                "<{p}txBody><{a}bodyPr vertOverflow=\"clip\" horzOverflow=\"clip\" rtlCol=\"0\" anchor=\"t\"/><{a}lstStyle/>{paragraphs}</{p}txBody>"
            ),
            edits,
        ),
    }
    Ok(())
}

/// The properties of the runs of an object's text, and of the ends of its paragraphs.
fn runs<'a, 'd>(obj: &'a El<'d>) -> Vec<&'a El<'d>> {
    let mut out = Vec::new();
    if let Some(body) = obj.child("txBody") {
        super::find_all(body, "rPr", &mut out);
        super::find_all(body, "endParaRPr", &mut out);
    }
    out
}

/// Sets one thing of an object of the drawing.
#[allow(clippy::too_many_arguments)]
fn apply(
    doc: &str,
    anchor: &El,
    obj: &El,
    a: &str,
    key: &str,
    value: &Json,
    edits: &mut Edits,
) -> Result<()> {
    let props = obj.child("spPr");
    let no_shape = || {
        format!(
            "`{key}` は、このオブジェクトには設定できません（図形・テキスト ボックス・線に設定できます）"
        )
    };
    match key {
        "name" => {
            let name = text_of(key, value)?;
            if name.trim().is_empty() {
                return Err("名前が空です".into());
            }
            let pr = nv_pr(obj).ok_or("オブジェクトの名前を読み取れません")?;
            scan::set_attr(doc, pr.open.clone(), "name", name.trim(), edits);
        }
        "macro" => {
            let name = macro_ref(&text_of(key, value)?);
            scan::set_attr(doc, obj.open.clone(), "macro", &name, edits);
        }
        "text" => set_text(doc, obj, a, &text_of(key, value)?, edits)?,
        "fill" => {
            let props = props.filter(|_| obj.local() == "sp").ok_or_else(no_shape)?;
            let colour = colour_of(key, value)?;
            set_fill(
                doc,
                props,
                a,
                colour.as_deref(),
                &["xfrm", "custGeom", "prstGeom"],
                edits,
            );
        }
        "line" | "lineWidth" => {
            let props = props
                .filter(|_| matches!(obj.local(), "sp" | "cxnSp" | "pic"))
                .ok_or_else(no_shape)?;
            let line = props.child("ln");
            const AFTER: &[&str] = &["effectLst", "effectDag", "scene3d", "sp3d", "extLst"];
            if key == "line" {
                let colour = colour_of(key, value)?;
                match line {
                    Some(ln) => set_fill(doc, ln, a, colour.as_deref(), &[], edits),
                    None => insert_before(
                        doc,
                        props,
                        AFTER,
                        format!("<{a}ln>{}</{a}ln>", fill_xml(a, colour.as_deref())),
                        edits,
                    ),
                }
            } else {
                let pt = value
                    .as_f64()
                    .filter(|w| (0.0..=1584.0).contains(w))
                    .ok_or_else(|| wrong(key, value))?;
                let w = ((pt * 12700.0).round() as i64).to_string();
                match line {
                    Some(ln) => scan::set_attr(doc, ln.open.clone(), "w", &w, edits),
                    None => insert_before(doc, props, AFTER, format!("<{a}ln w=\"{w}\"/>"), edits),
                }
            }
        }
        "color" | "fontSize" | "bold" | "italic" => {
            let list = runs(obj);
            if list.is_empty() {
                return Err(format!(
                    "`{key}` は、文字のないオブジェクトには設定できません（先に `text` を設定してください）"
                ));
            }
            for run in list {
                match key {
                    "color" => {
                        let colour = colour_of(key, value)?.ok_or_else(|| wrong(key, value))?;
                        set_fill(doc, run, a, Some(&colour), &["ln"], edits);
                    }
                    "fontSize" => {
                        let pt = value
                            .as_f64()
                            .filter(|s| (1.0..=4000.0).contains(s))
                            .ok_or_else(|| wrong(key, value))?;
                        let sz = ((pt * 100.0).round() as i64).to_string();
                        scan::set_attr(doc, run.open.clone(), "sz", &sz, edits);
                    }
                    _ => {
                        let on = if flag_of(key, value)? { "1" } else { "0" };
                        let attr = if key == "bold" { "b" } else { "i" };
                        scan::set_attr(doc, run.open.clone(), attr, on, edits);
                    }
                }
            }
        }
        "align" => {
            let algn = match value.as_str() {
                Some("left") => "l",
                Some("center") => "ctr",
                Some("right") => "r",
                _ => return Err(wrong(key, value)),
            };
            let body = obj.child("txBody").ok_or_else(|| {
                format!("`{key}` は、文字のないオブジェクトには設定できません（先に `text` を設定してください）")
            })?;
            for p in body.children.iter().filter(|c| c.local() == "p") {
                match p.child("pPr") {
                    Some(pr) => scan::set_attr(doc, pr.open.clone(), "algn", algn, edits),
                    None if p.end == p.open.end => {
                        append(doc, p, format!("<{a}pPr algn=\"{algn}\"/>"), edits)
                    }
                    None => edits.insert(p.inner.start, format!("<{a}pPr algn=\"{algn}\"/>")),
                }
            }
        }
        "valign" => {
            let to = match value.as_str() {
                Some("top") => "t",
                Some("middle") => "ctr",
                Some("bottom") => "b",
                _ => return Err(wrong(key, value)),
            };
            let pr = obj
                .child("txBody")
                .and_then(|b| b.child("bodyPr"))
                .ok_or_else(|| {
                    format!("`{key}` は、文字のないオブジェクトには設定できません（先に `text` を設定してください）")
                })?;
            scan::set_attr(doc, pr.open.clone(), "anchor", to, edits);
        }
        "placement" => {
            if anchor.local() != "twoCellAnchor" {
                return Err(
                    "このオブジェクトは、セルに合わせた移動やサイズ変更を設定できません".into(),
                );
            }
            let mode = match placement_of(key, value)? {
                (true, true) => "twoCell",
                (true, false) => "oneCell",
                _ => "absolute",
            };
            scan::set_attr(doc, anchor.open.clone(), "editAs", mode, edits);
        }
        _ => {
            return Err(format!(
                "図形の設定 `{key}` はありません（name / text / macro / fill / line / lineWidth / color / fontSize / bold / italic / align / valign / placement）"
            ));
        }
    }
    Ok(())
}

/// Sets what `settings` says of the object of the drawing at `index` ([`items`]), one after
/// the other: each is made in the drawing as the one before left it.
pub(super) fn set(
    pkg: &mut Pkg,
    _sheet: &Sheet,
    part: &str,
    index: usize,
    settings: &Settings,
) -> Result<()> {
    for (key, value) in settings {
        pkg.change(part, &[], |doc, root, edits| {
            let (a, _) = dml_prefix(doc, root);
            let list = items(doc, root);
            let item = list.get(index).ok_or("オブジェクトが見つかりません")?;
            for anchor in &item.anchors {
                if let Some(obj) = object_of(anchor) {
                    apply(doc, anchor, obj, &a, key, value, edits)?;
                }
            }
            Ok(())
        })?;
    }
    Ok(())
}
