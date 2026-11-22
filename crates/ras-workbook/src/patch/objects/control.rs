//! The controls of a sheet: form controls (the shapes of the VML drawing, with the properties
//! and the mirrors newer Excel adds) and ActiveX controls (the sheet's `<controls>`, the
//! persisted control, its VML shape).

use serde_json::Value as Json;

use super::{
    Ctl, Order, Pkg, REL, Rect, Result, Settings, Sheet, colour_of, find_all, flag_of, int_of,
    macro_ref, placement_of, reference_of, rel_prefix, sheet_insert, short, text_of, vml_part,
    wrong,
};
use crate::oforms::{self, Record};
use crate::patch::scan::{self, Edits, El};
use crate::view::controls::spid_number;

const VML_TYPE: &str = "application/vnd.openxmlformats-officedocument.vmlDrawing";
const AX_XML_TYPE: &str = "application/vnd.ms-office.activeX+xml";
const AX_BIN_TYPE: &str = "application/vnd.ms-office.activeX";
const AX_BIN_REL: &str =
    "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary";
const AX_NS: &str = "http://schemas.microsoft.com/office/2006/activeX";

/// The shape type of every control.
const SHAPE_TYPE: &str = "<v:shapetype id=\"_x0000_t201\" coordsize=\"21600,21600\" o:spt=\"201\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/><v:path shadowok=\"f\" o:extrusionok=\"f\" strokeok=\"f\" fillok=\"f\" o:connecttype=\"rect\"/><o:lock v:ext=\"edit\" shapetype=\"t\"/></v:shapetype>";

/// The form controls: (what the edit calls it, the object type of its VML shape, the name
/// Excel gives a new one, it has a text).
const FORMS: &[(&str, &str, &str, bool)] = &[
    ("button", "Button", "ボタン", true),
    ("checkBox", "Checkbox", "チェック", true),
    ("optionButton", "Radio", "オプション", true),
    ("groupBox", "GBox", "グループ", true),
    ("label", "Label", "ラベル", true),
    ("listBox", "List", "リスト", false),
    ("comboBox", "Drop", "ドロップ", false),
    ("scrollBar", "Scroll", "スクロール バー", false),
    ("spinner", "Spin", "スピン", false),
];

/// The name Excel gives a form control of an object type without a name of its own.
pub(crate) fn form_base(object: &str) -> &'static str {
    FORMS
        .iter()
        .find(|f| f.1 == object)
        .map(|f| f.2)
        .unwrap_or(match object {
            "Edit" => "エディット",
            "Pict" => "オブジェクト",
            _ => "コントロール",
        })
}

// ---------------------------------------------------------------------------- VML shapes

/// The shape id of a VML shape (1025 of `_x0000_s1025`).
pub(super) fn shape_id(doc: &str, sh: &El) -> Option<u32> {
    scan::attr(doc, sh.open.clone(), "spid")
        .or_else(|| scan::attr(doc, sh.open.clone(), "id"))
        .and_then(|s| spid_number(&s))
}

/// The name of a VML shape that has one of its own: its `id` when it is not its shape id.
pub(super) fn shape_name(doc: &str, sh: &El) -> Option<String> {
    let id = scan::attrs(doc, sh.open.clone())
        .into_iter()
        .find(|(n, _)| *n == "id")
        .map(|(_, v)| scan::unescape(&doc[v]))?;
    (!id.starts_with("_x0000_s") && !id.is_empty()).then(|| decode_name(&id))
}

/// `Button_x0020_1` → `Button 1`.
pub(crate) fn decode_name(id: &str) -> String {
    let mut out = String::new();
    let mut rest = id;
    while let Some(i) = rest.find("_x") {
        let code = rest.get(i + 2..i + 6);
        let ch = code
            .filter(|_| rest[i + 6..].starts_with('_'))
            .and_then(|c| u32::from_str_radix(c, 16).ok())
            .and_then(char::from_u32);
        match ch {
            Some(c) => {
                out.push_str(&rest[..i]);
                out.push(c);
                rest = &rest[i + 7..];
            }
            None => {
                out.push_str(&rest[..i + 2]);
                rest = &rest[i + 2..];
            }
        }
    }
    out.push_str(rest);
    out
}

/// A name as the `id` of a VML shape: what an id cannot hold is written `_xHHHH_`.
fn encode_name(name: &str) -> String {
    let mut out = String::new();
    for (i, c) in name.chars().enumerate() {
        let plain = c.is_alphanumeric() && !(i == 0 && c.is_ascii_digit()) || c == '_' && i > 0;
        if plain || (i > 0 && matches!(c, '-' | '.')) {
            out.push(c);
        } else {
            let mut units = [0u16; 2];
            for u in c.encode_utf16(&mut units) {
                out.push_str(&format!("_x{u:04X}_"));
            }
        }
    }
    out
}

fn shape_of<'a, 'd>(doc: &str, root: &'a El<'d>, spid: u32) -> Option<&'a El<'d>> {
    let mut shapes = Vec::new();
    find_all(root, "shape", &mut shapes);
    shapes
        .into_iter()
        .find(|sh| shape_id(doc, sh) == Some(spid))
}

/// The style of a VML shape with some of its properties given other values.
fn restyle(style: &str, set: &[(&str, String)]) -> String {
    let mut props: Vec<(String, String)> = style
        .split(';')
        .filter_map(|p| {
            let (k, v) = p.split_once(':')?;
            Some((k.trim().to_string(), v.trim().to_string()))
        })
        .collect();
    for (key, value) in set {
        match props.iter_mut().find(|(k, _)| k == key) {
            Some((_, v)) => *v = value.clone(),
            None => props.push((key.to_string(), value.clone())),
        }
    }
    props
        .iter()
        .map(|(k, v)| format!("{k}:{v}"))
        .collect::<Vec<_>>()
        .join(";")
}

fn box_style(rect: &Rect) -> Vec<(&'static str, String)> {
    let pt = |px: f64| format!("{}pt", short(px * 0.75));
    vec![
        ("margin-left", pt(rect.x)),
        ("margin-top", pt(rect.y)),
        ("width", pt(rect.w)),
        ("height", pt(rect.h)),
    ]
}

/// `x:Anchor`: columns and rows with the pixels into them.
fn vml_anchor(rect: &Rect) -> String {
    let (f, t) = (rect.from, rect.to);
    [
        f.0 as f64, f.1, f.2 as f64, f.3, t.0 as f64, t.1, t.2 as f64, t.3,
    ]
    .iter()
    .map(|v| (v.round() as i64).to_string())
    .collect::<Vec<_>>()
    .join(", ")
}

/// Sets, adds or takes away a value of a shape's `x:ClientData` (`Some("")`: a flag).
fn client(doc: &str, data: &El, name: &str, value: Option<&str>, edits: &mut Edits) {
    let p = scan::prefix(data.name);
    let old = data.child(name);
    let Some(value) = value else {
        if let Some(old) = old {
            edits.replace(old.span(), "");
        }
        return;
    };
    let xml = if value.is_empty() {
        format!("<{p}{name}/>")
    } else {
        format!("<{p}{name}>{}</{p}{name}>", scan::escape(value))
    };
    match old {
        Some(old) => edits.replace(old.span(), xml),
        None => super::shape::append(doc, data, xml, edits),
    }
}

/// Takes an attribute away from a start tag.
fn drop_attr(doc: &str, el: &El, name: &str, edits: &mut Edits) {
    for (n, value) in scan::attrs(doc, el.open.clone()) {
        if n == name || scan::local(n) == name && !n.starts_with("xmlns") {
            let start = doc[..value.start].rfind(n).unwrap_or(value.start);
            let start = if doc[..start].ends_with(' ') {
                start - 1
            } else {
                start
            };
            edits.replace(start..value.end + 1, "");
            return;
        }
    }
}

fn set_or_drop(doc: &str, el: &El, name: &str, value: Option<&str>, edits: &mut Edits) {
    match value {
        Some(v) => scan::set_attr(doc, el.open.clone(), name, v, edits),
        None => drop_attr(doc, el, name, edits),
    }
}

/// The entries of the control in the sheet's `<controls>` (for newer Excel and for older).
fn entries<'a, 'd>(doc: &str, root: &'a El<'d>, spid: u32) -> Vec<&'a El<'d>> {
    let mut all = Vec::new();
    find_all(root, "control", &mut all);
    all.retain(|c| {
        c.attr(doc, "shapeId")
            .and_then(|s| spid_number(&s))
            .is_some_and(|id| id == spid)
    });
    all
}

// ---------------------------------------------------------------------------- place, delete, order

pub(super) fn place(pkg: &mut Pkg, sheet: &Sheet, ctl: &Ctl, rect: &Rect) -> Result<()> {
    if let Some(part) = vml_part(pkg, sheet)? {
        pkg.change(&part, &[], |doc, root, edits| {
            let Some(sh) = shape_of(doc, root, ctl.spid) else {
                return Ok(());
            };
            if let Some(span) = scan::attr_span(doc, sh.open.clone(), "style") {
                let style = scan::unescape(&doc[span.clone()]);
                edits.replace(span, scan::escape(&restyle(&style, &box_style(rect))));
            }
            if let Some(data) = sh.child("ClientData") {
                client(doc, data, "Anchor", Some(&vml_anchor(rect)), edits);
            }
            Ok(())
        })?;
    }
    pkg.change(&sheet.part, &["sheetData"], |doc, root, edits| {
        for entry in entries(doc, root, ctl.spid) {
            if let Some(anchor) = super::find(entry, "anchor") {
                for (name, m) in [("from", rect.from), ("to", rect.to)] {
                    if let Some(el) = anchor.child(name) {
                        super::set_marker(doc, el, m, edits);
                    }
                }
            }
        }
        Ok(())
    })?;
    // An ActiveX control keeps its size itself.
    if ctl.activex()
        && let Ok(Some(mut ax)) = ActiveX::read(pkg, ctl)
        && let Some(record) = &mut ax.record
    {
        record.set_size(himetric(rect));
        ax.write(pkg)?;
    }
    Ok(())
}

fn himetric(rect: &Rect) -> (i32, i32) {
    let h = |px: f64| (px * 2540.0 / 96.0).round() as i32;
    (h(rect.w), h(rect.h))
}

pub(super) fn delete(pkg: &mut Pkg, sheet: &Sheet, ctl: &Ctl) -> Result<()> {
    if let Some(part) = vml_part(pkg, sheet)? {
        pkg.change(&part, &[], |doc, root, edits| {
            if let Some(sh) = shape_of(doc, root, ctl.spid) {
                edits.replace(sh.span(), "");
            }
            Ok(())
        })?;
    }
    // Its entries, each with what wraps it alone.
    fn units<'a, 'd>(doc: &str, el: &'a El<'d>, spid: u32, out: &mut Vec<&'a El<'d>>) {
        for c in &el.children {
            let mut inside = Vec::new();
            find_all(c, "control", &mut inside);
            let is = |x: &El| {
                x.attr(doc, "shapeId")
                    .and_then(|s| spid_number(&s))
                    .is_some_and(|id| id == spid)
            };
            let wraps = c.local() == "AlternateContent"
                && !inside.is_empty()
                && inside.iter().all(|x| is(x))
                && super::find(c, "controls").is_none();
            if wraps || (c.local() == "control" && is(c)) {
                out.push(c);
            } else {
                units(doc, c, spid, out);
            }
        }
    }
    pkg.change(&sheet.part, &["sheetData"], |doc, root, edits| {
        let mut list = Vec::new();
        units(doc, root, ctl.spid, &mut list);
        for el in list {
            edits.replace(el.span(), "");
        }
        Ok(())
    })?;
    // `<controls>` without a control is not valid: it goes with its last one.
    pkg.change(&sheet.part, &["sheetData"], |_, root, edits| {
        for c in &root.children {
            let holds = c.local() == "controls" || super::find(c, "controls").is_some();
            let mut inside = Vec::new();
            find_all(c, "control", &mut inside);
            if holds && inside.is_empty() {
                edits.replace(c.span(), "");
            }
        }
        Ok(())
    })?;
    if let Some(rel) = &ctl.rel {
        pkg.unrelate(&sheet.part, &rel.id)?;
        for inner in pkg.rels(&rel.part)? {
            if inner.kind == AX_BIN_REL {
                pkg.discard(&inner.part)?;
            }
        }
        pkg.discard(&rel.part)?;
    }
    Ok(())
}

/// A control without a mirror in the drawing: among the shapes of the VML drawing.
pub(super) fn order(pkg: &mut Pkg, sheet: &Sheet, ctl: &Ctl, to: Order) -> Result<()> {
    let Some(part) = vml_part(pkg, sheet)? else {
        return Ok(());
    };
    pkg.change(&part, &[], |doc, root, edits| {
        let mut shapes = Vec::new();
        find_all(root, "shape", &mut shapes);
        shapes.retain(|sh| {
            sh.child("ClientData")
                .is_some_and(|d| d.attr(doc, "ObjectType").as_deref() != Some("Note"))
        });
        let Some(at) = shapes
            .iter()
            .position(|sh| shape_id(doc, sh) == Some(ctl.spid))
        else {
            return Ok(());
        };
        let z = |sh: &El| {
            scan::attr(doc, sh.open.clone(), "style").and_then(|s| {
                s.split(';').find_map(|p| {
                    p.trim()
                        .strip_prefix("z-index:")?
                        .trim()
                        .parse::<i64>()
                        .ok()
                })
            })
        };
        let mut levels: Vec<i64> = shapes
            .iter()
            .enumerate()
            .map(|(i, sh)| z(sh).unwrap_or(i as i64 + 1))
            .collect();
        levels.sort();
        let new = super::reorder(shapes.len(), at, to);
        for (slot, from) in new.iter().enumerate() {
            let sh = shapes[*from];
            let text = &doc[sh.span()];
            // The shape at this place, at the level of the place.
            let moved = match scan::attr_span(text, 0..sh.open.len(), "style") {
                Some(span) => {
                    let style = restyle(
                        &scan::unescape(&text[span.clone()]),
                        &[("z-index", levels[slot].to_string())],
                    );
                    format!(
                        "{}{}{}",
                        &text[..span.start],
                        scan::escape(&style),
                        &text[span.end..]
                    )
                }
                None => text.to_string(),
            };
            if moved != doc[shapes[slot].span()] {
                edits.replace(shapes[slot].span(), moved);
            }
        }
        Ok(())
    })
}

// ---------------------------------------------------------------------------- ActiveX data

/// The persisted data of an ActiveX control.
struct ActiveX {
    /// The class as VBA names it (`CommandButton`); empty for a control that is not MSForms'.
    kind: &'static str,
    display: u8,
    bin: Option<String>,
    /// The class id that starts the stream.
    head: Vec<u8>,
    record: Option<Record>,
}

impl ActiveX {
    fn read(pkg: &mut Pkg, ctl: &Ctl) -> Result<Option<ActiveX>> {
        let Some(rel) = ctl.rel.as_ref().filter(|_| ctl.activex()) else {
            return Ok(None);
        };
        let Some(doc) = pkg.text(&rel.part)? else {
            return Ok(None);
        };
        let Some(tag) = scan::tags(&doc, 0).find(|t| !t.is_close() && t.local() == "ocx") else {
            return Ok(None);
        };
        let span = tag.start..tag.end;
        let class = scan::attr(&doc, span.clone(), "classid").unwrap_or_default();
        let (kind, record_name, display) = oforms::class_of(&class).unwrap_or(("", "", 0));
        let mut ax = ActiveX {
            kind,
            display,
            bin: None,
            head: Vec::new(),
            record: None,
        };
        let stream = matches!(
            scan::attr(&doc, span.clone(), "persistence").as_deref(),
            Some("persistStreamInit" | "persistStream")
        );
        let bin = scan::attr(&doc, span, "id")
            .and_then(|id| pkg.target(&rel.part, &id).transpose())
            .transpose()?;
        if let (true, Some(bin)) = (stream, bin)
            && let Some(data) = pkg.bytes(&bin.part)?
        {
            let skip = if data.len() > 16 && data[16..].starts_with(&[0, 2]) {
                16
            } else {
                0
            };
            ax.record = Record::parse(record_name, &data[skip..]);
            ax.head = data[..skip].to_vec();
            ax.bin = Some(bin.part);
            if let Some(r) = &ax.record
                && record_name == "MorphData"
            {
                ax.display = r.get("DisplayStyle").map_or(ax.display, |d| d as u8);
            }
        }
        Ok(Some(ax))
    }

    fn write(&self, pkg: &mut Pkg) -> Result<()> {
        if let (Some(bin), Some(record)) = (&self.bin, &self.record) {
            let mut data = self.head.clone();
            data.extend(record.encode());
            pkg.put(bin, data);
        }
        Ok(())
    }
}

/// `RRGGBB` as an OLE colour (`0x00BBGGRR`).
fn ole(rgb: &str) -> u32 {
    let v = u32::from_str_radix(rgb, 16).unwrap_or(0);
    ((v & 0xff) << 16) | (v & 0xff00) | (v >> 16)
}

/// Sets one thing of the persisted data of an ActiveX control. False: the data holds no such
/// thing.
fn set_activex(ax: &mut ActiveX, key: &str, value: &Json) -> Result<bool> {
    let kind = ax.kind;
    let display = ax.display;
    let Some(r) = &mut ax.record else {
        return Err(
            "この ActiveX コントロールのデータは書き換えられません（Microsoft Forms 2.0 のコントロールで、Excel が保存した形式のものを設定できます）".into(),
        );
    };
    let number = |r: &mut Record, name: &str| -> Result<bool> {
        Ok(r.set(name, int_of(key, value)? as i32 as u32))
    };
    Ok(match key {
        "text" if matches!(kind, "TextBox" | "ComboBox") => {
            r.set_string("Value", &text_of(key, value)?)
        }
        "text" => r.set_string("Caption", &text_of(key, value)?),
        "value" if matches!(kind, "ScrollBar" | "SpinButton") => number(r, "Position")?,
        "value" if matches!(display, 4..=6) => {
            let on = match value {
                Json::Bool(b) => *b,
                Json::Number(n) => n.as_f64() != Some(0.0),
                _ => return Err(wrong(key, value)),
            };
            r.set_string("Value", if on { "1" } else { "0" })
        }
        "value" => r.set_string("Value", &text_of(key, value)?),
        "color" | "fill" => {
            let name = if key == "color" {
                "ForeColor"
            } else {
                "BackColor"
            };
            match colour_of(key, value)? {
                Some(rgb) => r.set(name, ole(&rgb)),
                None => {
                    r.clear(name);
                    r.has(name)
                }
            }
        }
        "min" => number(r, "Min")?,
        "max" => number(r, "Max")?,
        "step" => number(r, "SmallChange")?,
        "page" => number(r, "LargeChange")?,
        "fontName" | "fontSize" | "bold" | "italic" | "align" => {
            let Some(font) = &mut r.text else {
                return Ok(false);
            };
            match key {
                "fontName" => font.set_string("FontName", text_of(key, value)?.trim()),
                "fontSize" => {
                    let pt = value
                        .as_f64()
                        .filter(|s| (1.0..=400.0).contains(s))
                        .ok_or_else(|| wrong(key, value))?;
                    font.set("FontHeight", (pt * 20.0).round() as u32)
                }
                "align" => font.set(
                    "ParagraphAlign",
                    match value.as_str() {
                        Some("left") => 1,
                        Some("right") => 2,
                        Some("center") => 3,
                        _ => return Err(wrong(key, value)),
                    },
                ),
                _ => {
                    let on = flag_of(key, value)?;
                    let bit = if key == "bold" { 1 } else { 2 };
                    let effects = font.get("FontEffects").unwrap_or(0);
                    font.set(
                        "FontEffects",
                        if on { effects | bit } else { effects & !bit },
                    );
                    if key == "bold" {
                        font.set("FontWeight", if on { 700 } else { 400 });
                    }
                    true
                }
            }
        }
        _ => false,
    })
}

// ---------------------------------------------------------------------------- settings

const FORM_KEYS: &str = "name / text / macro / linkedCell / inputRange / checked / value / min / max / step / page / dropLines / threeD / horizontal / fontSize / color / placement";
const ACTIVEX_KEYS: &str = "name / text / value / linkedCell / inputRange / color / fill / fontName / fontSize / bold / italic / align / min / max / step / page / placement";

/// The text of a control's VML shape, written anew with the font it had.
fn vml_text(
    doc: &str,
    sh: &El,
    font: &(String, f64),
    text: Option<&str>,
    size: Option<f64>,
    colour: Option<&str>,
    edits: &mut Edits,
) -> Result<()> {
    let Some(tb) = sh.child("textbox") else {
        return Err("このコントロールには文字がありません".into());
    };
    let div = tb.child("div");
    let old = super::find(tb, "font");
    let (mut face, mut sz, mut color) = (font.0.clone(), font.1 * 20.0, "auto".to_string());
    if let Some(f) = old {
        face = f.attr(doc, "face").unwrap_or(face);
        sz = f
            .attr(doc, "size")
            .and_then(|v| v.parse().ok())
            .unwrap_or(sz);
        color = f.attr(doc, "color").unwrap_or(color);
    }
    if let Some(pt) = size {
        sz = pt * 20.0;
    }
    if let Some(c) = colour {
        color = format!("#{c}");
    }
    let body = match text {
        Some(t) => scan::escape(t),
        None => match old {
            Some(f) => doc[f.inner.clone()].to_string(),
            None => String::new(),
        },
    };
    let style = div
        .and_then(|d| d.attr(doc, "style"))
        .unwrap_or_else(|| "text-align:left".into());
    let xml = format!(
        "<div style='{}'><font face=\"{}\" size=\"{}\" color=\"{}\">{body}</font></div>",
        scan::escape(&style).replace('\'', "&apos;"),
        scan::escape(&face),
        sz.round() as i64,
        scan::escape(&color)
    );
    if tb.end == tb.open.end {
        super::shape::append(doc, tb, xml, edits);
    } else {
        edits.replace(tb.inner.clone(), xml);
    }
    Ok(())
}

/// What one setting changes: in the VML shape's client data, in the properties of the form
/// control, in the entries of the sheet's `<controls>`, in the mirror of the drawing.
#[derive(Default)]
struct Change {
    /// (`x:ClientData` child, its value; None: taken away).
    client: Vec<(&'static str, Option<String>)>,
    /// (attribute of `formControlPr`, its value).
    props: Vec<(&'static str, Option<String>)>,
    /// (attribute of `controlPr`, its value).
    entry: Vec<(&'static str, Option<String>)>,
    /// The name of the entries.
    name: Option<String>,
    /// How the anchor of the entries follows the cells.
    follows: Option<(bool, bool)>,
    text: Option<String>,
    size: Option<f64>,
    colour: Option<String>,
    /// What the mirror is set to (the settings of a shape).
    mirror: Settings,
}

pub(super) fn set(
    pkg: &mut Pkg,
    sheet: &Sheet,
    ctl: &Ctl,
    mirror: Option<usize>,
    settings: &Settings,
) -> Result<()> {
    let activex = ctl.activex();
    let object = ctl.object.as_str();
    let mut ax = if activex {
        ActiveX::read(pkg, ctl)?
    } else {
        None
    };
    for (key, value) in settings {
        let key = key.as_str();
        let mut c = Change::default();
        let some = |s: String| if s.is_empty() { None } else { Some(s) };
        let both = |c: &mut Change, vml: &'static str, prop: &'static str, v: Option<String>| {
            c.client.push((vml, v.clone()));
            c.props.push((prop, v));
        };
        match key {
            "name" => {
                let name = text_of(key, value)?.trim().to_string();
                if name.is_empty() {
                    return Err("名前が空です".into());
                }
                let taken = super::names(pkg, sheet)?;
                if taken.contains(&name.to_lowercase())
                    && name.to_lowercase() != ctl.name.to_lowercase()
                {
                    return Err(format!("名前 `{name}` のオブジェクトが既にあります"));
                }
                c.mirror.push((key.into(), value.clone()));
                c.name = Some(name);
            }
            "placement" => {
                let (moves, sizes) = placement_of(key, value)?;
                // (`<x:MoveWithCells/>` says that the object does not move with the cells.)
                c.client.push(("MoveWithCells", (!moves).then(String::new)));
                c.client.push(("SizeWithCells", (!sizes).then(String::new)));
                c.follows = Some((moves, sizes));
                c.mirror.push((key.into(), value.clone()));
            }
            "linkedCell" => {
                let v = some(reference_of(key, value)?);
                c.entry.push(("linkedCell", v.clone()));
                both(&mut c, "FmlaLink", "fmlaLink", v);
            }
            "inputRange" => {
                let v = some(reference_of(key, value)?);
                c.entry.push(("listFillRange", v.clone()));
                both(&mut c, "FmlaRange", "fmlaRange", v);
            }
            _ if activex => {
                let done = match &mut ax {
                    Some(ax) => set_activex(ax, key, value)?,
                    None => false,
                };
                if !done {
                    return Err(format!(
                        "ActiveX コントロール `{}` の設定 `{key}` はありません（{ACTIVEX_KEYS}。コントロールの種類によって使えるものが違います）",
                        ctl.name
                    ));
                }
            }
            "macro" => {
                let v = some(macro_ref(&text_of(key, value)?));
                c.client.push(("FmlaMacro", v.clone()));
                c.entry.push(("macro", v.clone()));
                c.mirror
                    .push((key.into(), Json::String(v.unwrap_or_default())));
            }
            "text" => {
                c.text = Some(text_of(key, value)?);
                c.mirror.push((key.into(), value.clone()));
            }
            "fontSize" => {
                c.size = Some(
                    value
                        .as_f64()
                        .filter(|s| (1.0..=400.0).contains(s))
                        .ok_or_else(|| wrong(key, value))?,
                );
                c.mirror.push((key.into(), value.clone()));
            }
            "color" => {
                c.colour = Some(colour_of(key, value)?.ok_or_else(|| wrong(key, value))?);
                c.mirror.push((key.into(), value.clone()));
            }
            "checked" if matches!(object, "Checkbox" | "Radio") => {
                let (vml, prop) = match value {
                    Json::Bool(true) => (Some("1"), Some("Checked")),
                    Json::Bool(false) => (None, None),
                    Json::String(s) if s == "mixed" && object == "Checkbox" => {
                        (Some("2"), Some("Mixed"))
                    }
                    _ => return Err(wrong(key, value)),
                };
                c.client.push(("Checked", vml.map(str::to_string)));
                c.props.push(("checked", prop.map(str::to_string)));
            }
            "value" if matches!(object, "List" | "Drop") => {
                let v = Some(int_of(key, value)?.to_string());
                both(&mut c, "Sel", "sel", v);
            }
            "value" | "min" | "max" | "step" | "page"
                if matches!(object, "Scroll" | "Spin" | "List" | "Drop") =>
            {
                let (vml, prop) = match key {
                    "value" => ("Val", "val"),
                    "min" => ("Min", "min"),
                    "max" => ("Max", "max"),
                    "step" => ("Inc", "inc"),
                    _ => ("Page", "page"),
                };
                let n = int_of(key, value)?;
                if !(0..=30000).contains(&n) {
                    return Err(format!(
                        "`{key}` は 0 から 30000 までの整数で指定してください"
                    ));
                }
                both(&mut c, vml, prop, Some(n.to_string()));
            }
            "dropLines" if object == "Drop" => {
                let v = Some(int_of(key, value)?.max(1).to_string());
                both(&mut c, "DropLines", "dropLines", v);
            }
            "threeD" => {
                let flat = !flag_of(key, value)?;
                let (vml, prop) = if matches!(object, "List" | "Drop" | "Edit") {
                    ("NoThreeD2", "noThreeD2")
                } else {
                    ("NoThreeD", "noThreeD")
                };
                c.client.push((vml, flat.then(String::new)));
                c.props.push((prop, flat.then(|| "1".to_string())));
            }
            "horizontal" if matches!(object, "Scroll" | "Spin") => {
                let on = flag_of(key, value)?;
                c.client.push(("Horiz", on.then(String::new)));
                c.props.push(("horiz", on.then(|| "1".to_string())));
            }
            _ => {
                return Err(format!(
                    "コントロール `{}` の設定 `{key}` はありません（{FORM_KEYS}。コントロールの種類によって使えるものが違います）",
                    ctl.name
                ));
            }
        }
        write(pkg, sheet, ctl, mirror, &c)?;
    }
    if let Some(ax) = &ax {
        ax.write(pkg)?;
    }
    Ok(())
}

fn write(pkg: &mut Pkg, sheet: &Sheet, ctl: &Ctl, mirror: Option<usize>, c: &Change) -> Result<()> {
    // The entries of the sheet's `<controls>` say the name; without them, the VML shape does.
    let mut listed = false;
    pkg.change(&sheet.part, &["sheetData"], |doc, root, edits| {
        for entry in entries(doc, root, ctl.spid) {
            listed = true;
            if let Some(name) = &c.name {
                scan::set_attr(doc, entry.open.clone(), "name", name, edits);
            }
            let Some(pr) = entry.child("controlPr") else {
                continue;
            };
            for (attr, value) in &c.entry {
                set_or_drop(doc, pr, attr, value.as_deref(), edits);
            }
            if let (Some((moves, sizes)), Some(anchor)) = (c.follows, pr.child("anchor")) {
                set_or_drop(doc, anchor, "moveWithCells", moves.then_some("1"), edits);
                set_or_drop(doc, anchor, "sizeWithCells", sizes.then_some("1"), edits);
            }
        }
        Ok(())
    })?;
    if let Some(part) = vml_part(pkg, sheet)? {
        let font = sheet.font.clone();
        pkg.change(&part, &[], |doc, root, edits| {
            let Some(sh) = shape_of(doc, root, ctl.spid) else {
                return Ok(());
            };
            if let Some(name) = c
                .name
                .as_ref()
                .filter(|_| !listed || shape_name(doc, sh).is_some())
            {
                let spid = format!("_x0000_s{}", ctl.spid);
                scan::set_attr(doc, sh.open.clone(), "id", &encode_name(name), edits);
                if scan::attr(doc, sh.open.clone(), "spid").is_none() {
                    let o = "o:";
                    scan::set_attr(doc, sh.open.clone(), &format!("{o}spid"), &spid, edits);
                }
            }
            if let Some(data) = sh.child("ClientData") {
                for (name, value) in &c.client {
                    client(doc, data, name, value.as_deref(), edits);
                }
            }
            if c.text.is_some() || c.size.is_some() || c.colour.is_some() {
                vml_text(
                    doc,
                    sh,
                    &font,
                    c.text.as_deref(),
                    c.size,
                    c.colour.as_deref(),
                    edits,
                )?;
            }
            Ok(())
        })?;
    }
    if let Some(rel) = ctl.rel.as_ref().filter(|r| r.kind.ends_with("/ctrlProp"))
        && !c.props.is_empty()
        && pkg.has(&rel.part)
    {
        pkg.change(&rel.part, &[], |doc, root, edits| {
            for (attr, value) in &c.props {
                set_or_drop(doc, root, attr, value.as_deref(), edits);
            }
            Ok(())
        })?;
    }
    if let (Some(index), false) = (mirror, c.mirror.is_empty())
        && let Some(part) = super::drawing_part(pkg, sheet)?
    {
        // (The mirror of an ActiveX control has no text; what it cannot take is left.)
        for one in &c.mirror {
            let _ = super::shape::set(pkg, sheet, &part, index, &vec![one.clone()]);
        }
    }
    Ok(())
}

// ---------------------------------------------------------------------------- adding

/// The VML drawing of the sheet; made when the sheet has none.
fn ensure_vml(pkg: &mut Pkg, sheet: &Sheet) -> Result<String> {
    if let Some(part) = vml_part(pkg, sheet)? {
        return Ok(part);
    }
    // A block of shape ids no other VML drawing of the workbook uses.
    let mut used: Vec<u32> = Vec::new();
    for name in pkg.names() {
        if !(name.starts_with("xl/drawings/") && name.ends_with(".vml")) {
            continue;
        }
        let Some(doc) = pkg.text(&name)? else {
            continue;
        };
        for tag in scan::tags(&doc, 0).filter(|t| !t.is_close()) {
            let span = tag.start..tag.end;
            match tag.local() {
                "idmap" => used.extend(
                    scan::attr(&doc, span, "data")
                        .unwrap_or_default()
                        .split(',')
                        .filter_map(|n| n.trim().parse::<u32>().ok()),
                ),
                "shape" => used.extend(
                    scan::attr(&doc, span.clone(), "spid")
                        .or_else(|| scan::attr(&doc, span, "id"))
                        .and_then(|s| spid_number(&s))
                        .map(|id| id >> 10),
                ),
                _ => {}
            }
        }
    }
    let block = (1..).find(|n| !used.contains(n)).expect("a block is free");
    let part = pkg.free("xl/drawings/vmlDrawing", ".vml");
    pkg.put(
        &part,
        format!("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"{block}\"/></o:shapelayout>{SHAPE_TYPE}</xml>").into_bytes(),
    );
    pkg.content_type(None, "vml", VML_TYPE)?;
    let id = pkg.relate(&sheet.part, &format!("{REL}/vmlDrawing"), &part)?;
    pkg.change(&sheet.part, &["sheetData"], |doc, root, edits| {
        let (r, declare) = rel_prefix(doc, root);
        let p = scan::prefix(root.name);
        sheet_insert(
            root,
            "legacyDrawing",
            format!("<{p}legacyDrawing{declare} {r}id=\"{id}\"/>"),
            edits,
        );
        Ok(())
    })?;
    Ok(part)
}

/// Adds a shape to the VML drawing: `make` writes it from its shape id, its level among the
/// shapes and the declarations of namespaces it needs. The shape id.
fn add_shape(
    pkg: &mut Pkg,
    sheet: &Sheet,
    make: impl FnOnce(&str, u32, usize, &str) -> String,
) -> Result<u32> {
    let part = ensure_vml(pkg, sheet)?;
    let mut spid = 0;
    pkg.change(&part, &[], |doc, root, edits| {
        let mut shapes = Vec::new();
        find_all(root, "shape", &mut shapes);
        let block = super::find(root, "idmap")
            .and_then(|m| m.attr(doc, "data"))
            .and_then(|d| d.split(',').next()?.trim().parse::<u32>().ok())
            .unwrap_or(1);
        spid = shapes
            .iter()
            .filter_map(|sh| shape_id(doc, sh))
            .max()
            .map_or(block * 1024 + 1, |m| m + 1);
        // The prefixes the drawing gives VML's namespaces are those Excel writes, or the shape
        // says them itself.
        let declared = |prefix: &str| {
            scan::attrs(doc, root.open.clone())
                .iter()
                .any(|(n, _)| *n == format!("xmlns:{prefix}"))
        };
        let mut declare = String::new();
        for (prefix, ns) in [
            ("v", "urn:schemas-microsoft-com:vml"),
            ("o", "urn:schemas-microsoft-com:office:office"),
            ("x", "urn:schemas-microsoft-com:office:excel"),
        ] {
            if !declared(prefix) {
                declare.push_str(&format!(" xmlns:{prefix}=\"{ns}\""));
            }
        }
        let mut xml = String::new();
        if !doc.contains("_x0000_t201\"") {
            xml.push_str(&SHAPE_TYPE.replacen(
                "<v:shapetype",
                &format!("<v:shapetype{declare}"),
                1,
            ));
        }
        xml.push_str(&make(doc, spid, shapes.len() + 1, &declare));
        super::shape::append(doc, root, xml, edits);
        Ok(())
    })?;
    Ok(spid)
}

fn shape_style(rect: &Rect, level: usize) -> String {
    let mut props = vec![("position", "absolute".to_string())];
    props.extend(box_style(rect));
    props.push(("z-index", level.to_string()));
    props.push(("mso-wrap-style", "tight".to_string()));
    restyle("", &props)
}

/// Adds a form control; its name.
pub(super) fn add_form(pkg: &mut Pkg, sheet: &Sheet, kind: &str, rect: &Rect) -> Result<String> {
    let &(_, object, base, _) = FORMS
        .iter()
        .find(|f| f.0.eq_ignore_ascii_case(kind))
        .ok_or_else(|| {
            format!(
                "フォーム コントロール `{kind}` はありません（{}）",
                FORMS.iter().map(|f| f.0).collect::<Vec<_>>().join(" / ")
            )
        })?;
    let taken = super::names(pkg, sheet)?;
    let font = sheet.font.clone();
    let mut name = String::new();
    add_shape(pkg, sheet, |doc, spid, level, declare| {
        name = super::fresh_name(&taken, base, spid & 1023, true);
        let text = |align: &str, colour: &str| {
            format!(
                "<v:textbox style='mso-direction-alt:auto' o:singleclick=\"f\"><div style='text-align:{align}'><font face=\"{}\" size=\"{}\" color=\"{colour}\">{}</font></div></v:textbox>",
                scan::escape(&font.0),
                (font.1 * 20.0).round() as i64,
                scan::escape(&name)
            )
        };
        let anchor = format!("<x:Anchor>{}</x:Anchor>", vml_anchor(rect));
        let list = "<x:Val>0</x:Val><x:Min>0</x:Min><x:Max>0</x:Max><x:Inc>1</x:Inc><x:Page>1</x:Page><x:Dx>16</x:Dx><x:Sel>0</x:Sel><x:NoThreeD2/><x:SelType>Single</x:SelType><x:LCT>Normal</x:LCT>";
        let (attrs, body, data) = match object {
            "Button" => (
                " o:button=\"t\" fillcolor=\"buttonFace [67]\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\"",
                format!(
                    "<v:fill color2=\"buttonFace [67]\" o:detectmouseclick=\"t\"/><o:lock v:ext=\"edit\" rotation=\"t\"/>{}",
                    text("center", "#000000")
                ),
                format!(
                    "{anchor}<x:PrintObject>False</x:PrintObject><x:AutoFill>False</x:AutoFill><x:TextHAlign>Center</x:TextHAlign><x:TextVAlign>Center</x:TextVAlign>"
                ),
            ),
            "Checkbox" | "Radio" => {
                // The first option button of a sheet starts its group.
                let first = object == "Radio" && !doc.contains("ObjectType=\"Radio\"");
                (
                    " filled=\"f\" fillcolor=\"window [65]\" stroked=\"f\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\"",
                    format!(
                        "<v:path shadowok=\"t\" strokeok=\"t\" fillok=\"t\"/><o:lock v:ext=\"edit\" rotation=\"t\"/>{}",
                        text("left", "auto")
                    ),
                    format!(
                        "<x:SizeWithCells/>{anchor}<x:AutoFill>False</x:AutoFill><x:AutoLine>False</x:AutoLine><x:TextVAlign>Center</x:TextVAlign><x:NoThreeD/>{}",
                        if first { "<x:FirstButton/>" } else { "" }
                    ),
                )
            }
            "GBox" | "Label" => (
                " filled=\"f\" fillcolor=\"window [65]\" stroked=\"f\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\"",
                format!(
                    "<v:path shadowok=\"t\" strokeok=\"t\" fillok=\"t\"/><o:lock v:ext=\"edit\" rotation=\"t\"/>{}",
                    text("left", "auto")
                ),
                format!(
                    "<x:SizeWithCells/>{anchor}<x:AutoFill>False</x:AutoFill><x:AutoLine>False</x:AutoLine>{}",
                    if object == "GBox" {
                        "<x:NoThreeD/>"
                    } else {
                        ""
                    }
                ),
            ),
            "List" => (
                " o:insetmode=\"auto\"",
                "<o:lock v:ext=\"edit\" rotation=\"t\" text=\"t\"/>".to_string(),
                format!("<x:SizeWithCells/>{anchor}<x:AutoLine>False</x:AutoLine>{list}"),
            ),
            "Drop" => (
                " o:insetmode=\"auto\"",
                "<o:lock v:ext=\"edit\" rotation=\"t\" text=\"t\"/>".to_string(),
                format!(
                    "<x:SizeWithCells/>{anchor}<x:AutoLine>False</x:AutoLine>{list}<x:DropStyle>Combo</x:DropStyle><x:DropLines>8</x:DropLines>"
                ),
            ),
            _ => (
                " o:insetmode=\"auto\"",
                "<o:lock v:ext=\"edit\" rotation=\"t\" text=\"t\"/>".to_string(),
                format!(
                    "<x:SizeWithCells/>{anchor}<x:Val>0</x:Val><x:Min>0</x:Min><x:Max>{}</x:Max><x:Inc>1</x:Inc><x:Page>10</x:Page>{}<x:Dx>16</x:Dx>",
                    if object == "Spin" { 30000 } else { 100 },
                    if rect.w > rect.h { "<x:Horiz/>" } else { "" }
                ),
            ),
        };
        format!(
            "<v:shape{declare} id=\"{}\" o:spid=\"_x0000_s{spid}\" type=\"#_x0000_t201\" style='{}'{attrs}>{body}<x:ClientData ObjectType=\"{object}\">{data}</x:ClientData></v:shape>",
            scan::escape(&encode_name(&name)),
            shape_style(rect, level)
        )
    })?;
    Ok(name)
}

/// The persisted data of a new control of a class, as Excel writes one.
fn fresh(
    kind: &str,
    record: &str,
    display: u8,
    name: &str,
    rect: &Rect,
    font: &(String, f64),
) -> Option<Record> {
    let mut r = Record::new(record)?;
    r.set_size(himetric(rect));
    if record == "MorphData" {
        r.set("DisplayStyle", display as u32);
        match display {
            // A list scrolls both ways; a combo box shows its button and completes what is typed.
            2 => {
                r.set("ScrollBars", 3);
                r.set("MatchEntry", 0);
            }
            3 => {
                r.set("ShowDropButtonWhen", 2);
                r.set("MatchEntry", 1);
            }
            4..=6 => {
                // The colours of a button's face and text; not checked.
                r.set("BackColor", 0x8000_000F);
                r.set("ForeColor", 0x8000_0012);
                r.set_string("Value", "0");
            }
            _ => {}
        }
    }
    if matches!(
        kind,
        "CommandButton" | "Label" | "CheckBox" | "OptionButton" | "ToggleButton"
    ) {
        r.set_string("Caption", name);
    }
    if let Some(text) = &mut r.text {
        text.set_string("FontName", &font.0);
        text.set("FontHeight", (font.1 * 20.0).round() as u32);
        // SHIFTJIS_CHARSET for a Japanese font, else the default one; a variable pitch.
        let japanese = !font.0.is_ascii()
            || ["Yu ", "MS ", "Meiryo"]
                .iter()
                .any(|p| font.0.starts_with(p));
        text.set("FontCharSet", if japanese { 128 } else { 1 });
        text.set("FontPitchAndFamily", 2);
        if matches!(kind, "CommandButton" | "ToggleButton") {
            text.set("ParagraphAlign", 3);
        }
    }
    Some(r)
}

/// Adds an ActiveX control; its name.
pub(super) fn add_activex(
    pkg: &mut Pkg,
    sheet: &Sheet,
    class: &str,
    rect: &Rect,
) -> Result<String> {
    let &(clsid, kind, record, display) = oforms::clsid_of(class).ok_or_else(|| {
        format!(
            "ActiveX コントロール `{class}` はありません（{}）",
            oforms::CLASSES
                .iter()
                .map(|c| c.1)
                .collect::<Vec<_>>()
                .join(" / ")
        )
    })?;
    let taken = super::names(pkg, sheet)?;
    let name = super::fresh_name(&taken, kind, 1, false);
    let data = fresh(kind, record, display, &name, rect, &sheet.font)
        .ok_or("コントロールのデータを作れません")?;
    let mut bin = oforms::clsid_bytes(clsid)
        .ok_or("クラス ID を読み取れません")?
        .to_vec();
    bin.extend(data.encode());
    // The control: its class and its data.
    let xml_part = pkg.free("xl/activeX/activeX", ".xml");
    let bin_part = format!("{}.bin", xml_part.trim_end_matches(".xml"));
    pkg.put(&bin_part, bin);
    let bin_id = pkg.relate(&xml_part, AX_BIN_REL, &bin_part)?;
    pkg.put(
        &xml_part,
        format!("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>\n<ax:ocx ax:classid=\"{{{clsid}}}\" ax:persistence=\"persistStreamInit\" r:id=\"{bin_id}\" xmlns:ax=\"{AX_NS}\" xmlns:r=\"{REL}\"/>").into_bytes(),
    );
    pkg.content_type(Some(&xml_part), "xml", AX_XML_TYPE)?;
    pkg.content_type(Some(&bin_part), "bin", AX_BIN_TYPE)?;
    // Where it is: a shape of the VML drawing.
    let spid = add_shape(pkg, sheet, |_, spid, level, declare| {
        format!(
            "<v:shape{declare} id=\"{}\" o:spid=\"_x0000_s{spid}\" type=\"#_x0000_t201\" style='{}' o:insetmode=\"auto\"><x:ClientData ObjectType=\"Pict\"><x:SizeWithCells/><x:Anchor>{}</x:Anchor><x:CF>Pict</x:CF><x:AutoPict/></x:ClientData></v:shape>",
            scan::escape(&encode_name(&name)),
            shape_style(rect, level),
            vml_anchor(rect)
        )
    })?;
    // The sheet lists it.
    let id = pkg.relate(&sheet.part, &format!("{REL}/control"), &xml_part)?;
    pkg.change(&sheet.part, &["sheetData"], |doc, root, edits| {
        let (r, declare) = rel_prefix(doc, root);
        let mut lists = Vec::new();
        find_all(root, "controls", &mut lists);
        let entry = |p: &str| {
            format!(
                "<{p}control{declare} shapeId=\"{spid}\" {r}id=\"{id}\" name=\"{}\"/>",
                scan::escape(&name)
            )
        };
        if lists.is_empty() {
            let p = scan::prefix(root.name);
            sheet_insert(
                root,
                "controls",
                format!("<{p}controls>{}</{p}controls>", entry(p)),
                edits,
            );
        }
        for list in lists {
            super::shape::append(doc, list, entry(scan::prefix(list.name)), edits);
        }
        Ok(())
    })?;
    Ok(name)
}
