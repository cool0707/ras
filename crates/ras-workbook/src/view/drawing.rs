//! The drawings of a sheet for the grid: shapes, text boxes, lines and connectors, groups,
//! pictures (EMF too: `emf.rs`), WordArt and SmartArt as HTML/SVG in the order of the sheet's
//! drawing, its charts (read in `chart.rs`, drawn by the page), the form and ActiveX controls and
//! the embedded objects (`controls.rs`): everything a macro can be assigned to.
//!
//! Drawn as xljs prints sheets (`src/api/xlsx_shape.rs` there). The parts are read from the
//! workbook file itself: umya-spreadsheet's model leaves out much of what drawings hold. What
//! xljs and RAS share of reading them is umya-spreadsheet's `drawingml` (in the fork): the
//! geometry of shapes, colours, anchors and transforms, the relationships of the parts.

use std::cell::{Cell, RefCell};
use std::collections::HashMap;
use std::fmt::Write as _;
use std::io::{Cursor, Read};

use umya_spreadsheet::drawingml::EMU_PT;
use umya_spreadsheet::drawingml::anchor::{Anchor, Xfrm, anchor_of, xfrm};
use umya_spreadsheet::drawingml::color::{alpha_of, color_of, fill_color, style_ref};
use umya_spreadsheet::drawingml::geometry as geom;
use umya_spreadsheet::drawingml::package::Package;

use super::controls::{self, Controls};
use super::xml::{Node, parse};

/// The largest part read (a picture, a drawing).
const MAX_PART: u64 = 256 << 20;

/// The parts of a workbook file that describe its drawings (shapes, pictures, charts, controls)
/// and how sheets reach them. Empty when the file has none.
#[derive(Default, Debug)]
pub struct Parts {
    files: HashMap<String, Vec<u8>>,
}

fn read_part(zr: &mut zip::ZipArchive<Cursor<&[u8]>>, name: &str) -> Option<Vec<u8>> {
    let f = zr.by_name(name).ok()?;
    if f.size() > MAX_PART {
        return None;
    }
    let mut out = Vec::with_capacity(f.size() as usize);
    f.take(MAX_PART).read_to_end(&mut out).ok()?;
    Some(out)
}

impl Parts {
    /// The parts of the workbook file `data` (the bytes of an .xlsx / .xlsm).
    pub fn read(data: &[u8]) -> Parts {
        let mut files = HashMap::new();
        let Ok(mut zr) = zip::ZipArchive::new(Cursor::new(data)) else {
            return Parts { files };
        };
        let names: Vec<String> = zr.file_names().map(str::to_string).collect();
        for name in ["xl/workbook.xml", "xl/_rels/workbook.xml.rels"] {
            if let Some(b) = read_part(&mut zr, name) {
                files.insert(name.into(), b);
            }
        }
        if !names.iter().any(|n| n.starts_with("xl/drawings/")) {
            return Parts { files };
        }
        for name in &names {
            let keep = name == "xl/workbook.xml"
                || name.ends_with(".rels")
                || name.starts_with("xl/drawings/")
                || name.starts_with("xl/charts/")
                || name.starts_with("xl/diagrams/")
                || name.starts_with("xl/media/")
                || name.starts_with("xl/theme/")
                || name.starts_with("xl/activeX/")
                || name.starts_with("xl/ctrlProps/");
            if keep && let Some(b) = read_part(&mut zr, name) {
                files.insert(name.clone(), b);
            }
        }
        // The controls of the sheets that have any (`<controls>` of the sheet's XML).
        let sheets: Vec<String> = names
            .iter()
            .filter(|n| n.starts_with("xl/worksheets/_rels/"))
            .filter(|n| {
                files.get(*n).is_some_and(|r| {
                    ["/ctrlProp\"", "/control\"", "/oleObject\""]
                        .iter()
                        .any(|k| String::from_utf8_lossy(r).contains(k))
                })
            })
            .map(|n| {
                format!(
                    "xl/worksheets/{}",
                    n.trim_start_matches("xl/worksheets/_rels/")
                        .trim_end_matches(".rels")
                )
            })
            .collect();
        for sheet in sheets {
            let Some(xml) = read_part(&mut zr, &sheet) else {
                continue;
            };
            let xml = String::from_utf8_lossy(&xml);
            if let Some(c) = element_text(&xml, "controls") {
                files.insert(format!("{sheet}#controls"), c.into_bytes());
            }
            if let Some(c) = element_text(&xml, "oleObjects") {
                files.insert(format!("{sheet}#oleObjects"), c.into_bytes());
            }
        }
        Parts { files }
    }

    pub(crate) fn bytes(&self, part: &str) -> Option<&[u8]> {
        self.files.get(part).map(Vec::as_slice)
    }
}

/// The parts as XML (as much as can be read: VML is not always well-formed), their
/// relationships, the sheets' parts and the theme's colours.
impl Package for Parts {
    fn xml(&self, part: &str) -> Option<Node> {
        parse(self.files.get(part)?)
    }
}

/// The text of the first element `name` (with any prefix) of an XML document, to its first end
/// (the elements read so are not nested in their own kind).
fn element_text(xml: &str, name: &str) -> Option<String> {
    let open = xml.match_indices(name).map(|(i, _)| i).find(|&i| {
        let lt = xml[..i].rfind('<').unwrap_or(0);
        let before = &xml[lt + 1..i];
        !before.starts_with('/')
            && before.chars().all(|c| c.is_alphanumeric() || c == ':')
            && xml[i + name.len()..].starts_with(['>', ' ', '/'])
    })?;
    let start = xml[..open].rfind('<')?;
    let close = xml[open..]
        .match_indices(name)
        .map(|(i, _)| open + i)
        .find(|&i| {
            xml[..i]
                .rfind('<')
                .is_some_and(|lt| xml[lt..].starts_with("</"))
                && xml[i + name.len()..].starts_with('>')
        });
    Some(match close {
        Some(c) => xml[start..c + name.len() + 1].to_string(),
        None => xml[start..].to_string(),
    })
}

// ---------------------------------------------------------------------------------------- text

pub(crate) fn esc(s: &str) -> String {
    let mut o = String::with_capacity(s.len());
    for c in s.chars() {
        match c {
            '&' => o.push_str("&amp;"),
            '<' => o.push_str("&lt;"),
            '>' => o.push_str("&gt;"),
            '"' => o.push_str("&quot;"),
            c => o.push(c),
        }
    }
    o
}

/// A quoted CSS string. Besides the quote and the backslash, < > and control characters are
/// escaped, so that the string cannot end the attribute it is in.
fn css_string(s: &str) -> String {
    let mut o = String::with_capacity(s.len() + 2);
    o.push('"');
    for c in s.chars() {
        match c {
            '"' | '\\' | '<' | '>' | '\0'..='\u{1f}' | '\u{7f}' => {
                let _ = write!(o, "\\{:X} ", c as u32);
            }
            c => o.push(c),
        }
    }
    o.push('"');
    o
}

/// A number for CSS (at most 3 decimals).
pub(crate) fn n(v: f64) -> String {
    let s = format!("{v:.3}");
    let s = s.trim_end_matches('0').trim_end_matches('.');
    if s == "-0" { "0".into() } else { s.to_string() }
}

/// A font-family list for an Excel font name (Japanese names also by their English names, which
/// every Windows knows).
pub(crate) fn font_family(name: &str) -> String {
    let alias = match name {
        "游ゴシック" => "Yu Gothic",
        "游ゴシック Light" => "Yu Gothic Light",
        "游ゴシック Medium" => "Yu Gothic Medium",
        "游明朝" => "Yu Mincho",
        "ＭＳ Ｐゴシック" => "MS PGothic",
        "ＭＳ ゴシック" => "MS Gothic",
        "ＭＳ Ｐ明朝" => "MS PMincho",
        "ＭＳ 明朝" => "MS Mincho",
        "メイリオ" => "Meiryo",
        "BIZ UDPゴシック" => "BIZ UDPGothic",
        "BIZ UDゴシック" => "BIZ UDGothic",
        "BIZ UDP明朝" => "BIZ UDPMincho",
        "BIZ UD明朝" => "BIZ UDMincho",
        _ => "",
    };
    let mut out = css_string(name);
    if !alias.is_empty() {
        out.push(',');
        out.push_str(&css_string(alias));
    }
    if name.contains("明朝") || name.contains("Mincho") {
        out.push_str(",\"Yu Mincho\",serif");
    } else {
        out.push_str(",\"Yu Gothic\",Meiryo,sans-serif");
    }
    out
}

// ------------------------------------------------------------------------------------- objects

/// One object of a sheet's drawing.
#[derive(Debug)]
pub struct Obj {
    pub anchor: Anchor,
    /// The HTML of the object, filling a box of the anchor's size.
    pub html: String,
    /// Its name (as the name box of Excel shows it) and the macro it runs when clicked.
    pub name: String,
    pub macro_name: String,
    /// "shape", "picture", "group", "chart", "control".
    pub kind: &'static str,
    /// What the page draws of a chart (`chart::chart_spec`).
    pub chart: Option<serde_json::Value>,
    /// An ActiveX control: its class (`CommandButton`; empty when it is not one of Microsoft
    /// Forms 2.0), whose events the sheet's module receives.
    pub activex: Option<&'static str>,
    /// A macro can be assigned to it (`OnAction`): a shape, a picture, a chart, a form control.
    pub assignable: bool,
    /// A control's shape (`_x0000_s1025` in the VML drawing: 1025).
    pub shape_id: Option<u32>,
}

/// The pictures of a page: each one once (`data:` URLs), the objects refer to them by number
/// (`<img data-img="N">`).
#[derive(Default)]
pub struct Images {
    pub list: Vec<String>,
    by_part: HashMap<String, usize>,
}

impl Images {
    pub(crate) fn add(&mut self, part: &str, data: &[u8], mime: &str) -> usize {
        if let Some(i) = self.by_part.get(part) {
            return *i;
        }
        self.list
            .push(format!("data:{mime};base64,{}", base64(data)));
        self.by_part.insert(part.to_string(), self.list.len() - 1);
        self.list.len() - 1
    }
}

pub(crate) fn base64(data: &[u8]) -> String {
    const T: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut out = String::with_capacity(data.len().div_ceil(3) * 4);
    for c in data.chunks(3) {
        let b = [c[0], *c.get(1).unwrap_or(&0), *c.get(2).unwrap_or(&0)];
        let v = (b[0] as u32) << 16 | (b[1] as u32) << 8 | b[2] as u32;
        out.push(T[(v >> 18) as usize & 63] as char);
        out.push(T[(v >> 12) as usize & 63] as char);
        out.push(if c.len() > 1 {
            T[(v >> 6) as usize & 63] as char
        } else {
            '='
        });
        out.push(if c.len() > 2 {
            T[v as usize & 63] as char
        } else {
            '='
        });
    }
    out
}

/// The type of a picture by its part's name; None for what a browser cannot show (EMF, WMF).
pub(crate) fn image_mime(part: &str) -> Option<&'static str> {
    let ext = part.rsplit('.').next().unwrap_or("").to_ascii_lowercase();
    Some(match ext.as_str() {
        "png" => "image/png",
        "jpg" | "jpeg" | "jpe" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" | "dib" => "image/bmp",
        "svg" => "image/svg+xml",
        "webp" => "image/webp",
        "ico" => "image/x-icon",
        "tif" | "tiff" => "image/tiff",
        _ => return None,
    })
}

/// What the drawing of a sheet needs from the workbook.
pub struct Env<'a> {
    /// The size (pt) of an anchor's box on the sheet.
    pub size: &'a dyn Fn(&Anchor) -> Option<(f64, f64)>,
    /// The texts of the cells of a reference ("$A$1:$A$5", "Sheet2!A1:A3"), for the lists of
    /// controls.
    pub texts: &'a dyn Fn(&str) -> Vec<String>,
    /// The text of a cell, for the controls linked to one.
    pub value: &'a dyn Fn(&str) -> Option<String>,
    /// The cells of one area of a reference (rows of cells), for the series of charts; None
    /// when they cannot be read (another workbook, a sheet that is not there).
    pub range: &'a dyn Fn(&str) -> Option<Vec<Vec<super::chart::Point>>>,
    /// The workbook's default font (name, size): the Normal style's.
    pub font: (String, f64),
}

/// What drawing a shape needs.
pub(crate) struct Ctx<'a> {
    pub parts: &'a Parts,
    pub rels: HashMap<String, (String, String)>,
    pub theme: HashMap<String, String>,
    pub font: (String, f64),
    pub ids: &'a Cell<usize>,
    pub images: &'a RefCell<Images>,
}

const OBJECTS: [&str; 5] = ["sp", "cxnSp", "grpSp", "pic", "graphicFrame"];

/// The objects of a sheet's drawing in drawing order (later ones on top), with its controls.
pub fn sheet_objects(
    parts: &Parts,
    sheet: &str,
    env: &Env,
    ids: &Cell<usize>,
    images: &RefCell<Images>,
) -> Vec<Obj> {
    let mut out = Vec::new();
    let Some(sheet_part) = parts.sheet_part(sheet) else {
        return out;
    };
    let theme = parts.theme_colors();
    let font = env.font.clone();
    let mut controls = Controls::read(parts, &sheet_part, &font);
    let mut drawings: Vec<String> = parts
        .rels(&sheet_part)
        .into_values()
        .filter(|(t, _)| t.ends_with("/drawing"))
        .map(|(_, p)| p)
        .collect();
    drawings.sort();
    for d in drawings {
        let Some(drawing) = parts.xml(&d) else {
            continue;
        };
        let ctx = Ctx {
            parts,
            rels: parts.rels(&d),
            theme: theme.clone(),
            font: font.clone(),
            ids,
            images,
        };
        for a in &drawing.children {
            if a.name.ends_with("Anchor") {
                anchor_object(&ctx, a, env, &mut controls, &mut out);
            } else if a.name == "AlternateContent" {
                // What newer Excel wraps: the mirror of a form or ActiveX control (drawn from the
                // control), a slicer, a 3D model (the fallback, a picture).
                let choice: Vec<&Node> = a
                    .all("Choice")
                    .flat_map(|c| c.children.iter().filter(|x| x.name.ends_with("Anchor")))
                    .collect();
                let mirror = choice.iter().find_map(|x| {
                    x.find("compatExt")
                        .and_then(|c| c.attr("spid"))
                        .and_then(controls::spid_number)
                        .map(|id| (*x, id))
                });
                if let Some((x, id)) = mirror {
                    let sp = x
                        .children
                        .iter()
                        .find(|c| OBJECTS.contains(&c.name.as_str()));
                    if let Some(o) = controls.take(id, Some(anchor_of(x)), sp, env, &ctx) {
                        out.push(o);
                    }
                    continue;
                }
                for x in a
                    .all("Fallback")
                    .flat_map(|f| f.children.iter().filter(|x| x.name.ends_with("Anchor")))
                {
                    anchor_object(&ctx, x, env, &mut controls, &mut out);
                }
            }
        }
    }
    // The controls without a mirror in the drawing (files of Excel 2007, ActiveX controls).
    let ctx = Ctx {
        parts,
        rels: HashMap::new(),
        theme: theme.clone(),
        font: font.clone(),
        ids,
        images,
    };
    out.extend(controls.rest(env, &ctx));
    out
}

fn anchor_object(ctx: &Ctx, a: &Node, env: &Env, controls: &mut Controls, out: &mut Vec<Obj>) {
    let Some(obj) = a
        .children
        .iter()
        .find(|c| OBJECTS.contains(&c.name.as_str()))
        .or_else(|| {
            // Shapes that newer Excel wraps in mc:AlternateContent.
            a.find("Fallback").and_then(|f| {
                f.children
                    .iter()
                    .find(|c| OBJECTS.contains(&c.name.as_str()))
            })
        })
    else {
        return;
    };
    let anchor = anchor_of(a);
    let nv = obj
        .children
        .iter()
        .find(|c| c.name.starts_with("nv"))
        .and_then(|c| c.child("cNvPr"));
    let name = nv.and_then(|c| c.attr("name")).unwrap_or("").to_string();
    // A control's mirror without mc:AlternateContent.
    if let Some(id) = nv
        .and_then(|c| c.find("compatExt"))
        .and_then(|c| c.attr("spid"))
        .and_then(controls::spid_number)
    {
        if let Some(o) = controls.take(id, Some(anchor.clone()), Some(obj), env, ctx) {
            out.push(o);
        }
        return;
    }
    // "[0]!Macro1": a macro of this workbook.
    let macro_name = obj
        .attr("macro")
        .unwrap_or("")
        .trim_start_matches("[0]!")
        .to_string();
    if obj.name == "graphicFrame" {
        let size = (env.size)(&anchor);
        // SmartArt: the shapes Excel keeps drawn.
        if let Some(html) = diagram_html(ctx, obj, size) {
            out.push(Obj {
                anchor,
                html,
                name,
                macro_name,
                kind: "group",
                chart: None,
                activex: None,
                assignable: true,
                shape_id: None,
            });
            return;
        }
        // A chart: drawn by the page; one the page does not draw is a box with its title.
        let (html, chart) = chart_html(ctx, obj, &name, env);
        if !html.is_empty() {
            out.push(Obj {
                anchor,
                html,
                name,
                macro_name,
                kind: "chart",
                chart,
                activex: None,
                assignable: true,
                shape_id: None,
            });
        }
        return;
    }
    // The box comes from the anchor (the shape's xfrm says the same in EMU).
    let ext = (env.size)(&anchor)
        .or_else(|| xfrm(obj).map(|x| x.ext))
        .filter(|e| e.0 > 0.0 && e.1 > 0.0);
    let html = object_html(ctx, obj, None, ext);
    if !html.is_empty() {
        let kind = match obj.name.as_str() {
            "pic" => "picture",
            "grpSp" => "group",
            _ => "shape",
        };
        out.push(Obj {
            anchor,
            html,
            name,
            macro_name,
            kind,
            chart: None,
            activex: None,
            // (A group has no macro of its own in the file.)
            assignable: obj.name != "grpSp",
            shape_id: None,
        });
    }
}

/// A chart: its spec for the page, or a box with its title when the page does not draw it (the
/// newer kinds: waterfall, treemap, ...; other graphics).
fn chart_html(
    ctx: &Ctx,
    frame: &Node,
    name: &str,
    env: &Env,
) -> (String, Option<serde_json::Value>) {
    if frame
        .find("cNvPr")
        .and_then(|c| c.attr("hidden"))
        .is_some_and(|h| h == "1" || h == "true")
    {
        return (String::new(), None);
    }
    let Some(data) = frame.find("graphicData") else {
        return (String::new(), None);
    };
    let chart = data
        .children
        .iter()
        .find(|c| c.name == "chart")
        .and_then(|c| c.attr("id"))
        .and_then(|id| ctx.rels.get(id));
    if let Some((_, part)) = chart
        && let Some(spec) = super::chart::chart_spec(ctx.parts, part, &ctx.theme, env)
    {
        return ("<div class=\"chartbox\"></div>".into(), Some(spec));
    }
    // chartEx (cx:chart) and the like: the title of the chart, if it has one.
    let part = chart.map(|(_, p)| p.clone()).or_else(|| {
        let id = data.children.iter().find_map(|c| c.attr("id"))?;
        ctx.rels.get(id).map(|(_, p)| p.clone())
    });
    let title = part.and_then(|p| ctx.parts.xml(&p)).and_then(|root| {
        let t = root.find("chart")?.child("title")?;
        let mut runs = Vec::new();
        t.find_all("t", &mut runs);
        let s: String = runs.iter().map(|r| r.text.as_str()).collect();
        (!s.trim().is_empty()).then(|| s.trim().to_string())
    });
    let what = if chart.is_some() || data.attr("uri").is_some_and(|u| u.contains("chart")) {
        "このグラフ"
    } else {
        "この図表"
    };
    let title = title.unwrap_or_else(|| name.to_string());
    let html = format!(
        "<div class=\"chartnote\"><svg viewBox=\"0 0 24 24\" aria-hidden=\"true\"><path d=\"M3 21h18M6 18V11M11 18V6M16 18v-9M21 18v-4\"/></svg><b>{}</b><span>{what}の種類は表示できません</span></div>",
        esc(&title)
    );
    (html, None)
}

/// SmartArt (a graphic frame of a diagram) as the shapes Excel keeps drawn for it (the drawing
/// of its data model, `xl/diagrams/drawingN.xml`); None for other graphics. `size`: the frame's
/// box (pt).
fn diagram_html(ctx: &Ctx, frame: &Node, size: Option<(f64, f64)>) -> Option<String> {
    let ids = frame.find("relIds")?;
    let (_, data) = ctx.rels.get(ids.attr("dm")?)?;
    let model = ctx.parts.xml(data)?;
    let drawing = model
        .find("dataModelExt")
        .and_then(|e| e.attr("relId"))
        .and_then(|id| ctx.rels.get(id))
        .map(|(_, p)| p.clone())
        .or_else(|| {
            // Without the extension: the diagram drawing of the data part, by its name.
            let n = data
                .rsplit('/')
                .next()?
                .trim_start_matches("data")
                .trim_end_matches(".xml")
                .to_string();
            Some(format!("xl/diagrams/drawing{n}.xml"))
        })?;
    let tree = ctx.parts.xml(&drawing)?;
    let tree = tree.find("spTree")?;
    // The frame's size in the diagram's coordinates (pt): its xfrm, else its box.
    let ext = frame
        .child("xfrm")
        .and_then(|x| x.child("ext"))
        .and_then(|e| {
            Some((
                e.attr("cx")?.parse::<f64>().ok()? / EMU_PT,
                e.attr("cy")?.parse::<f64>().ok()? / EMU_PT,
            ))
        });
    let (w, h) = size.or(ext)?;
    let (fw, fh) = ext.unwrap_or((w, h));
    if fw <= 0.0 || fh <= 0.0 {
        return None;
    }
    let mut out = String::new();
    let mut put = |c: &Node| {
        let Some(cx) = xfrm(c) else { return };
        let place = (cx.off.0 / fw, cx.off.1 / fh, cx.ext.0 / fw, cx.ext.1 / fh);
        out.push_str(&object_html(
            ctx,
            c,
            Some(place),
            Some((cx.ext.0 * w / fw, cx.ext.1 * h / fh)),
        ));
    };
    for c in tree
        .children
        .iter()
        .filter(|c| matches!(c.name.as_str(), "sp" | "cxnSp" | "grpSp" | "pic"))
    {
        // The text in the box SmartArt gives it (txXfrm), not in the shape's own text box.
        match (c.child("txXfrm"), c.child("txBody")) {
            (Some(tx), Some(body)) => {
                let mut shape = c.clone();
                shape
                    .children
                    .retain(|x| x.name != "txBody" && x.name != "txXfrm");
                put(&shape);
                let props = Node {
                    name: "spPr".into(),
                    children: vec![
                        Node {
                            name: "xfrm".into(),
                            ..tx.clone()
                        },
                        Node {
                            name: "noFill".into(),
                            ..Default::default()
                        },
                    ],
                    ..Default::default()
                };
                put(&Node {
                    name: "sp".into(),
                    children: vec![props, body.clone()],
                    ..Default::default()
                });
            }
            _ => put(c),
        }
    }
    (!out.is_empty()).then(|| {
        format!("<div class=\"shape\" style=\"left:0;top:0;width:100%;height:100%\">{out}</div>")
    })
}

/// The HTML of a shape, connector, picture or group. Without `place` it fills its container (a
/// top-level object, sized by its anchor); in a group, `place` is its box as fractions of the
/// group's box. `size` is the object's size in pt.
pub(crate) fn object_html(
    ctx: &Ctx,
    obj: &Node,
    place: Option<(f64, f64, f64, f64)>,
    size: Option<(f64, f64)>,
) -> String {
    let nv = obj
        .children
        .iter()
        .find(|c| c.name.starts_with("nv"))
        .and_then(|c| c.child("cNvPr"));
    if nv
        .and_then(|c| c.attr("hidden"))
        .is_some_and(|h| h == "1" || h == "true")
    {
        return String::new();
    }
    let x = xfrm(obj).unwrap_or_default();
    let (w, h) = size.unwrap_or(x.ext);
    let pct = |v: f64| n(v * 100.0);
    let mut style = match place {
        Some((l, t, pw, ph)) => format!(
            "left:{}%;top:{}%;width:{}%;height:{}%;",
            pct(l),
            pct(t),
            pct(pw),
            pct(ph)
        ),
        None => "left:0;top:0;width:100%;height:100%;".to_string(),
    };
    if x.rot != 0.0 {
        let _ = write!(style, "transform:rotate({}deg);", n(x.rot));
    }
    // The shape's effects (shadow, glow, reflection) on all of it, its text too (as Excel shows
    // them on the text of a shape without fill).
    if obj.name != "grpSp" {
        let fx = obj.child("spPr").and_then(|p| p.child("effectLst"));
        style.push_str(&effect_css(fx, &ctx.theme, x.rot));
        style.push_str(&reflect_css(fx));
    }
    let inner = match obj.name.as_str() {
        "grpSp" => group_html(ctx, obj, &x, w, h),
        "pic" => picture_html(ctx, obj, &x),
        "graphicFrame" => String::new(),
        _ => shape_html(ctx, obj, &x, w, h),
    };
    if inner.is_empty() {
        return String::new();
    }
    format!("<div class=\"shape\" style=\"{style}\">{inner}</div>")
}

fn group_html(ctx: &Ctx, g: &Node, x: &Xfrm, w: f64, h: f64) -> String {
    // Children are in the group's child coordinates (chOff, chExt).
    let (cw, ch) = (
        if x.ch_ext.0 > 0.0 {
            x.ch_ext.0
        } else {
            x.ext.0.max(0.01)
        },
        if x.ch_ext.1 > 0.0 {
            x.ch_ext.1
        } else {
            x.ext.1.max(0.01)
        },
    );
    let mut out = String::new();
    for c in g
        .children
        .iter()
        .filter(|c| matches!(c.name.as_str(), "sp" | "cxnSp" | "grpSp" | "pic"))
    {
        let Some(cx) = xfrm(c) else { continue };
        let place = (
            (cx.off.0 - x.ch_off.0) / cw,
            (cx.off.1 - x.ch_off.1) / ch,
            cx.ext.0 / cw,
            cx.ext.1 / ch,
        );
        out.push_str(&object_html(
            ctx,
            c,
            Some(place),
            Some((cx.ext.0 * w / cw, cx.ext.1 * h / ch)),
        ));
    }
    if x.flip_h || x.flip_v {
        return format!(
            "<div style=\"position:absolute;inset:0;transform:scale({},{})\">{out}</div>",
            if x.flip_h { -1 } else { 1 },
            if x.flip_v { -1 } else { 1 }
        );
    }
    out
}

fn picture_html(ctx: &Ctx, pic: &Node, x: &Xfrm) -> String {
    // An SVG picture is a PNG fallback in a:blip with the SVG itself in the blip's extension
    // (asvg:svgBlip); use the SVG so that it stays sharp.
    let blip = pic.find("blip");
    let found = |b: Option<&Node>| {
        let id = b?.attr("embed")?;
        let (_, part) = ctx.rels.get(id)?;
        Some((part.clone(), ctx.parts.bytes(part)?))
    };
    let Some((part, data)) = found(blip.and_then(|b| b.find("svgBlip"))).or_else(|| found(blip))
    else {
        return String::new();
    };
    let flip = if x.flip_h || x.flip_v {
        format!(
            "transform:scale({},{});",
            if x.flip_h { -1 } else { 1 },
            if x.flip_v { -1 } else { 1 }
        )
    } else {
        String::new()
    };
    picture_of(ctx, &part, data, &flip)
}

/// An `<img>` filling its box for the picture `part`, or a box saying that it cannot be shown.
pub(crate) fn picture_of(ctx: &Ctx, part: &str, data: &[u8], style: &str) -> String {
    // EMF (a range copied as a picture, a linked picture, the preview of an object): as SVG.
    if part.to_ascii_lowercase().ends_with(".emf")
        && let Some(svg) = super::emf::to_svg(data)
    {
        let i = ctx
            .images
            .borrow_mut()
            .add(part, svg.as_bytes(), "image/svg+xml");
        return format!(
            "<img alt=\"\" data-img=\"{i}\" style=\"position:absolute;inset:0;width:100%;height:100%;{style}\">"
        );
    }
    let Some(mime) = image_mime(part) else {
        let kind = part.rsplit('.').next().unwrap_or("").to_ascii_uppercase();
        return format!(
            "<div class=\"noimage\">表示できない形式の画像（{}）</div>",
            esc(&kind)
        );
    };
    let svg;
    let data = if mime == "image/svg+xml" {
        svg = svg_stretched(data);
        svg.as_slice()
    } else {
        data
    };
    let i = ctx.images.borrow_mut().add(part, data, mime);
    format!(
        "<img alt=\"\" data-img=\"{i}\" style=\"position:absolute;inset:0;width:100%;height:100%;{style}\">"
    )
}

/// An SVG that fills the picture's box as Excel draws it: stretched (preserveAspectRatio="none")
/// and scaled (a viewBox from width and height when it has none; without one a browser draws it
/// at its own size).
fn svg_stretched(data: &[u8]) -> Vec<u8> {
    let s = String::from_utf8_lossy(data);
    let Some(start) = s
        .match_indices("<svg")
        .map(|(i, _)| i)
        .find(|&i| s[i + 4..].starts_with(|c: char| c.is_whitespace() || c == '>' || c == '/'))
    else {
        return data.to_vec();
    };
    let Some(len) = s[start..].find('>') else {
        return data.to_vec();
    };
    let mut end = start + len;
    if s[..end].ends_with('/') {
        end -= 1;
    }
    let tag = &s[start..end];
    let attr = |name: &str| {
        let i = tag.match_indices(name).map(|(i, _)| i).find(|&i| {
            tag[..i].ends_with(char::is_whitespace)
                && tag[i + name.len()..].trim_start().starts_with('=')
        })?;
        let rest = tag[i + name.len()..].trim_start()[1..].trim_start();
        let q = rest.chars().next().filter(|c| *c == '"' || *c == '\'')?;
        rest[1..].split(q).next()
    };
    let mut add = String::new();
    if attr("viewBox").is_none() {
        let size = |name| {
            attr(name)
                .and_then(|v| v.trim().trim_end_matches("px").trim().parse::<f64>().ok())
                .filter(|v| *v > 0.0)
        };
        if let (Some(w), Some(h)) = (size("width"), size("height")) {
            let _ = write!(add, " viewBox=\"0 0 {} {}\"", n(w), n(h));
        }
    }
    if attr("preserveAspectRatio").is_none() {
        add.push_str(" preserveAspectRatio=\"none\"");
    }
    format!("{}{add}{}", &s[..end], &s[end..]).into_bytes()
}

/// An arrowhead of a line end (headEnd / tailEnd), as an SVG marker.
fn marker(end: Option<&Node>, color: &str, id: &str, width: f64) -> Option<String> {
    let e = end?;
    let kind = e.attr("type")?;
    let size = |k: &str| match e.attr(k) {
        Some("sm") => 2.0,
        Some("lg") => 5.0,
        _ => 3.0,
    };
    // As Excel: in proportion to the line, but not tiny on thin lines.
    let unit = width.max(2.0);
    let (mw, mh) = (size("w") * unit, size("len") * unit);
    let shape = match kind {
        "triangle" => "<path d=\"M0,0 L10,5 L0,10 Z\"/>".to_string(),
        "stealth" => "<path d=\"M0,0 L10,5 L0,10 L3,5 Z\"/>".to_string(),
        "arrow" => format!(
            "<path d=\"M0,0 L10,5 L0,10\" fill=\"none\" stroke=\"{color}\" stroke-width=\"1.5\"/>"
        ),
        "diamond" => "<path d=\"M0,5 L5,0 L10,5 L5,10 Z\"/>".to_string(),
        "oval" => "<circle cx=\"5\" cy=\"5\" r=\"5\"/>".to_string(),
        _ => return None,
    };
    let refx = if matches!(kind, "diamond" | "oval") {
        5
    } else {
        9
    };
    Some(format!(
        "<marker id=\"{id}\" viewBox=\"0 0 10 10\" refX=\"{refx}\" refY=\"5\" markerWidth=\"{}\" markerHeight=\"{}\" orient=\"auto-start-reverse\" markerUnits=\"userSpaceOnUse\" fill=\"{color}\" overflow=\"visible\">{shape}</marker>",
        n(mh),
        n(mw)
    ))
}

fn dash(ln: Option<&Node>, width: f64) -> String {
    let Some(d) = ln
        .and_then(|l| l.child("prstDash"))
        .and_then(|p| p.attr("val"))
    else {
        return String::new();
    };
    let k = width.max(0.5);
    let pattern: &[f64] = match d {
        "dash" => &[4.0, 3.0],
        "sysDash" => &[3.0, 1.0],
        "dot" | "sysDot" => &[1.0, 1.0],
        "lgDash" => &[8.0, 3.0],
        "dashDot" | "sysDashDot" => &[4.0, 3.0, 1.0, 3.0],
        "lgDashDot" => &[8.0, 3.0, 1.0, 3.0],
        "lgDashDotDot" | "sysDashDotDot" => &[8.0, 3.0, 1.0, 3.0, 1.0, 3.0],
        _ => return String::new(),
    };
    let v: Vec<String> = pattern.iter().map(|x| n(x * k)).collect();
    format!(" stroke-dasharray=\"{}\"", v.join(" "))
}

fn shape_html(ctx: &Ctx, sp: &Node, x: &Xfrm, w: f64, h: f64) -> String {
    let props = sp.child("spPr");
    let geom = if let Some(g) = props.and_then(|p| p.child("prstGeom")) {
        // Presets that are not defined are drawn as their rectangle.
        geom::preset(g.attr("prst").unwrap_or("rect"), Some(g), w, h)
            .or_else(|| geom::preset("rect", None, w, h))
    } else if let Some(g) = props.and_then(|p| p.child("custGeom")) {
        Some(geom::custom(g, w, h))
    } else if sp.name == "cxnSp" {
        geom::preset("line", None, w, h)
    } else {
        None
    };
    let paths = geom
        .as_ref()
        .map(|g| g.paths.as_slice())
        .unwrap_or_default();
    let ln = props.and_then(|p| p.child("ln"));
    // Fill: the shape's own, else the theme's by its style.
    let fill = fill_color(props, &ctx.theme)
        .or_else(|| style_ref(sp, "fillRef", &ctx.theme).map(|(c, _)| (c, 1.0)))
        .unwrap_or(("none".into(), 1.0));
    let stroke = fill_color(ln, &ctx.theme)
        .or_else(|| style_ref(sp, "lnRef", &ctx.theme).map(|(c, _)| (c, 1.0)))
        .unwrap_or(("none".into(), 1.0));
    let width = ln
        .and_then(|l| l.attr("w"))
        .and_then(|v| v.parse::<f64>().ok())
        .map(|v| v / EMU_PT)
        .unwrap_or_else(
            || match style_ref(sp, "lnRef", &ctx.theme).map(|(_, i)| i) {
                Some(1) => 0.5,
                Some(3) => 1.5,
                Some(_) => 1.0,
                None => 0.75,
            },
        );
    let op = |name: &str, a: f64| {
        if a < 1.0 {
            format!(" {name}-opacity=\"{}\"", n(a))
        } else {
            String::new()
        }
    };
    let mut body = String::new();
    let mut defs = String::new();
    // The arrowheads (made once, for the open paths).
    let mut ends: Option<String> = None;
    for p in paths {
        let f = if p.fill == geom::Fill::None || fill.0 == "none" {
            "none".to_string()
        } else {
            geom::shade(&fill.0, p.fill)
        };
        let s = if p.stroke { stroke.0.as_str() } else { "none" };
        if p.d.is_empty() || (f == "none" && s == "none") {
            continue;
        }
        let _ = write!(body, "<path d=\"{}\" fill=\"{f}\"", p.d);
        if f != "none" {
            body.push_str(&op("fill", fill.1));
        }
        if s == "none" {
            body.push_str(" stroke=\"none\"");
        } else {
            let _ = write!(
                body,
                " stroke=\"{s}\" stroke-width=\"{}\"{}{} stroke-linejoin=\"round\"",
                n(width),
                op("stroke", stroke.1),
                dash(ln, width)
            );
            if p.open {
                let e = ends.get_or_insert_with(|| {
                    let id = ctx.ids.get();
                    ctx.ids.set(id + 1);
                    let mut e = String::new();
                    for (k, attr) in [("headEnd", "marker-start"), ("tailEnd", "marker-end")] {
                        let mid = format!("ras-m{id}{k}");
                        if let Some(m) = marker(ln.and_then(|l| l.child(k)), s, &mid, width) {
                            defs.push_str(&m);
                            let _ = write!(e, " {attr}=\"url(#{mid})\"");
                        }
                    }
                    e
                });
                body.push_str(e);
            }
        }
        body.push_str("/>");
    }
    let mut svg = String::new();
    if !body.is_empty() {
        let flip = if x.flip_h || x.flip_v {
            format!(
                " transform=\"translate({},{}) scale({},{})\"",
                if x.flip_h { n(w) } else { "0".into() },
                if x.flip_v { n(h) } else { "0".into() },
                if x.flip_h { -1 } else { 1 },
                if x.flip_v { -1 } else { 1 }
            )
        } else {
            String::new()
        };
        let _ = write!(
            svg,
            "<svg viewBox=\"0 0 {w} {h}\" preserveAspectRatio=\"none\" style=\"position:absolute;inset:0;width:100%;height:100%;overflow:visible\">{}<g{flip}>{body}</g></svg>",
            if defs.is_empty() {
                String::new()
            } else {
                format!("<defs>{defs}</defs>")
            },
            w = n(w.max(0.01)),
            h = n(h.max(0.01)),
        );
    }
    // The text box of the geometry, mirrored with the shape.
    let rect = geom.as_ref().and_then(|g| g.text).map(|(l, t, r, b)| {
        let (l, r) = if x.flip_h { (w - r, w - l) } else { (l, r) };
        let (t, b) = if x.flip_v { (h - b, h - t) } else { (t, b) };
        (l.max(0.0), t.max(0.0), (w - r).max(0.0), (h - b).max(0.0))
    });
    let text = sp
        .child("txBody")
        .map(|t| text_html(ctx, sp, t, rect.unwrap_or_default(), (w, h), x.rot))
        .unwrap_or_default();
    format!("{svg}{text}")
}

/// The text of a shape (txBody). `bx` is the distance of the geometry's text box from the
/// shape's edges (left, top, right, bottom); `size` is the shape's size and `rot` its rotation
/// (degrees).
///
/// WordArt is text with effects (text_fill, the outline, effect_css) and maybe a text warp
/// (bodyPr's prstTxWarp): its lines are put in a .warp box with the warp's paths, and the page
/// lays the characters out along them (`warpText` in view.js).
pub(crate) fn text_html(
    ctx: &Ctx,
    sp: &Node,
    tx: &Node,
    bx: (f64, f64, f64, f64),
    size: (f64, f64),
    rot: f64,
) -> String {
    let body = tx.child("bodyPr");
    let ins = |k: &str, d: f64| {
        body.and_then(|b| b.attr(k))
            .and_then(|v| v.parse::<f64>().ok())
            .unwrap_or(d)
            / EMU_PT
    };
    let (l, t, r, b) = (
        bx.0 + ins("lIns", 91440.0),
        bx.1 + ins("tIns", 45720.0),
        bx.2 + ins("rIns", 91440.0),
        bx.3 + ins("bIns", 45720.0),
    );
    let anchor = match body.and_then(|b| b.attr("anchor")) {
        Some("ctr") => "center",
        Some("b") => "flex-end",
        _ => "flex-start",
    };
    let vert = matches!(
        body.and_then(|b| b.attr("vert")),
        Some("vert" | "eaVert" | "wordArtVert" | "mongolianVert")
    );
    let nowrap = body.and_then(|b| b.attr("wrap")) == Some("none");
    let base_color = style_ref(sp, "fontRef", &ctx.theme)
        .map(|(c, _)| c)
        .unwrap_or_else(|| "#000000".into());
    // Defaults from the list style (lvl1pPr/defRPr).
    let lst = tx.child("lstStyle").and_then(|l| l.child("lvl1pPr"));
    let mut paras = String::new();
    let mut any = false;
    // The reflection of the text (of its first run with one).
    let mut reflect = String::new();
    for p in tx.all("p") {
        let ppr = p.child("pPr");
        let algn = ppr
            .and_then(|x| x.attr("algn"))
            .or_else(|| lst.and_then(|x| x.attr("algn")));
        let align = match algn {
            Some("ctr") => "center",
            Some("r") => "right",
            Some("just") | Some("dist") => "justify",
            _ => "left",
        };
        let def = ppr
            .and_then(|x| x.child("defRPr"))
            .or_else(|| lst.and_then(|x| x.child("defRPr")));
        let mut runs = String::new();
        if let Some(bu) = ppr
            .and_then(|x| x.child("buChar"))
            .and_then(|c| c.attr("char"))
        {
            let _ = write!(runs, "{}&nbsp;", esc(bu));
        }
        let mut size_last = def
            .and_then(|d| d.attr("sz"))
            .and_then(|v| v.parse::<f64>().ok())
            .map(|v| v / 100.0);
        for c in &p.children {
            match c.name.as_str() {
                "r" | "fld" => {
                    let s = c.child("t").map(|t| t.text.clone()).unwrap_or_default();
                    if s.is_empty() {
                        continue;
                    }
                    any = true;
                    let rpr = c.child("rPr");
                    let css = run_css(ctx, rpr, def, &base_color, rot);
                    if reflect.is_empty() {
                        reflect = reflect_css(
                            rpr.and_then(|r| r.child("effectLst"))
                                .or_else(|| def.and_then(|d| d.child("effectLst"))),
                        );
                    }
                    if let Some(sz) = rpr
                        .and_then(|x| x.attr("sz"))
                        .and_then(|v| v.parse::<f64>().ok())
                    {
                        size_last = Some(sz / 100.0);
                    }
                    let _ = write!(runs, "<span style=\"{css}\">{}</span>", esc(&s));
                }
                "br" => runs.push_str("<br>"),
                _ => {}
            }
        }
        if runs.is_empty() {
            // An empty paragraph keeps its line.
            let sz = p
                .child("endParaRPr")
                .and_then(|x| x.attr("sz"))
                .and_then(|v| v.parse::<f64>().ok())
                .map(|v| v / 100.0)
                .or(size_last)
                .unwrap_or(11.0);
            runs = format!("<span style=\"font-size:{}pt\">&nbsp;</span>", n(sz));
        }
        let _ = write!(paras, "<p style=\"margin:0;text-align:{align}\">{runs}</p>");
    }
    if !any {
        return String::new();
    }
    let family = esc(&font_family(&ctx.font.0));
    // A text warp: in the text box less its insets.
    let (ww, wh) = (size.0 - l - r, size.1 - t - b);
    if let Some(warp) = body.and_then(|b| b.child("prstTxWarp"))
        && ww > 0.0
        && wh > 0.0
        && let Some(paths) = geom::text_warp(warp.attr("prst").unwrap_or(""), Some(warp), ww, wh)
    {
        let paths: Vec<String> = paths
            .iter()
            .map(|p| {
                format!(
                    "[{}]",
                    p.iter()
                        .map(|(x, y)| format!("{},{}", n(*x), n(*y)))
                        .collect::<Vec<_>>()
                        .join(",")
                )
            })
            .collect();
        let spec = format!(
            "{{\"w\":{},\"h\":{},\"paths\":[{}]}}",
            n(ww),
            n(wh),
            paths.join(",")
        );
        return format!(
            "<div class=\"warp\" data-warp=\"{}\" style=\"position:absolute;left:{}pt;top:{}pt;width:{}pt;height:{}pt;overflow:visible;line-height:1.2;white-space:pre;font-family:{family};font-size:11pt;color:{base_color};{reflect}\">{paras}</div>",
            esc(&spec),
            n(l),
            n(t),
            n(ww),
            n(wh)
        );
    }
    if !reflect.is_empty() {
        // Reflected below the text, not below the box.
        paras = format!("<div style=\"{reflect}\">{paras}</div>");
    }
    let mut style = format!(
        "position:absolute;inset:0;box-sizing:border-box;padding:{}pt {}pt {}pt {}pt;display:flex;flex-direction:column;justify-content:{anchor};overflow:visible;line-height:1.2;font-family:{};font-size:11pt;color:{base_color};",
        n(t),
        n(r),
        n(b),
        n(l),
        family
    );
    style.push_str(if nowrap {
        "white-space:pre;"
    } else {
        "white-space:pre-wrap;overflow-wrap:anywhere;"
    });
    if vert {
        style.push_str("writing-mode:vertical-rl;");
    }
    format!("<div style=\"{style}\">{paras}</div>")
}

fn run_css(
    ctx: &Ctx,
    rpr: Option<&Node>,
    def: Option<&Node>,
    base_color: &str,
    rot: f64,
) -> String {
    let get = |k: &str| {
        rpr.and_then(|x| x.attr(k))
            .or_else(|| def.and_then(|x| x.attr(k)))
    };
    let mut css = String::new();
    if let Some(sz) = get("sz").and_then(|v| v.parse::<f64>().ok()) {
        let _ = write!(css, "font-size:{}pt;", n(sz / 100.0));
    }
    if matches!(get("b"), Some("1" | "true")) {
        css.push_str("font-weight:bold;");
    }
    if matches!(get("i"), Some("1" | "true")) {
        css.push_str("font-style:italic;");
    }
    let mut deco = Vec::new();
    if get("u").is_some_and(|u| u != "none") {
        deco.push("underline");
    }
    if get("strike").is_some_and(|s| s != "noStrike") {
        deco.push("line-through");
    }
    if !deco.is_empty() {
        let _ = write!(css, "text-decoration:{};", deco.join(" "));
    }
    match get("cap") {
        Some("all") => css.push_str("text-transform:uppercase;"),
        Some("small") => css.push_str("font-variant:small-caps;"),
        _ => {}
    }
    if let Some(spc) = get("spc")
        .and_then(|v| v.parse::<f64>().ok())
        .filter(|v| *v != 0.0)
    {
        let _ = write!(css, "letter-spacing:{}pt;", n(spc / 100.0));
    }
    let face = |n0: Option<&Node>| -> Option<String> {
        let n0 = n0?;
        for k in ["ea", "latin"] {
            if let Some(f) = n0.child(k).and_then(|c| c.attr("typeface"))
                && !f.is_empty()
                && !f.starts_with('+')
            {
                return Some(f.to_string());
            }
        }
        None
    };
    if let Some(f) = face(rpr).or_else(|| face(def)) {
        let _ = write!(css, "font-family:{};", esc(&font_family(&f)));
    }
    css.push_str(&text_fill(rpr, def, &ctx.theme, base_color));
    // The outline of the characters.
    let ln = rpr
        .and_then(|x| x.child("ln"))
        .or_else(|| def.and_then(|x| x.child("ln")));
    if let Some((c, a)) = fill_color(ln, &ctx.theme).filter(|(c, _)| c != "none") {
        let w = ln
            .and_then(|l| l.attr("w"))
            .and_then(|v| v.parse::<f64>().ok())
            .map(|v| v / EMU_PT)
            .unwrap_or(0.75)
            .max(0.25);
        let _ = write!(css, "-webkit-text-stroke:{}pt {};", n(w), rgba(&c, a));
    }
    let fx = rpr
        .and_then(|x| x.child("effectLst"))
        .or_else(|| def.and_then(|x| x.child("effectLst")));
    css.push_str(&effect_css(fx, &ctx.theme, rot));
    css
}

/// The fill of a run's characters: a color, a gradient (as the background of the text) or none
/// (only the outline shows).
fn text_fill(
    rpr: Option<&Node>,
    def: Option<&Node>,
    theme: &HashMap<String, String>,
    base_color: &str,
) -> String {
    fn fill(r: Option<&Node>) -> Option<&Node> {
        r?.children.iter().find(|c| {
            matches!(
                c.name.as_str(),
                "solidFill" | "gradFill" | "noFill" | "pattFill"
            )
        })
    }
    let Some(f) = fill(rpr).or_else(|| fill(def)) else {
        return String::new();
    };
    match f.name.as_str() {
        "noFill" => "color:transparent;".into(),
        "gradFill" => match gradient_css(f, theme) {
            Some(g) => format!(
                "background-image:{g};-webkit-background-clip:text;background-clip:text;color:transparent;"
            ),
            None => String::new(),
        },
        _ => {
            let f = if f.name == "pattFill" {
                f.child("fgClr").unwrap_or(f)
            } else {
                f
            };
            match color_of(f, theme) {
                Some(c) if c != base_color || alpha_of(f) < 1.0 => {
                    format!("color:{};", rgba(&c, alpha_of(f)))
                }
                _ => String::new(),
            }
        }
    }
}

/// A CSS color of "#RRGGBB" with an alpha.
fn rgba(color: &str, alpha: f64) -> String {
    if alpha >= 1.0 {
        return color.to_string();
    }
    let v = color
        .strip_prefix('#')
        .and_then(|h| u32::from_str_radix(h, 16).ok())
        .unwrap_or(0);
    format!(
        "rgba({},{},{},{})",
        (v >> 16) & 0xff,
        (v >> 8) & 0xff,
        v & 0xff,
        n(alpha)
    )
}

/// The CSS color of a node with a color (a fill, a gradient stop, an effect).
fn css_color(node: &Node, theme: &HashMap<String, String>) -> Option<String> {
    Some(rgba(&color_of(node, theme)?, alpha_of(node)))
}

/// A gradient fill (gradFill) as a CSS gradient: linear at its angle, or radial from the center
/// of its fillToRect.
fn gradient_css(g: &Node, theme: &HashMap<String, String>) -> Option<String> {
    let mut stops: Vec<(f64, String)> = g
        .child("gsLst")?
        .all("gs")
        .filter_map(|s| {
            Some((
                s.attr("pos")
                    .and_then(|v| v.parse::<f64>().ok())
                    .unwrap_or(0.0)
                    / 1000.0,
                css_color(s, theme)?,
            ))
        })
        .collect();
    if stops.is_empty() {
        return None;
    }
    stops.sort_by(|a, b| a.0.total_cmp(&b.0));
    let list = stops
        .iter()
        .map(|(p, c)| format!("{c} {}%", n(*p)))
        .collect::<Vec<_>>()
        .join(",");
    if let Some(path) = g.child("path") {
        let r = path.child("fillToRect");
        let v = |k: &str| {
            r.and_then(|r| r.attr(k))
                .and_then(|v| v.parse::<f64>().ok())
                .unwrap_or(50000.0)
                / 1000.0
        };
        let (cx, cy) = (
            (v("l") + 100.0 - v("r")) / 2.0,
            (v("t") + 100.0 - v("b")) / 2.0,
        );
        let shape = if path.attr("path") == Some("circle") {
            "circle"
        } else {
            "ellipse"
        };
        return Some(format!(
            "radial-gradient({shape} farthest-corner at {}% {}%,{list})",
            n(cx),
            n(cy)
        ));
    }
    // DrawingML: 0 is left to right, clockwise; CSS: 90deg is left to right.
    let ang = g
        .child("lin")
        .and_then(|l| l.attr("ang"))
        .and_then(|v| v.parse::<f64>().ok())
        .unwrap_or(0.0)
        / 60000.0;
    Some(format!("linear-gradient({}deg,{list})", n(ang + 90.0)))
}

/// The glow and outer shadow of an effect list as a CSS filter. `rot` is the rotation of the
/// shape (a shadow that does not turn with it is turned back).
fn effect_css(fx: Option<&Node>, theme: &HashMap<String, String>, rot: f64) -> String {
    let Some(fx) = fx else { return String::new() };
    let emu =
        |e: &Node, k: &str| e.attr(k).and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0) / EMU_PT;
    let mut f = Vec::new();
    // The glow first: the shadow falls from the glowing text.
    if let Some(g) = fx.child("glow")
        && let Some(c) = css_color(g, theme)
    {
        let r = n(emu(g, "rad") / 2.0);
        f.push(format!(
            "drop-shadow(0 0 {r}pt {c}) drop-shadow(0 0 {r}pt {c})"
        ));
    }
    if let Some(s) = fx.child("outerShdw")
        && let Some(c) = css_color(s, theme)
    {
        let mut dir = s
            .attr("dir")
            .and_then(|v| v.parse::<f64>().ok())
            .unwrap_or(0.0)
            / 60000.0;
        if matches!(s.attr("rotWithShape"), Some("0" | "false")) {
            dir -= rot;
        }
        let (dist, dir) = (emu(s, "dist"), dir.to_radians());
        f.push(format!(
            "drop-shadow({}pt {}pt {}pt {c})",
            n(dist * dir.cos()),
            n(dist * dir.sin()),
            n(emu(s, "blurRad") / 2.0)
        ));
    }
    if f.is_empty() {
        String::new()
    } else {
        format!("filter:{};", f.join(" "))
    }
}

/// The reflection of an effect list as -webkit-box-reflect (below, fading from stA at stPos to
/// endA at endPos).
fn reflect_css(fx: Option<&Node>) -> String {
    let Some(r) = fx.and_then(|f| f.child("reflection")) else {
        return String::new();
    };
    let num = |k: &str, d: f64| r.attr(k).and_then(|v| v.parse::<f64>().ok()).unwrap_or(d);
    let (st_a, st_pos, end_a, end_pos) = (
        num("stA", 100000.0) / 100000.0,
        num("stPos", 0.0) / 1000.0,
        num("endA", 0.0) / 100000.0,
        num("endPos", 100000.0) / 1000.0,
    );
    format!(
        "-webkit-box-reflect:below {}pt linear-gradient(rgba(0,0,0,{}) {}%,rgba(0,0,0,{}) {}%,transparent {}%);",
        n(num("dist", 0.0) / EMU_PT),
        n(st_a),
        n(st_pos),
        n(end_a),
        n(end_pos),
        n(end_pos)
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    fn ctx<'a>(
        parts: &'a Parts,
        rels: HashMap<String, (String, String)>,
        ids: &'a Cell<usize>,
        images: &'a RefCell<Images>,
    ) -> Ctx<'a> {
        Ctx {
            parts,
            rels,
            theme: HashMap::new(),
            font: ("Calibri".into(), 11.0),
            ids,
            images,
        }
    }

    #[test]
    fn base64_of_bytes() {
        assert_eq!(base64(b""), "");
        assert_eq!(base64(b"f"), "Zg==");
        assert_eq!(base64(b"fo"), "Zm8=");
        assert_eq!(base64(b"foo"), "Zm9v");
        assert_eq!(base64(b"foob"), "Zm9vYg==");
    }

    /// An SVG picture is drawn from the SVG, not from its PNG fallback; each picture once.
    #[test]
    fn svg_picture() {
        let svg = br#"<?xml version="1.0"?><svg xmlns="http://www.w3.org/2000/svg" width="20px" height="10"><rect width="20" height="10"/></svg>"#;
        let files = HashMap::from([
            ("xl/media/image1.png".to_string(), b"PNG".to_vec()),
            ("xl/media/image2.svg".to_string(), svg.to_vec()),
        ]);
        let parts = Parts { files };
        let (ids, images) = (Cell::new(0), RefCell::new(Images::default()));
        let rels = HashMap::from([
            (
                "rId1".to_string(),
                ("image".to_string(), "xl/media/image1.png".to_string()),
            ),
            (
                "rId2".to_string(),
                ("image".to_string(), "xl/media/image2.svg".to_string()),
            ),
        ]);
        let ctx = ctx(&parts, rels, &ids, &images);
        let pic = parse(br#"<xdr:pic xmlns:xdr="x" xmlns:a="a" xmlns:r="r" xmlns:asvg="s"><xdr:blipFill><a:blip r:embed="rId1"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip r:embed="rId2"/></a:ext></a:extLst></a:blip></xdr:blipFill></xdr:pic>"#).unwrap();
        let html = picture_html(&ctx, &pic, &Xfrm::default());
        assert!(html.contains("data-img=\"0\""), "{html}");
        assert!(images.borrow().list[0].starts_with("data:image/svg+xml;base64,"));
        // Without the extension: the PNG; the same picture again is the same entry.
        let pic = parse(
            br#"<xdr:pic xmlns:xdr="x" xmlns:a="a" xmlns:r="r"><a:blip r:embed="rId1"/></xdr:pic>"#,
        )
        .unwrap();
        assert!(picture_html(&ctx, &pic, &Xfrm::default()).contains("data-img=\"1\""));
        assert!(picture_html(&ctx, &pic, &Xfrm::default()).contains("data-img=\"1\""));
        assert_eq!(images.borrow().list.len(), 2);
        assert_eq!(images.borrow().list[1], "data:image/png;base64,UE5H");
        // An SVG that says how to fit is kept as it is.
        let keep = br#"<svg viewBox="0 0 1 1" preserveAspectRatio="xMidYMid"/>"#;
        assert_eq!(svg_stretched(keep), keep.to_vec());
        // What a browser cannot show is said.
        assert!(
            picture_of(&ctx, "xl/media/image3.emf", b"", "")
                .contains("表示できない形式の画像（EMF）")
        );
    }

    /// An oval as Excel makes it: the theme's fill and line by its style, its text in the
    /// middle.
    #[test]
    fn oval_with_text() {
        let parts = Parts::default();
        let (ids, images) = (Cell::new(0), RefCell::new(Images::default()));
        let ctx = ctx(&parts, HashMap::new(), &ids, &images);
        let sp = parse(
            br#"<xdr:sp xmlns:xdr="x" xmlns:a="a"><xdr:nvSpPr><xdr:cNvPr id="2" name="Oval 1"/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></xdr:spPr><xdr:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></xdr:style><xdr:txBody><a:bodyPr anchor="ctr"/><a:p><a:pPr algn="ctr"/><a:r><a:t>OK</a:t></a:r></a:p></xdr:txBody></xdr:sp>"#,
        )
        .unwrap();
        let html = object_html(&ctx, &sp, None, Some((100.0, 50.0)));
        assert!(
            html.contains("<path d=\"M0,25 A50,25")
                && html.contains("fill=\"#4472C4\"")
                && html.contains("stroke=\"#223962\""),
            "{html}"
        );
        assert!(
            html.contains("justify-content:center")
                && html.contains("color:#FFFFFF")
                && html.contains(">OK</span>"),
            "{html}"
        );
    }
}
