//! The controls of a sheet for the grid: form controls (buttons, check boxes, option buttons,
//! list and combo boxes, spin buttons, scroll bars, labels, group boxes) and ActiveX controls
//! (MSForms: CommandButton, CheckBox, TextBox, ...), drawn as Excel shows them.
//!
//! A control is in several places of the file: the sheet's `<controls>` (its name, where it is,
//! its properties: `xl/ctrlProps/*.xml` for a form control, `xl/activeX/*.xml` and the persisted
//! control for an ActiveX one), the sheet's VML drawing (where it is, its text and font: what
//! Excel 2007 reads) and, from Excel 2010 on, a hidden mirror in the sheet's drawing (its place
//! among the shapes, its text). The ActiveX controls are drawn from their persisted data
//! (MS-OFORMS): their caption, value, colors and font.

use std::collections::HashMap;
use std::fmt::Write as _;

use umya_spreadsheet::drawingml::EMU_PX;
use umya_spreadsheet::drawingml::anchor::{Anchor, anchor_of};
use umya_spreadsheet::drawingml::color::color_of;
use umya_spreadsheet::drawingml::package::Package;

use super::drawing::{Ctx, Env, Obj, Parts, esc, font_family, n};
use super::xml::Node;

/// 1025 of a shape id "_x0000_s1025" (VML, the mirror's compatExt), or of "1025".
pub(crate) fn spid_number(s: &str) -> Option<u32> {
    let digits = s.rsplit(|c: char| !c.is_ascii_digit()).next()?;
    digits.parse().ok()
}

/// A text as a control shows it.
#[derive(Clone, Debug, Default)]
struct Text {
    text: String,
    font: Option<String>,
    size: Option<f64>,
    color: Option<String>,
    bold: bool,
    italic: bool,
    underline: bool,
    strike: bool,
    /// "left", "center", "right".
    align: Option<&'static str>,
    /// "top", "center", "bottom".
    valign: Option<&'static str>,
}

/// What the persisted data of an ActiveX control says.
#[derive(Clone, Debug, Default)]
struct Ax {
    class: &'static str,
    caption: Option<String>,
    value: Option<String>,
    fore: Option<String>,
    back: Option<String>,
    /// The back is drawn (BackStyle opaque).
    opaque: bool,
    font: Option<String>,
    size: Option<f64>,
    bold: bool,
    italic: bool,
    underline: bool,
    align: Option<&'static str>,
    /// MorphData's DisplayStyle (1 text, 2 list, 3 combo, 4 check box, 5 option button, 6 toggle
    /// button, 7 drop-down list).
    display: u8,
    min: i64,
    max: i64,
    position: i64,
    /// Orientation: -1 automatic, 0 vertical, 1 horizontal.
    orientation: i64,
}

#[derive(Debug, Default)]
struct Control {
    id: u32,
    name: String,
    /// The form control's object type ("Button", "CheckBox", ...) or "ActiveX".
    kind: String,
    anchor: Option<Anchor>,
    /// The form control's properties (`formControlPr`).
    props: HashMap<String, String>,
    /// The VML shape's `x:ClientData` values (`Checked`, `FmlaMacro`, ...).
    client: HashMap<String, String>,
    text: Option<Text>,
    ax: Option<Ax>,
    /// The picture Excel shows for it (an embedded object's, an ActiveX control's preview).
    image: Option<String>,
    hidden: bool,
    done: bool,
}

pub(crate) struct Controls {
    list: Vec<Control>,
    font: (String, f64),
}

impl Controls {
    /// The controls of the sheet whose part is `sheet_part`.
    pub(crate) fn read(parts: &Parts, sheet_part: &str, font: &(String, f64)) -> Controls {
        let mut list: Vec<Control> = Vec::new();
        let rels = parts.rels(sheet_part);
        // The sheet's <controls>: the name, the anchor and the properties.
        if let Some(root) = parts.xml(&format!("{sheet_part}#controls")) {
            let mut found = Vec::new();
            root.find_all("control", &mut found);
            for c in found {
                let Some(id) = c.attr("shapeId").and_then(spid_number) else {
                    continue;
                };
                let at = match list.iter().position(|x| x.id == id) {
                    Some(i) => i,
                    None => {
                        list.push(Control {
                            id,
                            ..Default::default()
                        });
                        list.len() - 1
                    }
                };
                let ctl = &mut list[at];
                if ctl.name.is_empty() {
                    ctl.name = c.attr("name").unwrap_or("").to_string();
                }
                if ctl.anchor.is_none()
                    && let Some(a) = c.find("anchor")
                {
                    let an = anchor_of(a);
                    if an.from.is_some() {
                        ctl.anchor = Some(an);
                    }
                }
                // The preview of an ActiveX control.
                if ctl.image.is_none()
                    && let Some((_, part)) = c
                        .find("controlPr")
                        .and_then(|p| p.attr("id"))
                        .and_then(|r| rels.get(r))
                {
                    ctl.image = Some(part.clone());
                }
                let Some((kind, part)) = c.attr("id").and_then(|r| rels.get(r)) else {
                    continue;
                };
                if kind.ends_with("/ctrlProp") {
                    if let Some(p) = parts.xml(part) {
                        ctl.props = p.attrs.iter().cloned().collect();
                        if let Some(t) = p.attr("objectType") {
                            ctl.kind = t.to_string();
                        }
                    }
                } else if kind.ends_with("/control") && ctl.ax.is_none() {
                    ctl.kind = "ActiveX".into();
                    ctl.ax = Some(activex(parts, part));
                }
            }
        }
        // Embedded objects (OLE): where they are and their pictures.
        if let Some(root) = parts.xml(&format!("{sheet_part}#oleObjects")) {
            let mut found = Vec::new();
            root.find_all("oleObject", &mut found);
            for o in found {
                let Some(id) = o.attr("shapeId").and_then(spid_number) else {
                    continue;
                };
                let at = match list.iter().position(|x| x.id == id) {
                    Some(i) => i,
                    None => {
                        list.push(Control {
                            id,
                            kind: "Pict".into(),
                            ..Default::default()
                        });
                        list.len() - 1
                    }
                };
                let ctl = &mut list[at];
                if ctl.name.is_empty() {
                    ctl.name = o.attr("progId").unwrap_or("").to_string();
                }
                if let Some(pr) = o.child("objectPr") {
                    if ctl.anchor.is_none()
                        && let Some(a) = pr
                            .find("anchor")
                            .map(anchor_of)
                            .filter(|a| a.from.is_some())
                    {
                        ctl.anchor = Some(a);
                    }
                    if ctl.image.is_none()
                        && let Some((_, part)) = pr.attr("id").and_then(|r| rels.get(r))
                    {
                        ctl.image = Some(part.clone());
                    }
                }
            }
        }
        // The VML drawing: anchors, texts, the state of form controls.
        let vml: Vec<String> = rels
            .values()
            .filter(|(t, _)| t.ends_with("/vmlDrawing"))
            .map(|(_, p)| p.clone())
            .collect();
        for part in vml {
            let Some(root) = parts.xml(&part) else {
                continue;
            };
            let vml_rels = parts.rels(&part);
            let mut shapes = Vec::new();
            root.find_all("shape", &mut shapes);
            for sh in shapes {
                let Some(data) = sh.child("ClientData") else {
                    continue;
                };
                let object = data.attr("ObjectType").unwrap_or("");
                // Comments are not drawn (the grid shows cells, not their notes).
                if object == "Note" {
                    continue;
                }
                let Some(id) = sh
                    .attr("spid")
                    .or_else(|| sh.attr("id"))
                    .and_then(spid_number)
                else {
                    continue;
                };
                let at = match list.iter().position(|x| x.id == id) {
                    Some(i) => i,
                    None => {
                        list.push(Control {
                            id,
                            ..Default::default()
                        });
                        list.len() - 1
                    }
                };
                let ctl = &mut list[at];
                if ctl.kind.is_empty() {
                    ctl.kind = form_kind(object).to_string();
                }
                ctl.client = data
                    .children
                    .iter()
                    .map(|c| (c.name.clone(), c.all_text().trim().to_string()))
                    .collect();
                ctl.hidden = sh
                    .attr("style")
                    .is_some_and(|s| s.replace(' ', "").contains("visibility:hidden"));
                if ctl.anchor.is_none() {
                    ctl.anchor = ctl.client.get("Anchor").and_then(|a| vml_anchor(a));
                }
                if let Some(tb) = sh.child("textbox") {
                    ctl.text = Some(vml_text(tb, &ctl.client));
                }
                // The picture of an object or an ActiveX control.
                if ctl.image.is_none()
                    && let Some((_, p)) = sh
                        .child("imagedata")
                        .and_then(|d| d.attr("relid"))
                        .and_then(|r| vml_rels.get(r))
                {
                    ctl.image = Some(p.clone());
                }
            }
        }
        Controls {
            list,
            font: font.clone(),
        }
    }

    /// The control `id` drawn at `anchor` (the place of its mirror in the drawing).
    pub(crate) fn take(
        &mut self,
        id: u32,
        anchor: Option<Anchor>,
        mirror: Option<&Node>,
        env: &Env,
        ctx: &Ctx,
    ) -> Option<Obj> {
        let i = self.list.iter().position(|c| c.id == id && !c.done)?;
        self.list[i].done = true;
        if let Some(a) = anchor.filter(|a| a.from.is_some() || a.pos.is_some()) {
            self.list[i].anchor = Some(a);
        }
        if let Some(sp) = mirror {
            let c = &mut self.list[i];
            if c.name.is_empty() {
                c.name = sp
                    .find("cNvPr")
                    .and_then(|x| x.attr("name"))
                    .unwrap_or("")
                    .to_string();
            }
            // The text of the mirror (DrawingML) is read better than VML's.
            if c.kind != "ActiveX"
                && let Some(t) = sp.child("txBody").and_then(|t| mirror_text(t, ctx))
            {
                let old = c.text.take().unwrap_or_default();
                c.text = Some(Text {
                    align: t.align.or(old.align),
                    valign: old.valign.or(t.valign),
                    ..t
                });
            }
            if c.client.get("FmlaMacro").is_none_or(String::is_empty)
                && let Some(m) = sp.attr("macro").filter(|m| !m.is_empty())
            {
                c.client.insert("FmlaMacro".into(), m.to_string());
            }
        }
        self.draw(i, env, ctx)
    }

    /// The controls not drawn yet (no mirror in the drawing), in the order of the sheet.
    pub(crate) fn rest(&mut self, env: &Env, ctx: &Ctx) -> Vec<Obj> {
        let mut out = Vec::new();
        for i in 0..self.list.len() {
            if self.list[i].done {
                continue;
            }
            self.list[i].done = true;
            if let Some(o) = self.draw(i, env, ctx) {
                out.push(o);
            }
        }
        out
    }

    fn draw(&self, i: usize, env: &Env, ctx: &Ctx) -> Option<Obj> {
        let c = &self.list[i];
        if c.hidden {
            return None;
        }
        let anchor = c.anchor.clone()?;
        let picture = || {
            c.image
                .as_ref()
                .and_then(|p| Some(super::drawing::picture_of(ctx, p, ctx.parts.bytes(p)?, "")))
        };
        let html = match &c.ax {
            // An ActiveX control the grid draws itself; another one: its preview.
            Some(ax) if !ax.class.is_empty() => activex_html(c, ax, &self.font),
            Some(ax) => picture().unwrap_or_else(|| activex_html(c, ax, &self.font)),
            // An embedded object: its picture.
            None if c.kind == "Pict" => picture()?,
            None => form_html(c, env, &self.font)?,
        };
        let macro_name = c
            .client
            .get("FmlaMacro")
            .map(|m| m.trim_start_matches("[0]!").to_string())
            .unwrap_or_default();
        Some(Obj {
            anchor,
            html,
            name: c.name.clone(),
            macro_name,
            kind: "control",
            chart: None,
            activex: c.ax.as_ref().map(msforms_type),
            // Embedded objects run no macro.
            assignable: c.ax.is_none() && c.kind != "Pict",
            shape_id: Some(c.id),
        })
    }
}

/// The object type of `formControlPr` for a VML `ObjectType`.
fn form_kind(vml: &str) -> &str {
    match vml {
        "Checkbox" => "CheckBox",
        "Radio" => "Radio",
        "Drop" => "Drop",
        "List" => "List",
        "Spin" => "Spin",
        "Scroll" => "Scroll",
        "GBox" => "GBox",
        "Edit" => "EditBox",
        "Pict" => "Pict",
        other => other,
    }
}

/// A VML `x:Anchor` (columns and rows with offsets in pixels) as a drawing anchor.
fn vml_anchor(s: &str) -> Option<Anchor> {
    let v: Vec<i64> = s
        .split(',')
        .map(|x| x.trim().parse::<i64>())
        .collect::<Result<_, _>>()
        .ok()?;
    if v.len() != 8 || v.iter().any(|x| *x < 0) {
        return None;
    }
    let px = |p: i64| (p as f64 * EMU_PX) as i64;
    Some(Anchor {
        from: Some((v[0] as u32, px(v[1]), v[2] as u32, px(v[3]))),
        to: Some((v[4] as u32, px(v[5]), v[6] as u32, px(v[7]))),
        ..Default::default()
    })
}

fn align_of(s: &str) -> Option<&'static str> {
    match s.to_ascii_lowercase().as_str() {
        "left" | "l" | "justify" | "distributed" => Some("left"),
        "center" | "ctr" | "centered" => Some("center"),
        "right" | "r" => Some("right"),
        "top" | "t" => Some("top"),
        "bottom" | "b" => Some("bottom"),
        _ => None,
    }
}

/// The text of a VML text box: its lines (`div`) and the font of its first run.
fn vml_text(tb: &Node, client: &HashMap<String, String>) -> Text {
    let mut divs = Vec::new();
    tb.find_all("div", &mut divs);
    let lines: Vec<String> = if divs.is_empty() {
        vec![tb.all_text()]
    } else {
        divs.iter().map(|d| d.all_text()).collect()
    };
    let mut t = Text {
        text: lines.join("\n").trim_matches(['\r', '\n']).to_string(),
        ..Default::default()
    };
    if let Some(f) = tb.find("font") {
        t.font = f.attr("face").map(str::to_string);
        // In twentieths of a point.
        t.size = f
            .attr("size")
            .and_then(|v| v.parse::<f64>().ok())
            .filter(|v| *v > 0.0)
            .map(|v| v / 20.0);
        t.color = f
            .attr("color")
            .filter(|c| c.starts_with('#') && c.len() == 7)
            .map(str::to_string);
        t.bold = f.find("b").is_some();
        t.italic = f.find("i").is_some();
        t.underline = f.find("u").is_some();
        t.strike = f.find("s").is_some();
    }
    let style = divs.first().and_then(|d| d.attr("style")).unwrap_or("");
    t.align = style
        .split(';')
        .find_map(|p| p.trim().strip_prefix("text-align:"))
        .and_then(|a| align_of(a.trim()));
    if let Some(a) = client.get("TextHAlign").and_then(|a| align_of(a)) {
        t.align = Some(a);
    }
    t.valign = client.get("TextVAlign").and_then(|a| align_of(a));
    t
}

/// The text of a control's mirror (txBody): its paragraphs and the font of its first run.
fn mirror_text(tx: &Node, ctx: &Ctx) -> Option<Text> {
    let mut lines = Vec::new();
    let mut first: Option<&Node> = None;
    for p in tx.all("p") {
        let mut line = String::new();
        for r in &p.children {
            match r.name.as_str() {
                "r" | "fld" => {
                    line.push_str(&r.child("t").map(|t| t.text.clone()).unwrap_or_default());
                    first = first.or(r.child("rPr"));
                }
                "br" => line.push('\n'),
                _ => {}
            }
        }
        lines.push(line);
    }
    let text = lines.join("\n");
    if text.trim().is_empty() {
        return None;
    }
    let mut t = Text {
        text,
        ..Default::default()
    };
    let algn = tx
        .all("p")
        .next()
        .and_then(|p| p.child("pPr"))
        .and_then(|p| p.attr("algn"));
    t.align = algn.and_then(align_of);
    t.valign = tx
        .child("bodyPr")
        .and_then(|b| b.attr("anchor"))
        .and_then(align_of);
    if let Some(r) = first {
        t.size = r
            .attr("sz")
            .and_then(|v| v.parse::<f64>().ok())
            .map(|v| v / 100.0);
        t.bold = matches!(r.attr("b"), Some("1" | "true"));
        t.italic = matches!(r.attr("i"), Some("1" | "true"));
        t.underline = r.attr("u").is_some_and(|u| u != "none");
        t.strike = r.attr("strike").is_some_and(|s| s != "noStrike");
        t.font = ["ea", "latin"].iter().find_map(|k| {
            r.child(k)
                .and_then(|c| c.attr("typeface"))
                .filter(|f| !f.is_empty() && !f.starts_with('+'))
                .map(str::to_string)
        });
        t.color = r.child("solidFill").and_then(|f| color_of(f, &ctx.theme));
    }
    Some(t)
}

/// The CSS of a text's font (what it does not say: `default`'s name and size).
fn text_css(t: &Text, default: &(String, f64)) -> String {
    let mut css = format!(
        "font-family:{};font-size:{}pt;",
        esc(&font_family(t.font.as_deref().unwrap_or(&default.0))),
        n(t.size.unwrap_or(default.1))
    );
    if let Some(c) = &t.color {
        let _ = write!(css, "color:{c};");
    }
    if t.bold {
        css.push_str("font-weight:bold;");
    }
    if t.italic {
        css.push_str("font-style:italic;");
    }
    match (t.underline, t.strike) {
        (true, true) => css.push_str("text-decoration:underline line-through;"),
        (true, false) => css.push_str("text-decoration:underline;"),
        (false, true) => css.push_str("text-decoration:line-through;"),
        _ => {}
    }
    css
}

/// The flex alignment of a text box: (justify-content across, align-items down).
fn flex(align: Option<&str>, valign: Option<&str>, default: (&str, &str)) -> String {
    let h = match align.unwrap_or(default.0) {
        "center" => "center",
        "right" => "flex-end",
        _ => "flex-start",
    };
    let v = match valign.unwrap_or(default.1) {
        "center" => "center",
        "bottom" => "flex-end",
        _ => "flex-start",
    };
    let t = match align.unwrap_or(default.0) {
        "center" => "center",
        "right" => "right",
        _ => "left",
    };
    format!("justify-content:{h};align-items:{v};text-align:{t};")
}

const CHECK: &str = "<svg viewBox=\"0 0 12 12\"><path d=\"M2.5 6.2l2.4 2.4 4.6-5\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"1.6\"/></svg>";

fn check_html(state: u8, round: bool) -> String {
    let mark = match (state, round) {
        (0, _) => "",
        (_, true) => "<i class=\"dot\"></i>",
        (1, false) => CHECK,
        _ => "<i class=\"mixed\"></i>",
    };
    format!(
        "<span class=\"{}\">{mark}</span>",
        if round { "radio" } else { "box" }
    )
}

/// The arrow of a scroll bar's or spin button's button, pointing `dir` (up, down, left, right).
fn arrow(dir: &str) -> String {
    let d = match dir {
        "up" => "M2 7h6L5 3z",
        "down" => "M2 4h6L5 8z",
        "left" => "M7 2v6L3 5z",
        _ => "M4 2v6l4-3z",
    };
    format!("<span class=\"btn\"><svg viewBox=\"0 0 10 10\"><path d=\"{d}\"/></svg></span>")
}

/// A scroll bar with its thumb at `pos` from `min` to `max`.
fn scroll_html(horiz: bool, min: i64, max: i64, pos: i64) -> String {
    let span = (max - min).max(0) as f64;
    let at = if span > 0.0 {
        ((pos - min) as f64 / span).clamp(0.0, 1.0)
    } else {
        0.0
    };
    let (a, b, dir) = if horiz {
        ("left", "right", "h")
    } else {
        ("up", "down", "v")
    };
    format!(
        "<div class=\"ctl scroll {dir}\">{}<span class=\"track\"><i class=\"thumb\" style=\"--at:{}\"></i></span>{}</div>",
        arrow(a),
        n(at),
        arrow(b)
    )
}

fn spin_html(horiz: bool) -> String {
    if horiz {
        format!(
            "<div class=\"ctl spin h\">{}{}</div>",
            arrow("left"),
            arrow("right")
        )
    } else {
        format!(
            "<div class=\"ctl spin v\">{}{}</div>",
            arrow("up"),
            arrow("down")
        )
    }
}

/// The HTML of a form control; None for kinds that are not drawn (dialog frames and the like).
fn form_html(c: &Control, env: &Env, font: &(String, f64)) -> Option<String> {
    let prop = |k: &str| c.props.get(k).map(String::as_str).filter(|v| !v.is_empty());
    let client = |k: &str| {
        c.client
            .get(k)
            .map(String::as_str)
            .filter(|v| !v.is_empty())
    };
    let num = |p: &str, v: &str| {
        prop(p)
            .or_else(|| client(v))
            .and_then(|x| x.parse::<i64>().ok())
    };
    let text = c.text.clone().unwrap_or_default();
    let label = esc(&text.text);
    let linked = prop("fmlaLink")
        .or_else(|| client("FmlaLink"))
        .and_then(|r| (env.value)(r));
    let html = match c.kind.as_str() {
        "Button" => format!(
            "<div class=\"ctl button\" style=\"{}{}\"><span>{label}</span></div>",
            flex(text.align, text.valign, ("center", "center")),
            text_css(&text, font)
        ),
        "CheckBox" | "Radio" => {
            // The state of the linked cell, else the one saved.
            let saved = match prop("checked").or_else(|| client("Checked")) {
                Some("Checked" | "1") => 1,
                Some("Mixed" | "2") => 2,
                _ => 0,
            };
            let state = match linked.as_deref().map(str::to_ascii_uppercase).as_deref() {
                Some("TRUE") if c.kind == "CheckBox" => 1,
                Some("FALSE") if c.kind == "CheckBox" => 0,
                Some("#N/A") if c.kind == "CheckBox" => 2,
                _ => saved,
            };
            format!(
                "<div class=\"ctl check{}\" style=\"{}{}\">{}<span class=\"label\">{label}</span></div>",
                if prop("noThreeD").is_some() || c.client.contains_key("NoThreeD") {
                    " flat"
                } else {
                    ""
                },
                flex(text.align, text.valign, ("left", "center")),
                text_css(&text, font),
                check_html(state, c.kind == "Radio")
            )
        }
        "Label" => format!(
            "<div class=\"ctl label\" style=\"{}{}\"><span>{label}</span></div>",
            flex(text.align, text.valign, ("left", "top")),
            text_css(&text, font)
        ),
        "GBox" => format!(
            "<div class=\"ctl gbox\" style=\"{}\"><i class=\"frame\"></i><span class=\"label\">{label}</span></div>",
            text_css(&text, font)
        ),
        "EditBox" => format!(
            "<div class=\"ctl edit\" style=\"{}\"><span>{label}</span></div>",
            text_css(&text, font)
        ),
        "List" | "Drop" => {
            let items = prop("fmlaRange")
                .or_else(|| client("FmlaRange"))
                .map(|r| (env.texts)(r))
                .unwrap_or_default();
            let sel = linked
                .as_deref()
                .and_then(|v| v.parse::<f64>().ok())
                .map(|v| v as i64)
                .or_else(|| num("sel", "Sel"))
                .unwrap_or(0);
            let css = text_css(
                &Text {
                    size: Some(font.1.min(11.0)),
                    ..Default::default()
                },
                font,
            );
            if c.kind == "Drop" {
                let shown = usize::try_from(sel - 1)
                    .ok()
                    .and_then(|i| items.get(i))
                    .map(|s| esc(s))
                    .unwrap_or_default();
                format!(
                    "<div class=\"ctl drop\" style=\"{css}\"><span class=\"value\">{shown}</span>{}</div>",
                    arrow("down")
                )
            } else {
                let mut list = String::new();
                for (i, item) in items.iter().take(200).enumerate() {
                    let _ = write!(
                        list,
                        "<span{}>{}</span>",
                        if i as i64 + 1 == sel {
                            " class=\"sel\""
                        } else {
                            ""
                        },
                        esc(item)
                    );
                }
                format!("<div class=\"ctl list\" style=\"{css}\">{list}</div>")
            }
        }
        "Spin" | "Scroll" => {
            let horiz = prop("horiz").is_some_and(|h| h == "1" || h == "true")
                || c.client.contains_key("Horiz");
            if c.kind == "Spin" {
                spin_html(horiz)
            } else {
                let pos = linked
                    .as_deref()
                    .and_then(|v| v.parse::<f64>().ok())
                    .map(|v| v as i64)
                    .or_else(|| num("val", "Val"))
                    .unwrap_or(0);
                scroll_html(
                    horiz,
                    num("min", "Min").unwrap_or(0),
                    num("max", "Max").unwrap_or(100),
                    pos,
                )
            }
        }
        _ => return None,
    };
    Some(html)
}

/// An OLE_COLOR as CSS: a system color (0x80 in the high byte; as Windows 10 has them) or
/// 0x00BBGGRR.
fn ole_color(v: u32) -> String {
    if v & 0xff00_0000 == 0x8000_0000 {
        return match v & 0xff {
            0x00 => "#c8c8c8",
            0x04 | 0x0f => "#f0f0f0",
            0x05 | 0x0e | 0x14 => "#ffffff",
            0x06 => "#646464",
            0x0a => "#b4b4b4",
            0x0b => "#f4f7fc",
            0x0c => "#ababab",
            0x0d => "#0078d7",
            0x10 => "#a0a0a0",
            0x11 => "#6d6d6d",
            0x15 => "#696969",
            0x16 => "#e3e3e3",
            0x18 => "#ffffe1",
            _ => "#000000",
        }
        .to_string();
    }
    format!(
        "#{:02x}{:02x}{:02x}",
        v & 0xff,
        (v >> 8) & 0xff,
        (v >> 16) & 0xff
    )
}

/// The HTML of an ActiveX control.
fn activex_html(c: &Control, ax: &Ax, default: &(String, f64)) -> String {
    // The default font of MSForms controls in Excel.
    let fallback = (default.0.clone(), 9.0);
    let font = Text {
        font: ax.font.clone(),
        size: ax.size,
        color: ax.fore.clone(),
        bold: ax.bold,
        italic: ax.italic,
        underline: ax.underline,
        ..Default::default()
    };
    let css = text_css(&font, &fallback);
    let caption = esc(ax.caption.as_deref().unwrap_or(&c.name));
    let back = |d: &str| format!("background:{};", ax.back.as_deref().unwrap_or(d));
    let value = esc(ax.value.as_deref().unwrap_or(""));
    let on = matches!(
        ax.value.as_deref().map(str::to_ascii_lowercase).as_deref(),
        Some("1" | "-1" | "true")
    );
    match (ax.class, ax.display) {
        ("CommandButton", _) | ("MorphData", 6) => format!(
            "<div class=\"ctl button{}\" style=\"{}{}{}\"><span>{caption}</span></div>",
            if ax.class != "CommandButton" && on {
                " down"
            } else {
                ""
            },
            flex(ax.align.or(Some("center")), None, ("center", "center")),
            css,
            back("#f0f0f0")
        ),
        ("Label", _) => format!(
            "<div class=\"ctl label\" style=\"{}{}{}\"><span>{caption}</span></div>",
            flex(ax.align, None, ("left", "top")),
            css,
            if ax.opaque {
                back("#f0f0f0")
            } else {
                String::new()
            }
        ),
        ("MorphData", 4 | 5) => format!(
            "<div class=\"ctl check flat\" style=\"{}{}{}\">{}<span class=\"label\">{caption}</span></div>",
            flex(ax.align, None, ("left", "center")),
            css,
            if ax.opaque {
                back("#ffffff")
            } else {
                String::new()
            },
            check_html(u8::from(on), ax.display == 5)
        ),
        ("MorphData", 3 | 7) => format!(
            "<div class=\"ctl drop\" style=\"{}{}\"><span class=\"value\">{value}</span>{}</div>",
            css,
            back("#ffffff"),
            arrow("down")
        ),
        ("MorphData", 2) => format!(
            "<div class=\"ctl list\" style=\"{}{}\"></div>",
            css,
            back("#ffffff")
        ),
        ("MorphData", _) => format!(
            "<div class=\"ctl edit\" style=\"{}{}\"><span>{value}</span></div>",
            css,
            back("#ffffff")
        ),
        ("SpinButton", _) => spin_html(ax.orientation == 1),
        ("ScrollBar", _) => scroll_html(ax.orientation == 1, ax.min, ax.max, ax.position),
        _ => format!(
            "<div class=\"ctl other\"><span>{}</span></div>",
            esc(&c.name)
        ),
    }
}

/// The class of an ActiveX control by its CLSID, as the grid draws it, with MorphData's
/// DisplayStyle (TextBox 1, ListBox 2, ComboBox 3, CheckBox 4, OptionButton 5, ToggleButton 6).
fn class_of(clsid: &str) -> (&'static str, u8) {
    match clsid.trim_matches(['{', '}']).to_ascii_uppercase().as_str() {
        "D7053240-CE69-11CD-A777-00DD01143C57" => ("CommandButton", 0),
        "978C9E23-D4B0-11CE-BF2D-00AA003F40D0" => ("Label", 0),
        "8BD21D10-EC42-11CE-9E0D-00AA006002F3" => ("MorphData", 1),
        "8BD21D20-EC42-11CE-9E0D-00AA006002F3" => ("MorphData", 2),
        "8BD21D30-EC42-11CE-9E0D-00AA006002F3" => ("MorphData", 3),
        "8BD21D40-EC42-11CE-9E0D-00AA006002F3" => ("MorphData", 4),
        "8BD21D50-EC42-11CE-9E0D-00AA006002F3" => ("MorphData", 5),
        "8BD21D60-EC42-11CE-9E0D-00AA006002F3" => ("MorphData", 6),
        "79176FB0-B7F2-11CE-97EF-00AA006D2776" => ("SpinButton", 0),
        "DFD181E0-5E2F-11CE-A449-00AA004A803D" => ("ScrollBar", 0),
        _ => ("", 0),
    }
}

/// The Microsoft Forms 2.0 type of an ActiveX control (`TextBox`; empty for another control).
fn msforms_type(ax: &Ax) -> &'static str {
    match (ax.class, ax.display) {
        ("MorphData", 1) => "TextBox",
        ("MorphData", 2) => "ListBox",
        ("MorphData", 3) => "ComboBox",
        ("MorphData", 4) => "CheckBox",
        ("MorphData", 5) => "OptionButton",
        ("MorphData", 6) => "ToggleButton",
        ("MorphData", _) => "",
        (class, _) => class,
    }
}

/// What an ActiveX control (`xl/activeX/activeXN.xml`) says of itself.
fn activex(parts: &Parts, part: &str) -> Ax {
    let Some(ocx) = parts.xml(part) else {
        return Ax::default();
    };
    let (class, display) = class_of(ocx.attr("classid").unwrap_or(""));
    let mut ax = Ax {
        class,
        display,
        max: 100,
        orientation: -1,
        ..Default::default()
    };
    // persistPropertyBag: the properties as they are.
    for p in ocx.all("ocxPr") {
        let v = p.attr("value").unwrap_or("").to_string();
        match p.attr("name").unwrap_or("") {
            "Caption" => ax.caption = Some(v),
            "Value" => ax.value = Some(v),
            _ => {}
        }
    }
    let bin = ocx
        .attr("id")
        .and_then(|id| parts.rels(part).remove(id))
        .and_then(|(_, p)| parts.bytes(&p).map(<[u8]>::to_vec));
    if let Some(data) = bin {
        // persistStream: the CLSID, then the control's stream (MS-OFORMS); a storage (CFB) is not
        // read.
        let stream = if data.len() > 16 && data[16..].starts_with(&[0, 2]) {
            &data[16..]
        } else {
            &data[..]
        };
        read_stream(&mut ax, stream);
    }
    ax
}

/// A field of a record's DataBlock: (mask bit, name, size).
type Field = (u8, &'static str, usize);

/// How a class's record is read: its fields, whether its mask has 64 bits, the strings and sizes
/// of its ExtraDataBlock (mask bit, name), the mask bits of its pictures.
type Spec = (
    &'static [Field],
    bool,
    &'static [(u8, &'static str)],
    &'static [u8],
);

const COMMAND_BUTTON: &[Field] = &[
    (0, "ForeColor", 4),
    (1, "BackColor", 4),
    (2, "VariousPropertyBits", 4),
    (3, "Caption", 4),
    (4, "PicturePosition", 4),
    (6, "MousePointer", 1),
    (7, "Picture", 2),
    (8, "Accelerator", 2),
    (10, "MouseIcon", 2),
];
const LABEL: &[Field] = &[
    (0, "ForeColor", 4),
    (1, "BackColor", 4),
    (2, "VariousPropertyBits", 4),
    (3, "Caption", 4),
    (4, "PicturePosition", 4),
    (6, "MousePointer", 1),
    (7, "BorderColor", 4),
    (8, "BorderStyle", 2),
    (9, "SpecialEffect", 2),
    (10, "Picture", 2),
    (11, "Accelerator", 2),
    (12, "MouseIcon", 2),
];
const MORPH_DATA: &[Field] = &[
    (0, "VariousPropertyBits", 4),
    (1, "BackColor", 4),
    (2, "ForeColor", 4),
    (3, "MaxLength", 4),
    (4, "BorderStyle", 1),
    (5, "ScrollBars", 1),
    (6, "DisplayStyle", 1),
    (7, "MousePointer", 1),
    (9, "PasswordChar", 2),
    (10, "ListWidth", 4),
    (11, "BoundColumn", 2),
    (12, "TextColumn", 2),
    (13, "ColumnCount", 2),
    (14, "ListRows", 2),
    (15, "cColumnInfo", 2),
    (16, "MatchEntry", 1),
    (17, "ListStyle", 1),
    (18, "ShowDropButtonWhen", 1),
    (20, "DropButtonStyle", 1),
    (21, "MultiSelect", 1),
    (22, "Value", 4),
    (23, "Caption", 4),
    (24, "PicturePosition", 4),
    (25, "BorderColor", 4),
    (26, "SpecialEffect", 4),
    (27, "MouseIcon", 2),
    (28, "Picture", 2),
    (29, "Accelerator", 2),
    (32, "GroupName", 4),
];
const SPIN_BUTTON: &[Field] = &[
    (0, "ForeColor", 4),
    (1, "BackColor", 4),
    (2, "VariousPropertyBits", 4),
    (5, "Min", 4),
    (6, "Max", 4),
    (7, "Position", 4),
    (8, "PrevEnabled", 4),
    (9, "NextEnabled", 4),
    (10, "SmallChange", 4),
    (11, "Orientation", 4),
];
const SCROLL_BAR: &[Field] = &[
    (0, "ForeColor", 4),
    (1, "BackColor", 4),
    (2, "VariousPropertyBits", 4),
    (4, "MousePointer", 1),
    (5, "Min", 4),
    (6, "Max", 4),
    (7, "Position", 4),
    (9, "PrevEnabled", 4),
    (10, "NextEnabled", 4),
    (11, "SmallChange", 4),
    (12, "LargeChange", 4),
    (13, "Orientation", 4),
];
const TEXT_PROPS: &[Field] = &[
    (0, "FontName", 4),
    (1, "FontEffects", 4),
    (2, "FontHeight", 4),
    (4, "FontCharSet", 1),
    (5, "FontPitchAndFamily", 1),
    (6, "ParagraphAlign", 1),
    (7, "FontWeight", 2),
];

/// A record read: its values, its strings (in the order of `strings`), where it ends.
struct Record {
    mask: u64,
    values: HashMap<&'static str, u32>,
    strings: HashMap<&'static str, String>,
    /// Where the StreamData after the record starts.
    end: usize,
}

/// Reads a record of MS-OFORMS (minor 0, major 2): its DataBlock, then the strings of its
/// ExtraDataBlock (`extra`: the strings and sizes in their order, a size as "#Size").
fn record(
    data: &[u8],
    fields: &[Field],
    mask64: bool,
    extra: &[(u8, &'static str)],
) -> Option<Record> {
    let u16_at = |i: usize| data.get(i..i + 2).map(|b| u16::from_le_bytes([b[0], b[1]]));
    let u32_at = |i: usize| {
        data.get(i..i + 4)
            .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
    };
    if data.len() < 8 || data[0] != 0 || data[1] != 2 {
        return None;
    }
    let cb = u16_at(2)? as usize;
    let end = 4 + cb;
    if end > data.len() {
        return None;
    }
    let mut mask = u32_at(4)? as u64;
    let mut pos: usize = 8;
    if mask64 {
        mask |= (u32_at(8)? as u64) << 32;
        pos = 12;
    }
    let mut values = HashMap::new();
    for &(bit, name, size) in fields {
        if mask & (1u64 << bit) == 0 {
            continue;
        }
        pos = pos.next_multiple_of(size);
        let v = match size {
            1 => *data.get(pos)? as u32,
            2 => u16_at(pos)? as u32,
            _ => u32_at(pos)?,
        };
        values.insert(name, v);
        pos += size;
    }
    pos = pos.next_multiple_of(4);
    let mut strings = HashMap::new();
    for &(bit, name) in extra {
        if mask & (1u64 << bit) == 0 {
            continue;
        }
        if name == "#Size" {
            pos += 8;
            continue;
        }
        let packed = values.get(name).copied().unwrap_or(0);
        let len = (packed & 0x7fff_ffff) as usize;
        let raw = data.get(pos..pos.checked_add(len)?)?;
        let s = if packed & 0x8000_0000 != 0 {
            raw.iter().map(|b| *b as char).collect()
        } else {
            let units: Vec<u16> = raw
                .chunks_exact(2)
                .map(|c| u16::from_le_bytes([c[0], c[1]]))
                .collect();
            String::from_utf16_lossy(&units)
        };
        strings.insert(name, s);
        pos = (pos + len).next_multiple_of(4);
    }
    if pos > end {
        return None;
    }
    Some(Record {
        mask,
        values,
        strings,
        end,
    })
}

/// Skips the pictures (GuidAndPicture) of a record's StreamData; returns where they end.
fn skip_pictures(data: &[u8], mut pos: usize, count: usize) -> Option<usize> {
    for _ in 0..count {
        let size = data
            .get(pos + 20..pos + 24)
            .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]))? as usize;
        pos = pos.checked_add(24 + size)?;
        if pos > data.len() {
            return None;
        }
    }
    Some(pos)
}

/// The control's stream: its caption, value, colors and font.
fn read_stream(ax: &mut Ax, data: &[u8]) {
    const SIZE: &str = "#Size";
    let (fields, mask64, extra, pictures): Spec = match ax.class {
        "CommandButton" => (
            COMMAND_BUTTON,
            false,
            &[(3, "Caption"), (5, SIZE)],
            &[7, 10],
        ),
        "Label" => (LABEL, false, &[(3, "Caption"), (5, SIZE)], &[10, 12]),
        "MorphData" => (
            MORPH_DATA,
            true,
            &[(8, SIZE), (22, "Value"), (23, "Caption"), (32, "GroupName")],
            &[27, 28],
        ),
        "SpinButton" => (SPIN_BUTTON, false, &[(3, SIZE)], &[]),
        "ScrollBar" => (SCROLL_BAR, false, &[(3, SIZE)], &[]),
        _ => return,
    };
    let Some(r) = record(data, fields, mask64, extra) else {
        return;
    };
    let v = |k: &str| r.values.get(k).copied();
    ax.caption = r.strings.get("Caption").cloned().or(ax.caption.take());
    ax.value = r.strings.get("Value").cloned().or(ax.value.take());
    ax.fore = v("ForeColor").map(ole_color);
    ax.back = v("BackColor").map(ole_color);
    // fBackStyle: opaque unless said otherwise (the default bits of each class have it).
    ax.opaque = v("VariousPropertyBits").is_none_or(|b| b & 0x8 != 0);
    if ax.class == "MorphData" {
        ax.display = v("DisplayStyle").map(|d| d as u8).unwrap_or(ax.display);
    }
    let signed = |k: &str| v(k).map(|x| x as i32 as i64);
    ax.min = signed("Min").unwrap_or(0);
    ax.max = signed("Max").unwrap_or(if ax.class == "ScrollBar" { 32767 } else { 100 });
    ax.position = signed("Position").unwrap_or(0);
    ax.orientation = signed("Orientation").unwrap_or(-1);
    if ax.orientation == -1 {
        // Automatic: by the shape of the control, which is not known here; vertical.
        ax.orientation = 0;
    }
    // TextProps after the pictures.
    if !matches!(ax.class, "CommandButton" | "Label" | "MorphData") {
        return;
    }
    let count = pictures
        .iter()
        .filter(|b| r.mask & (1u64 << **b) != 0)
        .count();
    let Some(at) = skip_pictures(data, r.end, count) else {
        return;
    };
    let Some(t) = record(&data[at..], TEXT_PROPS, false, &[(0, "FontName")]) else {
        return;
    };
    ax.font = t.strings.get("FontName").cloned().filter(|s| !s.is_empty());
    ax.size = t
        .values
        .get("FontHeight")
        .map(|h| *h as f64 / 20.0)
        .filter(|s| *s > 0.0);
    if let Some(e) = t.values.get("FontEffects") {
        ax.bold = e & 1 != 0;
        ax.italic = e & 2 != 0;
        ax.underline = e & 4 != 0;
    }
    if let Some(w) = t.values.get("FontWeight") {
        ax.bold = *w > 550;
    }
    ax.align = match t.values.get("ParagraphAlign") {
        Some(1) => Some("left"),
        Some(2) => Some("right"),
        Some(3) => Some("center"),
        _ => None,
    };
}

/// The stream of a CommandButton (MS-OFORMS 2.2.1) with a caption and a font.
#[cfg(test)]
pub(crate) fn command_button_for_tests(caption: &str, font: &str, height: u32) -> Vec<u8> {
    let utf16 = |s: &str| {
        s.encode_utf16()
            .flat_map(|u| u.to_le_bytes())
            .collect::<Vec<u8>>()
    };
    let pad = |v: &mut Vec<u8>| {
        while v.len() % 4 != 0 {
            v.push(0);
        }
    };
    let cap = utf16(caption);
    // ForeColor, BackColor, VariousPropertyBits, Caption; Size.
    let mask: u32 = 0b10_1111;
    let mut block = Vec::new();
    block.extend(0x8000_0012u32.to_le_bytes());
    block.extend(0x8000_000Fu32.to_le_bytes());
    block.extend(0x0000_001Bu32.to_le_bytes());
    block.extend((cap.len() as u32).to_le_bytes());
    block.extend(&cap);
    pad(&mut block);
    block.extend(2000i32.to_le_bytes());
    block.extend(600i32.to_le_bytes());
    let mut out = vec![0, 2];
    out.extend(((4 + block.len()) as u16).to_le_bytes());
    out.extend(mask.to_le_bytes());
    out.extend(block);
    // TextProps: FontName, FontEffects, FontHeight, FontWeight.
    let name = utf16(font);
    let mut tp = Vec::new();
    tp.extend((name.len() as u32).to_le_bytes());
    tp.extend(1u32.to_le_bytes());
    tp.extend(height.to_le_bytes());
    tp.extend(700u16.to_le_bytes());
    pad(&mut tp);
    tp.extend(&name);
    pad(&mut tp);
    out.extend([0, 2]);
    out.extend(((4 + tp.len()) as u16).to_le_bytes());
    out.extend(0b1000_0111u32.to_le_bytes());
    out.extend(tp);
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn shape_ids() {
        assert_eq!(spid_number("_x0000_s1025"), Some(1025));
        assert_eq!(spid_number("2049"), Some(2049));
        assert_eq!(spid_number("CommandButton1"), Some(1));
        assert_eq!(spid_number("x"), None);
    }

    #[test]
    fn vml_anchors() {
        let a = vml_anchor(" 5, 10, 4, 0, 7, 0, 6, 12").unwrap();
        assert_eq!(a.from, Some((5, 95250, 4, 0)));
        assert_eq!(a.to, Some((7, 0, 6, 114300)));
        assert!(vml_anchor("1, 2, 3").is_none());
    }

    /// A CommandButton as Excel persists it: caption, colors and font.
    #[test]
    fn command_button_stream() {
        let data = command_button_for_tests("実行", "ＭＳ Ｐゴシック", 220);
        let mut ax = Ax {
            class: "CommandButton",
            ..Default::default()
        };
        read_stream(&mut ax, &data);
        assert_eq!(ax.caption.as_deref(), Some("実行"));
        assert_eq!(
            (ax.fore.as_deref(), ax.back.as_deref()),
            (Some("#000000"), Some("#f0f0f0"))
        );
        assert_eq!(
            (ax.font.as_deref(), ax.size),
            (Some("ＭＳ Ｐゴシック"), Some(11.0))
        );
        assert!(ax.bold);
    }
}
