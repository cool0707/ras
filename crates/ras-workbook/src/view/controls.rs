//! The controls of a sheet for the grid: form controls (buttons, check boxes, option buttons,
//! list and combo boxes, spin buttons, scroll bars, labels, group boxes) and ActiveX controls
//! (MSForms: CommandButton, CheckBox, TextBox, ...), drawn as Excel shows them.
//!
//! What a control is (its kind, where it is, its text, its state, the cells it is linked to, what
//! an ActiveX control persists) is read by umya-spreadsheet's `drawingml::controls`, shared with
//! xljs; this draws it.

use std::fmt::Write as _;

use umya_spreadsheet::drawingml::anchor::Anchor;
pub(crate) use umya_spreadsheet::drawingml::controls::spid_number;
use umya_spreadsheet::drawingml::controls::{
    self as read, ActiveX as Ax, CheckState, Control, ControlText as Text,
};

use super::drawing::{Ctx, Env, Obj, Parts, esc, font_family, n};
use super::xml::Node;

pub(crate) struct Controls {
    list: Vec<Control>,
    /// Those drawn already.
    done: Vec<bool>,
    font: (String, f64),
    /// An ActiveX control's settings in the sheet (`controlPr`: its linked cell and items) and
    /// in its data (its small and large change), by its shape.
    activex: std::collections::HashMap<u32, AxSettings>,
    /// The items of the list and combo boxes, by name, as the stand-in for Excel has them
    /// (`state`): drawn in place of those of the file.
    lists: std::collections::HashMap<String, Vec<String>>,
    /// The text boxes of the form controls (by their shapes): 余白 自動, the insets (points).
    boxes: std::collections::HashMap<u32, (bool, [f64; 4])>,
}

/// The insets of a VML text box (`inset`: left, top, right, bottom) in points; those it does
/// not give: VML's (0.1in, 0.05in).
fn vml_insets(inset: Option<&str>) -> [f64; 4] {
    let mut out = [7.2, 3.6, 7.2, 3.6];
    for (i, v) in inset.unwrap_or("").split(',').take(4).enumerate() {
        let v = v.trim();
        let (num, unit) = v
            .find(|c: char| c.is_ascii_alphabetic())
            .map_or((v, ""), |k| (&v[..k], &v[k..]));
        let Ok(n) = num.trim().parse::<f64>() else {
            continue;
        };
        out[i] = match unit {
            "pt" => n,
            "in" => n * 72.0,
            "mm" => n * 72.0 / 25.4,
            "cm" => n * 72.0 / 2.54,
            "px" => n * 0.75,
            // (A number alone: EMU.)
            _ => n / 12700.0,
        };
    }
    out
}

/// What an ActiveX control keeps besides what umya-spreadsheet reads.
#[derive(Clone, Default)]
struct AxSettings {
    link: String,
    items: String,
    small: Option<i64>,
    large: Option<i64>,
    /// `VariousPropertyBits` (MS-OFORMS): bit 1 is `Enabled`.
    bits: Option<i64>,
    group: Option<String>,
    /// Its font struck through (`FontEffects`).
    strike: bool,
    /// The sheet's settings of it (`controlPr`, its `anchor`): オブジェクトを印刷する, ロック,
    /// 代替テキスト, whether it moves and sizes with the cells (None: not said there).
    print: Option<bool>,
    locked: Option<bool>,
    alt: String,
    moves: Option<(bool, bool)>,
}

impl Controls {
    /// The controls of the sheet whose part is `sheet_part`.
    pub(crate) fn read(parts: &Parts, sheet_part: &str, font: &(String, f64)) -> Controls {
        use umya_spreadsheet::drawingml::package::Package;
        let list = read::controls(parts, sheet_part);
        let mut activex = std::collections::HashMap::new();
        let rels = parts.rels(sheet_part);
        if let Some(root) = parts.sheet_element(sheet_part, "controls") {
            let mut found = Vec::new();
            root.find_all("control", &mut found);
            for c in found {
                let (Some(id), Some((kind, part))) = (
                    c.attr("shapeId").and_then(|v| v.parse::<u32>().ok()),
                    c.attr("id").and_then(|r| rels.get(r)),
                ) else {
                    continue;
                };
                if activex.contains_key(&id) {
                    continue;
                }
                let pr = c.find("controlPr");
                let attr = |k: &str| pr.and_then(|p| p.attr(k)).unwrap_or("").to_string();
                // An ActiveX control's small and large change, ..., from its data (a form
                // control: its sheet's settings only).
                let ax = kind.ends_with("/control");
                let stream = parts.xml(part).filter(|_| ax).and_then(|ocx| {
                    let (class, _) = crate::activex::class_of(ocx.attr("classid").unwrap_or(""));
                    let bin = ocx.attr("id").and_then(|r| parts.rels(part).remove(r))?.1;
                    crate::activex::Stream::read(class, Parts::bytes(parts, &bin)?)
                });
                let n = |k: &str| stream.as_ref().and_then(|st| st.number(k));
                let flag = |k: &str| {
                    pr.and_then(|p| p.attr(k))
                        .map(|v| !matches!(v, "0" | "false"))
                };
                let anchor = pr.and_then(|p| p.find("anchor"));
                let moves = anchor.map(|a| {
                    let on = |k: &str| matches!(a.attr(k), Some("1" | "true"));
                    (on("moveWithCells"), on("sizeWithCells"))
                });
                activex.insert(
                    id,
                    AxSettings {
                        strike: stream
                            .as_ref()
                            .and_then(crate::activex::Stream::font)
                            .and_then(|f| f.number("FontEffects"))
                            .is_some_and(|e| e & crate::activex::STRIKE != 0),
                        print: flag("print"),
                        locked: flag("locked"),
                        alt: attr("altText"),
                        moves,
                        link: attr("linkedCell"),
                        items: attr("listFillRange"),
                        small: n("SmallChange"),
                        large: n("LargeChange"),
                        bits: n("VariousPropertyBits"),
                        group: stream
                            .as_ref()
                            .and_then(|st| st.string("GroupName"))
                            .filter(|g| !g.is_empty()),
                    },
                );
            }
        }
        // The text boxes of the form controls' VML shapes: 余白 自動 (`o:insetmode`), and their
        // insets (points).
        let mut boxes = std::collections::HashMap::new();
        let mut vml: Vec<&String> = rels
            .values()
            .filter(|(t, _)| t.ends_with("/vmlDrawing"))
            .map(|(_, p)| p)
            .collect();
        vml.sort();
        for part in vml {
            let Some(root) = parts.xml(part) else {
                continue;
            };
            let mut shapes = Vec::new();
            root.find_all("shape", &mut shapes);
            for sh in shapes {
                let (Some(id), Some(tb)) = (
                    sh.attr("spid")
                        .or_else(|| sh.attr("id"))
                        .and_then(spid_number),
                    sh.child("textbox"),
                ) else {
                    continue;
                };
                let auto = sh.attr("insetmode") != Some("custom");
                boxes.insert(id, (auto, vml_insets(tb.attr("inset"))));
            }
        }
        Controls {
            boxes,
            done: vec![false; list.len()],
            list,
            font: font.clone(),
            activex,
            lists: std::collections::HashMap::new(),
        }
    }

    /// What the controls are and hold, for the stand-in for Excel (the debug run's objects of
    /// the controls: `OLEObjects`, `Shapes`, `Sheet1.CheckBox1`): one object a control.
    pub(crate) fn state(&self) -> Vec<serde_json::Value> {
        use serde_json::json;
        let mut out = Vec::new();
        for c in &self.list {
            if c.kind == "Pict" {
                continue;
            }
            let mut j = match &c.activex {
                Some(ax) => {
                    let s = self.activex.get(&c.shape_id).cloned().unwrap_or_default();
                    json!({
                        "activex": true,
                        "type": ax.msforms_type(),
                        "caption": ax.caption,
                        "value": ax.value,
                        "min": ax.min,
                        "max": ax.max,
                        "position": ax.position,
                        "smallChange": s.small.unwrap_or(1),
                        "largeChange": s.large.unwrap_or(1),
                        "linkedCell": s.link,
                        "listFillRange": s.items,
                        "enabled": s.bits.is_none_or(|b| b & 2 != 0),
                        "groupName": s.group,
                    })
                }
                None => {
                    let (val, min, max) = c.value_range();
                    let num = |p: &str, v: &str, d: i64| {
                        c.property(p, v)
                            .and_then(|x| x.trim().parse::<i64>().ok())
                            .unwrap_or(d)
                    };
                    json!({
                        "activex": false,
                        "type": c.kind,
                        "text": c.text.as_ref().map(|t| t.text.clone()),
                        "checked": match c.checked() {
                            CheckState::Unchecked => 0,
                            CheckState::Checked => 1,
                            CheckState::Mixed => 2,
                        },
                        "sel": c.selected(),
                        "value": val,
                        "min": min,
                        "max": max,
                        "smallChange": num("inc", "Inc", 1),
                        "largeChange": num("page", "Page", 10),
                        "linkedCell": c.linked_cell().unwrap_or(""),
                        "listFillRange": c.list_range().unwrap_or(""),
                        "macro": c.macro_name().unwrap_or(""),
                        "enabled": !c.client.contains_key("Disabled"),
                    })
                }
            };
            j["name"] = json!(c.name);
            j["shapeId"] = json!(c.shape_id);
            j["visible"] = json!(!c.hidden);
            out.push(j);
        }
        out
    }

    /// Takes the state the stand-in for Excel gives the controls (by name: what `state` gives,
    /// changed by the program), in place of what the file holds.
    pub(crate) fn apply(&mut self, state: &serde_json::Value) {
        let Some(list) = state.as_array() else {
            return;
        };
        for s in list {
            let Some(name) = s["name"].as_str() else {
                continue;
            };
            let Some(c) = self.list.iter_mut().find(|c| c.name == name) else {
                continue;
            };
            if let Some(v) = s["visible"].as_bool() {
                c.hidden = !v;
            }
            let text = |k: &str| s[k].as_str().map(str::to_string);
            let int = |k: &str| s[k].as_f64().map(|v| v as i64);
            if let Some(items) = s["list"].as_array() {
                self.lists.insert(
                    name.to_string(),
                    items
                        .iter()
                        .map(|i| i.as_str().unwrap_or_default().to_string())
                        .collect(),
                );
            }
            match &mut c.activex {
                Some(ax) => {
                    if let Some(t) = text("caption") {
                        ax.caption = Some(t);
                    }
                    if !s["value"].is_null() {
                        ax.value = text("value");
                    }
                    if let Some(v) = int("position") {
                        ax.position = v;
                    }
                    if let Some(v) = int("min") {
                        ax.min = v;
                    }
                    if let Some(v) = int("max") {
                        ax.max = v;
                    }
                }
                None => {
                    if let Some(t) = text("text") {
                        c.text.get_or_insert_with(Default::default).text = t;
                    }
                    if let Some(v) = int("checked") {
                        let on = ["Unchecked", "Checked", "Mixed"][v.clamp(0, 2) as usize];
                        c.props.insert("checked".into(), on.into());
                    }
                    for (key, prop) in [
                        ("sel", "sel"),
                        ("value", "val"),
                        ("min", "min"),
                        ("max", "max"),
                    ] {
                        if let Some(v) = int(key) {
                            c.props.insert(prop.into(), v.to_string());
                        }
                    }
                }
            }
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
        let i = (0..self.list.len()).find(|&i| self.list[i].shape_id == id && !self.done[i])?;
        self.done[i] = true;
        if let Some(a) = anchor.filter(|a| a.from.is_some() || a.pos.is_some()) {
            self.list[i].anchor = Some(a);
        }
        if let Some(sp) = mirror {
            self.list[i].apply_mirror(sp, &ctx.theme);
        }
        self.draw(i, env, ctx)
    }

    /// The controls not drawn yet (no mirror in the drawing), in the order of the sheet.
    pub(crate) fn rest(&mut self, env: &Env, ctx: &Ctx) -> Vec<Obj> {
        let mut out = Vec::new();
        for i in 0..self.list.len() {
            if self.done[i] {
                continue;
            }
            self.done[i] = true;
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
        // The items of a list or combo box: the stand-in's, else those of its range.
        let items: Vec<String> = self.lists.get(&c.name).cloned().unwrap_or_else(|| {
            let range = match &c.activex {
                Some(_) => self.activex.get(&c.shape_id).map(|s| s.items.as_str()),
                None => c.list_range(),
            };
            range
                .filter(|r| !r.is_empty())
                .map(|r| (env.texts)(r))
                .unwrap_or_default()
        });
        let items = items.as_slice();
        let settings = self.activex.get(&c.shape_id).cloned().unwrap_or_default();
        let html = match &c.activex {
            // An ActiveX control the grid draws itself; another one: its preview.
            Some(ax) if !ax.class.is_empty() => activex_html(c, ax, &settings, &self.font, items),
            Some(ax) => {
                picture().unwrap_or_else(|| activex_html(c, ax, &settings, &self.font, items))
            }
            // An embedded object: its picture.
            None if c.kind == "Pict" => picture()?,
            None => form_html(c, env, &self.font, items)?,
        };
        // A form control's settings, and its text where it shows one.
        let form = (c.activex.is_none() && c.kind != "Pict").then(|| {
            let (val, min, max) = c.value_range();
            let num = |p: &str, v: &str, d: i64| {
                c.property(p, v)
                    .and_then(|x| x.trim().parse::<i64>().ok())
                    .unwrap_or(d)
            };
            serde_json::json!({
                "type": c.kind,
                "link": c.linked_cell().unwrap_or(""),
                "items": c.list_range().unwrap_or(""),
                "checked": c.checked() == CheckState::Checked,
                "sel": c.selected(),
                "val": val,
                "min": min,
                "max": max,
                "inc": num("inc", "Inc", 1),
                "page": num("page", "Page", 10),
                "list": items,
                "altText": settings.alt,
            })
        });
        // A button's font, 配置 and 余白 (as Excel's コントロールの書式設定 has them).
        let form = form.map(|mut j| {
            if c.kind == "Button" {
                let t = c.text.clone().unwrap_or_default();
                let raw = |k: &str, d: Option<&str>| {
                    c.client
                        .get(k)
                        .map(|v| v.trim().to_ascii_lowercase())
                        .filter(|v| !v.is_empty())
                        .or_else(|| d.map(str::to_string))
                };
                let (auto, inset) = self
                    .boxes
                    .get(&c.shape_id)
                    .copied()
                    .unwrap_or((true, [7.2, 3.6, 7.2, 3.6]));
                j["font"] = serde_json::json!({
                    "name": t.font.clone().unwrap_or_else(|| self.font.0.clone()),
                    "size": t.size.unwrap_or(self.font.1),
                    "bold": t.bold,
                    "italic": t.italic,
                    "underline": t.underline,
                    "strike": t.strike,
                });
                j["foreColor"] = serde_json::json!(t.color);
                j["textAlign"] =
                    serde_json::json!(raw("TextHAlign", t.align).unwrap_or_else(|| "left".into()));
                j["textVAlign"] =
                    serde_json::json!(raw("TextVAlign", t.valign).unwrap_or_else(|| "top".into()));
                j["insetAuto"] = serde_json::json!(auto);
                j["inset"] = serde_json::json!(inset.map(|v| (v * 100.0).round() / 100.0));
            }
            j
        });
        let shows_text = matches!(
            c.kind.as_str(),
            "Button" | "CheckBox" | "Radio" | "Label" | "GBox"
        );
        // An ActiveX control's: its Microsoft Forms type, its text (a caption, or the value of
        // a text, list or combo box), its state, linked cell, items, value in a range.
        let (form, text) = match (&c.activex, form) {
            (Some(ax), _) => {
                let forms = crate::activex::forms_type(ax.class, ax.display);
                let AxSettings {
                    link,
                    items: fill,
                    small: inc,
                    large: page,
                    bits,
                    strike,
                    alt,
                    ..
                } = settings.clone();
                let bit = |b: i64, d: bool| bits.map_or(d, |v| v & b != 0);
                let value_text = matches!(forms, "TextBox" | "ListBox" | "ComboBox");
                let text = if value_text {
                    ax.value.clone()
                } else {
                    ax.caption.clone()
                };
                let shows = !matches!(forms, "SpinButton" | "ScrollBar" | "");
                (
                    (!forms.is_empty()).then(|| {
                        serde_json::json!({
                            "type": forms,
                            "activex": true,
                            "link": link,
                            "items": fill,
                            "checked": ax.value.as_deref() == Some("1"),
                            "val": ax.position,
                            "min": ax.min,
                            "max": ax.max,
                            "inc": inc.unwrap_or(1),
                            "page": page.unwrap_or(1),
                            "list": items,
                            // Its properties (Microsoft Forms): its font, colors, ...
                            "font": (!matches!(forms, "SpinButton" | "ScrollBar")).then(|| serde_json::json!({
                                "name": ax.font.clone().unwrap_or_else(|| self.font.0.clone()),
                                "size": ax.size.unwrap_or(9.0),
                                "bold": ax.bold,
                                "italic": ax.italic,
                                "underline": ax.underline,
                                "strike": strike,
                            })),
                            "foreColor": ax.fore,
                            "backColor": ax.back,
                            "opaque": ax.opaque,
                            "enabled": bit(crate::activex::ENABLED, true),
                            "wordWrap": bit(crate::activex::WORD_WRAP, ax.class == "Label"),
                            "autoSize": bit(crate::activex::AUTO_SIZE, false),
                            "altText": alt,
                        })
                    }),
                    (shows && !forms.is_empty()).then(|| text.unwrap_or_default()),
                )
            }
            (None, form) => {
                let text = (form.is_some() && shows_text)
                    .then(|| c.text.as_ref().map(|t| t.text.clone()).unwrap_or_default());
                (form, text)
            }
        };
        Some(Obj {
            anchor,
            html,
            name: c.name.clone(),
            macro_name: c.macro_name().unwrap_or_default().to_string(),
            kind: "control",
            chart: None,
            activex: c.activex.as_ref().map(Ax::msforms_type),
            // Embedded objects run no macro.
            assignable: c.activex.is_none() && c.kind != "Pict",
            shape_id: Some(c.shape_id),
            text,
            element: "",
            rot: 0.0,
            flip: (false, false),
            form,
            adjust: None,
            paint: None,
            text_box: None,
            // An ActiveX control's プロパティ: as the sheet says them, else its VML shape.
            props: (c.kind != "Pict").then(|| {
                let vml = |k: &str| {
                    c.client
                        .get(k)
                        .map(|v| !v.trim().eq_ignore_ascii_case("false"))
                };
                let (moves, sizes) = settings.moves.unwrap_or((
                    !c.client.contains_key("MoveWithCells"),
                    !c.client.contains_key("SizeWithCells"),
                ));
                serde_json::json!({
                    "placement": match (moves, sizes) {
                        (true, true) => "twoCell",
                        (true, false) => "oneCell",
                        _ => "absolute",
                    },
                    "print": settings.print.or_else(|| vml("PrintObject")).unwrap_or(true),
                    "locked": settings.locked.or_else(|| vml("Locked")).unwrap_or(true),
                    // 文字列のロック: a form control with a text (VML's `LockText`, its
                    // properties' `lockText`).
                    "lockText": (c.activex.is_none() && shows_text).then(|| {
                        c.client
                            .get("LockText")
                            .map(|v| !v.trim().eq_ignore_ascii_case("false"))
                            .or_else(|| c.props.get("lockText").map(|v| matches!(v.as_str(), "1" | "true")))
                            .unwrap_or(true)
                    }),
                })
            }),
        })
    }
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
fn form_html(c: &Control, env: &Env, font: &(String, f64), items: &[String]) -> Option<String> {
    let text = c.text.clone().unwrap_or_default();
    let label = esc(&text.text);
    let linked = c.linked_cell().and_then(|r| (env.value)(r));
    let html = match c.kind.as_str() {
        "Button" => format!(
            "<div class=\"ctl button\" style=\"{}{}\"><span>{label}</span></div>",
            flex(text.align, text.valign, ("center", "center")),
            text_css(&text, font)
        ),
        "CheckBox" | "Radio" => {
            // The state of the linked cell, else the one saved.
            let saved = match c.checked() {
                CheckState::Unchecked => 0,
                CheckState::Checked => 1,
                CheckState::Mixed => 2,
            };
            let state = match linked.as_deref().map(str::to_ascii_uppercase).as_deref() {
                Some("TRUE") if c.kind == "CheckBox" => 1,
                Some("FALSE") if c.kind == "CheckBox" => 0,
                Some("#N/A") if c.kind == "CheckBox" => 2,
                _ => saved,
            };
            format!(
                "<div class=\"ctl check{}\" style=\"{}{}\">{}<span class=\"label\">{label}</span></div>",
                if c.is_flat() { " flat" } else { "" },
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
            let sel = linked
                .as_deref()
                .and_then(|v| v.parse::<f64>().ok())
                .map(|v| v as i64)
                .unwrap_or_else(|| c.selected());
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
            let horiz = c.is_horizontal();
            if c.kind == "Spin" {
                spin_html(horiz)
            } else {
                let (saved, min, max) = c.value_range();
                let pos = linked
                    .as_deref()
                    .and_then(|v| v.parse::<f64>().ok())
                    .map(|v| v as i64)
                    .unwrap_or(saved);
                scroll_html(horiz, min, max, pos)
            }
        }
        _ => return None,
    };
    Some(html)
}

/// The HTML of an ActiveX control.
fn activex_html(
    c: &Control,
    ax: &Ax,
    s: &AxSettings,
    default: &(String, f64),
    items: &[String],
) -> String {
    // The default font of MSForms controls in Excel.
    let fallback = (default.0.clone(), 9.0);
    let enabled = s.bits.is_none_or(|b| b & crate::activex::ENABLED != 0);
    let font = Text {
        font: ax.font.clone(),
        size: ax.size,
        // (Disabled: its text grayed.)
        color: if enabled {
            ax.fore.clone()
        } else {
            Some("#a0a0a0".into())
        },
        bold: ax.bold,
        italic: ax.italic,
        underline: ax.underline,
        strike: s.strike,
        ..Default::default()
    };
    let mut css = text_css(&font, &fallback);
    // WordWrap off: its text on one line (a label's is on unless said otherwise).
    let wraps = s
        .bits
        .map_or(ax.class == "Label", |b| b & crate::activex::WORD_WRAP != 0);
    if !wraps {
        css.push_str("white-space:pre;");
    }
    let caption = esc(ax.caption.as_deref().unwrap_or(&c.name));
    // (BackStyle transparent: no back.)
    let back = |d: &str| {
        if ax.opaque {
            format!("background:{};", ax.back.as_deref().unwrap_or(d))
        } else {
            "background:transparent;".to_string()
        }
    };
    let value = esc(ax.value.as_deref().unwrap_or(""));
    let on = ax.is_on();
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
        ("MorphData", 2) => {
            let mut list = String::new();
            let selected = ax.value.as_deref().unwrap_or("");
            for item in items.iter().take(200) {
                let _ = write!(
                    list,
                    "<span{}>{}</span>",
                    if !selected.is_empty() && item == selected {
                        " class=\"sel\""
                    } else {
                        ""
                    },
                    esc(item)
                );
            }
            format!(
                "<div class=\"ctl list\" style=\"{}{}\">{list}</div>",
                css,
                back("#ffffff")
            )
        }
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
