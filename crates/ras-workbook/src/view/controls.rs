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
}

impl Controls {
    /// The controls of the sheet whose part is `sheet_part`.
    pub(crate) fn read(parts: &Parts, sheet_part: &str, font: &(String, f64)) -> Controls {
        let list = read::controls(parts, sheet_part);
        Controls {
            done: vec![false; list.len()],
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
        let html = match &c.activex {
            // An ActiveX control the grid draws itself; another one: its preview.
            Some(ax) if !ax.class.is_empty() => activex_html(c, ax, &self.font),
            Some(ax) => picture().unwrap_or_else(|| activex_html(c, ax, &self.font)),
            // An embedded object: its picture.
            None if c.kind == "Pict" => picture()?,
            None => form_html(c, env, &self.font)?,
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
                "val": val,
                "min": min,
                "max": max,
                "inc": num("inc", "Inc", 1),
                "page": num("page", "Page", 10),
            })
        });
        let shows_text = matches!(
            c.kind.as_str(),
            "Button" | "CheckBox" | "Radio" | "Label" | "GBox"
        );
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
            text: (form.is_some() && shows_text)
                .then(|| c.text.as_ref().map(|t| t.text.clone()).unwrap_or_default()),
            element: "",
            rot: 0.0,
            flip: (false, false),
            form,
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
fn form_html(c: &Control, env: &Env, font: &(String, f64)) -> Option<String> {
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
            let items = c.list_range().map(|r| (env.texts)(r)).unwrap_or_default();
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
