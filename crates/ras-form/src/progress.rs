//! The progress bar (`ProgressBar`), a control RAS composes from MSForms labels
//! (RAS language specification §14.11.13).
//!
//! MSForms has no progress bar (the one of Windows Common Controls is an ActiveX
//! control that is not on every PC), so a `ProgressBar` of a FormSpec is written
//! as four labels that sit on top of each other:
//!
//! ```text
//! <name>_Caption   Upload progress                     56%   <name>_Value
//! <name>_Track     ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
//! <name>_Bar       ━━━━━━━━━━━━━━━━━━━━━━━━
//! ```
//!
//! The track label keeps what the labels cannot show in its `Tag`
//! (`RAS.ProgressBar|min=0|max=100|value=56|showValue=1|accent=1`). The bundled
//! RAS class `ProgressBar` (`ras-compiler/stdlib/progressbar.ras`) reads it when
//! the form loads and moves the labels when the code sets `value`, `caption` and
//! the other properties. Without `barColor`, the bar takes the accent color of
//! Windows when the form loads (`highlight` until then).
//!
//! The layout here is the one of the class's `layout()` and of the debug run's
//! `progressParts` (`ras-compiler/src/js/forms.js`); the three must agree.

use serde_json::{json, Map, Value};

use crate::spec::{FormSpec, FormSpecColor, FormSpecControl};

/// The FormSpec type of a progress bar.
pub const TYPE: &str = "ProgressBar";
/// The first item of the track label's `Tag`.
pub const TAG_PREFIX: &str = "RAS.ProgressBar";
/// The labels a progress bar is made of, by the suffix of their names, back to front.
pub const PARTS: [&str; 4] = ["_Track", "_Bar", "_Caption", "_Value"];

/// Width when omitted, in points.
pub const DEFAULT_WIDTH: f64 = 200.0;
/// Height when omitted, in points.
pub const DEFAULT_HEIGHT: f64 = 24.0;
/// Height of the bar when omitted, in points (6 pixels).
pub const DEFAULT_BAR_HEIGHT: f64 = 4.5;
/// Space between the text and the bar, in points.
pub const GAP: f64 = 3.0;
/// Width of the label of the percentage, in points.
pub const VALUE_WIDTH: f64 = 36.0;
/// Color of the track when omitted.
pub const DEFAULT_TRACK_COLOR: &str = "#E5E5E5";

/// Whether `control_type` is the progress bar.
pub fn is_progress_bar(control_type: &str) -> bool {
    control_type.trim().eq_ignore_ascii_case(TYPE)
}

/// The settings of a progress bar, with the defaults filled in.
#[derive(Clone, Debug, PartialEq)]
pub struct Settings {
    /// Least value.
    pub min: f64,
    /// Greatest value.
    pub max: f64,
    /// Current value (not clamped).
    pub value: f64,
    /// The percentage is shown at the right end.
    pub show_value: bool,
    /// Height of the bar, in points.
    pub bar_height: f64,
}

fn number(control: &FormSpecControl, name: &str) -> Option<f64> {
    control.scalar(name).and_then(Value::as_f64)
}

impl Settings {
    /// The settings of `control`.
    pub fn of(control: &FormSpecControl) -> Settings {
        Settings {
            min: number(control, "min").unwrap_or(0.0),
            max: number(control, "max").unwrap_or(100.0),
            value: control
                .value
                .as_ref()
                .and_then(Value::as_f64)
                .unwrap_or(0.0),
            show_value: control
                .scalar("showValue")
                .and_then(Value::as_bool)
                .unwrap_or(true),
            bar_height: number(control, "barHeight").unwrap_or(DEFAULT_BAR_HEIGHT),
        }
    }

    /// The share of the range the value covers (0 to 1).
    pub fn fraction(&self) -> f64 {
        if self.max <= self.min {
            return 0.0;
        }
        ((self.value - self.min) / (self.max - self.min)).clamp(0.0, 1.0)
    }

    /// The text of the percentage (`56%`), rounded half up.
    pub fn percent_text(&self) -> String {
        format!("{}%", (self.fraction() * 100.0 + 0.5).floor() as i64)
    }
}

/// A length in points rounded to whole pixels at 96 DPI (0.75 pt).
fn pixels(points: f64) -> f64 {
    (points / 0.75).round() * 0.75
}

/// Where the parts go: `(left, top, width, height)` of the track, the bar, the caption
/// and the percentage, in that order (as [`PARTS`]).
pub fn layout(
    left: f64,
    top: f64,
    width: f64,
    height: f64,
    settings: &Settings,
) -> [(f64, f64, f64, f64); 4] {
    let bar_height = settings.bar_height.clamp(0.75, height.max(0.75));
    let bar_top = top + height - bar_height;
    let text_height = (height - bar_height - GAP).max(0.0);
    let value_width = if settings.show_value {
        VALUE_WIDTH.min(width)
    } else {
        0.0
    };
    [
        (left, bar_top, width, bar_height),
        (
            left,
            bar_top,
            pixels(width * settings.fraction()),
            bar_height,
        ),
        (left, top, (width - value_width).max(0.0), text_height),
        (left + width - value_width, top, value_width, text_height),
    ]
}

/// The `Tag` of the track label: what the class needs that the labels do not show.
pub fn tag(settings: &Settings, accent: bool) -> String {
    format!(
        "{TAG_PREFIX}|min={}|max={}|value={}|showValue={}|accent={}",
        settings.min,
        settings.max,
        settings.value,
        u8::from(settings.show_value),
        u8::from(accent)
    )
}

fn color(value: &Value) -> Option<FormSpecColor> {
    serde_json::from_value(value.clone()).ok()
}

fn rgb(hex: &str) -> FormSpecColor {
    FormSpecColor {
        kind: "rgb".into(),
        value: hex.into(),
    }
}

fn system(name: &str) -> FormSpecColor {
    FormSpecColor {
        kind: "system".into(),
        value: name.into(),
    }
}

/// The labels `control` (a progress bar) is written as.
pub fn parts(control: &FormSpecControl) -> Vec<FormSpecControl> {
    let settings = Settings::of(control);
    let left = control.left.unwrap_or(0.0);
    let top = control.top.unwrap_or(0.0);
    let width = control.width.unwrap_or(DEFAULT_WIDTH);
    let height = control.height.unwrap_or(DEFAULT_HEIGHT);
    let bar_color = control.scalar("barColor").and_then(color);
    let accent = bar_color.is_none();
    let bar_color = bar_color.unwrap_or_else(|| system("highlight"));
    let track_color = control
        .scalar("trackColor")
        .and_then(color)
        .unwrap_or_else(|| rgb(DEFAULT_TRACK_COLOR));
    let value_color = control
        .scalar("valueColor")
        .and_then(color)
        .unwrap_or_else(|| system("grayText"));
    let boxes = layout(left, top, width, height, &settings);
    let visible = control.visible.unwrap_or(true);
    let fraction = settings.fraction();
    let label = |suffix: &str, index: usize| {
        let (l, t, w, h) = boxes[index];
        let mut scalars = Map::new();
        scalars.insert("borderStyle".into(), json!(0));
        FormSpecControl {
            id: format!("{}{suffix}", control.id),
            parent_id: control.parent_id.clone(),
            z_index: control.z_index,
            control_type: "Label".into(),
            name: format!("{}{suffix}", control.name),
            caption: Some(String::new()),
            control_tip_text: control.control_tip_text.clone(),
            left: Some(l),
            top: Some(t),
            width: Some(w),
            height: Some(h),
            enabled: control.enabled,
            visible: Some(visible),
            scalars,
            ..FormSpecControl::default()
        }
    };
    let mut track = label(PARTS[0], 0);
    track.back_color = Some(track_color);
    track.tag = Some(tag(&settings, accent));
    let mut bar = label(PARTS[1], 1);
    bar.back_color = Some(bar_color);
    // A label is at least a pixel wide: an empty bar is hidden instead.
    if fraction <= 0.0 {
        bar.visible = Some(false);
        bar.width = Some(0.75);
    }
    let text = |mut part: FormSpecControl, caption: String| {
        part.caption = Some(caption);
        part.font = control.font.clone();
        part.scalars.insert("backStyle".into(), json!(0));
        part.scalars.insert("wordWrap".into(), json!(false));
        if part.height == Some(0.0) || part.width == Some(0.0) {
            part.visible = Some(false);
            part.height = Some(part.height.unwrap_or(0.0).max(0.75));
            part.width = Some(part.width.unwrap_or(0.0).max(0.75));
        }
        part
    };
    let mut caption = text(
        label(PARTS[2], 2),
        control.caption.clone().unwrap_or_default(),
    );
    caption.fore_color = control.fore_color.clone();
    caption.tag = control.tag.clone();
    let mut value = text(label(PARTS[3], 3), settings.percent_text());
    value.fore_color = Some(value_color);
    value.text_align = Some(3);
    if !settings.show_value {
        value.visible = Some(false);
    }
    vec![track, bar, caption, value]
}

/// Replaces each progress bar of `spec` by its labels, in its place.
pub fn expand(spec: &FormSpec) -> FormSpec {
    if !spec
        .controls
        .iter()
        .any(|c| is_progress_bar(&c.control_type))
    {
        return spec.clone();
    }
    let mut out = spec.clone();
    out.controls = spec
        .controls
        .iter()
        .flat_map(|c| {
            if is_progress_bar(&c.control_type) {
                parts(c)
            } else {
                vec![c.clone()]
            }
        })
        .collect();
    out
}

/// The settings recorded in the `Tag` of a track label, if it is one.
fn read_tag(tag: &str) -> Option<(Settings, bool)> {
    let mut items = tag.split('|');
    if items.next()? != TAG_PREFIX {
        return None;
    }
    let mut settings = Settings {
        min: 0.0,
        max: 100.0,
        value: 0.0,
        show_value: true,
        bar_height: DEFAULT_BAR_HEIGHT,
    };
    let mut accent = true;
    for item in items {
        let Some((key, value)) = item.split_once('=') else {
            continue;
        };
        let Ok(n) = value.trim().parse::<f64>() else {
            continue;
        };
        match key {
            "min" => settings.min = n,
            "max" => settings.max = n,
            "value" => settings.value = n,
            "showValue" => settings.show_value = n != 0.0,
            "accent" => accent = n != 0.0,
            _ => {}
        }
    }
    Some((settings, accent))
}

/// Turns the labels of each progress bar of a form read from a workbook (`ras form pull`)
/// back into the `ProgressBar`, in the place of its track label.
pub fn collapse(spec: &mut FormSpec) {
    loop {
        let found = spec.controls.iter().enumerate().find_map(|(i, c)| {
            let base = c.name.strip_suffix(PARTS[0])?;
            let tag = read_tag(c.tag.as_deref()?)?;
            let part = |suffix: &str| {
                spec.controls.iter().position(|o| {
                    o.name.eq_ignore_ascii_case(&format!("{base}{suffix}"))
                        && o.control_type.eq_ignore_ascii_case("Label")
                        && o.parent_id == c.parent_id
                })
            };
            Some((
                i,
                base.to_string(),
                tag,
                part(PARTS[1])?,
                part(PARTS[2])?,
                part(PARTS[3])?,
            ))
        });
        let Some((track_at, base, (settings, accent), bar_at, caption_at, value_at)) = found else {
            return;
        };
        let track = &spec.controls[track_at];
        let bar = &spec.controls[bar_at];
        let caption = &spec.controls[caption_at];
        let value = &spec.controls[value_at];
        let top = pixels(caption.top.unwrap_or(0.0));
        let bar_top = track.top.unwrap_or(0.0);
        let bar_height = pixels(track.height.unwrap_or(DEFAULT_BAR_HEIGHT));
        let mut scalars = Map::new();
        let mut put = |name: &str, v: Value| {
            scalars.insert(name.to_string(), v);
        };
        if settings.min != 0.0 {
            put("min", json!(settings.min));
        }
        if settings.max != 100.0 {
            put("max", json!(settings.max));
        }
        if !settings.show_value {
            put("showValue", json!(false));
        }
        if bar_height != DEFAULT_BAR_HEIGHT {
            put("barHeight", json!(bar_height));
        }
        let differs = |c: &Option<FormSpecColor>, default: &FormSpecColor| {
            c.as_ref().filter(|c| *c != default).map(|c| json!(c))
        };
        if !accent {
            if let Some(c) = &bar.back_color {
                put("barColor", json!(c));
            }
        }
        if let Some(c) = differs(&track.back_color, &rgb(DEFAULT_TRACK_COLOR)) {
            put("trackColor", c);
        }
        if let Some(c) = differs(&value.fore_color, &system("grayText")) {
            put("valueColor", c);
        }
        let number = |n: f64| {
            if n.fract() == 0.0 && n.abs() < 1e15 {
                json!(n as i64)
            } else {
                json!(n)
            }
        };
        let bar_control = FormSpecControl {
            id: base.to_lowercase(),
            parent_id: track.parent_id.clone(),
            z_index: track.z_index,
            control_type: TYPE.into(),
            name: base,
            caption: caption.caption.clone(),
            tag: caption.tag.clone(),
            control_tip_text: caption.control_tip_text.clone(),
            value: (settings.value != 0.0).then(|| number(settings.value)),
            left: track.left.map(pixels),
            top: Some(top),
            width: track.width.map(pixels),
            height: Some(pixels(bar_top + track.height.unwrap_or(0.0) - top)),
            enabled: caption.enabled,
            visible: track.visible,
            font: caption.font.clone(),
            fore_color: caption.fore_color.clone(),
            scalars,
            ..FormSpecControl::default()
        };
        let mut parts = [track_at, bar_at, caption_at, value_at];
        parts.sort_unstable();
        // Where the track was, once the parts in front of it are gone.
        let at = track_at - parts.iter().filter(|&&i| i < track_at).count();
        for i in parts.into_iter().rev() {
            spec.controls.remove(i);
        }
        spec.controls.insert(at, bar_control);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn spec(control: &str) -> FormSpec {
        let text = format!(
            r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"F"}},"controls":[{control}]}}"#
        );
        let parsed = crate::parse(&text);
        assert!(!parsed.has_errors(), "{:?}", parsed.issues);
        parsed.spec.unwrap()
    }

    #[test]
    fn a_progress_bar_is_four_labels() {
        let spec = spec(
            r#"{"id":"pb","name":"pbUpload","type":"ProgressBar","caption":"Upload progress","value":56,"left":12,"top":12,"width":240,"height":24}"#,
        );
        let expanded = expand(&spec);
        let names: Vec<&str> = expanded.controls.iter().map(|c| c.name.as_str()).collect();
        assert_eq!(
            names,
            [
                "pbUpload_Track",
                "pbUpload_Bar",
                "pbUpload_Caption",
                "pbUpload_Value"
            ]
        );
        let [track, bar, caption, value] = &expanded.controls[..] else {
            unreachable!()
        };
        assert_eq!(
            track.tag.as_deref(),
            Some("RAS.ProgressBar|min=0|max=100|value=56|showValue=1|accent=1")
        );
        assert_eq!((track.top, track.height), (Some(31.5), Some(4.5)));
        // 56% of 240 pt is 134.4 pt: 179.2 px, written as 179 px.
        assert_eq!(bar.width, Some(134.25));
        assert_eq!(bar.back_color.as_ref().unwrap().value, "highlight");
        assert_eq!(caption.caption.as_deref(), Some("Upload progress"));
        assert_eq!(caption.width, Some(204.0));
        assert_eq!(value.caption.as_deref(), Some("56%"));
        assert_eq!(value.text_align, Some(3));
        // The labels are generated as any other.
        crate::compile(&spec, 932).unwrap();
    }

    #[test]
    fn colors_and_range() {
        let spec = spec(
            r##"{"id":"pb","name":"pb","type":"ProgressBar","min":10,"max":20,"value":10,"showValue":false,"barColor":{"kind":"rgb","value":"#000000"}}"##,
        );
        let expanded = expand(&spec);
        let track = &expanded.controls[0];
        assert_eq!(
            track.tag.as_deref(),
            Some("RAS.ProgressBar|min=10|max=20|value=10|showValue=0|accent=0")
        );
        assert_eq!(expanded.controls[1].visible, Some(false));
        assert_eq!(
            expanded.controls[1].back_color.as_ref().unwrap().value,
            "#000000"
        );
        assert_eq!(expanded.controls[3].visible, Some(false));
        assert_eq!(expanded.controls[2].width, Some(DEFAULT_WIDTH));
        crate::compile(&spec, 932).unwrap();
    }

    #[test]
    fn labels_read_from_a_workbook_are_a_progress_bar_again() {
        let spec = spec(
            r##"{"id":"no","name":"btnNo","type":"CommandButton","top":90},{"id":"pb","name":"pbUpload","type":"ProgressBar","caption":"Upload progress","value":56,"left":12,"top":12,"width":240,"height":30,"showValue":false,"barHeight":6,"barColor":{"kind":"rgb","value":"#000000"}},{"id":"ok","name":"btnOK","type":"CommandButton","top":60}"##,
        );
        let designer = crate::compile(&spec, 932).unwrap();
        let mut read = crate::snapshot("F", &designer.streams, &designer.storages, 932)
            .unwrap()
            .spec;
        assert_eq!(read.controls.len(), 6);
        collapse(&mut read);
        assert_eq!(read.controls.len(), 3);
        assert_eq!(read.controls[0].name, "btnNo");
        let bar = &read.controls[1];
        assert_eq!(bar.control_type, TYPE);
        assert_eq!(bar.name, "pbUpload");
        assert_eq!(bar.caption.as_deref(), Some("Upload progress"));
        assert_eq!(bar.value, Some(json!(56)));
        assert_eq!(
            (bar.left, bar.top, bar.width, bar.height),
            (Some(12.0), Some(12.0), Some(240.0), Some(30.0))
        );
        assert_eq!(bar.scalar("showValue"), Some(&json!(false)));
        assert_eq!(bar.scalar("barHeight"), Some(&json!(6.0)));
        assert_eq!(
            bar.scalar("barColor"),
            Some(&json!({"kind": "rgb", "value": "#000000"}))
        );
        assert_eq!(bar.scalar("trackColor"), None);
        assert_eq!(read.controls[2].name, "btnOK");
    }

    #[test]
    fn checks() {
        for (control, field) in [
            (
                r#"{"id":"a","name":"pb","type":"ProgressBar","min":5,"max":5}"#,
                "controls[0].max",
            ),
            (
                r#"{"id":"a","name":"pb","type":"ProgressBar","value":101}"#,
                "controls[0].value",
            ),
            (
                r#"{"id":"a","name":"pb","type":"ProgressBar","barColor":{"kind":"rgb","value":"red"}}"#,
                "controls[0].barColor",
            ),
            (
                r#"{"id":"a","name":"pb","type":"ProgressBar"},{"id":"b","name":"pb_Bar","type":"Label"}"#,
                "controls[1].name",
            ),
        ] {
            let text = format!(
                r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"F"}},"controls":[{control}]}}"#
            );
            let parsed = crate::parse(&text);
            assert!(
                parsed
                    .issues
                    .iter()
                    .any(|i| i.code == "UFV024" && i.field == field),
                "{control}: {:?}",
                parsed.issues
            );
        }
    }
}
