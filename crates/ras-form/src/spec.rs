//! The FormSpec document model (`forms/<Name>.form.json`).
//!
//! Port of xlflow's `internal/vba/userforms/spec` data types (Copyright (c)
//! 2026 harumiWeb, MIT License; see `NOTICE.md`). Field names and meaning
//! follow xlflow's FormSpec version 1, so a spec written for xlflow in JSON
//! is accepted as is.

use serde::{Deserialize, Serialize};
use serde_json::{Map, Value};

/// `kind` written for new RAS forms.
pub const KIND: &str = "ras.userform";
/// `kind` of xlflow's FormSpec, accepted as well.
pub const XLFLOW_KIND: &str = "xlflow.userform";

/// A whole FormSpec document.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FormSpec {
    /// Always `1`.
    pub schema_version: i64,
    /// `ras.userform` or `xlflow.userform`.
    pub kind: String,
    /// Always `designer`.
    pub basis: String,
    /// `points` or `parent-relative` (both mean parent-relative points).
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub coordinate_system: String,
    /// The form itself.
    pub form: FormSpecForm,
    /// Flat list of controls; containment is expressed with `parentId`.
    #[serde(default)]
    pub controls: Vec<FormSpecControl>,
    /// Snapshot-only capture warnings.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub warnings: Option<Value>,
}

/// The `form` object.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FormSpecForm {
    /// VBA component name of the form.
    pub name: String,
    /// Caption; defaults to the name.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub caption: Option<String>,
    /// Persisted client width (`DisplayedSize`) in points; default 240.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    /// Persisted client height in points; default 180.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
    /// Partial font.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<FormSpecFont>,
    /// Text color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fore_color: Option<FormSpecColor>,
    /// Background color.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub back_color: Option<FormSpecColor>,
    /// `0` none, `1` single.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_style: Option<i64>,
    /// Snapshot-only observations (never authoring input).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub observed: Option<Value>,
}

/// One control.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FormSpecControl {
    /// Stable identifier used by `parentId`.
    #[serde(default)]
    pub id: String,
    /// Id of the containing Frame, MultiPage or Page.
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub parent_id: String,
    /// Sibling order (z-order); defaults to the position in `controls`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub z_index: Option<i64>,
    /// Built-in type (`Label`, `TextBox`, ...) or a custom type.
    #[serde(rename = "type", default)]
    pub control_type: String,
    /// VBA name of the control.
    #[serde(default)]
    pub name: String,
    /// ProgID; only needed for custom controls.
    #[serde(default, skip_serializing_if = "String::is_empty")]
    pub prog_id: String,
    #[allow(missing_docs)]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub caption: Option<String>,
    #[allow(missing_docs)]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tag: Option<String>,
    #[allow(missing_docs)]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub control_tip_text: Option<String>,
    #[allow(missing_docs)]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub accelerator: Option<String>,
    #[allow(missing_docs)]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    /// Any JSON value (string, Boolean or number depending on the type).
    #[serde(
        default,
        skip_serializing_if = "Option::is_none",
        deserialize_with = "non_null"
    )]
    pub value: Option<Value>,
    #[allow(missing_docs)]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<f64>,
    #[allow(missing_docs)]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<f64>,
    #[allow(missing_docs)]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[allow(missing_docs)]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub height: Option<f64>,
    #[allow(missing_docs)]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tab_index: Option<i64>,
    /// MultiPage / TabStrip selection (zero based, `-1` none).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub selected_index: Option<i64>,
    #[allow(missing_docs)]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub enabled: Option<bool>,
    #[allow(missing_docs)]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub visible: Option<bool>,
    /// ComboBox / ListBox items (observed only; not generated).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub list: Option<Vec<String>>,
    /// TabStrip tabs; `None` = unspecified, `Some([])` = no tabs.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tabs: Option<Vec<FormSpecTab>>,
    /// Snapshot-only list of unmodelled properties.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub unsupported: Option<Vec<String>>,
    /// Image picture action.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture: Option<Value>,
    /// Legacy nested children (normalised to the flat list).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub controls: Option<Vec<FormSpecControl>>,
    /// Unchecked property bag.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub properties: Option<Map<String, Value>>,
    /// Snapshot-only observations.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub observed: Option<Value>,
    #[allow(missing_docs)]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font: Option<FormSpecFont>,
    #[allow(missing_docs)]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fore_color: Option<FormSpecColor>,
    #[allow(missing_docs)]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub back_color: Option<FormSpecColor>,
    /// `1` left, `2` center, `3` right.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_align: Option<i64>,
    #[allow(missing_docs)]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border_style: Option<i64>,
    #[allow(missing_docs)]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture_alignment: Option<i64>,
    #[allow(missing_docs)]
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub picture_size_mode: Option<i64>,
    /// ComboBox `Style`: `0` DropDownCombo, `2` DropDownList.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style: Option<i64>,
    /// ListBox `MultiSelect`: `0` single, `1` multi, `2` extended.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub multi_select: Option<i64>,
    /// ComboBox / ListBox `ListStyle`: `0` plain, `1` option.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub list_style: Option<i64>,
    /// ComboBox / ListBox `MatchEntry`: `0` first letter, `1` complete, `2` none.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub match_entry: Option<i64>,
    /// ComboBox `ShowDropButtonWhen`: `0` never, `1` focus, `2` always.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub show_drop_button_when: Option<i64>,
    /// ComboBox `ListRows`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub list_rows: Option<i64>,
}

fn non_null<'de, D: serde::Deserializer<'de>>(d: D) -> Result<Option<Value>, D::Error> {
    Ok(Option::<Value>::deserialize(d)?.filter(|v| !v.is_null()))
}

/// One TabStrip tab.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
#[allow(missing_docs)]
pub struct FormSpecTab {
    pub name: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub caption: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub control_tip_text: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tag: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub accelerator: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub enabled: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub visible: Option<bool>,
}

/// A partial font; omitted members keep the default.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
#[allow(missing_docs)]
pub struct FormSpecFont {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub name: Option<String>,
    /// Points.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub size: Option<f64>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strike_through: Option<bool>,
    /// Windows charset number (0–255).
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub charset: Option<i64>,
}

/// A color: `{kind: "rgb", value: "#RRGGBB"}`, `{kind: "system", value:
/// "buttonFace"}` or `{kind: "ole", value: "0x8000000F"}`.
#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[allow(missing_docs)]
pub struct FormSpecColor {
    pub kind: String,
    pub value: String,
}

const SYSTEM_COLORS: &[(&str, u32)] = &[
    ("scrollBar", 0),
    ("desktop", 1),
    ("activeCaption", 2),
    ("inactiveCaption", 3),
    ("menu", 4),
    ("window", 5),
    ("windowFrame", 6),
    ("menuText", 7),
    ("windowText", 8),
    ("captionText", 9),
    ("activeBorder", 10),
    ("inactiveBorder", 11),
    ("appWorkspace", 12),
    ("highlight", 13),
    ("highlightText", 14),
    ("buttonFace", 15),
    ("buttonShadow", 16),
    ("grayText", 17),
    ("buttonText", 18),
    ("inactiveCaptionText", 19),
    ("buttonHighlight", 20),
    ("buttonDarkShadow", 21),
    ("buttonLight", 22),
    ("infoText", 23),
    ("infoBackground", 24),
    ("hotLight", 26),
    ("gradientActiveCaption", 27),
    ("gradientInactiveCaption", 28),
    ("menuHighlight", 29),
    ("menuBar", 30),
];

/// Names of the system colors accepted by `{kind: "system"}`.
pub fn system_color_names() -> impl Iterator<Item = &'static str> {
    SYSTEM_COLORS.iter().map(|(name, _)| *name)
}

impl FormSpecColor {
    /// The color spelled the way xlflow projects an OLE color.
    pub fn from_ole(value: u32) -> FormSpecColor {
        if value & 0xffff_ff00 == 0x8000_0000 {
            if let Some((name, _)) = SYSTEM_COLORS
                .iter()
                .find(|(_, index)| value & 255 == *index)
            {
                return FormSpecColor {
                    kind: "system".into(),
                    value: name.to_string(),
                };
            }
        }
        if value & 0xff00_0000 == 0 {
            return FormSpecColor {
                kind: "rgb".into(),
                value: format!(
                    "#{:02X}{:02X}{:02X}",
                    value & 255,
                    value >> 8 & 255,
                    value >> 16 & 255
                ),
            };
        }
        FormSpecColor {
            kind: "ole".into(),
            value: format!("0x{value:08X}"),
        }
    }

    /// The OLE color value (`OLE_COLOR`).
    pub fn ole(&self) -> Result<u32, String> {
        let bad = || {
            Err(format!(
                "色の指定 {} {:?} が不正です",
                self.kind, self.value
            ))
        };
        match self.kind.as_str() {
            "rgb" => {
                let v = &self.value;
                if v.len() == 7
                    && v.starts_with('#')
                    && v[1..].chars().all(|c| c.is_ascii_hexdigit())
                {
                    let n = u32::from_str_radix(&v[1..], 16).map_err(|_| "".to_string())?;
                    return Ok((n & 255) << 16 | (n & 0xff00) | (n >> 16));
                }
                bad()
            }
            "system" => match SYSTEM_COLORS.iter().find(|(name, _)| *name == self.value) {
                Some((_, index)) => Ok(0x8000_0000 | index),
                None => bad(),
            },
            "ole" => {
                let v = &self.value;
                if v.len() == 10
                    && v.starts_with("0x")
                    && v[2..].chars().all(|c| c.is_ascii_hexdigit())
                {
                    return u32::from_str_radix(&v[2..], 16).map_err(|_| "".to_string());
                }
                bad()
            }
            _ => bad(),
        }
    }
}

impl FormSpec {
    /// The minimal spec written by `ras form new`.
    pub fn new_form(name: &str) -> FormSpec {
        FormSpec {
            schema_version: 1,
            kind: KIND.to_string(),
            basis: "designer".to_string(),
            form: FormSpecForm {
                name: name.to_string(),
                caption: Some(name.to_string()),
                width: Some(240.0),
                height: Some(180.0),
                ..FormSpecForm::default()
            },
            ..FormSpec::default()
        }
    }

    /// Pretty JSON text (two-space indentation, trailing newline).
    pub fn to_json(&self) -> String {
        let mut text = serde_json::to_string_pretty(self).expect("FormSpec serialises");
        text.push('\n');
        text
    }
}

/// Flattens legacy nested `controls`, fills missing ids and `zIndex`
/// (`NormalizeFormSpec`).
pub(crate) fn normalize(spec: &mut FormSpec) {
    struct State {
        next_id: usize,
        used: std::collections::HashSet<String>,
    }
    fn unique_id(state: &mut State, name: &str) -> String {
        let base = name.trim();
        let base = if base.is_empty() {
            "control".to_string()
        } else {
            base.replace(' ', "_").to_lowercase()
        };
        if state.used.insert(base.clone()) {
            return base;
        }
        loop {
            let id = format!("{base}_{:03}", state.next_id);
            state.next_id += 1;
            if state.used.insert(id.clone()) {
                return id;
            }
        }
    }
    fn visit(
        state: &mut State,
        mut control: FormSpecControl,
        inherited: &str,
        index: usize,
        out: &mut Vec<FormSpecControl>,
    ) {
        let id = control.id.trim().to_string();
        control.id = if id.is_empty() {
            unique_id(state, &control.name)
        } else {
            state.used.insert(id.clone());
            id
        };
        if control.parent_id.trim().is_empty() {
            control.parent_id = inherited.to_string();
        }
        if control.z_index.is_none() {
            control.z_index = Some(index as i64);
        }
        let children = control.controls.take().unwrap_or_default();
        let parent = control.id.clone();
        out.push(control);
        for (i, child) in children.into_iter().enumerate() {
            visit(state, child, &parent, i, out);
        }
    }
    let mut state = State {
        next_id: 1,
        used: Default::default(),
    };
    let mut out = Vec::new();
    for (index, control) in std::mem::take(&mut spec.controls).into_iter().enumerate() {
        visit(&mut state, control, "", index, &mut out);
    }
    spec.controls = out;
}
