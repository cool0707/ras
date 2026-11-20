//! FormSpec → designer storage (`compiler.CompileNew`).
//!
//! Port of xlflow's `internal/vba/userforms/compiler/generate.go` and the
//! generation part of `typed.go` / `compiler.go` (Copyright (c) 2026
//! harumiWeb, MIT License; see `NOTICE.md`).

use std::collections::{BTreeMap, HashMap};
use std::path::Path;

use serde_json::Value;

use crate::contract;
use crate::oforms::{
    self, ControlDefinition, Definition, Designer, Edit, FontEdit, PValue, PersistError, Tab,
    TabStripState,
};
use crate::spec::{FormSpec, FormSpecControl, FormSpecFont, FormSpecTab};

/// Why a valid FormSpec could not be turned into a designer storage.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CompileError {
    /// `userform_generation_invalid`, `userform_generation_unsupported` or
    /// `userform_generation_conflict`.
    pub code: &'static str,
    /// Name of the control concerned (empty for the form).
    pub control: String,
    /// JSON path of the field concerned (may be empty).
    pub property: String,
    /// Message in Japanese.
    pub reason: String,
}

impl std::fmt::Display for CompileError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match (self.control.is_empty(), self.property.is_empty()) {
            (true, true) => write!(f, "{}", self.reason),
            (true, false) => write!(f, "{}: {}", self.property, self.reason),
            (false, true) => write!(f, "{}: {}", self.control, self.reason),
            (false, false) => write!(f, "{} ({}): {}", self.property, self.control, self.reason),
        }
    }
}

impl std::error::Error for CompileError {}

pub(crate) const INVALID: &str = "userform_generation_invalid";
pub(crate) const UNSUPPORTED: &str = "userform_generation_unsupported";
pub(crate) const CONFLICT: &str = "userform_generation_conflict";

fn fail(
    code: &'static str,
    control: &str,
    property: &str,
    reason: impl Into<String>,
) -> CompileError {
    CompileError {
        code,
        control: control.to_string(),
        property: property.to_string(),
        reason: reason.into(),
    }
}

fn from_persist(error: PersistError) -> CompileError {
    match error {
        PersistError::Invalid(m) => fail(INVALID, "", "", m),
        PersistError::Unsupported(m) => fail(UNSUPPORTED, "", "", m),
    }
}

/// Class index and default size (points) of each generatable type.
fn defaults(control_type: &str) -> Option<(u16, f64, f64)> {
    Some(match control_type.to_lowercase().as_str() {
        "label" => (21, 72.0, 18.0),
        "textbox" => (23, 120.0, 18.0),
        "commandbutton" => (17, 72.0, 24.0),
        "checkbox" => (26, 72.0, 18.0),
        "optionbutton" => (27, 72.0, 18.0),
        "togglebutton" => (28, 72.0, 18.0),
        "combobox" => (25, 120.0, 18.0),
        "listbox" => (24, 120.0, 72.0),
        "spinbutton" => (16, 18.0, 36.0),
        "scrollbar" => (47, 120.0, 18.0),
        "image" => (12, 72.0, 72.0),
        "frame" => (14, 144.0, 108.0),
        "multipage" => (57, 240.0, 180.0),
        "page" => (7, 0.0, 0.0),
        "tabstrip" => (18, 240.0, 48.0),
        _ => return None,
    })
}

/// Points → HIMETRIC.
pub(crate) fn himetric(points: f64, property: &str) -> Result<i32, String> {
    if !points.is_finite() || ((property == "width" || property == "height") && points < 0.0) {
        return Err("座標・サイズが不正です".into());
    }
    let h = (points * 2540.0 / 72.0).round();
    if h < i32::MIN as f64 || h > i32::MAX as f64 {
        return Err("座標・サイズが HIMETRIC の範囲を超えています".into());
    }
    Ok(h as i32)
}

fn valid_generation_name(name: &str) -> bool {
    let mut chars = name.chars();
    match chars.next() {
        Some(c) if c.is_alphabetic() => {
            chars.all(|c| c.is_alphabetic() || c.is_numeric() || c == '_')
        }
        _ => false,
    }
}

/// `ValidateWritableComponentIdentity` for the form name.
fn validate_component_name(name: &str, codepage: u16) -> Result<(), String> {
    if name.is_empty() {
        return Err("フォーム名が空です".into());
    }
    if name.contains(['\0', '\r', '\n']) {
        return Err(format!("フォーム名 {name:?} に NUL や改行は使えません"));
    }
    ras_ovba::encode_text(codepage, name).map_err(|e| e.to_string())?;
    if name.encode_utf16().count() > 31 {
        return Err(format!("フォーム名 {name:?} が長すぎます（31 文字まで）"));
    }
    if name.eq_ignore_ascii_case("dir") || name.eq_ignore_ascii_case("_VBA_PROJECT") {
        return Err(format!("フォーム名 {name:?} は予約されています"));
    }
    Ok(())
}

/// A converted property value.
#[derive(Clone, Debug, PartialEq)]
enum Converted {
    Str(String),
    Int(i64),
    Geometry(i32),
    TabIndex(i16),
    Bool(bool),
}

fn number(value: &Value) -> Option<f64> {
    value.as_f64()
}

/// `convertValue`.
fn convert_value(kind: &str, property: &str, value: &Value) -> Result<Converted, String> {
    match property {
        "left" | "top" | "width" | "height" => {
            let v = number(value).ok_or("座標・サイズが不正です")?;
            Ok(Converted::Geometry(himetric(v, property)?))
        }
        "tabIndex" => match value.as_i64() {
            Some(v) if (0..=i16::MAX as i64).contains(&v) => Ok(Converted::TabIndex(v as i16)),
            _ => Err("tabIndex は 0〜32767 にしてください".into()),
        },
        "value" | "text" => {
            if kind.eq_ignore_ascii_case("CheckBox") || kind.eq_ignore_ascii_case("OptionButton") {
                let state = match value {
                    Value::Bool(b) => Some(*b),
                    Value::Number(n) => match n.as_f64() {
                        Some(v) if v == 1.0 || v == -1.0 => Some(true),
                        Some(0.0) => Some(false),
                        _ => None,
                    },
                    Value::String(s) => match s.trim().to_lowercase().as_str() {
                        "true" | "1" | "-1" => Some(true),
                        "false" | "0" => Some(false),
                        _ => None,
                    },
                    _ => None,
                };
                return match state {
                    Some(true) => Ok(Converted::Str("1".into())),
                    Some(false) => Ok(Converted::Str("0".into())),
                    None => Err(
                        "オンオフの状態は真偽値か 0 / 1 / -1 にしてください（3 状態は未対応）"
                            .into(),
                    ),
                };
            }
            match value {
                Value::String(s) => Ok(Converted::Str(s.clone())),
                _ => Err("text / value はこのコントロールでは文字列にしてください".into()),
            }
        }
        _ => Err(format!("{property} は変換できません")),
    }
}

/// `bagValue`.
fn bag_value(kind: &str, key: &str, value: &Value) -> Result<Converted, String> {
    match key {
        "tag" | "controltiptext" | "groupname" | "caption" | "accelerator" => match value {
            Value::String(s) => Ok(Converted::Str(s.clone())),
            _ => Err("文字列にしてください".into()),
        },
        "backcolor" | "forecolor" | "bordercolor" | "borderstyle" | "maxlength" => {
            let v = number(value)
                .filter(|v| v.is_finite() && v.trunc() == *v && *v >= 0.0 && *v <= u32::MAX as f64);
            let Some(v) = v else {
                return Err("32 ビットの符号なし整数にしてください".into());
            };
            if key == "borderstyle" && v > 1.0 {
                return Err("borderStyle は 0 か 1 にしてください".into());
            }
            if key == "maxlength" && v > i32::MAX as f64 {
                return Err("maxLength が Excel の範囲を超えています".into());
            }
            Ok(Converted::Int(v as i64))
        }
        "picturealignment" | "picturesizemode" => {
            let v = number(value)
                .filter(|v| v.is_finite() && v.trunc() == *v && (0.0..=4.0).contains(v));
            let Some(v) = v else {
                return Err(format!("{key} は対応範囲の整数にしてください"));
            };
            if key == "picturesizemode" && v != 0.0 && v != 1.0 && v != 3.0 {
                return Err("pictureSizeMode は 0、1、3 のいずれかにしてください".into());
            }
            Ok(Converted::Int(v as i64))
        }
        "enabled" | "visible" => value
            .as_bool()
            .map(Converted::Bool)
            .ok_or_else(|| "真偽値にしてください".into()),
        "tabindex" => {
            let v =
                number(value).filter(|v| v.trunc() == *v && (0.0..=i16::MAX as f64).contains(v));
            v.map(|v| Converted::TabIndex(v as i16))
                .ok_or_else(|| "tabIndex は 0〜32767 にしてください".into())
        }
        _ => convert_value(kind, key, value),
    }
}

/// `generationValue`.
fn generation_value(kind: &str, key: &str, value: &Value) -> Result<Converted, String> {
    if value.is_null() {
        return Err("null による既定値への戻しには対応していません".into());
    }
    if key == "tabindex" {
        return if value.is_i64() || value.is_u64() {
            convert_value(kind, "tabIndex", value)
        } else {
            bag_value(kind, key, value)
        };
    }
    if key == "value"
        && (kind.eq_ignore_ascii_case("SpinButton") || kind.eq_ignore_ascii_case("ScrollBar"))
    {
        let v =
            number(value).filter(|v| v.is_finite() && v.trunc() == *v && (0.0..=100.0).contains(v));
        return v
            .map(|v| Converted::Int(v as i64))
            .ok_or_else(|| "value は既定の範囲 0〜100 の整数にしてください".into());
    }
    if (key == "value" || key == "text") && kind.eq_ignore_ascii_case("ToggleButton") {
        return convert_value("CheckBox", key, value);
    }
    bag_value(kind, key, value)
}

fn property_name(key: &str) -> Option<&'static str> {
    Some(match key {
        "accelerator" => "Accelerator",
        "tag" => "Tag",
        "controltiptext" => "ControlTipText",
        "groupname" => "GroupName",
        "backcolor" => "BackColor",
        "forecolor" => "ForeColor",
        "bordercolor" => "BorderColor",
        "borderstyle" => "BorderStyle",
        "maxlength" => "MaxLength",
        "caption" => "Caption",
        "text" | "value" => "Value",
        "left" => "Left",
        "top" => "Top",
        "width" => "Width",
        "height" => "Height",
        "tabindex" => "TabIndex",
        "enabled" => "Enabled",
        "visible" => "Visible",
        "picturealignment" => "PictureAlignment",
        "picturesizemode" => "PictureSizeMode",
        _ => return None,
    })
}

fn applicable_property(kind: &str, key: &str) -> bool {
    match key {
        "picturealignment" | "picturesizemode" => kind.eq_ignore_ascii_case("Image"),
        "groupname" => kind.eq_ignore_ascii_case("OptionButton"),
        "maxlength" => {
            kind.eq_ignore_ascii_case("TextBox") || kind.eq_ignore_ascii_case("ComboBox")
        }
        "caption" | "text" | "value" => {
            if kind.eq_ignore_ascii_case("ListBox") && (key == "text" || key == "value") {
                return false;
            }
            contract::has_control_property(kind, key)
        }
        _ => true,
    }
}

/// Typed appearance fields present on a control (`hasTypedProperty`).
fn has_typed_property(control: &FormSpecControl, key: &str) -> bool {
    match key {
        "font" => control.font.is_some(),
        "forecolor" => control.fore_color.is_some(),
        "backcolor" => control.back_color.is_some(),
        "textalign" => control.text_align.is_some(),
        "borderstyle" => control.border_style.is_some(),
        "picturealignment" => control.picture_alignment.is_some(),
        "picturesizemode" => control.picture_size_mode.is_some(),
        _ => false,
    }
}

fn control_definition(
    control: &FormSpecControl,
    tab_index: usize,
    root: Option<&Path>,
) -> Result<ControlDefinition, (&'static str, String, String)> {
    let failure =
        |code: &'static str, property: &str, reason: String| (code, property.to_string(), reason);
    let Some((class, default_width, default_height)) = defaults(&control.control_type) else {
        return Err(failure(
            UNSUPPORTED,
            "type",
            "このコントロールの種類は生成できません".into(),
        ));
    };
    let built_in =
        contract::control_type(&control.control_type).expect("generatable types are built in");
    let kind = built_in.name;
    if !control.prog_id.is_empty() && !control.prog_id.eq_ignore_ascii_case(built_in.prog_id) {
        return Err(failure(
            UNSUPPORTED,
            "progId",
            "カスタム ProgID のコントロールは生成できません".into(),
        ));
    }
    if !valid_generation_name(&control.name) {
        return Err(failure(
            INVALID,
            "name",
            "VBA の識別子にしてください".into(),
        ));
    }
    if control.picture.is_none()
        && control
            .unsupported
            .as_ref()
            .is_some_and(|u| u.iter().any(|v| v.trim().eq_ignore_ascii_case("picture")))
    {
        return Err(failure(
            UNSUPPORTED,
            "picture",
            "既存の画像を保持したままの生成には対応していません".into(),
        ));
    }
    // The items of a ComboBox (25) or ListBox (24) are not part of the
    // designer storage; they are filled by generated code (`codebehind`).
    let list_control = class == 24 || class == 25;
    if !list_control
        && (control.list.as_ref().is_some_and(|l| !l.is_empty())
            || (control.selected_index.is_some() && class != 57 && class != 18))
    {
        return Err(failure(
            UNSUPPORTED,
            "",
            "list / selectedIndex は ComboBox と ListBox でだけ使えます".into(),
        ));
    }
    if tab_index > i16::MAX as usize {
        return Err(failure(INVALID, "", "コントロールが多すぎます".into()));
    }
    // Merge the property bag with the typed fields (`propertyMap`).
    let mut props: BTreeMap<String, Value> = BTreeMap::new();
    if let Some(bag) = &control.properties {
        let mut keys: Vec<&String> = bag.keys().collect();
        keys.sort();
        for key in keys {
            let name = key.trim().to_lowercase();
            if props.contains_key(&name) {
                return Err(failure(
                    CONFLICT,
                    "properties",
                    format!("プロパティ {key:?} が重複しています"),
                ));
            }
            props.insert(name, bag[key].clone());
        }
    }
    let typed: [(&str, Option<Value>); 13] = [
        ("caption", control.caption.clone().map(Value::from)),
        ("text", control.text.clone().map(Value::from)),
        ("value", control.value.clone()),
        ("tag", control.tag.clone().map(Value::from)),
        (
            "controltiptext",
            control.control_tip_text.clone().map(Value::from),
        ),
        ("accelerator", control.accelerator.clone().map(Value::from)),
        ("left", control.left.map(Value::from)),
        ("top", control.top.map(Value::from)),
        ("width", control.width.map(Value::from)),
        ("height", control.height.map(Value::from)),
        ("tabindex", control.tab_index.map(Value::from)),
        ("enabled", control.enabled.map(Value::from)),
        ("visible", control.visible.map(Value::from)),
    ];
    for (key, value) in typed {
        let Some(value) = value else { continue };
        if let Some(old) = props.get(key) {
            let a = generation_value(kind, key, old)
                .map_err(|e| failure(INVALID, &format!("properties.{key}"), e))?;
            let b = generation_value(kind, key, &value).map_err(|e| failure(INVALID, key, e))?;
            if a != b {
                return Err(failure(
                    CONFLICT,
                    key,
                    "プロパティの別名同士で値が食い違っています".into(),
                ));
            }
        }
        props.insert(key.to_string(), value);
    }
    let mut values: HashMap<&'static str, Converted> = HashMap::new();
    for (key, value) in &props {
        if has_typed_property(control, key) {
            continue;
        }
        let name = property_name(key).filter(|_| applicable_property(kind, key));
        let Some(mut name) = name else {
            return Err(failure(
                UNSUPPORTED,
                &format!("properties.{key}"),
                "このコントロールでは使えないプロパティです".into(),
            ));
        };
        let converted = generation_value(kind, key, value).map_err(|e| failure(INVALID, key, e))?;
        if (kind == "SpinButton" || kind == "ScrollBar") && key == "value" {
            name = "Position";
        }
        if let Some(old) = values.get(name) {
            if *old != converted {
                return Err(failure(
                    CONFLICT,
                    "",
                    "text と value の値が食い違っています".into(),
                ));
            }
        }
        values.insert(name, converted);
    }
    let size = |w: f64, h: f64| -> Result<(i32, i32), String> {
        Ok((himetric(w, "width")?, himetric(h, "height")?))
    };
    let mut item = ControlDefinition {
        name: control.name.clone(),
        class,
        size: size(default_width, default_height).map_err(|e| failure(INVALID, "", e))?,
        position: (0, 0),
        tab_index: tab_index as i16,
        visible: true,
        properties: BTreeMap::new(),
        controls: Vec::new(),
        tabs: None,
        internal: false,
    };
    if class == 7
        && (control.left.is_some()
            || control.top.is_some()
            || control.width.is_some()
            || control.height.is_some())
    {
        return Err(failure(
            UNSUPPORTED,
            "geometry",
            "Page の位置とサイズは MultiPage から決まります".into(),
        ));
    }
    if class == 18 {
        item.tabs = Some(
            generation_tabs(
                control.tabs.as_deref().unwrap_or(&[]),
                control.selected_index,
            )
            .map_err(|e| failure(INVALID, "selectedIndex", e))?,
        );
    }
    let mut enabled = None;
    for (name, value) in &values {
        match (*name, value) {
            ("Left", Converted::Geometry(v)) => item.position.0 = *v,
            ("Top", Converted::Geometry(v)) => item.position.1 = *v,
            ("Width", Converted::Geometry(v)) => item.size.0 = *v,
            ("Height", Converted::Geometry(v)) => item.size.1 = *v,
            ("TabIndex", Converted::TabIndex(v)) => item.tab_index = *v,
            ("Visible", Converted::Bool(v)) => item.visible = *v,
            ("Enabled", Converted::Bool(v)) => enabled = Some(*v),
            (name, Converted::Str(s)) => {
                item.properties
                    .insert(name.to_string(), PValue::Str(s.clone()));
            }
            (name, Converted::Int(n)) => {
                item.properties.insert(name.to_string(), PValue::Int(*n));
            }
            (name, other) => {
                return Err(failure(
                    INVALID,
                    name,
                    format!("{other:?} はこのプロパティに使えません"),
                ))
            }
        }
    }
    if contract::has_control_property(kind, "caption") && !item.properties.contains_key("Caption") {
        item.properties
            .insert("Caption".into(), PValue::Str(control.name.clone()));
    }
    if contract::has_control_property(kind, "value")
        && !matches!(kind, "ListBox" | "SpinButton" | "ScrollBar")
        && !item.properties.contains_key("Value")
    {
        let value = if (26..=28).contains(&class) { "0" } else { "" };
        item.properties
            .insert("Value".into(), PValue::Str(value.into()));
    }
    if (class == 16 || class == 47) && !item.properties.contains_key("Position") {
        item.properties.insert("Position".into(), PValue::Int(0));
    }
    if matches!(class, 7 | 14 | 57) {
        let mut bits: i64 = if class == 57 { 0xc004 } else { 0x8004 };
        if enabled == Some(false) {
            bits &= !4;
        }
        if class == 57 {
            item.properties
                .insert("MultiPageEnabled".into(), PValue::Bool(bits & 4 != 0));
        } else {
            item.properties
                .insert("BooleanProperties".into(), PValue::Int(bits));
        }
    } else {
        let mut bits = oforms::default_various_property_bits(class).unwrap_or(0x1b);
        if class == 23 || class == 25 {
            bits = 0x2c80481b;
        }
        if enabled == Some(false) {
            bits &= !2;
        }
        item.properties
            .insert("VariousPropertyBits".into(), PValue::Int(bits));
    }
    if let Some(picture) = &control.picture {
        if picture.get("remove").and_then(Value::as_bool) != Some(true) {
            let path = picture
                .get("path")
                .and_then(Value::as_str)
                .filter(|p| !p.trim().is_empty());
            let Some(path) = path else {
                return Err(failure(
                    INVALID,
                    "picture.path",
                    "画像ファイルのパスを指定してください".into(),
                ));
            };
            let Some(root) = root else {
                return Err(failure(
                    INVALID,
                    "picture.path",
                    "プロジェクトのルートが分からないため、画像を読めません".into(),
                ));
            };
            // A picture MS-Forms cannot store is converted to a bitmap here;
            // its transparent parts show the control's background color.
            let background = crate::picture::background_of(control.back_color.as_ref());
            let raw = crate::picture::load_with(root, path, background)
                .and_then(|picture| crate::picture::encode(&picture.data))
                .map_err(|e| failure(INVALID, "picture.path", e))?;
            item.properties
                .insert("Picture".into(), PValue::Picture(raw));
        }
    }
    Ok(item)
}

fn generation_tabs(tabs: &[FormSpecTab], index: Option<i64>) -> Result<TabStripState, String> {
    if index.is_some_and(|i| i < -1 || i > i32::MAX as i64) {
        return Err("selectedIndex が int32 の範囲を超えています".into());
    }
    let mut state = TabStripState {
        tabs: Vec::new(),
        selected_index: if tabs.is_empty() { -1 } else { 0 },
    };
    for input in tabs {
        let mut tab = Tab {
            name: input.name.clone(),
            caption: input.name.clone(),
            enabled: true,
            visible: true,
            ..Tab::default()
        };
        if let Some(v) = &input.caption {
            tab.caption = v.clone();
        }
        if let Some(v) = &input.control_tip_text {
            tab.control_tip_text = v.clone();
        }
        if let Some(v) = &input.tag {
            tab.tag = v.clone();
        }
        if let Some(v) = &input.accelerator {
            tab.accelerator = v.clone();
        }
        if let Some(v) = input.enabled {
            tab.enabled = v;
        }
        if let Some(v) = input.visible {
            tab.visible = v;
        }
        state.tabs.push(tab);
    }
    if let Some(index) = index {
        state.selected_index = index as i32;
    }
    let n = tabs.len() as i32;
    let s = state.selected_index;
    if s < -1 || (n == 0 && s != -1) || (n > 0 && (s < 0 || s >= n)) {
        return Err("selectedIndex がタブの範囲外です".into());
    }
    Ok(state)
}

/// `promoteLegacyAppearance`: typed appearance given only in the property
/// bag is treated as the typed field.
fn promote_legacy_appearance(control: &mut FormSpecControl) {
    let Some(bag) = control.properties.clone() else {
        return;
    };
    for (key, raw) in &bag {
        let lower = key.trim().to_lowercase();
        match lower.as_str() {
            "font" => {
                if control.font.is_some() {
                    continue;
                }
                let Some(members) = raw.as_object() else {
                    continue;
                };
                if members.keys().all(|k| {
                    matches!(
                        k.as_str(),
                        "name"
                            | "size"
                            | "bold"
                            | "italic"
                            | "underline"
                            | "strikeThrough"
                            | "charset"
                    )
                }) {
                    if let Ok(font) = serde_json::from_value::<FormSpecFont>(raw.clone()) {
                        control.font = Some(font);
                    }
                }
            }
            "forecolor" | "backcolor" => {
                let Some(n) = raw
                    .as_f64()
                    .filter(|n| n.trunc() == *n && *n >= 0.0 && *n <= u32::MAX as f64)
                else {
                    continue;
                };
                let color = crate::spec::FormSpecColor::from_ole(n as u32);
                if key.eq_ignore_ascii_case("foreColor") && control.fore_color.is_none() {
                    control.fore_color = Some(color.clone());
                }
                if key.eq_ignore_ascii_case("backColor") && control.back_color.is_none() {
                    control.back_color = Some(color);
                }
            }
            "textalign" | "borderstyle" | "picturealignment" | "picturesizemode" => {
                let Some(n) = raw
                    .as_f64()
                    .filter(|n| n.trunc() == *n && (0.0..=4.0).contains(n))
                else {
                    continue;
                };
                let target = match lower.as_str() {
                    "textalign" => &mut control.text_align,
                    "borderstyle" => &mut control.border_style,
                    "picturealignment" => &mut control.picture_alignment,
                    _ => &mut control.picture_size_mode,
                };
                if target.is_none() {
                    *target = Some(n as i64);
                }
            }
            _ => {}
        }
    }
}

fn font_edit(font: &FormSpecFont) -> FontEdit {
    FontEdit {
        name: font.name.clone(),
        size: font.size,
        bold: font.bold,
        italic: font.italic,
        underline: font.underline,
        strike_through: font.strike_through,
        charset: font.charset,
    }
}

/// Generates the designer storage of a validated, normalised spec. Picture
/// paths are resolved against `root`, the project root.
pub(crate) fn compile_new(
    spec: &FormSpec,
    codepage: u16,
    root: Option<&Path>,
) -> Result<Designer, CompileError> {
    let form_name = &spec.form.name;
    if !spec.coordinate_system.is_empty()
        && spec.coordinate_system != "points"
        && spec.coordinate_system != "parent-relative"
    {
        return Err(fail(
            INVALID,
            "",
            "coordinateSystem",
            "points か parent-relative にしてください",
        ));
    }
    validate_component_name(form_name, codepage).map_err(|e| fail(INVALID, "", "form.name", e))?;
    let mut root_size = (
        himetric(240.0, "width").expect("default"),
        himetric(180.0, "height").expect("default"),
    );
    if let Some(w) = spec.form.width {
        root_size.0 = himetric(w, "width").map_err(|e| fail(INVALID, "", "form.width", e))?;
    }
    if let Some(h) = spec.form.height {
        root_size.1 = himetric(h, "height").map_err(|e| fail(INVALID, "", "form.height", e))?;
    }
    let caption = spec
        .form
        .caption
        .clone()
        .unwrap_or_else(|| form_name.clone());
    if !valid_generation_name(form_name) {
        return Err(fail(INVALID, "", "form.name", "VBA の識別子にしてください"));
    }
    let mut promoted = spec.clone();
    for control in &mut promoted.controls {
        promote_legacy_appearance(control);
    }
    if let Some(issue) = crate::validate::validate_strict(&promoted)
        .into_iter()
        .find(|i| i.severity == crate::Severity::Error)
    {
        return Err(fail(INVALID, "", &issue.field, issue.message));
    }
    let spec = &promoted;
    let mut original_index = HashMap::new();
    let mut children: HashMap<&str, Vec<&FormSpecControl>> = HashMap::new();
    for (i, control) in spec.controls.iter().enumerate() {
        original_index.insert(control.id.as_str(), i);
        children
            .entry(control.parent_id.as_str())
            .or_default()
            .push(control);
    }
    for list in children.values_mut() {
        list.sort_by_key(|c| c.z_index.unwrap_or(0));
    }
    fn build(
        parent: &str,
        depth: usize,
        children: &HashMap<&str, Vec<&FormSpecControl>>,
        original_index: &HashMap<&str, usize>,
        root: Option<&Path>,
    ) -> Result<Vec<ControlDefinition>, CompileError> {
        let mut result = Vec::new();
        for (sibling, control) in children
            .get(parent)
            .map(Vec::as_slice)
            .unwrap_or(&[])
            .iter()
            .enumerate()
        {
            let path = format!("controls[{}]", original_index[control.id.as_str()]);
            let mut item =
                control_definition(control, sibling, root).map_err(|(code, property, reason)| {
                    let property = if property.is_empty() {
                        path.clone()
                    } else {
                        format!("{path}.{property}")
                    };
                    fail(code, &control.name, &property, reason)
                })?;
            if children
                .get(control.id.as_str())
                .is_some_and(|c| !c.is_empty())
            {
                if !contract::control_type(&control.control_type).is_some_and(|t| t.container) {
                    return Err(fail(
                        UNSUPPORTED,
                        &control.name,
                        &format!("{path}.controls"),
                        "このコントロールには子コントロールを入れられません",
                    ));
                }
                if depth + 1 >= oforms::MAX_NESTING_DEPTH {
                    return Err(fail(
                        INVALID,
                        &control.name,
                        &format!("{path}.controls"),
                        "入れ子が深すぎます",
                    ));
                }
                item.controls = build(&control.id, depth + 1, children, original_index, root)?;
            }
            if item.class == 57 {
                let mut state = TabStripState {
                    tabs: Vec::new(),
                    selected_index: if item.controls.is_empty() { -1 } else { 0 },
                };
                if let Some(index) = control.selected_index {
                    if index < -1 || index > i32::MAX as i64 {
                        return Err(fail(
                            INVALID,
                            &control.name,
                            &format!("{path}.selectedIndex"),
                            "selectedIndex が int32 の範囲を超えています",
                        ));
                    }
                    state.selected_index = index as i32;
                }
                for page in &item.controls {
                    let text = |name: &str| match page.properties.get(name) {
                        Some(PValue::Str(s)) => Some(s.clone()),
                        _ => None,
                    };
                    let mut tab = Tab {
                        name: format!("Tab_{}", page.name),
                        caption: page.name.clone(),
                        enabled: true,
                        visible: page.visible,
                        ..Tab::default()
                    };
                    if let Some(v) = text("Caption") {
                        tab.caption = v;
                    }
                    if let Some(v) = text("ControlTipText") {
                        tab.control_tip_text = v;
                    }
                    if let Some(v) = text("Accelerator") {
                        tab.accelerator = v;
                    }
                    if let Some(PValue::Int(bits)) = page.properties.get("BooleanProperties") {
                        tab.enabled = bits & 4 != 0;
                    }
                    state.tabs.push(tab);
                }
                let n = state.tabs.len() as i32;
                let s = state.selected_index;
                if (n == 0 && s != -1) || (n > 0 && (s < 0 || s >= n)) {
                    return Err(fail(
                        INVALID,
                        &control.name,
                        &format!("{path}.selectedIndex"),
                        "selectedIndex が Page の範囲外です",
                    ));
                }
                item.tabs = Some(state);
            }
            if control.control_type.eq_ignore_ascii_case("Frame")
                && depth + 1 >= oforms::MAX_NESTING_DEPTH
            {
                return Err(fail(INVALID, &control.name, &path, "入れ子が深すぎます"));
            }
            result.push(item);
        }
        Ok(result)
    }
    let controls = build("", 0, &children, &original_index, root)?;
    let definition = Definition {
        name: form_name.clone(),
        caption,
        size: root_size,
        controls,
    };
    let mut form = oforms::new_form(&definition, codepage).map_err(|e| {
        let mut error = from_persist(e);
        if error.reason.contains("caption") {
            error.property = "form.caption".into();
        }
        error
    })?;

    // Typed appearance (`generationTypedEdits` + `ApplyEdits`).
    let typed_failure = |e: PersistError| fail(INVALID, "", "typed properties", e.to_string());
    let mut apply = |control: Option<&str>, property: &'static str, edit: Edit| {
        form.apply(control, property, &edit).map_err(typed_failure)
    };
    let color = |c: &crate::spec::FormSpecColor| {
        c.ole()
            .map(i64::from)
            .map_err(|e| fail(INVALID, "", "typed properties", e))
    };
    if let Some(font) = &spec.form.font {
        let edit = font_edit(font);
        if !edit.is_empty() {
            apply(None, "Font", Edit::Font(edit))?;
        }
    }
    if let Some(c) = &spec.form.fore_color {
        apply(None, "ForeColor", Edit::Int("ForeColor", color(c)?))?;
    }
    if let Some(c) = &spec.form.back_color {
        apply(None, "BackColor", Edit::Int("BackColor", color(c)?))?;
    }
    if let Some(b) = spec.form.border_style {
        apply(None, "BorderStyle", Edit::Int("BorderStyle", b))?;
    }
    for control in &spec.controls {
        let name = Some(control.name.as_str());
        if let Some(font) = &control.font {
            let edit = font_edit(font);
            if !edit.is_empty() {
                apply(name, "Font", Edit::Font(edit))?;
            }
        }
        if let Some(c) = &control.fore_color {
            apply(name, "ForeColor", Edit::Int("ForeColor", color(c)?))?;
        }
        if let Some(c) = &control.back_color {
            apply(name, "BackColor", Edit::Int("BackColor", color(c)?))?;
        }
        for (property, value) in [
            ("TextAlign", control.text_align),
            ("BorderStyle", control.border_style),
            ("PictureAlignment", control.picture_alignment),
            ("PictureSizeMode", control.picture_size_mode),
            // ComboBox Style is the record's DisplayStyle: combo (3) or drop-down list (7).
            ("DisplayStyle", control.style.map(|s| if s == 2 { 7 } else { 3 })),
            ("MultiSelect", control.multi_select),
            ("ListStyle", control.list_style),
            ("MatchEntry", control.match_entry),
            ("ShowDropButtonWhen", control.show_drop_button_when),
            ("ListRows", control.list_rows),
        ] {
            if let Some(v) = value {
                apply(name, property, Edit::Int(property, v))?;
            }
        }
    }
    form.serialize().map_err(from_persist)
}
