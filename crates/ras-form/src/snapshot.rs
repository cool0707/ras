//! Designer storage → FormSpec: taking an existing user form out of a
//! workbook.
//!
//! Port of xlflow's `internal/vba/userforms/projection` and the picture
//! extraction of `internal/filepull` (Copyright (c) 2026 harumiWeb, MIT
//! License; see `NOTICE.md`). On top of xlflow's projection, properties that
//! the generator can write but the projection reports as unmodelled (`Tag`,
//! `ControlTipText`, and `BorderColor` of the controls `scalars.rs` does not
//! cover) are carried over, and the plain properties of `scalars.rs` are
//! projected where they differ from a new control's, so that they survive a
//! round trip through `forms/`.

use std::collections::BTreeMap;

use serde_json::{json, Map, Value};

use crate::contract;
use crate::oforms::{default_various_property_bits, Tab};
use crate::picture::{self, Picture};
use crate::reader::{self, Control, Level, ReadError, Record};
use crate::scalars::{self, Store};
use crate::spec::{self, FormSpec, FormSpecColor, FormSpecControl, FormSpecFont, FormSpecTab};

/// `type` of a control the FormSpec cannot express (snapshot only).
pub const PLACEHOLDER_TYPE: &str = "Control";
const UNSUPPORTED_TYPE_PROPERTY: &str = "controlType";

/// Something a snapshot could not carry over.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SnapshotWarning {
    /// `unsupported_properties`, `unsupported_control_type`,
    /// `unnamed_control_placeholder` or `unsupported_picture_resource`.
    pub code: &'static str,
    /// Name of the control concerned; empty for the form.
    pub control: String,
    /// Message in Japanese.
    pub message: String,
}

/// The picture of an `Image` control found in a designer storage.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SnapshotPicture {
    /// Index of the control in [`FormSpec::controls`].
    pub control_index: usize,
    /// The picture file (BMP or JPEG) as stored.
    pub picture: Picture,
}

/// A user form read back from its designer storage.
#[derive(Clone, Debug, PartialEq)]
pub struct Snapshot {
    /// The form as a FormSpec. `picture` is not set on any control; see
    /// [`Snapshot::pictures`].
    pub spec: FormSpec,
    /// Pictures of `Image` controls, to be saved as files and referenced
    /// with `picture.path`.
    pub pictures: Vec<SnapshotPicture>,
    /// What could not be carried over.
    pub warnings: Vec<SnapshotWarning>,
}

/// Why a designer storage could not be turned into a FormSpec.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum SnapshotError {
    /// The storage is not a designer storage this crate can read.
    Read(ReadError),
    /// The storage was read but its contents contradict each other.
    Inconsistent(String),
}

impl std::fmt::Display for SnapshotError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            SnapshotError::Read(e) => write!(f, "{e}"),
            SnapshotError::Inconsistent(m) => write!(f, "デザイナストレージの内容が矛盾しています（{m}）"),
        }
    }
}

impl std::error::Error for SnapshotError {}

/// HIMETRIC → points.
pub(crate) fn points(value: i32) -> f64 {
    value as f64 * 72.0 / 2540.0
}

fn lower_first(name: &str) -> String {
    let mut chars = name.chars();
    match chars.next() {
        Some(first) => first.to_lowercase().chain(chars).collect(),
        None => String::new(),
    }
}

/// (kind, type, ProgID)
const IDENTITIES: &[(&str, &str, &str)] = &[
    ("MSForms.CheckBox", "CheckBox", "Forms.CheckBox.1"),
    ("MSForms.ComboBox", "ComboBox", "Forms.ComboBox.1"),
    ("MSForms.CommandButton", "CommandButton", "Forms.CommandButton.1"),
    ("MSForms.Form", "Page", "Forms.Page.1"),
    ("MSForms.Frame", "Frame", "Forms.Frame.1"),
    ("MSForms.Image", "Image", "Forms.Image.1"),
    ("MSForms.Label", "Label", "Forms.Label.1"),
    ("MSForms.ListBox", "ListBox", "Forms.ListBox.1"),
    ("MSForms.MultiPage", "MultiPage", "Forms.MultiPage.1"),
    ("MSForms.OptionButton", "OptionButton", "Forms.OptionButton.1"),
    ("MSForms.ScrollBar", "ScrollBar", "Forms.ScrollBar.1"),
    ("MSForms.SpinButton", "SpinButton", "Forms.SpinButton.1"),
    ("MSForms.TabStrip", "TabStrip", "Forms.TabStrip.1"),
    ("MSForms.TextBox", "TextBox", "Forms.TextBox.1"),
    ("MSForms.ToggleButton", "ToggleButton", "Forms.ToggleButton.1"),
];

/// Scalar record fields without a FormSpec counterpart and their public name.
const UNSUPPORTED_SCALARS: &[(&str, &str)] = &[
    ("Accelerator", "accelerator"),
    ("BackColor", "backColor"),
    ("BorderColor", "borderColor"),
    ("BorderStyle", "borderStyle"),
    ("BooleanProperties", "booleanProperties"),
    ("BoundColumn", "boundColumn"),
    ("ColumnCount", "columnCount"),
    ("Cycle", "cycle"),
    ("Delay", "delay"),
    ("DropButtonStyle", "dropButtonStyle"),
    ("ForeColor", "foreColor"),
    ("HelpContextID", "helpContextId"),
    ("LargeChange", "largeChange"),
    ("ListRows", "listRows"),
    ("ListStyle", "listStyle"),
    ("ListWidth", "listWidth"),
    ("MatchEntry", "matchEntry"),
    ("Max", "max"),
    ("MaxLength", "maxLength"),
    ("Min", "min"),
    ("MousePointer", "mousePointer"),
    ("MultiSelect", "multiSelect"),
    ("Orientation", "orientation"),
    ("PasswordChar", "passwordChar"),
    ("PictureAlignment", "pictureAlignment"),
    ("PicturePosition", "picturePosition"),
    ("PictureSizeMode", "pictureSizeMode"),
    ("Position", "position"),
    ("ProportionalThumb", "proportionalThumb"),
    ("ScrollBars", "scrollBars"),
    ("ShowDropButtonWhen", "showDropButtonWhen"),
    ("SmallChange", "smallChange"),
    ("SpecialEffect", "specialEffect"),
    ("TabFixedHeight", "tabFixedHeight"),
    ("TabFixedWidth", "tabFixedWidth"),
    ("TabOrientation", "tabOrientation"),
    ("TabStyle", "tabStyle"),
    ("TakeFocusOnClick", "takeFocusOnClick"),
    ("TextColumn", "textColumn"),
    ("Zoom", "zoom"),
];

fn has_property(control_type: &str, name: &str) -> bool {
    contract::has_control_property(control_type, name)
}

fn supports_selected_index(control_type: &str) -> bool {
    matches!(control_type, "ComboBox" | "ListBox" | "TabStrip")
}

fn supports_value(control_type: &str) -> bool {
    matches!(
        control_type,
        "TextBox"
            | "ComboBox"
            | "ListBox"
            | "CheckBox"
            | "OptionButton"
            | "ToggleButton"
            | "SpinButton"
            | "ScrollBar"
    )
}

fn supports_numeric_value(control_type: &str) -> bool {
    matches!(control_type, "SpinButton" | "ScrollBar")
}

fn default_site_flags(control_type: &str) -> i64 {
    match control_type {
        "Frame" | "MultiPage" | "Page" => 0x40023,
        "Label" => 0x32,
        _ => 0x33,
    }
}

fn default_various_bits(control_type: &str) -> Option<i64> {
    match control_type {
        // What Excel writes for a TextBox (WordWrap set).
        "TextBox" => Some(0x2c80_481b),
        "ComboBox" | "ListBox" | "CheckBox" | "OptionButton" | "ToggleButton" => {
            default_various_property_bits(23)
        }
        "Image" | "SpinButton" | "ScrollBar" | "TabStrip" => Some(0x1b),
        "Label" => default_various_property_bits(21),
        "CommandButton" => default_various_property_bits(17),
        _ => None,
    }
}

fn is_default_boolean_properties(control_type: &str, value: i64) -> bool {
    match control_type {
        "Form" => value == 0x4004,
        "Frame" | "Page" => value & !4 == 0x8000,
        "MultiPage" => value & !4 == 0xc000,
        _ => false,
    }
}

fn has_unprojected_text_properties(record: Option<&Record>) -> bool {
    let Some(record) = record else {
        return false;
    };
    let known = 1u64 | 1 << 1 | 1 << 2 | 1 << 4 | 1 << 6 | 1 << 7;
    record.mask & !known != 0
        || record
            .value("FontWeight")
            .is_some_and(|w| w != 400 && w != 700)
        || record.value("FontEffects").unwrap_or(0) & !15 != 0
        || !record.tail.is_empty()
}

/// Record contents that are not among `projected`.
fn unsupported_record_properties(record: &Record, projected: &[&str]) -> Vec<String> {
    let done = |name: &str| projected.contains(&name);
    let mut out = Vec::new();
    for name in record.values.keys() {
        if done(name) {
            continue;
        }
        if let Some((_, public)) = UNSUPPORTED_SCALARS.iter().find(|(n, _)| n == name) {
            out.push(public.to_string());
        }
    }
    let named = record
        .strings
        .keys()
        .chain(record.sizes.keys())
        .chain(record.arrays.keys());
    for name in named {
        if !done(name) {
            out.push(lower_first(name));
        }
    }
    for name in record.pictures.keys() {
        out.push(lower_first(name));
    }
    if has_unprojected_text_properties(record.text_props.as_deref()) {
        out.push("textProps".into());
    }
    if !record.tail.is_empty() {
        out.push("opaqueRecordTail".into());
    }
    out
}

fn unsupported_level_properties(level: Option<&Level>) -> Vec<String> {
    let Some(level) = level else {
        return Vec::new();
    };
    let mut out = Vec::new();
    if level.mouse_icon {
        out.push("mouseIcon".to_string());
    }
    if !level.font_raw.is_empty() && reader::level_font(level).is_err() {
        out.push("font".into());
    }
    if level.picture {
        out.push("picture".into());
    }
    for name in &level.extra_streams {
        out.push(format!("extraStream:{name}"));
    }
    out
}

fn unsupported_control_properties(control: &Control, control_type: &str) -> Vec<String> {
    let mut out = Vec::new();
    let site = &control.site;
    if site.values.contains_key("HelpContextID") {
        out.push("helpContextId".to_string());
    }
    if let Some(flags) = site.values.get("BitFlags") {
        // Visible, TabStop, Default and Cancel are projected.
        let projected = !2 & !scalars::described_site_flags(control_type);
        if flags & projected != default_site_flags(control_type) & projected {
            out.push("siteFlags".into());
        }
    }
    for name in ["Tag", "ControlTipText", "RuntimeLicKey", "ControlSource", "RowSource"] {
        if control_type == "Page" && (name == "Tag" || name == "ControlTipText") {
            continue;
        }
        if scalars::describes_site_text(control_type, name) {
            continue;
        }
        if site.strings.get(name).is_some_and(|v| !v.is_empty()) {
            out.push(lower_first(name));
        }
    }
    if let Some(record) = control.properties() {
        let mut projected = vec!["DisplayedSize", "Size", "VariousPropertyBits"];
        // An empty stored caption is a default, not state that gets lost.
        if has_property(control_type, "caption") || record.text("Caption").unwrap_or("").is_empty() {
            projected.push("Caption");
        }
        if supports_selected_index(control_type) {
            projected.push("ListIndex");
        }
        if supports_value(control_type) {
            projected.push("Value");
        }
        if supports_numeric_value(control_type) {
            projected.push("Position");
        }
        for name in ["BackColor", "ForeColor", "BorderStyle"] {
            if has_property(control_type, &lower_first(name)) {
                projected.push(name);
            }
        }
        if is_default_boolean_properties(control_type, record.value("BooleanProperties").unwrap_or(0)) {
            projected.push("BooleanProperties");
        }
        if control_type == "Image" {
            projected.extend(["PictureAlignment", "PictureSizeMode"]);
        }
        if control_type == "TabStrip" {
            projected.extend(["Items", "TipStrings", "TabNames", "Tags", "Accelerators"]);
        }
        // The plain properties (scalars.rs).
        for p in scalars::of_type(control_type) {
            if let Some(field) = p.field() {
                projected.push(field);
            }
            if let Store::Flag(flag, _) = p.store {
                projected.push(flag);
            }
        }
        // MS-OFORMS: a ListBox always has both scroll bars.
        if control_type == "ListBox" && record.value("ScrollBars") == Some(3) {
            projected.push("ScrollBars");
        }
        out.extend(unsupported_record_properties(record, &projected));
        if let Some(bits) = record.value("VariousPropertyBits") {
            // Enabled and the bits of properties are projected; a ComboBox is
            // Editable unless it is a drop-down list (its style).
            let mut done = 2 | scalars::described_bits(control_type);
            if control_type == "ComboBox" {
                done |= 1 << 14;
            }
            if default_various_bits(control_type).is_none_or(|d| bits & !done != d & !done) {
                out.push("variousPropertyBits".into());
            }
        }
        if record.value("ListIndex").is_some() && !supports_selected_index(control_type) {
            out.push("selectedIndex".into());
        }
    }
    if control.opaque {
        out.push("opaqueControlData".into());
    }
    out
}

fn spec_font(font: Option<reader::Font>) -> Option<FormSpecFont> {
    let font = font?;
    Some(FormSpecFont {
        name: font.name,
        size: font.size,
        bold: font.bold,
        italic: font.italic,
        underline: font.underline,
        strike_through: font.strike_through,
        charset: font.charset,
    })
}

/// `ReadControlFont`: a MultiPage shows the font of its hidden TabStrip.
fn control_font(control: &Control) -> Result<Option<reader::Font>, ()> {
    if let Some(multipage) = &control.multipage {
        return control_font(&multipage.hidden);
    }
    if let Some(level) = &control.level {
        return reader::level_font(level);
    }
    Ok(control
        .record
        .as_ref()
        .and_then(|r| reader::text_font(r.text_props.as_deref())))
}

fn color(value: i64) -> FormSpecColor {
    FormSpecColor::from_ole(value as u32)
}

fn project_typed_control(source: &Control, target: &mut FormSpecControl) {
    let control_type = target.control_type.clone();
    if let Ok(font) = control_font(source) {
        if has_property(&control_type, "font") {
            target.font = spec_font(font);
        }
    }
    let Some(mut record) = source.properties() else {
        return;
    };
    if let Some(multipage) = &source.multipage {
        match &multipage.hidden.record {
            Some(hidden) => record = hidden,
            None => return,
        }
    }
    if has_property(&control_type, "backColor") {
        target.back_color = record.value("BackColor").map(color);
    }
    if has_property(&control_type, "foreColor") {
        target.fore_color = record.value("ForeColor").map(color);
    }
    if has_property(&control_type, "borderStyle") {
        target.border_style = record.value("BorderStyle");
    }
    if has_property(&control_type, "textAlign") {
        if let Some(align) = record
            .text_props
            .as_ref()
            .and_then(|t| t.value("ParagraphAlign"))
        {
            // PARAFORMAT uses right = 2 / center = 3; MSForms the reverse.
            target.text_align = Some(match align {
                2 => 3,
                3 => 2,
                other => other,
            });
        }
    }
    if control_type == "Image" {
        target.picture_alignment = record.value("PictureAlignment");
        target.picture_size_mode = record.value("PictureSizeMode");
    }
    project_scalars(source, &control_type, record, target);
}

/// The plain properties (scalars.rs), each only where it differs from what
/// Excel stores for a new control (which is what the compiler writes).
fn project_scalars(source: &Control, control_type: &str, record: &Record, target: &mut FormSpecControl) {
    let Some(class) = scalars::class_of(control_type) else {
        return;
    };
    let bits = record
        .value("VariousPropertyBits")
        .or_else(|| default_various_property_bits(class))
        .unwrap_or(0x1b);
    let flags = source.site.values.get("BitFlags").copied().unwrap_or(0x33);
    for p in scalars::of_type(control_type) {
        let (raw, fresh) = match p.store {
            Store::Bit(..) | Store::Bits(..) => (
                scalars::get_bits(bits, p),
                scalars::get_bits(scalars::fresh_bits(class), p),
            ),
            Store::Site(bit) => (
                (flags >> bit) & 1,
                (scalars::fresh_site_flags(class) >> bit) & 1,
            ),
            Store::Flag(flag, set) => {
                let on = record.values.contains_key(flag) == set;
                (i64::from(on), i64::from(!set))
            }
            Store::Text(field) => {
                let text = record.text(field).unwrap_or("");
                if !text.is_empty() {
                    target.scalars.insert(p.name.into(), Value::from(text));
                }
                continue;
            }
            Store::SiteText(field) => {
                if let Some(text) = source.site.strings.get(field).filter(|t| !t.is_empty()) {
                    target.scalars.insert(p.name.into(), Value::from(text.as_str()));
                }
                continue;
            }
            _ => (
                record.value(p.field().expect("field")).unwrap_or(p.omitted),
                p.fresh.unwrap_or(p.omitted),
            ),
        };
        if raw != fresh {
            target.scalars.insert(p.name.into(), scalars::from_raw(p, raw));
        }
    }
}

fn project_control(source: &Control, target: &mut FormSpecControl) {
    project_typed_control(source, target);
    let control_type = target.control_type.clone();
    let page = control_type == "Page";
    if let Some((left, top)) = source.site.position.filter(|_| !page) {
        target.left = Some(points(left));
        target.top = Some(points(top));
    }
    let flags = source.site.values.get("BitFlags").copied().unwrap_or(0x33);
    target.visible = Some(flags & 2 != 0);
    target.tab_index = source.tab_index.map(i64::from);
    let Some(record) = source.properties() else {
        return;
    };
    let size = record
        .sizes
        .get("Size")
        .or_else(|| record.sizes.get("DisplayedSize"));
    if let Some((width, height)) = size.filter(|_| !page) {
        target.width = Some(points(*width));
        target.height = Some(points(*height));
    }
    if let Some(caption) = record.text("Caption") {
        if has_property(&control_type, "caption") {
            target.caption = Some(caption.to_string());
        }
    }
    if let Some(value) = record.text("Value").filter(|_| supports_value(&control_type)) {
        let state = match control_type.as_str() {
            "CheckBox" | "OptionButton" | "ToggleButton" => {
                match value.trim().to_lowercase().as_str() {
                    "1" | "-1" | "true" => Some("True"),
                    "0" | "false" => Some("False"),
                    _ => None,
                }
            }
            _ => None,
        };
        target.value = Some(Value::from(state.unwrap_or(value)));
        if control_type == "TextBox" {
            target.text = Some(value.to_string());
        }
    }
    if let Some(position) = record
        .value("Position")
        .filter(|_| supports_numeric_value(&control_type))
    {
        target.value = Some(Value::from(position));
    }
    if let Some(selected) = record
        .value("ListIndex")
        .filter(|_| supports_selected_index(&control_type))
    {
        target.selected_index = Some(selected);
    }
    target.enabled = Some(match record.value("VariousPropertyBits") {
        Some(bits) => bits & 2 != 0,
        // A Frame or Page keeps Enabled in bit 2 of BooleanProperties.
        None if record.spec.type_name == "Form" => {
            record.value("BooleanProperties").unwrap_or(4) & 4 != 0
        }
        None => true,
    });
}

fn project_tabs(tabs: &[Tab]) -> Vec<FormSpecTab> {
    tabs.iter()
        .map(|tab| FormSpecTab {
            name: tab.name.clone(),
            caption: Some(tab.caption.clone()),
            control_tip_text: Some(tab.control_tip_text.clone()),
            tag: Some(tab.tag.clone()),
            accelerator: Some(tab.accelerator.clone()),
            enabled: Some(tab.enabled),
            visible: Some(tab.visible),
        })
        .collect()
}

struct Projection {
    controls: Vec<FormSpecControl>,
    warnings: Vec<SnapshotWarning>,
    pictures: Vec<SnapshotPicture>,
    next_id: usize,
    unnamed: usize,
    /// Carry over what the generator can write and extract pictures.
    extras: bool,
}

fn unsupported_warning(control: &str, unsupported: &[String]) -> SnapshotWarning {
    SnapshotWarning {
        code: "unsupported_properties",
        control: control.to_string(),
        message: format!(
            "FormSpec で表せないプロパティがあります: {}（forms/ から書き戻すと既定値になります）",
            unsupported.join(", ")
        ),
    }
}

impl Projection {
    /// Moves properties the generator can write out of `unsupported`.
    fn carry_over(
        &mut self,
        source: &Control,
        control: &mut FormSpecControl,
        unsupported: &mut Vec<String>,
    ) {
        let control_type = control.control_type.clone();
        if control.caption.is_none() && has_property(&control_type, "caption") {
            // No stored caption is an empty one; generation would otherwise
            // fall back to the control name.
            control.caption = Some(String::new());
        }
        if matches!(control_type.as_str(), "ComboBox" | "ListBox") {
            // The items are not stored, so a stored selection names nothing.
            if control.selected_index.take().is_some_and(|i| i >= 0) {
                unsupported.push("selectedIndex".into());
                unsupported.sort();
            }
        }
        let mut take = |name: &str| {
            let before = unsupported.len();
            unsupported.retain(|u| u != name);
            unsupported.len() != before
        };
        if control_type != "Page" {
            if let Some(tag) = source.site.strings.get("Tag").filter(|_| take("tag")) {
                control.tag = Some(tag.clone());
            }
            if let Some(tip) = source
                .site
                .strings
                .get("ControlTipText")
                .filter(|_| take("controlTipText"))
            {
                control.control_tip_text = Some(tip.clone());
            }
        }
        let Some(record) = source.properties() else {
            return;
        };
        if matches!(control_type.as_str(), "CheckBox" | "OptionButton" | "ToggleButton")
            && record.text("Value").is_some_and(|v| {
                !matches!(v.trim().to_lowercase().as_str(), "0" | "1" | "-1" | "true" | "false")
            })
        {
            // Any other stored string is neither on nor off: Null.
            control.value = Some(Value::from("Null"));
        }
        let mut bag = Map::new();
        if control_type == "OptionButton" {
            if let Some(group) = record.text("GroupName").filter(|_| take("groupName")) {
                bag.insert("groupName".into(), Value::from(group));
            }
        }
        if matches!(control_type.as_str(), "TextBox" | "ComboBox") {
            if let Some(max) = record.value("MaxLength").filter(|_| take("maxLength")) {
                bag.insert("maxLength".into(), Value::from(max));
            }
        }
        if control_type == "ComboBox" && record.value("VariousPropertyBits").is_some_and(|b| b & !2 == 0x2c80_4819) {
            // What generation writes for a ComboBox; nothing is lost.
            take("variousPropertyBits");
        }
        if source.multipage.is_none() && control_type != PLACEHOLDER_TYPE {
            if let Some(border) = record.value("BorderColor").filter(|_| take("borderColor")) {
                bag.insert("borderColor".into(), Value::from(border));
            }
        }
        if !bag.is_empty() {
            control.properties = Some(bag);
        }
        if control_type == "Image" {
            if let Some(raw) = record.pictures.get("Picture") {
                match picture::decode(raw) {
                    Ok(picture) => {
                        take("picture");
                        self.pictures.push(SnapshotPicture {
                            control_index: self.controls.len(),
                            picture,
                        });
                    }
                    Err(reason) => self.warnings.push(SnapshotWarning {
                        code: "unsupported_picture_resource",
                        control: control.name.clone(),
                        message: format!(
                            "画像を BMP / JPEG として取り出せませんでした（{reason}）"
                        ),
                    }),
                }
            }
        }
    }

    fn controls(
        &mut self,
        controls: &[Control],
        parent_id: &str,
        page_tabs: Option<&[Tab]>,
    ) -> Result<(), SnapshotError> {
        for (index, source) in controls.iter().enumerate() {
            let inconsistent = |m: String| SnapshotError::Inconsistent(m);
            let strip = match &source.multipage {
                Some(multipage) => {
                    let Some(strip) = multipage.hidden.tab_strip.as_ref() else {
                        return Err(inconsistent(format!(
                            "MultiPage {:?} の内部 TabStrip を読めません",
                            source.name
                        )));
                    };
                    if !multipage.pages.is_empty() && strip.tabs.len() != multipage.pages.len() {
                        return Err(inconsistent(format!(
                            "MultiPage {:?} のタブ {} 個とページ {} 個が合いません",
                            source.name,
                            strip.tabs.len(),
                            multipage.pages.len()
                        )));
                    }
                    Some(strip)
                }
                None if source.kind == "MSForms.TabStrip" => {
                    let Some(strip) = source.tab_strip.as_ref() else {
                        return Err(inconsistent(format!(
                            "TabStrip {:?} のレコードがありません",
                            source.name
                        )));
                    };
                    Some(strip)
                }
                None => None,
            };

            self.next_id += 1;
            let id = format!("control_{:03}", self.next_id);
            let mut control = FormSpecControl {
                id: id.clone(),
                parent_id: parent_id.to_string(),
                z_index: Some(index as i64),
                name: source.name.clone(),
                ..FormSpecControl::default()
            };
            if control.name.trim().is_empty() {
                self.unnamed += 1;
                control.name = format!("<unnamed_{}>", self.unnamed);
                self.warnings.push(SnapshotWarning {
                    code: "unnamed_control_placeholder",
                    control: control.name.clone(),
                    message: "名前のないコントロールに仮の名前を付けました。VBA の識別子に直してください".into(),
                });
            }
            let identity = IDENTITIES.iter().find(|(kind, _, _)| *kind == source.kind);
            match identity {
                Some((_, type_name, prog_id)) => {
                    control.control_type = type_name.to_string();
                    control.prog_id = prog_id.to_string();
                }
                None => {
                    control.control_type = PLACEHOLDER_TYPE.to_string();
                    self.warnings.push(SnapshotWarning {
                        code: "unsupported_control_type",
                        control: control.name.clone(),
                        message: format!(
                            "コントロールの種類 {} は FormSpec で表せません（forms/ からは生成できません）",
                            source.kind
                        ),
                    });
                }
            }

            project_control(source, &mut control);
            if let Some(tabs) = page_tabs {
                let Some(tab) = tabs.get(index) else {
                    return Err(inconsistent(format!(
                        "Page {:?} に対応するタブがありません",
                        source.name
                    )));
                };
                control.caption = Some(tab.caption.clone());
                control.control_tip_text = Some(tab.control_tip_text.clone());
                control.accelerator = Some(tab.accelerator.clone());
                control.enabled = Some(tab.enabled);
                control.visible = Some(tab.visible);
                if let Some(tag) = source.site.strings.get("Tag") {
                    control.tag = Some(tag.clone());
                }
            }
            if let Some(strip) = strip {
                match &source.multipage {
                    Some(multipage) => {
                        let stored = multipage
                            .hidden
                            .record
                            .as_ref()
                            .and_then(|r| r.value("ListIndex"))
                            .unwrap_or(0);
                        control.selected_index =
                            Some(if multipage.pages.is_empty() { -1 } else { stored });
                        control.enabled = Some(multipage.properties.mask & (1 << 3) == 0);
                    }
                    None => {
                        control.tabs = Some(project_tabs(&strip.tabs));
                        control.selected_index = Some(if strip.tabs.is_empty() {
                            -1
                        } else {
                            strip.selected_index as i64
                        });
                    }
                }
            }

            let mut unsupported = unsupported_control_properties(source, &control.control_type);
            unsupported.extend(unsupported_level_properties(source.level.as_ref()));
            if control.control_type == "Page"
                && source.level.as_ref().is_some_and(|l| !l.font_raw.is_empty())
            {
                // A Page has no Font property; its stored font is not modelled.
                unsupported.push("font".into());
            }
            if identity.is_none() {
                unsupported.push(UNSUPPORTED_TYPE_PROPERTY.into());
            }
            unsupported.sort();
            unsupported.dedup();
            if self.extras {
                self.carry_over(source, &mut control, &mut unsupported);
            }
            if !unsupported.is_empty() {
                self.warnings
                    .push(unsupported_warning(&control.name, &unsupported));
                control.unsupported = Some(unsupported);
            }
            self.controls.push(control);

            match &source.multipage {
                Some(multipage) => {
                    let tabs = strip.map(|s| s.tabs.as_slice()).unwrap_or(&[]);
                    self.controls(&multipage.pages, &id, Some(tabs))?;
                }
                None => self.controls(&source.children, &id, None)?,
            }
        }
        Ok(())
    }
}

/// The `Caption` property of the designer block in the `VBFrame` stream.
fn frame_caption(streams: &BTreeMap<String, Vec<u8>>, codepage: u16) -> Option<String> {
    let frame = streams
        .iter()
        .find(|(path, _)| path.eq_ignore_ascii_case("\u{3}VBFrame"))?
        .1;
    let text = ras_ovba::decode_text(codepage, frame).ok()?;
    text.lines().find_map(|line| {
        let (key, value) = line.split_once('=')?;
        if !key.trim().eq_ignore_ascii_case("Caption") {
            return None;
        }
        let value = value.trim().strip_prefix('"')?.strip_suffix('"')?;
        Some(value.replace("\"\"", "\""))
    })
}

fn snapshot_with(
    name: &str,
    streams: &BTreeMap<String, Vec<u8>>,
    storages: &BTreeMap<String, [u8; 16]>,
    codepage: Option<u16>,
) -> Result<Snapshot, SnapshotError> {
    let extras = codepage.is_some();
    let form = reader::read_form(streams, storages).map_err(SnapshotError::Read)?;
    let record = &form.root.record;
    let mut spec = FormSpec {
        schema_version: 1,
        kind: spec::KIND.to_string(),
        basis: "designer".to_string(),
        coordinate_system: "parent-relative".to_string(),
        ..FormSpec::default()
    };
    spec.form.name = name.to_string();
    if let Ok(font) = reader::level_font(&form.root) {
        spec.form.font = spec_font(font);
    }
    spec.form.back_color = record.value("BackColor").map(color);
    spec.form.fore_color = record.value("ForeColor").map(color);
    spec.form.border_style = record.value("BorderStyle");
    spec.form.caption = record.text("Caption").map(str::to_string);
    if let (None, Some(codepage)) = (&spec.form.caption, codepage) {
        // Without a caption in the form record the form shows the one of
        // its designer block; without that one, none.
        spec.form.caption = Some(frame_caption(streams, codepage).unwrap_or_default());
    }
    if let Some((width, height)) = record.sizes.get("DisplayedSize") {
        spec.form.width = Some(points(*width));
        spec.form.height = Some(points(*height));
    }

    let mut projection = Projection {
        controls: Vec::new(),
        warnings: Vec::new(),
        pictures: Vec::new(),
        next_id: 0,
        unnamed: 0,
        extras,
    };
    let mut projected = vec!["Caption", "DisplayedSize", "BackColor", "ForeColor", "BorderStyle"];
    if is_default_boolean_properties("Form", record.value("BooleanProperties").unwrap_or(0)) {
        projected.push("BooleanProperties");
    }
    let mut unsupported = unsupported_record_properties(record, &projected);
    unsupported.extend(unsupported_level_properties(Some(&form.root)));
    unsupported.sort();
    unsupported.dedup();
    if extras && record.sizes.get("LogicalSize") == Some(&(0, 0)) {
        // An empty scroll area is what generation writes.
        unsupported.retain(|u| u != "logicalSize");
    }
    if !unsupported.is_empty() {
        projection
            .warnings
            .push(unsupported_warning("", &unsupported));
    }
    projection.controls(&form.controls, "", None)?;

    spec.controls = projection.controls;
    spec::normalize(&mut spec);
    let warnings: Vec<Value> = projection
        .warnings
        .iter()
        .map(|w| {
            let mut entry = json!({ "code": w.code, "message": w.message });
            if !w.control.is_empty() {
                entry["control"] = Value::from(w.control.clone());
            }
            entry
        })
        .collect();
    spec.warnings = Some(Value::Array(warnings));
    Ok(Snapshot {
        spec,
        pictures: projection.pictures,
        warnings: projection.warnings,
    })
}

/// Turns the designer storage of the form `name` into a FormSpec.
///
/// `streams` and `storages` are keyed by their path below the form storage,
/// as in [`crate::Designer`] and `ras_ovba::UserForm`; `codepage` is the
/// code page of the project the form comes from.
pub fn snapshot(
    name: &str,
    streams: &BTreeMap<String, Vec<u8>>,
    storages: &BTreeMap<String, [u8; 16]>,
    codepage: u16,
) -> Result<Snapshot, SnapshotError> {
    snapshot_with(name, streams, storages, Some(codepage))
}

/// xlflow's projection without the additions of [`snapshot`]; the result is
/// what xlflow produces for the same storage (apart from `kind` and the
/// language of the warnings). For comparing against xlflow.
#[doc(hidden)]
pub fn project(
    name: &str,
    streams: &BTreeMap<String, Vec<u8>>,
    storages: &BTreeMap<String, [u8; 16]>,
) -> Result<Snapshot, SnapshotError> {
    snapshot_with(name, streams, storages, None)
}
