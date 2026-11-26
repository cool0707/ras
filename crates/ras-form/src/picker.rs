//! The file / folder picker (`FilePicker`), a control RAS composes of a text box and a button
//! (RAS language specification §14.11.14).
//!
//! ```text
//! <name>_Text                                    <name>_Button
//! [C:\Users\me\Documents\report.xlsx          ][...]
//! ```
//!
//! The button keeps the settings of the dialog (`Application.FileDialog`) in its `Tag`
//! (`RAS.FilePicker|type=3|multi=0|title=…|filters=Excel,*.xlsx;…`, the values escaped with
//! `%XX` for `% | = ; ,`). The bundled RAS class `FilePicker`
//! (`ras-compiler/stdlib/filepicker.ras`) reads it when the form loads and shows the dialog when
//! the button is clicked.
//!
//! The layout here is the one of the class's `layout()` and of the debug run's `pickerParts`
//! (`ras-compiler/src/js/forms.js`); the three must agree.

use serde_json::{json, Map, Value};

use crate::spec::{FormSpec, FormSpecControl};

/// The FormSpec type of a file picker.
pub const TYPE: &str = "FilePicker";
/// The first item of the button's `Tag`.
pub const TAG_PREFIX: &str = "RAS.FilePicker";
/// The controls a file picker is made of, by the suffix of their names.
pub const PARTS: [&str; 2] = ["_Text", "_Button"];
/// Width when omitted, in points.
pub const DEFAULT_WIDTH: f64 = 200.0;
/// Height when omitted, in points.
pub const DEFAULT_HEIGHT: f64 = 18.0;
/// Width of the button when omitted, in points.
pub const DEFAULT_BUTTON_WIDTH: f64 = 24.0;
/// Caption of the button when omitted.
pub const DEFAULT_BUTTON_CAPTION: &str = "...";
/// Font size of the button when its caption is the default, in points: at 10 pt Excel clips
/// the "..." of a 24 pt wide button.
pub const DEFAULT_BUTTON_FONT_SIZE: f64 = 9.0;
/// `dialogType` when omitted (`msoFileDialogFilePicker`).
pub const DEFAULT_DIALOG_TYPE: i64 = 3;

/// Whether `control_type` is the file picker.
pub fn is_file_picker(control_type: &str) -> bool {
    control_type.trim().eq_ignore_ascii_case(TYPE)
}

/// One file type of the dialog (`FileDialog.Filters.Add`).
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Filter {
    /// `"Excel ブック"`.
    pub description: String,
    /// `"*.xlsx; *.xlsm"`.
    pub extensions: String,
}

/// The settings of the dialog of a file picker, with the defaults filled in.
#[derive(Clone, Debug, PartialEq)]
pub struct Settings {
    /// `msoFileDialogType`: 1 Open, 2 SaveAs, 3 FilePicker, 4 FolderPicker.
    pub dialog_type: i64,
    /// `AllowMultiSelect`.
    pub multi: bool,
    /// `Title` (empty: the dialog's own).
    pub title: String,
    /// `ButtonName` (empty: the dialog's own).
    pub button_name: String,
    /// `InitialFileName`.
    pub initial: String,
    /// `InitialView` (0: the dialog's own).
    pub view: i64,
    /// `FilterIndex`.
    pub filter_index: i64,
    /// `Filters`.
    pub filters: Vec<Filter>,
}

fn text(control: &FormSpecControl, name: &str) -> String {
    control
        .scalar(name)
        .and_then(Value::as_str)
        .unwrap_or_default()
        .to_string()
}

/// The filters written in `filters`, if it is an array of `{ description, extensions }`.
pub fn filters(value: &Value) -> Option<Vec<Filter>> {
    value
        .as_array()?
        .iter()
        .map(|f| {
            Some(Filter {
                description: f.get("description")?.as_str()?.to_string(),
                extensions: f.get("extensions")?.as_str()?.to_string(),
            })
        })
        .collect()
}

impl Settings {
    /// The settings of `control`.
    pub fn of(control: &FormSpecControl) -> Settings {
        let int = |name: &str, default: i64| {
            control
                .scalar(name)
                .and_then(Value::as_i64)
                .unwrap_or(default)
        };
        Settings {
            dialog_type: int("dialogType", DEFAULT_DIALOG_TYPE),
            multi: control
                .scalar("allowMultiSelect")
                .and_then(Value::as_bool)
                .unwrap_or(false),
            title: text(control, "title"),
            button_name: text(control, "buttonName"),
            initial: text(control, "initialFileName"),
            view: int("initialView", 0),
            filter_index: int("filterIndex", 1),
            filters: control
                .scalar("filters")
                .and_then(filters)
                .unwrap_or_default(),
        }
    }
}

/// `text` with `% | = ; ,` written as `%XX`.
pub fn escape(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '%' => out.push_str("%25"),
            '|' => out.push_str("%7C"),
            '=' => out.push_str("%3D"),
            ';' => out.push_str("%3B"),
            ',' => out.push_str("%2C"),
            c => out.push(c),
        }
    }
    out
}

/// The text [`escape`] made `text` of.
pub fn unescape(text: &str) -> String {
    text.replace("%7C", "|")
        .replace("%3D", "=")
        .replace("%3B", ";")
        .replace("%2C", ",")
        .replace("%25", "%")
}

/// The `Tag` of the button: the settings of the dialog.
pub fn tag(settings: &Settings) -> String {
    let filters: Vec<String> = settings
        .filters
        .iter()
        .map(|f| format!("{},{}", escape(&f.description), escape(&f.extensions)))
        .collect();
    format!(
        "{TAG_PREFIX}|type={}|multi={}|title={}|buttonName={}|initial={}|view={}|filterIndex={}|filters={}",
        settings.dialog_type,
        u8::from(settings.multi),
        escape(&settings.title),
        escape(&settings.button_name),
        escape(&settings.initial),
        settings.view,
        settings.filter_index,
        filters.join(";")
    )
}

/// The settings recorded in the `Tag` of a button, if it is one.
fn read_tag(tag: &str) -> Option<Settings> {
    let mut items = tag.split('|');
    if items.next()? != TAG_PREFIX {
        return None;
    }
    let mut s = Settings {
        dialog_type: DEFAULT_DIALOG_TYPE,
        multi: false,
        title: String::new(),
        button_name: String::new(),
        initial: String::new(),
        view: 0,
        filter_index: 1,
        filters: Vec::new(),
    };
    for item in items {
        let Some((key, value)) = item.split_once('=') else {
            continue;
        };
        let n = value.trim().parse::<i64>().ok();
        match key {
            "type" => s.dialog_type = n.unwrap_or(DEFAULT_DIALOG_TYPE),
            "multi" => s.multi = n.unwrap_or(0) != 0,
            "title" => s.title = unescape(value),
            "buttonName" => s.button_name = unescape(value),
            "initial" => s.initial = unescape(value),
            "view" => s.view = n.unwrap_or(0),
            "filterIndex" => s.filter_index = n.unwrap_or(1),
            "filters" => {
                s.filters = value
                    .split(';')
                    .filter_map(|pair| {
                        let (d, e) = pair.split_once(',')?;
                        Some(Filter {
                            description: unescape(d),
                            extensions: unescape(e),
                        })
                    })
                    .collect()
            }
            _ => {}
        }
    }
    Some(s)
}

/// Where the parts go: `(left, top, width, height)` of the text box and of the button.
pub fn layout(
    left: f64,
    top: f64,
    width: f64,
    height: f64,
    button_width: f64,
) -> [(f64, f64, f64, f64); 2] {
    let bw = button_width.min(width).max(0.75);
    let h = height.max(0.75);
    [
        (left, top, (width - bw).max(0.75), h),
        (left + width - bw, top, bw, h),
    ]
}

/// The text box and the button `control` (a file picker) is written as.
pub fn parts(control: &FormSpecControl) -> Vec<FormSpecControl> {
    let settings = Settings::of(control);
    let width = control.width.unwrap_or(DEFAULT_WIDTH);
    let height = control.height.unwrap_or(DEFAULT_HEIGHT);
    let button_width = control
        .scalar("buttonWidth")
        .and_then(Value::as_f64)
        .unwrap_or(DEFAULT_BUTTON_WIDTH);
    let [b, n] = layout(
        control.left.unwrap_or(0.0),
        control.top.unwrap_or(0.0),
        width,
        height,
        button_width,
    );
    let part = |suffix: &str, (l, t, w, h): (f64, f64, f64, f64)| FormSpecControl {
        id: format!("{}{suffix}", control.id),
        parent_id: control.parent_id.clone(),
        z_index: control.z_index,
        name: format!("{}{suffix}", control.name),
        control_tip_text: control.control_tip_text.clone(),
        left: Some(l),
        top: Some(t),
        width: Some(w),
        height: Some(h),
        enabled: control.enabled,
        visible: control.visible,
        font: control.font.clone(),
        ..FormSpecControl::default()
    };
    let mut text_box = part(PARTS[0], b);
    text_box.control_type = "TextBox".into();
    text_box.text = control.text.clone().or_else(|| {
        control
            .value
            .as_ref()
            .and_then(Value::as_str)
            .map(str::to_string)
    });
    text_box.tag = control.tag.clone();
    text_box.tab_index = control.tab_index;
    text_box.fore_color = control.fore_color.clone();
    text_box.back_color = control.back_color.clone();
    if let Some(locked) = control.scalar("locked") {
        text_box.scalars.insert("locked".into(), locked.clone());
    }
    let mut button = part(PARTS[1], n);
    button.control_type = "CommandButton".into();
    button.caption = Some(
        control
            .scalar("buttonCaption")
            .and_then(Value::as_str)
            .unwrap_or(DEFAULT_BUTTON_CAPTION)
            .to_string(),
    );
    if button.caption.as_deref() == Some(DEFAULT_BUTTON_CAPTION) {
        if let Some(font) = &mut button.font {
            if font.size.is_some() {
                font.size = Some(DEFAULT_BUTTON_FONT_SIZE);
            }
        }
    }
    button.tag = Some(tag(&settings));
    // A click on the button leaves the focus in the box.
    button
        .scalars
        .insert("takeFocusOnClick".into(), json!(false));
    vec![text_box, button]
}

/// Replaces each file picker of `spec` by its parts, in its place.
pub fn expand(spec: &FormSpec) -> FormSpec {
    if !spec
        .controls
        .iter()
        .any(|c| is_file_picker(&c.control_type))
    {
        return spec.clone();
    }
    let mut out = spec.clone();
    out.controls = spec
        .controls
        .iter()
        .flat_map(|c| {
            if is_file_picker(&c.control_type) {
                parts(c)
            } else {
                vec![c.clone()]
            }
        })
        .collect();
    out
}

/// A length in points rounded to whole pixels at 96 DPI (0.75 pt).
fn pixels(points: f64) -> f64 {
    (points / 0.75).round() * 0.75
}

/// Turns the parts of each file picker of a form read from a workbook (`ras form pull`) back
/// into the `FilePicker`, in the place of its text box.
pub fn collapse(spec: &mut FormSpec) {
    loop {
        let found = spec.controls.iter().enumerate().find_map(|(i, c)| {
            let base = c.name.strip_suffix(PARTS[1])?;
            if !c.control_type.eq_ignore_ascii_case("CommandButton") {
                return None;
            }
            let settings = read_tag(c.tag.as_deref()?)?;
            let text_at = spec.controls.iter().position(|o| {
                o.name.eq_ignore_ascii_case(&format!("{base}{}", PARTS[0]))
                    && o.control_type.eq_ignore_ascii_case("TextBox")
                    && o.parent_id == c.parent_id
            })?;
            Some((i, text_at, base.to_string(), settings))
        });
        let Some((button_at, text_at, base, s)) = found else {
            return;
        };
        let text_box = &spec.controls[text_at];
        let button = &spec.controls[button_at];
        let mut scalars = Map::new();
        let mut put = |name: &str, v: Value| {
            scalars.insert(name.to_string(), v);
        };
        if s.dialog_type != DEFAULT_DIALOG_TYPE {
            put("dialogType", json!(s.dialog_type));
        }
        if s.multi {
            put("allowMultiSelect", json!(true));
        }
        for (name, value) in [
            ("title", &s.title),
            ("buttonName", &s.button_name),
            ("initialFileName", &s.initial),
        ] {
            if !value.is_empty() {
                put(name, json!(value));
            }
        }
        if s.view != 0 {
            put("initialView", json!(s.view));
        }
        if !s.filters.is_empty() {
            put(
                "filters",
                Value::Array(
                    s.filters
                        .iter()
                        .map(
                            |f| json!({ "description": f.description, "extensions": f.extensions }),
                        )
                        .collect(),
                ),
            );
        }
        if s.filter_index != 1 {
            put("filterIndex", json!(s.filter_index));
        }
        let caption = button.caption.clone().unwrap_or_default();
        if caption != DEFAULT_BUTTON_CAPTION {
            put("buttonCaption", json!(caption));
        }
        let button_width = pixels(button.width.unwrap_or(DEFAULT_BUTTON_WIDTH));
        if button_width != DEFAULT_BUTTON_WIDTH {
            put("buttonWidth", json!(button_width));
        }
        if let Some(locked) = text_box.scalar("locked") {
            put("locked", locked.clone());
        }
        let left = text_box.left.map(pixels);
        let right = button.left.unwrap_or(0.0) + button.width.unwrap_or(0.0);
        let picker = FormSpecControl {
            id: base.to_lowercase(),
            parent_id: text_box.parent_id.clone(),
            z_index: text_box.z_index,
            control_type: TYPE.into(),
            name: base,
            text: text_box.text.clone().filter(|t| !t.is_empty()),
            tag: text_box.tag.clone(),
            control_tip_text: text_box.control_tip_text.clone(),
            left,
            top: text_box.top.map(pixels),
            width: Some(pixels(right - text_box.left.unwrap_or(0.0))),
            height: text_box.height.map(pixels),
            tab_index: text_box.tab_index,
            enabled: text_box.enabled,
            visible: text_box.visible,
            font: text_box.font.clone(),
            fore_color: text_box.fore_color.clone(),
            back_color: text_box.back_color.clone(),
            scalars,
            ..FormSpecControl::default()
        };
        let mut at = [text_at, button_at];
        at.sort_unstable();
        // Where the text box was, once the parts in front of it are gone.
        let place = text_at - at.iter().filter(|&&i| i < text_at).count();
        for i in at.into_iter().rev() {
            spec.controls.remove(i);
        }
        spec.controls.insert(place, picker);
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

    const PICKER: &str = r#"{"id":"fp","name":"fpInput","type":"FilePicker","left":12,"top":12,"width":240,"height":18,"text":"C:\\data\\a.xlsx","title":"ブックを選択","allowMultiSelect":true,"filters":[{"description":"Excel ブック","extensions":"*.xlsx; *.xlsm"},{"description":"CSV, TSV","extensions":"*.csv;*.tsv"}],"filterIndex":2,"buttonCaption":"参照","buttonWidth":36}"#;

    #[test]
    fn the_default_button_is_drawn_at_9_pt() {
        let font = r#""font":{"name":"Yu Gothic UI","size":10}"#;
        let spec_of = |extra: &str| {
            spec(&format!(
                r#"{{"id":"fp","name":"fp","type":"FilePicker",{font}{extra}}}"#
            ))
        };
        let expanded = expand(&spec_of(""));
        let size = |i: usize| expanded.controls[i].font.as_ref().and_then(|f| f.size);
        assert_eq!(
            (size(0), size(1)),
            (Some(10.0), Some(DEFAULT_BUTTON_FONT_SIZE))
        );
        assert_eq!(
            expanded.controls[1]
                .font
                .as_ref()
                .and_then(|f| f.name.as_deref()),
            Some("Yu Gothic UI")
        );
        // A caption of its own keeps the picker's font.
        let expanded = expand(&spec_of(r#","buttonCaption":"参照""#));
        assert_eq!(
            expanded.controls[1].font.as_ref().and_then(|f| f.size),
            Some(10.0)
        );
    }

    #[test]
    fn a_file_picker_is_a_text_box_and_a_button() {
        let spec = spec(PICKER);
        let expanded = expand(&spec);
        let [text_box, button] = &expanded.controls[..] else {
            unreachable!()
        };
        assert_eq!(text_box.name, "fpInput_Text");
        assert_eq!(text_box.control_type, "TextBox");
        assert_eq!(text_box.text.as_deref(), Some("C:\\data\\a.xlsx"));
        assert_eq!((text_box.left, text_box.width), (Some(12.0), Some(204.0)));
        assert_eq!(button.name, "fpInput_Button");
        assert_eq!((button.left, button.width), (Some(216.0), Some(36.0)));
        assert_eq!(button.caption.as_deref(), Some("参照"));
        assert_eq!(
            button.tag.as_deref(),
            Some("RAS.FilePicker|type=3|multi=1|title=ブックを選択|buttonName=|initial=|view=0|filterIndex=2|filters=Excel ブック,*.xlsx%3B *.xlsm;CSV%2C TSV,*.csv%3B*.tsv")
        );
        crate::compile(&spec, 932).unwrap();
    }

    #[test]
    fn parts_read_from_a_workbook_are_a_file_picker_again() {
        let spec = spec(&format!(
            r#"{{"id":"no","name":"btnNo","type":"CommandButton","top":90}},{PICKER},{{"id":"ok","name":"btnOK","type":"CommandButton","top":60}}"#
        ));
        let designer = crate::compile(&spec, 932).unwrap();
        let mut read = crate::snapshot("F", &designer.streams, &designer.storages, 932)
            .unwrap()
            .spec;
        collapse(&mut read);
        let names: Vec<&str> = read.controls.iter().map(|c| c.name.as_str()).collect();
        assert_eq!(names, ["btnNo", "fpInput", "btnOK"]);
        let picker = &read.controls[1];
        assert_eq!(picker.control_type, TYPE);
        assert_eq!(picker.name, "fpInput");
        assert_eq!(picker.text.as_deref(), Some("C:\\data\\a.xlsx"));
        assert_eq!(
            (picker.left, picker.top, picker.width, picker.height),
            (Some(12.0), Some(12.0), Some(240.0), Some(18.0))
        );
        let original = &spec.controls[1];
        for name in [
            "title",
            "allowMultiSelect",
            "filters",
            "filterIndex",
            "buttonCaption",
        ] {
            assert_eq!(picker.scalar(name), original.scalar(name), "{name}");
        }
        assert_eq!(
            picker.scalar("buttonWidth").and_then(Value::as_f64),
            Some(36.0)
        );
    }

    #[test]
    fn escaping_goes_round() {
        let text = "a%b|c=d;e,f%7C";
        assert_eq!(unescape(&escape(text)), text);
    }

    #[test]
    fn checks() {
        for (control, field) in [
            (
                r#"{"id":"a","name":"fp","type":"FilePicker","dialogType":5}"#,
                "controls[0].dialogType",
            ),
            (
                r#"{"id":"a","name":"fp","type":"FilePicker","filters":[{"description":"x"}]}"#,
                "controls[0].filters",
            ),
            (
                r#"{"id":"a","name":"fp","type":"FilePicker","dialogType":4,"filters":[{"description":"x","extensions":"*.x"}]}"#,
                "controls[0].filters",
            ),
            (
                r#"{"id":"a","name":"fp","type":"FilePicker","filters":[{"description":"x","extensions":"*.x"}],"filterIndex":2}"#,
                "controls[0].filterIndex",
            ),
            (
                r#"{"id":"a","name":"fp","type":"FilePicker","dialogType":2,"allowMultiSelect":true}"#,
                "controls[0].allowMultiSelect",
            ),
            (
                r#"{"id":"a","name":"fp","type":"FilePicker"},{"id":"b","name":"fp_Text","type":"Label"}"#,
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
                    .any(|i| (i.code == "UFV025" || i.code == "UFV003") && i.field == field),
                "{control}: {:?}",
                parsed.issues
            );
        }
    }
}
