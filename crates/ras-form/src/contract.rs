//! The FormSpec property contract: which fields exist where, their value
//! types and support levels.
//!
//! Port of xlflow's `internal/vba/userforms/spec/contract.go` and the editor
//! additions of `appearance.go` (Copyright (c) 2026 harumiWeb, MIT License;
//! see `NOTICE.md`).

/// JSON value type of a property.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[allow(missing_docs)]
pub enum ValueType {
    String,
    Number,
    Integer,
    Boolean,
    Any,
    StringArray,
    Object,
    ObjectArray,
}

impl ValueType {
    /// Name used in messages.
    pub fn label(self) -> &'static str {
        match self {
            ValueType::String => "文字列",
            ValueType::Number => "数値",
            ValueType::Integer => "整数",
            ValueType::Boolean => "真偽値",
            ValueType::Any => "任意の値",
            ValueType::StringArray => "文字列の配列",
            ValueType::Object => "オブジェクト",
            ValueType::ObjectArray => "オブジェクトの配列",
        }
    }
}

/// How well a property is supported by generation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[allow(missing_docs)]
pub enum Support {
    Supported,
    BestEffort,
    ObservedOnly,
    SnapshotOnly,
    CustomUnchecked,
}

impl Support {
    /// The xlflow spelling (`observed-only`, ...).
    pub fn label(self) -> &'static str {
        match self {
            Support::Supported => "supported",
            Support::BestEffort => "best-effort",
            Support::ObservedOnly => "observed-only",
            Support::SnapshotOnly => "snapshot-only",
            Support::CustomUnchecked => "custom/unchecked",
        }
    }
}

/// Contract of one property.
#[derive(Clone, Debug)]
pub struct Prop {
    /// Expected JSON type.
    pub value_type: ValueType,
    /// Must be present.
    pub required: bool,
    /// Support level.
    pub support: Support,
    /// Allowed values (compared with the JSON text of the value); empty = any.
    pub allowed: &'static [&'static str],
    /// Nested properties of an object value (font, color, picture).
    pub children: Option<Vec<(&'static str, Prop)>>,
}

const fn p(value_type: ValueType, required: bool, support: Support) -> Prop {
    Prop {
        value_type,
        required,
        support,
        allowed: &[],
        children: None,
    }
}

fn allowed(mut prop: Prop, values: &'static [&'static str]) -> Prop {
    prop.allowed = values;
    prop
}

use Support::{CustomUnchecked, ObservedOnly, SnapshotOnly, Supported};
use ValueType as T;

/// A property list keyed by name.
pub type Props = Vec<(&'static str, Prop)>;

/// Looks a property up, ignoring case.
pub fn lookup<'a>(props: &'a Props, name: &str) -> Option<&'a Prop> {
    props
        .iter()
        .find(|(k, _)| *k == name)
        .or_else(|| {
            props
                .iter()
                .find(|(k, _)| k.eq_ignore_ascii_case(name.trim()))
        })
        .map(|(_, v)| v)
}

fn font() -> Prop {
    let mut font = p(T::Object, false, Supported);
    font.children = Some(vec![
        ("name", p(T::String, false, Supported)),
        ("size", p(T::Number, false, Supported)),
        ("bold", p(T::Boolean, false, Supported)),
        ("italic", p(T::Boolean, false, Supported)),
        ("underline", p(T::Boolean, false, Supported)),
        ("strikeThrough", p(T::Boolean, false, Supported)),
        ("charset", p(T::Integer, false, Supported)),
    ]);
    font
}

fn color() -> Prop {
    let mut color = p(T::Object, false, Supported);
    color.children = Some(vec![
        (
            "kind",
            allowed(p(T::String, true, Supported), &["rgb", "system", "ole"]),
        ),
        ("value", p(T::String, true, Supported)),
    ]);
    color
}

fn border_style() -> Prop {
    allowed(p(T::Integer, false, Supported), &["0", "1"])
}

/// Fields of the document root.
pub fn document_properties() -> Props {
    vec![
        (
            "schemaVersion",
            allowed(p(T::Integer, true, Supported), &["1"]),
        ),
        (
            "kind",
            allowed(
                p(T::String, true, Supported),
                &["ras.userform", "xlflow.userform"],
            ),
        ),
        (
            "basis",
            allowed(p(T::String, true, Supported), &["designer"]),
        ),
        (
            "coordinateSystem",
            allowed(
                p(T::String, false, Supported),
                &["points", "parent-relative"],
            ),
        ),
        ("form", p(T::Object, true, Supported)),
        ("controls", p(T::ObjectArray, true, Supported)),
        ("warnings", p(T::ObjectArray, false, SnapshotOnly)),
    ]
}

/// Fields of `form`.
pub fn form_properties() -> Props {
    vec![
        ("name", p(T::String, true, Supported)),
        ("caption", p(T::String, false, Supported)),
        ("width", p(T::Number, false, Supported)),
        ("height", p(T::Number, false, Supported)),
        ("observed", p(T::Object, false, SnapshotOnly)),
        ("font", font()),
        ("foreColor", color()),
        ("backColor", color()),
        ("borderStyle", border_style()),
    ]
}

/// Fields of `form.observed`.
pub fn form_observed_properties() -> Props {
    [
        "caption",
        "width",
        "height",
        "insideWidth",
        "insideHeight",
        "clientWidth",
        "clientHeight",
    ]
    .into_iter()
    .map(|name| {
        (
            name,
            p(
                if name == "caption" {
                    T::String
                } else {
                    T::Number
                },
                false,
                SnapshotOnly,
            ),
        )
    })
    .collect()
}

/// Fields of a TabStrip tab.
pub fn tab_properties() -> Props {
    vec![
        ("name", p(T::String, true, Supported)),
        ("caption", p(T::String, false, Supported)),
        ("controlTipText", p(T::String, false, Supported)),
        ("tag", p(T::String, false, Supported)),
        ("accelerator", p(T::String, false, Supported)),
        ("enabled", p(T::Boolean, false, Supported)),
        ("visible", p(T::Boolean, false, Supported)),
    ]
}

/// Fields of a picture action.
pub fn picture_properties() -> Props {
    vec![
        ("path", p(T::String, false, Supported)),
        (
            "remove",
            allowed(p(T::Boolean, false, Supported), &["true"]),
        ),
    ]
}

/// Fields every control has.
pub fn common_control_properties() -> Props {
    vec![
        ("id", p(T::String, true, Supported)),
        ("name", p(T::String, true, Supported)),
        ("type", p(T::String, true, Supported)),
        ("progId", p(T::String, false, Supported)),
        ("parentId", p(T::String, false, Supported)),
        ("zIndex", p(T::Integer, false, Supported)),
        ("left", p(T::Number, false, Supported)),
        ("top", p(T::Number, false, Supported)),
        ("width", p(T::Number, false, Supported)),
        ("height", p(T::Number, false, Supported)),
        ("tabIndex", p(T::Integer, false, Supported)),
        ("enabled", p(T::Boolean, false, Supported)),
        ("visible", p(T::Boolean, false, Supported)),
        ("observed", p(T::Object, false, SnapshotOnly)),
        ("controls", p(T::ObjectArray, false, SnapshotOnly)),
        ("properties", p(T::Object, false, CustomUnchecked)),
        ("unsupported", p(T::StringArray, false, SnapshotOnly)),
    ]
}

/// A built-in control type.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ControlType {
    /// Canonical type name.
    pub name: &'static str,
    /// Default ProgID.
    pub prog_id: &'static str,
    /// Can be the parent of other controls.
    pub container: bool,
}

/// All built-in control types, in Toolbox order.
pub const CONTROL_TYPES: &[ControlType] = &[
    ControlType {
        name: "Label",
        prog_id: "Forms.Label.1",
        container: false,
    },
    ControlType {
        name: "TextBox",
        prog_id: "Forms.TextBox.1",
        container: false,
    },
    ControlType {
        name: "ComboBox",
        prog_id: "Forms.ComboBox.1",
        container: false,
    },
    ControlType {
        name: "ListBox",
        prog_id: "Forms.ListBox.1",
        container: false,
    },
    ControlType {
        name: "CommandButton",
        prog_id: "Forms.CommandButton.1",
        container: false,
    },
    ControlType {
        name: "CheckBox",
        prog_id: "Forms.CheckBox.1",
        container: false,
    },
    ControlType {
        name: "OptionButton",
        prog_id: "Forms.OptionButton.1",
        container: false,
    },
    ControlType {
        name: "ToggleButton",
        prog_id: "Forms.ToggleButton.1",
        container: false,
    },
    ControlType {
        name: "SpinButton",
        prog_id: "Forms.SpinButton.1",
        container: false,
    },
    ControlType {
        name: "ScrollBar",
        prog_id: "Forms.ScrollBar.1",
        container: false,
    },
    ControlType {
        name: "Image",
        prog_id: "Forms.Image.1",
        container: false,
    },
    ControlType {
        name: "Frame",
        prog_id: "Forms.Frame.1",
        container: true,
    },
    ControlType {
        name: "MultiPage",
        prog_id: "Forms.MultiPage.1",
        container: true,
    },
    ControlType {
        name: "Page",
        prog_id: "Forms.Page.1",
        container: true,
    },
    ControlType {
        name: "TabStrip",
        prog_id: "Forms.TabStrip.1",
        container: false,
    },
];

/// The built-in type `name` (ignoring case and surrounding blanks).
pub fn control_type(name: &str) -> Option<&'static ControlType> {
    let name = name.trim();
    CONTROL_TYPES
        .iter()
        .find(|t| t.name.eq_ignore_ascii_case(name))
}

/// The built-in type whose ProgID is `prog_id`.
pub fn control_type_by_prog_id(prog_id: &str) -> Option<&'static ControlType> {
    let prog_id = prog_id.trim();
    if prog_id.is_empty() {
        return None;
    }
    CONTROL_TYPES
        .iter()
        .find(|t| t.prog_id.eq_ignore_ascii_case(prog_id))
}

/// Type-specific properties of a built-in control (without the common ones).
pub fn type_properties(name: &str) -> Option<Props> {
    let t = control_type(name)?;
    let s = || p(T::String, false, Supported);
    let any = || p(T::Any, false, Supported);
    let mut props: Props = match t.name {
        "Label" | "CommandButton" | "Frame" => vec![("caption", s())],
        "TextBox" => vec![("text", s()), ("value", any())],
        "ComboBox" | "ListBox" => vec![
            ("text", s()),
            ("value", any()),
            ("list", p(T::StringArray, false, ObservedOnly)),
            ("selectedIndex", p(T::Integer, false, ObservedOnly)),
        ],
        "CheckBox" | "OptionButton" | "ToggleButton" => vec![("caption", s()), ("value", any())],
        "SpinButton" | "ScrollBar" => vec![("value", p(T::Integer, false, Supported))],
        "Image" => {
            let mut picture = p(T::Object, false, Supported);
            picture.children = Some(picture_properties());
            vec![
                ("picture", picture),
                (
                    "pictureAlignment",
                    allowed(p(T::Integer, false, Supported), &["0", "1", "2", "3", "4"]),
                ),
                (
                    "pictureSizeMode",
                    allowed(p(T::Integer, false, Supported), &["0", "1", "3"]),
                ),
            ]
        }
        "MultiPage" => vec![("selectedIndex", p(T::Integer, false, Supported))],
        "Page" => vec![
            ("caption", s()),
            ("controlTipText", s()),
            ("tag", s()),
            ("accelerator", s()),
        ],
        "TabStrip" => vec![
            ("selectedIndex", p(T::Integer, false, Supported)),
            ("tabs", p(T::ObjectArray, false, Supported)),
        ],
        _ => vec![],
    };
    let n = t.name;
    if !matches!(n, "Page" | "Image" | "SpinButton" | "ScrollBar") {
        props.push(("font", font()));
    }
    if n != "Page" {
        props.push(("backColor", color()));
    }
    if !matches!(n, "Image" | "Page") {
        props.push(("foreColor", color()));
    }
    if matches!(
        n,
        "Label" | "TextBox" | "ComboBox" | "ListBox" | "CheckBox" | "OptionButton" | "ToggleButton"
    ) {
        props.push((
            "textAlign",
            allowed(p(T::Integer, false, Supported), &["1", "2", "3"]),
        ));
    }
    if matches!(
        n,
        "Frame" | "TextBox" | "ComboBox" | "ListBox" | "Label" | "Image"
    ) {
        props.push(("borderStyle", border_style()));
    }
    Some(props)
}

/// All properties of a control type: the common ones (minus exclusions)
/// plus the type-specific ones.
pub fn control_properties(name: &str) -> Props {
    let mut props = common_control_properties();
    let Some(specific) = type_properties(name) else {
        return props;
    };
    if control_type(name).map(|t| t.name) == Some("Page") {
        props.retain(|(k, _)| !matches!(*k, "left" | "top" | "width" | "height"));
    }
    for (k, v) in specific {
        props.retain(|(existing, _)| *existing != k);
        props.push((k, v));
    }
    props
}

/// Whether `name` is a property of `type_name` (common or specific).
pub fn has_control_property(type_name: &str, name: &str) -> bool {
    lookup(&control_properties(type_name), name).is_some()
}

/// Whether a known parent type accepts a known child type. `None` when
/// neither is a built-in type.
pub fn parent_allows_child(
    parent: Option<&ControlType>,
    child: Option<&ControlType>,
) -> Option<bool> {
    let Some(parent) = parent else {
        // A Page can only be parented by a MultiPage.
        return child.filter(|c| c.name == "Page").map(|_| false);
    };
    if !parent.container {
        return Some(false);
    }
    if parent.name == "MultiPage" && child.map(|c| c.name) != Some("Page") {
        return Some(false);
    }
    if child.map(|c| c.name) == Some("Page") && parent.name != "MultiPage" {
        return Some(false);
    }
    Some(true)
}
