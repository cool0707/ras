//! The plain properties of the built-in controls: one value each, stored in a
//! bit of `VariousPropertyBits`, a field of the control's record, a flag of its
//! site, or a flag of its property mask ([MS-OFORMS] 2.2 and 2.5).
//!
//! One table drives the FormSpec contract, validation, generation and the
//! projection of an existing form, so that a property is described once. The
//! FormSpec value of each is the value of the VBA property of the same name.
//!
//! A property left out of a FormSpec is written as Excel writes it for a new
//! control ([`Scalar::fresh`]), and a pulled form shows only the properties that
//! differ from that.

use serde_json::Value;

use crate::spec::FormSpecColor;

/// The FormSpec value of a property.
#[derive(Clone, Copy, Debug)]
pub(crate) enum Kind {
    Bool,
    /// One of these integers.
    Enum(&'static [i64]),
    /// An integer in this range.
    Int(i64, i64),
    /// Points from 0 to this, stored in HIMETRIC.
    Points(f64),
    /// One character, or "" for none.
    Char,
    /// A FormSpec color.
    Color,
    /// A string.
    Text,
    /// `ColumnWidths`: widths separated by `;` (see [`parse_column_widths`]).
    Widths,
}

/// Where a property is stored.
#[derive(Clone, Copy, Debug)]
pub(crate) enum Store {
    /// A bit of `VariousPropertyBits`; inverted when the bit is set for value 0.
    Bit(u8, bool),
    /// Bits of `VariousPropertyBits` from this position and of this width.
    Bits(u8, u8),
    /// A record field holding the value.
    Field(&'static str),
    /// A record field holding the points in HIMETRIC.
    Himetric(&'static str),
    /// `ProportionalThumb`: -1 (0xFFFF) for true, 0 for false.
    Thumb,
    /// A property-mask bit without data; the flag is set for this value.
    Flag(&'static str, bool),
    /// A bit of the site's `BitFlags`.
    Site(u8),
    /// A record string field.
    Text(&'static str),
    /// A string of the control's site (`ControlSource`, `RowSource`).
    SiteText(&'static str),
    /// The `MorphDataColumnInfo` array after the record (`cColumnInfo` entries).
    Columns,
    /// A property of the form's designer block (the `VBFrame` stream, [MS-OVBA] 2.3.5).
    Frame(&'static str),
    /// Bits of a record field from this position and of this width.
    FieldBits(&'static str, u8, u8),
    /// The width (0) or height (1) of a record's size or position, in HIMETRIC.
    Extent(&'static str, u8),
}

/// One property.
#[derive(Debug)]
pub(crate) struct Scalar {
    /// FormSpec name (the VBA name in lower camel case).
    pub name: &'static str,
    /// Control types it applies to.
    pub types: &'static [&'static str],
    pub kind: Kind,
    pub store: Store,
    /// Stored value when the field is omitted (the file-format default).
    pub omitted: i64,
    /// Stored value Excel writes for a new control, when not `omitted`.
    pub fresh: Option<i64>,
}

const EDITS: &[&str] = &["TextBox", "ComboBox"];
/// The form itself (`FORM_SCALARS`).
const FORM: &[&str] = &["UserForm"];
/// The largest scroll area of a form, in points (HIMETRIC that fits an `i32`).
const SCROLL_SIZE: f64 = 1_000_000.0;
const LISTS: &[&str] = &["ComboBox", "ListBox"];
const FOCUSABLE: &[&str] = &[
    "TextBox",
    "ComboBox",
    "ListBox",
    "CommandButton",
    "CheckBox",
    "OptionButton",
    "ToggleButton",
    "SpinButton",
    "ScrollBar",
    "MultiPage",
    "TabStrip",
];
/// Controls whose value can be bound to a cell (`ControlSource`).
const BOUND: &[&str] = &[
    "TextBox",
    "ComboBox",
    "ListBox",
    "CheckBox",
    "OptionButton",
    "ToggleButton",
    "SpinButton",
    "ScrollBar",
];
/// fmMousePointer: 4 and 5 are not values of it ([MS-OFORMS] 2.5.x).
const POINTERS: &[i64] = &[0, 1, 2, 3, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 99];
const EFFECTS: &[i64] = &[0, 1, 2, 3, 6];
/// 254000 HIMETRIC, the most a tab can be.
const TAB_SIZE: f64 = 7200.0;

const fn s(
    name: &'static str,
    types: &'static [&'static str],
    kind: Kind,
    store: Store,
    omitted: i64,
    fresh: Option<i64>,
) -> Scalar {
    Scalar {
        name,
        types,
        kind,
        store,
        omitted,
        fresh,
    }
}

use Kind::{Bool, Char, Color, Enum, Int, Points, Text, Widths};
use Store::{
    Bit, Bits, Columns, Extent, Field, FieldBits, Flag, Frame, Himetric, Site, SiteText, Thumb,
};

/// Every plain property. A name can appear more than once for different types.
pub(crate) static SCALARS: &[Scalar] = &[
    // VariousPropertyBits (the bit values come from the control's default bits).
    s(
        "locked",
        &[
            "TextBox",
            "ComboBox",
            "ListBox",
            "CheckBox",
            "OptionButton",
            "ToggleButton",
            "CommandButton",
        ],
        Bool,
        Bit(2, false),
        0,
        None,
    ),
    s(
        "backStyle",
        &[
            "TextBox",
            "ComboBox",
            "CheckBox",
            "OptionButton",
            "ToggleButton",
            "CommandButton",
            "Label",
        ],
        Enum(&[0, 1]),
        Bit(3, false),
        0,
        None,
    ),
    s(
        "integralHeight",
        &["TextBox", "ListBox"],
        Bool,
        Bit(11, false),
        0,
        None,
    ),
    s(
        "matchRequired",
        &["ComboBox"],
        Bool,
        Bit(12, false),
        0,
        None,
    ),
    // fmAlignmentLeft (0) puts the caption to the left: the bit is set.
    s(
        "alignment",
        &["CheckBox", "OptionButton"],
        Enum(&[0, 1]),
        Bit(13, true),
        0,
        None,
    ),
    s(
        "imeMode",
        &["TextBox", "ComboBox", "ListBox"],
        Enum(&[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]),
        Bits(15, 4),
        0,
        None,
    ),
    s(
        "dragBehavior",
        EDITS,
        Enum(&[0, 1]),
        Bit(19, false),
        0,
        None,
    ),
    s(
        "enterKeyBehavior",
        &["TextBox"],
        Bool,
        Bit(20, false),
        0,
        None,
    ),
    s(
        "enterFieldBehavior",
        EDITS,
        Enum(&[0, 1]),
        Bit(21, false),
        0,
        None,
    ),
    s(
        "tabKeyBehavior",
        &["TextBox"],
        Bool,
        Bit(22, false),
        0,
        None,
    ),
    s(
        "wordWrap",
        &[
            "TextBox",
            "CheckBox",
            "OptionButton",
            "ToggleButton",
            "CommandButton",
            "Label",
        ],
        Bool,
        Bit(23, false),
        0,
        None,
    ),
    s("selectionMargin", EDITS, Bool, Bit(26, false), 0, None),
    s("autoWordSelect", EDITS, Bool, Bit(27, false), 0, None),
    s(
        "autoSize",
        &[
            "TextBox",
            "ComboBox",
            "CheckBox",
            "OptionButton",
            "ToggleButton",
            "CommandButton",
            "Label",
        ],
        Bool,
        Bit(28, false),
        0,
        None,
    ),
    s("hideSelection", EDITS, Bool, Bit(29, false), 0, None),
    s("autoTab", EDITS, Bool, Bit(30, false), 0, None),
    s("multiLine", &["TextBox"], Bool, Bit(31, false), 0, None),
    // Record fields of TextBox, ComboBox, ListBox and the toggles (MorphData).
    s(
        "maxLength",
        EDITS,
        Int(0, i32::MAX as i64),
        Field("MaxLength"),
        0,
        None,
    ),
    s(
        "scrollBars",
        &["TextBox"],
        Enum(&[0, 1, 2, 3]),
        Field("ScrollBars"),
        0,
        None,
    ),
    s(
        "passwordChar",
        &["TextBox"],
        Char,
        Field("PasswordChar"),
        0,
        None,
    ),
    s(
        "style",
        &["ComboBox"],
        Enum(&[0, 2]),
        Field("DisplayStyle"),
        3,
        None,
    ),
    s(
        "listStyle",
        &["ComboBox", "ListBox"],
        Enum(&[0, 1]),
        Field("ListStyle"),
        0,
        None,
    ),
    s(
        "matchEntry",
        &["ComboBox"],
        Enum(&[0, 1, 2]),
        Field("MatchEntry"),
        2,
        Some(1),
    ),
    s(
        "matchEntry",
        &["ListBox"],
        Enum(&[0, 1, 2]),
        Field("MatchEntry"),
        2,
        Some(0),
    ),
    s(
        "showDropButtonWhen",
        &["ComboBox"],
        Enum(&[0, 1, 2]),
        Field("ShowDropButtonWhen"),
        0,
        Some(2),
    ),
    s(
        "dropButtonStyle",
        &["ComboBox"],
        Enum(&[0, 1, 2, 3]),
        Field("DropButtonStyle"),
        1,
        None,
    ),
    s(
        "listRows",
        &["ComboBox"],
        Int(0, i16::MAX as i64),
        Field("ListRows"),
        8,
        None,
    ),
    s(
        "listWidth",
        &["ComboBox"],
        Points(TAB_SIZE * 10.0),
        Himetric("ListWidth"),
        0,
        None,
    ),
    s(
        "multiSelect",
        &["ListBox"],
        Enum(&[0, 1, 2]),
        Field("MultiSelect"),
        0,
        None,
    ),
    // The columns of the items (`list` gives an item's columns as an array).
    s(
        "columnCount",
        LISTS,
        Int(-1, i16::MAX as i64),
        Field("ColumnCount"),
        1,
        None,
    ),
    s("columnWidths", LISTS, Widths, Columns, 0, None),
    s(
        "boundColumn",
        LISTS,
        Int(0, i16::MAX as i64),
        Field("BoundColumn"),
        1,
        None,
    ),
    s(
        "textColumn",
        LISTS,
        Int(-1, i16::MAX as i64),
        Field("TextColumn"),
        -1,
        None,
    ),
    s("columnHeads", LISTS, Bool, Bit(10, false), 0, None),
    // The toggles keep TripleState in the field of MultiSelect (fmMultiSelectMulti for True),
    // as Excel writes it; [MS-OFORMS] names no place for it.
    s(
        "tripleState",
        &["CheckBox", "OptionButton", "ToggleButton"],
        Bool,
        Field("MultiSelect"),
        0,
        None,
    ),
    s(
        "groupName",
        &["OptionButton"],
        Text,
        Store::Text("GroupName"),
        0,
        None,
    ),
    s(
        "specialEffect",
        &["TextBox", "ComboBox", "ListBox"],
        Enum(EFFECTS),
        Field("SpecialEffect"),
        2,
        None,
    ),
    s(
        "specialEffect",
        &["CheckBox", "OptionButton"],
        Enum(&[0, 2]),
        Field("SpecialEffect"),
        2,
        None,
    ),
    s(
        "specialEffect",
        &["Label"],
        Enum(EFFECTS),
        Field("SpecialEffect"),
        0,
        None,
    ),
    s(
        "specialEffect",
        &["Image"],
        Enum(EFFECTS),
        Field("SpecialEffect"),
        0,
        None,
    ),
    // A Frame's own form record; Excel etches a new Frame.
    s(
        "specialEffect",
        &["Frame"],
        Enum(EFFECTS),
        Field("SpecialEffect"),
        0,
        Some(3),
    ),
    s(
        "borderColor",
        &["TextBox", "ComboBox", "ListBox", "Label", "Image", "Frame"],
        Color,
        Field("BorderColor"),
        0x8000_0006,
        None,
    ),
    s("cycle", &["Frame"], Enum(&[0, 2]), Field("Cycle"), 0, None),
    s(
        "mousePointer",
        &[
            "Label",
            "TextBox",
            "ComboBox",
            "ListBox",
            "CommandButton",
            "CheckBox",
            "OptionButton",
            "ToggleButton",
            "SpinButton",
            "ScrollBar",
            "Image",
            "Frame",
            "TabStrip",
        ],
        Enum(POINTERS),
        Field("MousePointer"),
        0,
        None,
    ),
    // Where the picture is beside the caption (fmPicturePosition 0–12, stored as MS-OFORMS
    // packs it: see PICTURE_POSITIONS).
    s(
        "picturePosition",
        &[
            "Label",
            "CommandButton",
            "CheckBox",
            "OptionButton",
            "ToggleButton",
        ],
        Enum(&[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]),
        Field("PicturePosition"),
        0x0007_0001,
        None,
    ),
    // The background picture of a Frame or Page repeated over it.
    s(
        "pictureTiling",
        &["Frame", "Page"],
        Bool,
        Flag("PictureTiling", true),
        0,
        None,
    ),
    // CommandButton.
    s("default", &["CommandButton"], Bool, Site(2), 0, None),
    s("cancel", &["CommandButton"], Bool, Site(3), 0, None),
    s(
        "takeFocusOnClick",
        &["CommandButton"],
        Bool,
        Flag("TakeFocusOnClick", false),
        0,
        None,
    ),
    s("tabStop", FOCUSABLE, Bool, Site(0), 0, None),
    // The cell the value is bound to, and the range the items come from (site strings).
    s(
        "controlSource",
        BOUND,
        Text,
        SiteText("ControlSource"),
        0,
        None,
    ),
    s(
        "rowSource",
        &["ComboBox", "ListBox"],
        Text,
        SiteText("RowSource"),
        0,
        None,
    ),
    // SpinButton and ScrollBar.
    s(
        "min",
        &["SpinButton", "ScrollBar"],
        Int(i32::MIN as i64, i32::MAX as i64),
        Field("Min"),
        0,
        None,
    ),
    s(
        "max",
        &["SpinButton"],
        Int(i32::MIN as i64, i32::MAX as i64),
        Field("Max"),
        100,
        None,
    ),
    s(
        "max",
        &["ScrollBar"],
        Int(i32::MIN as i64, i32::MAX as i64),
        Field("Max"),
        32767,
        None,
    ),
    s(
        "smallChange",
        &["SpinButton", "ScrollBar"],
        Int(i32::MIN as i64, i32::MAX as i64),
        Field("SmallChange"),
        1,
        None,
    ),
    s(
        "largeChange",
        &["ScrollBar"],
        Int(i32::MIN as i64, i32::MAX as i64),
        Field("LargeChange"),
        1,
        None,
    ),
    s(
        "orientation",
        &["SpinButton", "ScrollBar"],
        Enum(&[-1, 0, 1]),
        Field("Orientation"),
        -1,
        None,
    ),
    s("proportionalThumb", &["ScrollBar"], Bool, Thumb, -1, None),
    s(
        "delay",
        &["SpinButton", "ScrollBar"],
        Int(0, i32::MAX as i64),
        Field("Delay"),
        50,
        None,
    ),
    // TabStrip, and MultiPage through its hidden TabStrip.
    s(
        "style",
        &["TabStrip", "MultiPage"],
        Enum(&[0, 1, 2]),
        Field("TabStyle"),
        0,
        None,
    ),
    s(
        "tabOrientation",
        &["TabStrip", "MultiPage"],
        Enum(&[0, 1, 2, 3]),
        Field("TabOrientation"),
        0,
        None,
    ),
    s(
        "multiRow",
        &["TabStrip", "MultiPage"],
        Bool,
        Flag("MultiRow", true),
        0,
        None,
    ),
    s(
        "tabFixedWidth",
        &["TabStrip", "MultiPage"],
        Points(TAB_SIZE),
        Himetric("TabFixedWidth"),
        0,
        None,
    ),
    s(
        "tabFixedHeight",
        &["TabStrip", "MultiPage"],
        Points(TAB_SIZE),
        Himetric("TabFixedHeight"),
        0,
        None,
    ),
];

/// The plain properties of the form itself: those of its designer block (`VBFrame`) and
/// those of its own record (the `f` stream). `omitted` is the value the form has without
/// the line or the field; `fresh` what generation writes for a new form.
pub(crate) static FORM_SCALARS: &[Scalar] = &[
    s(
        "startUpPosition",
        FORM,
        Enum(&[0, 1, 2, 3]),
        Frame("StartUpPosition"),
        0,
        Some(1),
    ),
    s("showModal", FORM, Bool, Frame("ShowModal"), 1, None),
    s("tag", FORM, Text, Frame("Tag"), 0, None),
    s(
        "helpContextId",
        FORM,
        Int(i32::MIN as i64, i32::MAX as i64),
        Frame("HelpContextID"),
        0,
        None,
    ),
    s("rightToLeft", FORM, Bool, Frame("RightToLeft"), 0, None),
    s(
        "whatsThisButton",
        FORM,
        Bool,
        Frame("WhatsThisButton"),
        0,
        None,
    ),
    s("whatsThisHelp", FORM, Bool, Frame("WhatsThisHelp"), 0, None),
    s(
        "enabled",
        FORM,
        Bool,
        FieldBits("BooleanProperties", 2, 1),
        4,
        None,
    ),
    s(
        "borderColor",
        FORM,
        Color,
        Field("BorderColor"),
        0x8000_0012,
        None,
    ),
    s("cycle", FORM, Enum(&[0, 2]), Field("Cycle"), 0, None),
    s(
        "mousePointer",
        FORM,
        Enum(POINTERS),
        Field("MousePointer"),
        0,
        None,
    ),
    s(
        "specialEffect",
        FORM,
        Enum(EFFECTS),
        Field("SpecialEffect"),
        0,
        None,
    ),
    s("zoom", FORM, Int(10, 400), Field("Zoom"), 100, None),
    // MS-OFORMS states no default for DrawBuffer: every form stores it.
    s(
        "drawBuffer",
        FORM,
        Int(16000, 1_048_576),
        Field("DrawBuffer"),
        32000,
        None,
    ),
    s(
        "scrollBars",
        FORM,
        Enum(&[0, 1, 2, 3]),
        FieldBits("ScrollBars", 0, 2),
        0x0c,
        None,
    ),
    s(
        "keepScrollBarsVisible",
        FORM,
        Enum(&[0, 1, 2, 3]),
        FieldBits("ScrollBars", 2, 2),
        0x0c,
        None,
    ),
    // 0 is the size of the form; omitted, the scroll area is 4000 × 3000 HIMETRIC.
    s(
        "scrollWidth",
        FORM,
        Points(SCROLL_SIZE),
        Extent("LogicalSize", 0),
        4000,
        Some(0),
    ),
    s(
        "scrollHeight",
        FORM,
        Points(SCROLL_SIZE),
        Extent("LogicalSize", 1),
        3000,
        Some(0),
    ),
    s(
        "scrollLeft",
        FORM,
        Points(SCROLL_SIZE),
        Extent("ScrollPosition", 0),
        0,
        None,
    ),
    s(
        "scrollTop",
        FORM,
        Points(SCROLL_SIZE),
        Extent("ScrollPosition", 1),
        0,
        None,
    ),
    // The background picture (`form.picture`).
    s(
        "pictureAlignment",
        FORM,
        Enum(&[0, 1, 2, 3, 4]),
        Field("PictureAlignment"),
        2,
        None,
    ),
    s(
        "pictureSizeMode",
        FORM,
        Enum(&[0, 1, 3]),
        Field("PictureSizeMode"),
        0,
        None,
    ),
    s(
        "pictureTiling",
        FORM,
        Bool,
        Flag("PictureTiling", true),
        0,
        None,
    ),
];

/// fmPicturePosition 0 (LeftTop) to 12 (Center) as MS-OFORMS stores them.
const PICTURE_POSITIONS: [i64; 13] = [
    0x0002_0000, // LeftTop
    0x0005_0003, // LeftCenter
    0x0008_0006, // LeftBottom
    0x0000_0002, // RightTop
    0x0003_0005, // RightCenter
    0x0006_0008, // RightBottom
    0x0006_0000, // AboveLeft
    0x0007_0001, // AboveCenter
    0x0008_0002, // AboveRight
    0x0000_0006, // BelowLeft
    0x0001_0007, // BelowCenter
    0x0002_0008, // BelowRight
    0x0004_0004, // Center
];

/// `value` with `raw` put in the bits of a [`Store::FieldBits`] property.
pub(crate) fn put_field_bits(value: i64, p: &Scalar, raw: i64) -> i64 {
    match p.store {
        Store::FieldBits(_, at, width) => {
            let mask = ((1i64 << width) - 1) << at;
            (value & !mask) | ((raw << at) & mask)
        }
        _ => value,
    }
}

/// The value of a [`Store::FieldBits`] property in the field's `value`.
pub(crate) fn get_field_bits(value: i64, p: &Scalar) -> i64 {
    match p.store {
        Store::FieldBits(_, at, width) => (value >> at) & ((1 << width) - 1),
        _ => 0,
    }
}

/// A form property as its line of the designer block shows it: a VBABOOL ("0" / "-1"),
/// a number, or a quoted string.
pub(crate) fn frame_text(p: &Scalar, value: &Value, raw: i64) -> String {
    match p.kind {
        Kind::Bool => if raw != 0 { "-1" } else { "0" }.to_string(),
        Kind::Text => format!("\"{}\"", value.as_str().unwrap_or("").replace('"', "\"\"")),
        _ => raw.to_string(),
    }
}

/// The FormSpec value of a line of the designer block, if it can be read.
pub(crate) fn frame_value(p: &Scalar, text: &str) -> Option<Value> {
    match p.kind {
        Kind::Bool => match text.trim() {
            "0" => Some(Value::from(false)),
            "-1" | "1" => Some(Value::from(true)),
            _ => None,
        },
        Kind::Text => {
            let inner = text.trim().strip_prefix('"')?.strip_suffix('"')?;
            Some(Value::from(inner.replace("\"\"", "\"")))
        }
        _ => text.trim().parse::<i64>().ok().map(Value::from),
    }
}

/// The form's plain property `name`.
pub(crate) fn form_scalar(name: &str) -> Option<&'static Scalar> {
    FORM_SCALARS
        .iter()
        .find(|p| p.name.eq_ignore_ascii_case(name.trim()))
}

/// The record fields of a MultiPage's properties that its hidden TabStrip holds.
pub(crate) const TAB_STRIP_FIELDS: &[&str] = &[
    "TabStyle",
    "TabOrientation",
    "MultiRow",
    "TabFixedWidth",
    "TabFixedHeight",
];

impl Store {
    #[cfg(test)]
    fn is_text(self) -> bool {
        matches!(self, Store::Text(_) | Store::SiteText(_))
    }
}

impl Scalar {
    /// The record field the property is stored in, if it has one.
    pub(crate) fn field(&self) -> Option<&'static str> {
        match self.store {
            Store::Field(f) | Store::Himetric(f) | Store::Text(f) => Some(f),
            Store::Thumb => Some("ProportionalThumb"),
            _ => None,
        }
    }
}

/// The properties of `control_type`.
pub(crate) fn of_type(control_type: &str) -> impl Iterator<Item = &'static Scalar> + '_ {
    SCALARS
        .iter()
        .filter(move |p| p.types.iter().any(|t| t.eq_ignore_ascii_case(control_type)))
}

/// The property `name` of `control_type`.
pub(crate) fn find(control_type: &str, name: &str) -> Option<&'static Scalar> {
    of_type(control_type).find(|p| p.name.eq_ignore_ascii_case(name.trim()))
}

/// The class ID of a built-in control type.
pub(crate) fn class_of(control_type: &str) -> Option<u16> {
    Some(match control_type {
        "Page" => 7,
        "Image" => 12,
        "Frame" => 14,
        "SpinButton" => 16,
        "CommandButton" => 17,
        "TabStrip" => 18,
        "Label" => 21,
        "TextBox" => 23,
        "ListBox" => 24,
        "ComboBox" => 25,
        "CheckBox" => 26,
        "OptionButton" => 27,
        "ToggleButton" => 28,
        "ScrollBar" => 47,
        "MultiPage" => 57,
        _ => return None,
    })
}

/// `VariousPropertyBits` of a new control of the class, as Excel writes them.
pub(crate) fn fresh_bits(class: u16) -> i64 {
    match class {
        // Editable: a TextBox, and a ComboBox that is not a drop-down list.
        23 | 25 => 0x2c80_481b,
        21 => 0x0080_001b,
        24 | 26..=28 => 0x2c80_081b,
        _ => 0x1b,
    }
}

/// The site `BitFlags` of a new control of the class.
pub(crate) fn fresh_site_flags(class: u16) -> i64 {
    match class {
        21 => 0x32,
        7 | 14 | 57 => 0x40023,
        _ => 0x33,
    }
}

/// The bits of `VariousPropertyBits` that properties of `control_type` describe.
pub(crate) fn described_bits(control_type: &str) -> i64 {
    let mut mask = 0i64;
    for p in of_type(control_type) {
        match p.store {
            Store::Bit(bit, _) => mask |= 1 << bit,
            Store::Bits(at, width) => mask |= ((1 << width) - 1) << at,
            _ => {}
        }
    }
    mask
}

/// Whether the site string `name` (`ControlSource` ...) is a property of `control_type`.
pub(crate) fn describes_site_text(control_type: &str, name: &str) -> bool {
    of_type(control_type).any(|p| matches!(p.store, Store::SiteText(n) if n == name))
}

/// The site flags that properties of `control_type` describe.
pub(crate) fn described_site_flags(control_type: &str) -> i64 {
    of_type(control_type)
        .filter_map(|p| match p.store {
            Store::Site(bit) => Some(1i64 << bit),
            _ => None,
        })
        .fold(0, |a, b| a | b)
}

/// The stored value of a FormSpec value, or why it cannot be stored.
pub(crate) fn to_raw(p: &Scalar, value: &Value) -> Result<i64, String> {
    let name = p.name;
    let raw = match p.kind {
        Kind::Bool => i64::from(
            value
                .as_bool()
                .ok_or(format!("{name} は真偽値にしてください"))?,
        ),
        Kind::Enum(allowed) => {
            let n = value.as_i64().filter(|n| allowed.contains(n));
            n.ok_or(format!(
                "{name} は {} のいずれかにしてください",
                allowed
                    .iter()
                    .map(i64::to_string)
                    .collect::<Vec<_>>()
                    .join("、")
            ))?
        }
        Kind::Int(lo, hi) => value
            .as_i64()
            .filter(|n| (lo..=hi).contains(n))
            .ok_or(format!("{name} は {lo}〜{hi} の整数にしてください"))?,
        Kind::Points(max) => {
            let n = value
                .as_f64()
                .filter(|n| n.is_finite() && (0.0..=max).contains(n))
                .ok_or(format!(
                    "{name} は 0〜{max} の数値（ポイント）にしてください"
                ))?;
            (n * 2540.0 / 72.0).round() as i64
        }
        Kind::Char => {
            let text = value
                .as_str()
                .ok_or(format!("{name} は文字列にしてください"))?;
            let units: Vec<u16> = text.encode_utf16().collect();
            match units.as_slice() {
                [] => 0,
                [c] => i64::from(*c),
                _ => return Err(format!("{name} は 1 文字にしてください")),
            }
        }
        Kind::Color => {
            let color: FormSpecColor = serde_json::from_value(value.clone())
                .map_err(|_| format!("{name} は色（{{ kind, value }}）にしてください"))?;
            i64::from(color.ole()?)
        }
        Kind::Text => {
            value
                .as_str()
                .ok_or(format!("{name} は文字列にしてください"))?;
            0
        }
        Kind::Widths => {
            let text = value
                .as_str()
                .ok_or(format!("{name} は文字列にしてください"))?;
            parse_column_widths(text).map_err(|e| format!("{name}: {e}"))?;
            0
        }
    };
    Ok(match p.store {
        Store::Field("PicturePosition") => PICTURE_POSITIONS[raw as usize],
        // ComboBox Style is its DisplayStyle: combo (3) or drop-down list (7).
        Store::Field("DisplayStyle") => {
            if raw == 2 {
                7
            } else {
                3
            }
        }
        Store::Thumb => {
            if raw != 0 {
                -1
            } else {
                0
            }
        }
        _ => raw,
    })
}

/// The FormSpec value of a stored value.
pub(crate) fn from_raw(p: &Scalar, raw: i64) -> Value {
    match (p.kind, p.store) {
        (_, Store::Field("DisplayStyle")) => Value::from(if raw == 7 { 2 } else { 0 }),
        (_, Store::Field("PicturePosition")) => Value::from(
            PICTURE_POSITIONS
                .iter()
                .position(|v| *v == raw)
                .unwrap_or(7) as i64,
        ),
        (_, Store::Thumb) => Value::from(raw != 0),
        (Kind::Bool, _) => Value::from(raw != 0),
        (Kind::Points(_), _) => Value::from(crate::snapshot::points(raw as i32)),
        (Kind::Char, _) => Value::from(
            char::from_u32(raw as u32)
                .filter(|_| raw != 0)
                .map(String::from)
                .unwrap_or_default(),
        ),
        (Kind::Color, _) => {
            serde_json::to_value(FormSpecColor::from_ole(raw as u32)).expect("color")
        }
        _ => Value::from(raw),
    }
}

/// `ColumnWidths` (`"72 pt;;1 in"`) → the width of each column in HIMETRIC; -1 for a
/// column whose width the control works out (an empty entry or -1). A width is in points
/// unless it ends with `pt`, `in` or `cm`. An empty text gives no columns.
pub(crate) fn parse_column_widths(text: &str) -> Result<Vec<i32>, String> {
    if text.trim().is_empty() {
        return Ok(Vec::new());
    }
    let entries: Vec<&str> = text.split(';').collect();
    if entries.len() > i16::MAX as usize {
        return Err("列が多すぎます".into());
    }
    entries
        .into_iter()
        .map(|entry| {
            let entry = entry.trim().to_lowercase();
            if entry.is_empty() {
                return Ok(-1);
            }
            let (number, per_unit) = [("pt", 2540.0 / 72.0), ("in", 2540.0), ("cm", 1000.0)]
                .into_iter()
                .find_map(|(unit, scale)| entry.strip_suffix(unit).map(|n| (n.trim(), scale)))
                .unwrap_or((entry.as_str(), 2540.0 / 72.0));
            let value: f64 = number.parse().map_err(|_| {
                format!("幅 {entry:?} は数値（pt、in、cm の単位を付けられます）にしてください")
            })?;
            if value == -1.0 {
                return Ok(-1);
            }
            let himetric = (value * per_unit).round();
            if !value.is_finite() || value < 0.0 || himetric > (TAB_SIZE * 10.0 * 2540.0 / 72.0) {
                return Err(format!(
                    "幅 {entry:?} は 0〜{} pt にしてください",
                    TAB_SIZE * 10.0
                ));
            }
            Ok(himetric as i32)
        })
        .collect()
}

/// The widths of the columns (HIMETRIC, -1 worked out) as `ColumnWidths` shows them in points.
pub(crate) fn format_column_widths(widths: &[i32]) -> String {
    widths
        .iter()
        .map(|w| {
            if *w < 0 {
                return String::new();
            }
            let points = (*w as f64 * 72.0 / 2540.0 * 100.0).round() / 100.0;
            format!("{points} pt")
        })
        .collect::<Vec<_>>()
        .join(";")
}

/// The `MorphDataColumnInfo` records of `widths`: the width of a column is stored only
/// when it is not -1, the file-format default. Excel stores the last column first.
pub(crate) fn encode_column_infos(widths: &[i32]) -> Vec<u8> {
    let mut out = Vec::with_capacity(widths.len() * 12);
    for width in widths.iter().rev() {
        if *width < 0 {
            out.extend_from_slice(&[0, 2, 4, 0, 0, 0, 0, 0]);
        } else {
            out.extend_from_slice(&[0, 2, 8, 0, 1, 0, 0, 0]);
            out.extend_from_slice(&width.to_le_bytes());
        }
    }
    out
}

/// The widths in `count` `MorphDataColumnInfo` records that make up `raw` exactly, first
/// column first (the records run from the last column).
pub(crate) fn decode_column_infos(raw: &[u8], count: usize) -> Option<Vec<i32>> {
    let mut widths = Vec::with_capacity(count);
    let mut at = 0;
    for _ in 0..count {
        let header = raw.get(at..at + 8)?;
        if header[0] != 0 || header[1] != 2 {
            return None;
        }
        let cb = u16::from_le_bytes([header[2], header[3]]) as usize;
        let mask = u32::from_le_bytes([header[4], header[5], header[6], header[7]]);
        let width = match (mask, cb) {
            (0, 4) => -1,
            (1, 8) => i32::from_le_bytes(raw.get(at + 8..at + 12)?.try_into().ok()?),
            _ => return None,
        };
        widths.push(width);
        at += 4 + cb;
    }
    widths.reverse();
    (at == raw.len()).then_some(widths)
}

/// A bit (or bits) of `VariousPropertyBits` with the stored value of `p` put in.
pub(crate) fn put_bits(bits: i64, p: &Scalar, raw: i64) -> i64 {
    match p.store {
        Store::Bit(bit, inverted) => {
            let on = (raw != 0) != inverted;
            if on {
                bits | 1 << bit
            } else {
                bits & !(1 << bit)
            }
        }
        Store::Bits(at, width) => {
            let mask = ((1i64 << width) - 1) << at;
            (bits & !mask) | ((raw << at) & mask)
        }
        _ => bits,
    }
}

/// The stored value of `p` in `bits`.
pub(crate) fn get_bits(bits: i64, p: &Scalar) -> i64 {
    match p.store {
        Store::Bit(bit, inverted) => i64::from(((bits >> bit) & 1 != 0) != inverted),
        Store::Bits(at, width) => (bits >> at) & ((1 << width) - 1),
        _ => 0,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// A value of `p` other than the one a new control has.
    fn other_value(p: &Scalar, class: u16) -> Value {
        let fresh = match p.store {
            Store::Bit(..) | Store::Bits(..) => get_bits(fresh_bits(class), p),
            Store::Site(bit) => (fresh_site_flags(class) >> bit) & 1,
            Store::Flag(_, set) => i64::from(!set),
            _ => p.fresh.unwrap_or(p.omitted),
        };
        match p.kind {
            Kind::Bool => Value::from(fresh == 0),
            Kind::Enum(allowed) => {
                let current = from_raw(p, fresh);
                Value::from(*allowed.iter().find(|v| current != **v).unwrap())
            }
            Kind::Int(lo, _) => Value::from(if fresh > lo { fresh - 1 } else { fresh + 1 }),
            Kind::Points(_) => Value::from(36.0),
            Kind::Char => Value::from("*"),
            Kind::Color => serde_json::json!({ "kind": "rgb", "value": "#123456" }),
            Kind::Text => Value::from("group1"),
            Kind::Widths => Value::from("72;;36 pt"),
        }
    }

    /// The forms of `tools/excel-check` (checked against Excel on Windows) set every plain
    /// property of every control type, and of the form, to a value a new one does not have.
    #[test]
    fn the_excel_check_forms_cover_every_property() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../tools/excel-check");
        let mut specs = Vec::new();
        for entry in std::fs::read_dir(root.join("forms")).unwrap() {
            let path = entry.unwrap().path();
            if !path.to_string_lossy().ends_with(".form.json") {
                continue;
            }
            let parsed = crate::parse(&std::fs::read_to_string(&path).unwrap());
            assert!(
                !parsed.has_errors(),
                "{}: {:?}",
                path.display(),
                parsed.issues
            );
            let spec = parsed.spec.unwrap();
            crate::compile_in(&spec, 932, &root)
                .unwrap_or_else(|e| panic!("{}: {e}", path.display()));
            specs.push(spec);
        }
        let differs = |p: &Scalar, class: Option<u16>, value: &Value| -> bool {
            let Ok(raw) = to_raw(p, value) else {
                return false;
            };
            if matches!(p.kind, Kind::Text | Kind::Widths) {
                return value.as_str().is_some_and(|s| !s.is_empty());
            }
            match (p.store, class) {
                (Store::Bit(..) | Store::Bits(..), Some(class)) => {
                    raw != get_bits(fresh_bits(class), p)
                }
                (Store::Site(bit), Some(class)) => raw != (fresh_site_flags(class) >> bit) & 1,
                (Store::Flag(_, set), _) => raw != i64::from(!set),
                (Store::FieldBits(..), _) => {
                    get_field_bits(put_field_bits(p.omitted, p, raw), p)
                        != get_field_bits(p.fresh.unwrap_or(p.omitted), p)
                }
                _ => raw != p.fresh.unwrap_or(p.omitted),
            }
        };
        let mut missing = Vec::new();
        for p in SCALARS {
            for control_type in p.types {
                let class = class_of(control_type);
                let covered = specs.iter().flat_map(|s| &s.controls).any(|c| {
                    c.control_type == *control_type
                        && c.scalar(p.name).is_some_and(|v| differs(p, class, v))
                });
                if !covered {
                    missing.push(format!("{control_type}.{}", p.name));
                }
            }
        }
        for p in FORM_SCALARS {
            if !specs
                .iter()
                .any(|s| s.form.scalar(p.name).is_some_and(|v| differs(p, None, v)))
            {
                missing.push(format!("form.{}", p.name));
            }
        }
        assert!(
            missing.is_empty(),
            "not covered by tools/excel-check/forms: {missing:?}"
        );
    }

    #[test]
    fn every_property_goes_round_and_a_new_control_shows_none() {
        for p in SCALARS {
            for control_type in p.types {
                let class = class_of(control_type).unwrap();
                let value = other_value(p, class);
                // A Page goes into a MultiPage.
                let (owner, parent) = if *control_type == "Page" {
                    (
                        r#"{"id":"m","name":"m","type":"MultiPage"},"#,
                        r#","parentId":"m""#,
                    )
                } else {
                    ("", "")
                };
                let text = format!(
                    r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"F"}},"controls":[{owner}{{"id":"c","name":"c","type":"{control_type}"{parent},"{}":{value}}}]}}"#,
                    p.name
                );
                let parsed = crate::parse(&text);
                assert!(
                    !parsed.has_errors(),
                    "{control_type}.{}: {:?}",
                    p.name,
                    parsed.issues
                );
                let designer = crate::compile(parsed.spec.as_ref().unwrap(), 932)
                    .unwrap_or_else(|e| panic!("{control_type}.{}: {e}", p.name));
                let spec = crate::snapshot("F", &designer.streams, &designer.storages, 932)
                    .unwrap()
                    .spec;
                let control = spec.controls.iter().find(|c| c.name == "c").unwrap();
                let back = control.scalar(p.name).unwrap_or_else(|| {
                    panic!(
                        "{control_type}.{} = {value} was not read back: {:?}",
                        p.name, control.scalars
                    )
                });
                if p.store.is_text() {
                    assert_eq!(back, &value);
                } else if matches!(p.kind, Kind::Widths) {
                    let widths = |v: &Value| parse_column_widths(v.as_str().unwrap()).unwrap();
                    assert_eq!(widths(back), widths(&value), "{control_type}.{}", p.name);
                    assert_eq!(widths(back), [2540, -1, 1270]);
                } else {
                    assert_eq!(
                        to_raw(p, back),
                        to_raw(p, &value),
                        "{control_type}.{}",
                        p.name
                    );
                }
                // Only what was written shows; nothing is left unsupported by it.
                let unsupported = control.unsupported.clone().unwrap_or_default();
                assert!(
                    !unsupported
                        .iter()
                        .any(|u| ["variousPropertyBits", "siteFlags"].contains(&u.as_str())),
                    "{control_type}.{}: {unsupported:?}",
                    p.name
                );
            }
        }
        // A new control of each type: no plain property differs from Excel's.
        for t in crate::contract::CONTROL_TYPES
            .iter()
            .filter(|t| t.name != "Page")
        {
            let text = format!(
                r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"F"}},"controls":[{{"id":"c","name":"c","type":"{}"}}]}}"#,
                t.name
            );
            let parsed = crate::parse(&text);
            let designer = crate::compile(parsed.spec.as_ref().unwrap(), 932).unwrap();
            let spec = crate::snapshot("F", &designer.streams, &designer.storages, 932)
                .unwrap()
                .spec;
            let control = &spec.controls[0];
            let shown: Vec<_> = control.scalars.keys().collect();
            assert!(shown.is_empty(), "{}: {shown:?}", t.name);
            // Nor its font, colors or alignment: generation writes Excel's.
            assert_eq!(control.font, None, "{}", t.name);
            assert_eq!(control.back_color, None, "{}", t.name);
            assert_eq!(control.fore_color, None, "{}", t.name);
            assert_eq!(control.text_align, None, "{}", t.name);
        }
    }

    #[test]
    fn the_column_widths_are_stored_from_the_last_column() {
        let widths = parse_column_widths("30 pt;;60 pt").unwrap();
        let raw = encode_column_infos(&widths);
        // Excel stores "30 pt;;60 pt" as 60 pt, worked out, 30 pt.
        assert_eq!(&raw[..8], &[0, 2, 8, 0, 1, 0, 0, 0]);
        assert_eq!(
            i32::from_le_bytes(raw[8..12].try_into().unwrap()),
            widths[2]
        );
        assert_eq!(decode_column_infos(&raw, 3), Some(widths));
    }

    #[test]
    fn a_form_without_a_start_up_position_is_placed_by_hand() {
        let text = r#"{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{"name":"F"},"controls":[]}"#;
        let mut designer = crate::compile(crate::parse(text).spec.as_ref().unwrap(), 932).unwrap();
        let spec = crate::snapshot("F", &designer.streams, &designer.storages, 932)
            .unwrap()
            .spec;
        assert_eq!(spec.form.scalar("startUpPosition"), None);
        // The VBA editor leaves the line out for 0 (Manual).
        let frame = designer.streams.get_mut("\u{3}VBFrame").unwrap();
        let text = String::from_utf8(frame.clone()).unwrap();
        *frame = text
            .lines()
            .filter(|l| !l.contains("StartUpPosition"))
            .collect::<Vec<_>>()
            .join("\r\n")
            .into_bytes();
        let spec = crate::snapshot("F", &designer.streams, &designer.storages, 932)
            .unwrap()
            .spec;
        assert_eq!(spec.form.scalar("startUpPosition"), Some(&Value::from(0)));
    }
}
