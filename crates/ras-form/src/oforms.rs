//! Generation of the MS-OFORMS designer storage of a user form.
//!
//! Port of the generation path of xlflow's `internal/vba/userforms/oforms`
//! (`generate.go`, `mutate.go`, `specs.go`, `font.go`, `multipage.go`,
//! `tabstrip.go`, `caption.go`). Copyright (c) 2026 harumiWeb, MIT License;
//! see `NOTICE.md`.
//!
//! The model below is built in persistence units (HIMETRIC, OLE colors) and
//! encoded only at the very end, so the typed appearance edits that xlflow
//! applies to a freshly generated form operate on the same records here.

use std::collections::{BTreeMap, HashMap};

/// Persistence-level failure (`ErrInvalidEdit` / `ErrUnsupportedEdit`).
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) enum PersistError {
    Invalid(String),
    Unsupported(String),
}

impl std::fmt::Display for PersistError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            PersistError::Invalid(m) => write!(f, "不正な値です: {m}"),
            PersistError::Unsupported(m) => write!(f, "対応していない操作です: {m}"),
        }
    }
}

type Result<T> = std::result::Result<T, PersistError>;

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(PersistError::Invalid(message.into()))
}

fn unsupported<T>(message: impl Into<String>) -> Result<T> {
    Err(PersistError::Unsupported(message.into()))
}

const MAX_DESIGNER_STREAM_SIZE: usize = 64 << 20;
/// Nesting limit including the root form level.
pub(crate) const MAX_NESTING_DEPTH: usize = 64;
const MAX_SITES_PER_FORM: i64 = 65_535;
const FORM_MASK_FONT: u64 = 1 << 20;

/// The font effects of the text of a disabled control as Excel stores them: `CFE_DISABLED`
/// and `CFE_AUTOCOLOR` (the RichEdit `CHARFORMAT` flags that TextProps uses).
pub(crate) const DISABLED_TEXT_EFFECTS: i64 = 0x4000_2000;
/// The classes Excel stores them for: Label, TextBox, CommandButton, ComboBox, CheckBox,
/// OptionButton and ToggleButton (not a ListBox or TabStrip).
pub(crate) const DISABLED_TEXT_CLASSES: &[u16] = &[17, 21, 23, 25, 26, 27, 28];

// ---------------------------------------------------------------------------
// Record layouts (specs.go)

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum FieldKind {
    Unsigned,
    Signed,
    StringLength,
    Marker,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum ExtraKind {
    Size,
    Position,
    Str,
    Array,
}

#[derive(Debug)]
pub(crate) struct DataField {
    pub(crate) bit: u8,
    pub(crate) name: &'static str,
    pub(crate) size: usize,
    pub(crate) kind: FieldKind,
}

#[derive(Debug)]
pub(crate) struct ExtraField {
    pub(crate) bit: u8,
    pub(crate) name: &'static str,
    pub(crate) kind: ExtraKind,
}

#[derive(Debug)]
pub(crate) struct RecordSpec {
    /// Name of the record in messages (`MorphData`, `Form`, ...).
    pub(crate) type_name: &'static str,
    pub(crate) major: u8,
    pub(crate) mask64: bool,
    pub(crate) stop_after_extra: bool,
    pub(crate) text_props: bool,
    /// Bytes after the last known part are kept instead of being an error.
    pub(crate) raw_tail: bool,
    pub(crate) data: &'static [DataField],
    pub(crate) extra: &'static [ExtraField],
    /// Pictures (`GuidAndPicture`) that follow the record, in stream order.
    pub(crate) stream: &'static [(u8, &'static str)],
    /// Mask bits that carry a value by themselves: (bit, name, value).
    pub(crate) flags: &'static [(u8, &'static str, i64)],
}

const fn d(bit: u8, name: &'static str, size: usize, kind: FieldKind) -> DataField {
    DataField {
        bit,
        name,
        size,
        kind,
    }
}

const fn e(bit: u8, name: &'static str, kind: ExtraKind) -> ExtraField {
    ExtraField { bit, name, kind }
}

use ExtraKind as X;
use FieldKind::{Marker as M, Signed as S, StringLength as L, Unsigned as U};

pub(crate) static TEXT_PROPS: RecordSpec = RecordSpec {
    type_name: "TextProps",
    major: 2,
    mask64: false,
    stop_after_extra: true,
    text_props: false,
    raw_tail: false,
    stream: &[],
    flags: &[],
    data: &[
        d(0, "FontName", 4, L),
        d(1, "FontEffects", 4, U),
        d(2, "FontHeight", 4, U),
        d(4, "FontCharSet", 1, U),
        d(5, "FontPitchAndFamily", 1, U),
        d(6, "ParagraphAlign", 1, U),
        d(7, "FontWeight", 2, U),
    ],
    extra: &[e(0, "FontName", X::Str)],
};

pub(crate) static MORPH_DATA: RecordSpec = RecordSpec {
    type_name: "MorphData",
    major: 2,
    mask64: true,
    stop_after_extra: false,
    text_props: true,
    raw_tail: true,
    stream: &[(27, "MouseIcon"), (28, "Picture")],
    flags: &[],
    data: &[
        d(0, "VariousPropertyBits", 4, U),
        d(1, "BackColor", 4, U),
        d(2, "ForeColor", 4, U),
        d(3, "MaxLength", 4, U),
        d(4, "BorderStyle", 1, U),
        d(5, "ScrollBars", 1, U),
        d(6, "DisplayStyle", 1, U),
        d(7, "MousePointer", 1, U),
        d(9, "PasswordChar", 2, U),
        d(10, "ListWidth", 4, U),
        d(11, "BoundColumn", 2, U),
        d(12, "TextColumn", 2, S),
        d(13, "ColumnCount", 2, S),
        d(14, "ListRows", 2, U),
        d(15, "cColumnInfo", 2, U),
        d(16, "MatchEntry", 1, U),
        d(17, "ListStyle", 1, U),
        d(18, "ShowDropButtonWhen", 1, U),
        d(20, "DropButtonStyle", 1, U),
        d(21, "MultiSelect", 1, U),
        d(22, "Value", 4, L),
        d(23, "Caption", 4, L),
        d(24, "PicturePosition", 4, U),
        d(25, "BorderColor", 4, U),
        d(26, "SpecialEffect", 4, U),
        d(27, "MouseIcon", 2, M),
        d(28, "Picture", 2, M),
        d(29, "Accelerator", 2, U),
        d(32, "GroupName", 4, L),
    ],
    extra: &[
        e(8, "Size", X::Size),
        e(22, "Value", X::Str),
        e(23, "Caption", X::Str),
        e(32, "GroupName", X::Str),
    ],
};

pub(crate) static COMMAND_BUTTON: RecordSpec = RecordSpec {
    type_name: "CommandButton",
    major: 2,
    mask64: false,
    stop_after_extra: false,
    text_props: true,
    raw_tail: false,
    stream: &[(7, "Picture"), (10, "MouseIcon")],
    flags: &[(9, "TakeFocusOnClick", 0)],
    data: &[
        d(0, "ForeColor", 4, U),
        d(1, "BackColor", 4, U),
        d(2, "VariousPropertyBits", 4, U),
        d(3, "Caption", 4, L),
        d(4, "PicturePosition", 4, U),
        d(6, "MousePointer", 1, U),
        d(7, "Picture", 2, M),
        d(8, "Accelerator", 2, U),
        d(10, "MouseIcon", 2, M),
    ],
    extra: &[e(3, "Caption", X::Str), e(5, "Size", X::Size)],
};

pub(crate) static LABEL: RecordSpec = RecordSpec {
    type_name: "Label",
    major: 2,
    mask64: false,
    stop_after_extra: false,
    text_props: true,
    raw_tail: false,
    stream: &[(10, "Picture"), (12, "MouseIcon")],
    flags: &[],
    data: &[
        d(0, "ForeColor", 4, U),
        d(1, "BackColor", 4, U),
        d(2, "VariousPropertyBits", 4, U),
        d(3, "Caption", 4, L),
        d(4, "PicturePosition", 4, U),
        d(6, "MousePointer", 1, U),
        d(7, "BorderColor", 4, U),
        d(8, "BorderStyle", 2, U),
        d(9, "SpecialEffect", 2, U),
        d(10, "Picture", 2, M),
        d(11, "Accelerator", 2, U),
        d(12, "MouseIcon", 2, M),
    ],
    extra: &[e(3, "Caption", X::Str), e(5, "Size", X::Size)],
};

pub(crate) static IMAGE: RecordSpec = RecordSpec {
    type_name: "Image",
    major: 2,
    mask64: false,
    stop_after_extra: false,
    text_props: false,
    raw_tail: false,
    stream: &[(10, "Picture"), (14, "MouseIcon")],
    flags: &[],
    data: &[
        d(3, "BorderColor", 4, U),
        d(4, "BackColor", 4, U),
        d(5, "BorderStyle", 1, U),
        d(6, "MousePointer", 1, U),
        d(7, "PictureSizeMode", 1, U),
        d(8, "SpecialEffect", 1, U),
        d(10, "Picture", 2, M),
        d(11, "PictureAlignment", 1, U),
        d(13, "VariousPropertyBits", 4, U),
        d(14, "MouseIcon", 2, M),
    ],
    extra: &[e(9, "Size", X::Size)],
};

pub(crate) static SPIN_BUTTON: RecordSpec = RecordSpec {
    type_name: "SpinButton",
    major: 2,
    mask64: false,
    stop_after_extra: false,
    text_props: false,
    raw_tail: false,
    stream: &[(13, "MouseIcon")],
    flags: &[],
    data: &[
        d(0, "ForeColor", 4, U),
        d(1, "BackColor", 4, U),
        d(2, "VariousPropertyBits", 4, U),
        d(5, "Min", 4, S),
        d(6, "Max", 4, S),
        d(7, "Position", 4, S),
        d(8, "PrevEnabled", 4, S),
        d(9, "NextEnabled", 4, S),
        d(10, "SmallChange", 4, S),
        d(11, "Orientation", 4, S),
        d(12, "Delay", 4, U),
        d(13, "MouseIcon", 2, M),
        d(14, "MousePointer", 1, U),
    ],
    extra: &[e(3, "Size", X::Size)],
};

pub(crate) static SCROLL_BAR: RecordSpec = RecordSpec {
    type_name: "ScrollBar",
    major: 2,
    mask64: false,
    stop_after_extra: false,
    text_props: false,
    raw_tail: false,
    stream: &[(16, "MouseIcon")],
    flags: &[],
    data: &[
        d(0, "ForeColor", 4, U),
        d(1, "BackColor", 4, U),
        d(2, "VariousPropertyBits", 4, U),
        d(4, "MousePointer", 1, U),
        d(5, "Min", 4, S),
        d(6, "Max", 4, S),
        d(7, "Position", 4, S),
        d(9, "PrevEnabled", 4, S),
        d(10, "NextEnabled", 4, S),
        d(11, "SmallChange", 4, S),
        d(12, "LargeChange", 4, S),
        d(13, "Orientation", 4, S),
        d(14, "ProportionalThumb", 2, S),
        d(15, "Delay", 4, U),
        d(16, "MouseIcon", 2, M),
    ],
    extra: &[e(3, "Size", X::Size)],
};

pub(crate) static TAB_STRIP: RecordSpec = RecordSpec {
    type_name: "TabStrip",
    major: 2,
    mask64: false,
    stop_after_extra: false,
    text_props: true,
    raw_tail: true,
    stream: &[(24, "MouseIcon")],
    flags: &[(10, "MultiRow", 1), (13, "Tooltips", 0), (19, "NewVersion", 1)],
    data: &[
        d(0, "ListIndex", 4, S),
        d(1, "BackColor", 4, U),
        d(2, "ForeColor", 4, U),
        d(5, "ItemsSize", 4, U),
        d(6, "MousePointer", 1, U),
        d(8, "TabOrientation", 4, U),
        d(9, "TabStyle", 4, U),
        d(11, "TabFixedWidth", 4, U),
        d(12, "TabFixedHeight", 4, U),
        d(15, "TipStringsSize", 4, U),
        d(17, "NamesSize", 4, U),
        d(18, "VariousPropertyBits", 4, U),
        d(20, "TabsAllocated", 4, U),
        d(21, "TagsSize", 4, U),
        d(22, "TabData", 4, U),
        d(23, "AcceleratorsSize", 4, U),
        d(24, "MouseIcon", 2, M),
    ],
    extra: &[
        e(4, "Size", X::Size),
        e(5, "Items", X::Array),
        e(15, "TipStrings", X::Array),
        e(17, "TabNames", X::Array),
        e(21, "Tags", X::Array),
        e(23, "Accelerators", X::Array),
    ],
};

pub(crate) static FORM: RecordSpec = RecordSpec {
    type_name: "Form",
    major: 4,
    mask64: false,
    stop_after_extra: true,
    text_props: false,
    raw_tail: false,
    // Written by the level, around its font (FormStreamData: MouseIcon, Font, Picture).
    stream: &[],
    // [MS-OFORMS] stores no value for PictureTiling: the bit stands for True.
    flags: &[(24, "PictureTiling", 1)],
    data: &[
        d(1, "BackColor", 4, U),
        d(2, "ForeColor", 4, U),
        d(3, "NextAvailableID", 4, U),
        d(6, "BooleanProperties", 4, U),
        d(7, "BorderStyle", 1, U),
        d(8, "MousePointer", 1, U),
        d(9, "ScrollBars", 1, U),
        d(13, "GroupCnt", 4, S),
        d(15, "MouseIcon", 2, M),
        d(16, "Cycle", 1, U),
        d(17, "SpecialEffect", 1, U),
        d(18, "BorderColor", 4, U),
        d(19, "Caption", 4, L),
        d(20, "Font", 2, M),
        d(21, "Picture", 2, M),
        d(22, "Zoom", 4, U),
        d(23, "PictureAlignment", 1, U),
        d(25, "PictureSizeMode", 1, U),
        d(26, "ShapeCookie", 4, U),
        d(27, "DrawBuffer", 4, U),
    ],
    extra: &[
        e(10, "DisplayedSize", X::Size),
        e(11, "LogicalSize", X::Size),
        e(12, "ScrollPosition", X::Position),
        e(19, "Caption", X::Str),
    ],
};

pub(crate) static MULTI_PAGE_PROPERTIES: RecordSpec = RecordSpec {
    type_name: "MultiPageProperties",
    major: 2,
    mask64: false,
    stop_after_extra: true,
    text_props: false,
    raw_tail: false,
    stream: &[],
    flags: &[(3, "Flags", 0)],
    data: &[d(1, "PageCount", 4, S), d(2, "ID", 4, S)],
    extra: &[],
};

/// One `PageProperties` record of a MultiPage `x` stream.
pub(crate) static PAGE_PROPERTIES: RecordSpec = RecordSpec {
    type_name: "PageProperties",
    major: 2,
    mask64: false,
    stop_after_extra: true,
    text_props: false,
    raw_tail: false,
    stream: &[],
    flags: &[],
    data: &[d(1, "TransitionEffect", 4, U), d(2, "TransitionPeriod", 4, U)],
    extra: &[],
};

/// Name of the data field that holds the byte size of an array.
pub(crate) fn array_size_field(name: &str) -> &'static str {
    match name {
        "Items" => "ItemsSize",
        "TipStrings" => "TipStringsSize",
        "TabNames" => "NamesSize",
        "Tags" => "TagsSize",
        _ => "AcceleratorsSize",
    }
}

pub(crate) fn spec_for_class(class: u16) -> Option<&'static RecordSpec> {
    Some(match class {
        12 => &IMAGE,
        15 | 23..=28 => &MORPH_DATA,
        16 => &SPIN_BUTTON,
        17 => &COMMAND_BUTTON,
        18 => &TAB_STRIP,
        21 => &LABEL,
        47 => &SCROLL_BAR,
        _ => return None,
    })
}

pub(crate) fn is_container_class(class: u16) -> bool {
    matches!(class, 7 | 14 | 57)
}

/// The MS-OFORMS omission default of `VariousPropertyBits`.
pub(crate) fn default_various_property_bits(class: u16) -> Option<i64> {
    match class {
        17 | 18 => Some(0x1b),
        21 => Some(0x80001b),
        23..=28 => Some(0x2c80081b),
        _ => None,
    }
}

// ---------------------------------------------------------------------------
// Strings and records

/// A persisted string: compressed (one byte per UTF-16 unit) or UTF-16LE.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
struct Stored {
    raw: Vec<u8>,
    compressed: bool,
}

impl Stored {
    fn packed(&self) -> i64 {
        let mut n = self.raw.len() as u32;
        if self.compressed {
            n |= 0x8000_0000;
        }
        n as i64
    }
}

fn edited_string(text: &str) -> Result<Stored> {
    if text.len() > 3 * 65535 {
        return invalid("文字列が長すぎます");
    }
    if text.is_empty() {
        return Ok(Stored::default());
    }
    let units: Vec<u16> = text.encode_utf16().collect();
    if units.iter().all(|u| *u <= 0xff) {
        return Ok(Stored {
            raw: units.iter().map(|u| *u as u8).collect(),
            compressed: true,
        });
    }
    Ok(Stored {
        raw: units.iter().flat_map(|u| u.to_le_bytes()).collect(),
        compressed: false,
    })
}

/// A property value in persistence terms.
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum PValue {
    Str(String),
    Int(i64),
    Bool(bool),
    /// An encoded `StdPicture` (see [`crate::picture::encode`]).
    Picture(Vec<u8>),
    /// A size or position (HIMETRIC): `LogicalSize`, `ScrollPosition`.
    Pair(i32, i32),
}

#[derive(Clone)]
pub(crate) struct Record {
    spec: &'static RecordSpec,
    mask: u64,
    values: HashMap<&'static str, i64>,
    strings: HashMap<&'static str, Stored>,
    sizes: HashMap<&'static str, (i32, i32)>,
    arrays: HashMap<&'static str, Vec<u8>>,
    pictures: HashMap<&'static str, Vec<u8>>,
    text_props: Option<Box<Record>>,
    tail: Vec<u8>,
}

impl Record {
    fn new(spec: &'static RecordSpec) -> Record {
        Record {
            spec,
            mask: 0,
            values: HashMap::new(),
            strings: HashMap::new(),
            sizes: HashMap::new(),
            arrays: HashMap::new(),
            pictures: HashMap::new(),
            text_props: None,
            tail: Vec::new(),
        }
    }

    fn set(&mut self, name: &str, value: &PValue) -> Result<()> {
        // A property-mask bit without data (TakeFocusOnClick, MultiRow): the
        // bit is set when the property has the value the flag stands for.
        if let Some((bit, flag, set)) = self.spec.flags.iter().find(|(_, n, _)| *n == name) {
            let PValue::Bool(on) = value else {
                return invalid(format!("{name} は真偽値が必要です"));
            };
            if i64::from(*on) == *set {
                self.mask |= 1u64 << bit;
                self.values.insert(flag, *set);
            } else {
                self.mask &= !(1u64 << bit);
                self.values.remove(flag);
            }
            return Ok(());
        }
        if let PValue::Pair(width, height) = value {
            let Some(extra) = self.spec.extra.iter().find(|x| x.name == name) else {
                return unsupported(format!("プロパティ {name}"));
            };
            self.mask |= 1u64 << extra.bit;
            self.sizes.insert(extra.name, (*width, *height));
            return Ok(());
        }
        let Some(field) = self.spec.data.iter().find(|f| f.name == name) else {
            return unsupported(format!("プロパティ {name}"));
        };
        match field.kind {
            FieldKind::Marker => {
                // A picture or mouse icon (the font of a form is written by its level).
                if name != "Picture" && name != "MouseIcon" {
                    return unsupported(format!("リソース {name}"));
                }
                let PValue::Picture(raw) = value else {
                    return invalid(format!("{name} は StdPicture のバイト列が必要です"));
                };
                crate::picture::decode(raw)
                    .map_err(|e| PersistError::Invalid(format!("{name}: {e}")))?;
                self.pictures.insert(field.name, raw.clone());
                self.values.insert(field.name, 0xffff);
            }
            FieldKind::StringLength => {
                let PValue::Str(text) = value else {
                    return invalid(format!("{name} は文字列が必要です"));
                };
                let stored = edited_string(text)?;
                self.values.insert(field.name, stored.packed());
                self.strings.insert(field.name, stored);
            }
            FieldKind::Unsigned | FieldKind::Signed => {
                let PValue::Int(n) = value else {
                    return invalid(format!("{name} は整数が必要です"));
                };
                let n = *n;
                let bits = field.size as u32 * 8;
                let ok = if field.kind == FieldKind::Signed {
                    n >= -(1i64 << (bits - 1)) && n < (1i64 << (bits - 1))
                } else {
                    n >= 0 && (n as u64) < (1u64 << bits)
                };
                if !ok {
                    return invalid(format!("{name} の値 {n} が範囲外です"));
                }
                self.values.insert(field.name, n);
            }
        }
        self.mask |= 1u64 << field.bit;
        Ok(())
    }

    fn set_int(&mut self, name: &str, value: i64) -> Result<()> {
        self.set(name, &PValue::Int(value))
    }

    fn set_size(&mut self, name: &'static str, size: (i32, i32)) {
        if let Some(extra) = self.spec.extra.iter().find(|x| x.name == name) {
            self.mask |= 1u64 << extra.bit;
            self.sizes.insert(extra.name, size);
        }
    }

    fn encode(&self) -> Result<Vec<u8>> {
        let spec = self.spec;
        let mut b = vec![0u8, spec.major, 0, 0];
        b.extend_from_slice(&(self.mask as u32).to_le_bytes());
        if spec.mask64 {
            b.extend_from_slice(&((self.mask >> 32) as u32).to_le_bytes());
        }
        for field in spec.data {
            if self.mask & (1u64 << field.bit) == 0 {
                continue;
            }
            pad(&mut b, field.size);
            let value = self.values.get(field.name).copied().unwrap_or(0);
            match field.size {
                1 => b.push(value as u8),
                2 => b.extend_from_slice(&(value as u16).to_le_bytes()),
                _ => b.extend_from_slice(&(value as u32).to_le_bytes()),
            }
        }
        pad(&mut b, 4);
        for extra in spec.extra {
            if self.mask & (1u64 << extra.bit) == 0 {
                continue;
            }
            match extra.kind {
                ExtraKind::Size | ExtraKind::Position => {
                    let (w, h) = self.sizes.get(extra.name).copied().unwrap_or((0, 0));
                    b.extend_from_slice(&w.to_le_bytes());
                    b.extend_from_slice(&h.to_le_bytes());
                }
                ExtraKind::Str => {
                    if let Some(s) = self.strings.get(extra.name) {
                        b.extend_from_slice(&s.raw);
                    }
                    pad(&mut b, 4);
                }
                ExtraKind::Array => {
                    if let Some(a) = self.arrays.get(extra.name) {
                        b.extend_from_slice(a);
                    }
                }
            }
        }
        finish_header(&mut b)?;
        if spec.stop_after_extra {
            return Ok(b);
        }
        for (bit, name) in spec.stream {
            if self.mask & (1u64 << bit) != 0 {
                if let Some(picture) = self.pictures.get(name) {
                    b.extend_from_slice(picture);
                }
            }
        }
        if spec.text_props {
            let Some(text_props) = &self.text_props else {
                return invalid("TextProps がありません");
            };
            b.extend_from_slice(&text_props.encode()?);
        }
        b.extend_from_slice(&self.tail);
        Ok(b)
    }
}

fn pad(b: &mut Vec<u8>, size: usize) {
    let n = (size - b.len() % size) % size;
    b.resize(b.len() + n, 0);
}

fn finish_header(b: &mut [u8]) -> Result<()> {
    let cb = b.len() - 4;
    if cb > u16::MAX as usize {
        return invalid("レコードが長すぎます");
    }
    b[2..4].copy_from_slice(&(cb as u16).to_le_bytes());
    Ok(())
}

// ---------------------------------------------------------------------------
// Sites, levels, form

struct SiteField {
    bit: u8,
    name: &'static str,
    size: usize,
}

const SITE_FIELDS: &[SiteField] = &[
    SiteField {
        bit: 0,
        name: "NameData",
        size: 4,
    },
    SiteField {
        bit: 1,
        name: "TagData",
        size: 4,
    },
    SiteField {
        bit: 2,
        name: "ID",
        size: 4,
    },
    SiteField {
        bit: 3,
        name: "HelpContextID",
        size: 4,
    },
    SiteField {
        bit: 4,
        name: "BitFlags",
        size: 4,
    },
    SiteField {
        bit: 5,
        name: "ObjectStreamSize",
        size: 4,
    },
    SiteField {
        bit: 6,
        name: "TabIndex",
        size: 2,
    },
    SiteField {
        bit: 7,
        name: "ClsidCacheIndex",
        size: 2,
    },
    SiteField {
        bit: 9,
        name: "GroupID",
        size: 2,
    },
    SiteField {
        bit: 11,
        name: "ControlTipTextData",
        size: 4,
    },
    SiteField {
        bit: 13,
        name: "ControlSourceData",
        size: 4,
    },
    SiteField {
        bit: 14,
        name: "RowSourceData",
        size: 4,
    },
];

struct Site {
    mask: u32,
    values: HashMap<&'static str, i64>,
    name: Stored,
    tag: Option<Stored>,
    tip: Option<Stored>,
    /// `ControlSource` and `RowSource`, which follow `ControlTipText`.
    sources: [Option<Stored>; 2],
    position: (i32, i32),
}

impl Site {
    fn encode(&self) -> Result<Vec<u8>> {
        let mut b = vec![0u8, 0, 0, 0];
        b.extend_from_slice(&self.mask.to_le_bytes());
        for field in SITE_FIELDS {
            if self.mask & (1u32 << field.bit) == 0 {
                continue;
            }
            pad(&mut b, field.size);
            let value = self.values.get(field.name).copied().unwrap_or(0);
            match field.size {
                2 => b.extend_from_slice(&(value as u16).to_le_bytes()),
                _ => b.extend_from_slice(&(value as u32).to_le_bytes()),
            }
        }
        pad(&mut b, 4);
        b.extend_from_slice(&self.name.raw);
        pad(&mut b, 4);
        if let Some(tag) = &self.tag {
            b.extend_from_slice(&tag.raw);
            pad(&mut b, 4);
        }
        if self.mask & (1 << 8) != 0 {
            b.extend_from_slice(&self.position.0.to_le_bytes());
            b.extend_from_slice(&self.position.1.to_le_bytes());
        }
        for text in [&self.tip, &self.sources[0], &self.sources[1]].into_iter().flatten() {
            b.extend_from_slice(&text.raw);
            pad(&mut b, 4);
        }
        finish_header(&mut b)?;
        Ok(b)
    }
}

struct Control {
    name: String,
    class: u16,
    id: i64,
    site: Site,
    /// Property record of a non-container control (stored in the `o` stream).
    record: Option<Record>,
    /// Index of the child level of a container control.
    child: Option<usize>,
}

struct MultiPageState {
    hidden_id: i64,
    page_ids: Vec<i64>,
    disabled: bool,
}

struct Level {
    /// Storage path relative to the form storage (`""` for the form itself).
    path: String,
    record: Record,
    font_raw: Vec<u8>,
    class_table_raw: Vec<u8>,
    controls: Vec<Control>,
    depths_raw: Vec<u8>,
    trailing_raw: Vec<u8>,
    multipage: Option<MultiPageState>,
    compobj: Vec<u8>,
    vbframe: Option<Vec<u8>>,
    clsid: [u8; 16],
}

/// A generated designer storage, ready to be stored in a `vbaProject.bin`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Designer {
    /// Name of the form (and of its storage below the root of the project).
    pub name: String,
    /// Streams, keyed by their path relative to the form storage
    /// (`"f"`, `"o"`, `"\u{1}CompObj"`, `"\u{3}VBFrame"`, `"i01/f"`, ...).
    pub streams: BTreeMap<String, Vec<u8>>,
    /// Storages with their CLSID (wire byte order), keyed by their path
    /// relative to the form storage; `""` is the form storage itself.
    pub storages: BTreeMap<String, [u8; 16]>,
}

/// Generation input in persistence units (oforms `Definition`).
pub(crate) struct Definition {
    pub name: String,
    pub caption: String,
    pub size: (i32, i32),
    pub controls: Vec<ControlDefinition>,
    /// Fields of the form's own record (`Zoom`, `LogicalSize` ...).
    pub properties: BTreeMap<String, PValue>,
    /// Lines of the designer block besides Caption and the size (`ShowModal`, `Tag` ...), with
    /// their value as written; StartUpPosition is 1 unless given.
    pub frame: Vec<(&'static str, String)>,
}

#[derive(Clone)]
pub(crate) struct ControlDefinition {
    pub name: String,
    pub class: u16,
    pub size: (i32, i32),
    pub position: (i32, i32),
    pub tab_index: i16,
    pub visible: bool,
    pub properties: BTreeMap<String, PValue>,
    pub controls: Vec<ControlDefinition>,
    pub tabs: Option<TabStripState>,
    pub internal: bool,
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub(crate) struct Tab {
    pub name: String,
    pub caption: String,
    pub control_tip_text: String,
    pub tag: String,
    pub accelerator: String,
    pub enabled: bool,
    pub visible: bool,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct TabStripState {
    pub tabs: Vec<Tab>,
    pub selected_index: i32,
}

const FRAME_CLSID: [u8; 16] = [
    0x20, 0x20, 0x18, 0x6e, 0x60, 0xf4, 0xce, 0x11, 0x9b, 0xcd, 0x00, 0xaa, 0x00, 0x60, 0x8e, 0x01,
];
const MULTI_PAGE_CLSID: [u8; 16] = [
    0x70, 0x13, 0xe3, 0x46, 0x7a, 0x3f, 0xce, 0x11, 0xbe, 0xd6, 0, 0xaa, 0, 0x61, 0x10, 0x80,
];
const PAGE_CLSID: [u8; 16] = [
    0xf0, 0x69, 0x2a, 0xc6, 0xdc, 0x16, 0xce, 0x11, 0x9e, 0x98, 0, 0xaa, 0, 0x57, 0x4a, 0x4f,
];
const TEXT_PROPS_CLSID: [u8; 16] = [
    0x20, 0x09, 0xc2, 0xaf, 0x4e, 0xda, 0xce, 0x11, 0xb9, 0x43, 0x00, 0xaa, 0x00, 0x68, 0x87, 0xb4,
];

fn compobj_header(clsid: &[u8; 16]) -> Vec<u8> {
    let mut b = Vec::with_capacity(96);
    b.extend_from_slice(&0xfffe_0001u32.to_le_bytes());
    b.extend_from_slice(&0xa03u32.to_le_bytes());
    b.extend_from_slice(&u32::MAX.to_le_bytes());
    b.extend_from_slice(clsid);
    b
}

fn push_len_string(b: &mut Vec<u8>, text: &str) {
    b.extend_from_slice(&(text.len() as u32).to_le_bytes());
    b.extend_from_slice(text.as_bytes());
}

fn root_compobj() -> Vec<u8> {
    let mut b = compobj_header(&[0; 16]);
    push_len_string(&mut b, "Microsoft Forms 2.0 Form\0");
    push_len_string(&mut b, "Embedded Object\0");
    b.extend_from_slice(&0u32.to_le_bytes());
    b.extend_from_slice(&0x71b2_39f4u32.to_le_bytes());
    b.extend_from_slice(&[0; 12]);
    b
}

fn container_compobj(clsid: &[u8; 16], prog_id: &str) -> Vec<u8> {
    let mut b = compobj_header(clsid);
    let user_type = if *clsid == FRAME_CLSID {
        "Microsoft Forms 2.0 Frame\0"
    } else {
        "Microsoft Forms 2.0 Form\0"
    };
    push_len_string(&mut b, user_type);
    push_len_string(&mut b, "Embedded Object\0");
    push_len_string(&mut b, &format!("{prog_id}\0"));
    b.extend_from_slice(&0x71b2_39f4u32.to_le_bytes());
    b.extend_from_slice(&[0; 12]);
    b
}

/// HIMETRIC → twips, rounded half away from zero like Go's `math.Round`.
fn twips(himetric: i32) -> i64 {
    (himetric as f64 * 1440.0 / 2540.0).round() as i64
}

/// The designer properties of [MS-OVBA] 2.3.5.1, in the order they have to be written.
const DESIGNER_PROPERTIES: &[&str] = &[
    "Caption",
    "ClientHeight",
    "ClientLeft",
    "ClientTop",
    "ClientWidth",
    "Enabled",
    "HelpContextID",
    "RightToLeft",
    "ShowModal",
    "StartUpPosition",
    "Tag",
    "TypeInfoVer",
    "Visible",
    "WhatsThisButton",
    "WhatsThisHelp",
];

fn generated_vbframe(
    name: &str,
    caption: &str,
    size: (i32, i32),
    frame: &[(&'static str, String)],
    codepage: u16,
) -> Result<Vec<u8>> {
    if caption.chars().any(char::is_control) {
        return invalid("フォームの caption に制御文字は使えません");
    }
    let mut lines: Vec<(&str, String)> = vec![
        ("Caption", format!("\"{}\"", caption.replace('"', "\"\""))),
        ("ClientHeight", twips(size.1).to_string()),
        ("ClientWidth", twips(size.0).to_string()),
    ];
    if !frame.iter().any(|(key, _)| *key == "StartUpPosition") {
        lines.push(("StartUpPosition", "1".into()));
    }
    for (key, value) in frame {
        if value.chars().any(char::is_control) {
            return invalid(format!("フォームの {key} に制御文字は使えません"));
        }
        lines.push((key, value.clone()));
    }
    lines.sort_by_key(|(key, _)| DESIGNER_PROPERTIES.iter().position(|k| k == key));
    let mut text = format!("VERSION 5.00\r\nBegin {{C62A69F0-16DC-11CE-9E98-00AA00574A4F}} {name}\r\n");
    for (key, value) in lines {
        text.push_str(&format!("   {key} = {value}\r\n"));
    }
    text.push_str("End\r\n");
    ras_ovba::encode_text(codepage, &text).map_err(|e| PersistError::Invalid(e.to_string()))
}

/// Page box inside a new MultiPage: fixed 96 DPI, 2px border, 19px tab band on the side
/// of `tab_orientation` (Top, Bottom, Left, Right).
fn generated_page_box(size: (i32, i32), tab_orientation: i64) -> ((i32, i32), (i32, i32)) {
    let (w, h) = (size.0 as i64, size.1 as i64);
    let wide = 1.max((w as f64 * 96.0 / 2540.0).round() as i64);
    let tall = 1.max((h as f64 * 96.0 / 2540.0).round() as i64);
    let scale = |pixels: i64, value: i64, extent: i64| -> i32 {
        ((2 * pixels * value + extent) / (2 * extent)) as i32
    };
    // Pixels from the left, top, right and bottom edges.
    let (l, t, r, b) = match tab_orientation {
        1 => (2, 2, 2, 21),
        2 => (21, 2, 2, 2),
        3 => (2, 2, 21, 2),
        _ => (2, 21, 2, 2),
    };
    let left = scale(l, w, wide);
    let top = scale(t, h, tall);
    let right = scale(l.max(wide - r), w, wide);
    let bottom = scale(t.max(tall - b), h, tall);
    ((left, top), (0.max(right - left), 0.max(bottom - top)))
}

// ---------------------------------------------------------------------------
// TabStrip (tabstrip.go, encode side for a fresh record)

fn encode_array_string(value: &str) -> Vec<u8> {
    let units: Vec<u16> = value.encode_utf16().collect();
    let compressed = units.iter().all(|u| *u <= 0xff);
    let mut count = units.len() as u32;
    if compressed {
        count |= 0x8000_0000;
    }
    let mut out = count.to_le_bytes().to_vec();
    if compressed {
        out.extend(units.iter().map(|u| *u as u8));
    } else {
        out.extend(units.iter().flat_map(|u| u.to_le_bytes()));
    }
    while out.len() % 4 != 0 {
        out.push(0);
    }
    out
}

/// `SetTabStrip` for a record that holds no tab state yet.
fn set_tab_strip(record: &mut Record, state: &TabStripState) -> Result<()> {
    let tabs = &state.tabs;
    if tabs.is_empty() && state.selected_index == -1 {
        return Ok(());
    }
    let index = state.selected_index;
    if index < -1
        || (!tabs.is_empty() && index >= tabs.len() as i32)
        || (tabs.is_empty() && index >= 0)
    {
        return invalid(format!("ListIndex {index} がタブの範囲外です"));
    }
    type Getter = fn(&Tab) -> &str;
    let arrays: [(u8, &'static str, &'static str, Getter); 5] = [
        (5, "Items", "ItemsSize", |t| &t.caption),
        (15, "TipStrings", "TipStringsSize", |t| &t.control_tip_text),
        (17, "TabNames", "NamesSize", |t| &t.name),
        (21, "Tags", "TagsSize", |t| &t.tag),
        (23, "Accelerators", "AcceleratorsSize", |t| &t.accelerator),
    ];
    for (i, (bit, name, size_field, get)) in arrays.iter().enumerate() {
        let present = (i == 0 && !tabs.is_empty()) || tabs.iter().any(|t| !get(t).is_empty());
        if !present {
            continue;
        }
        let encoded: Vec<u8> = tabs
            .iter()
            .flat_map(|t| encode_array_string(get(t)))
            .collect();
        record.values.insert(size_field, encoded.len() as i64);
        record.arrays.insert(name, encoded);
        record.mask |= 1u64 << bit;
    }
    if !tabs.is_empty() {
        record.mask |= 1 << 20;
        record.values.insert("TabsAllocated", tabs.len() as i64);
        let mut tail = Vec::with_capacity(tabs.len() * 4);
        for tab in tabs {
            let word = u32::from(tab.visible) | (u32::from(tab.enabled) << 1);
            tail.extend_from_slice(&word.to_le_bytes());
        }
        record.mask |= 1 << 22;
        record.values.insert("TabData", tabs.len() as i64);
        record.tail = tail;
    }
    if index != -1 {
        record.mask |= 1;
        record.values.insert("ListIndex", index as i64);
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// Generation (generate.go)

struct Generation {
    levels: Vec<Level>,
    names: std::collections::HashSet<String>,
    next_id: i64,
}

/// Builds a new form from scratch (`oforms.NewForm`).
pub(crate) fn new_form(input: &Definition, codepage: u16) -> Result<Form> {
    if input.name.is_empty()
        || input
            .name
            .contains(['\\', '/', '\0', '\r', '\n', '"', '{', '}'])
        || input.name.eq_ignore_ascii_case("VBA")
    {
        return invalid("フォーム名が不正です");
    }
    if input.size.0 < 0 || input.size.1 < 0 || input.controls.len() > i16::MAX as usize + 1 {
        return invalid("サイズまたはコントロール数が不正です");
    }
    let vbframe = generated_vbframe(&input.name, &input.caption, input.size, &input.frame, codepage)?;
    let mut g = Generation {
        levels: Vec::new(),
        names: Default::default(),
        next_id: 1,
    };
    let root = g.level(
        String::new(),
        &input.caption,
        input.size,
        &input.properties,
        input.controls.clone(),
        0,
        0,
        None,
    )?;
    g.levels[root].vbframe = Some(vbframe);
    Ok(Form {
        name: input.name.clone(),
        levels: g.levels,
    })
}

impl Generation {
    #[allow(clippy::too_many_arguments)]
    fn level(
        &mut self,
        path: String,
        caption: &str,
        size: (i32, i32),
        properties: &BTreeMap<String, PValue>,
        mut controls: Vec<ControlDefinition>,
        depth: usize,
        class: u16,
        tabs: Option<TabStripState>,
    ) -> Result<usize> {
        if depth >= MAX_NESTING_DEPTH || controls.len() > i16::MAX as usize + 1 {
            return invalid("入れ子またはコントロール数が上限を超えています");
        }
        let first_id = self.next_id;
        let mut root = Record::new(&FORM);
        root.mask = 1 << 3 | 1 << 10 | 1 << 11 | 1 << 26 | 1 << 27;
        root.values.insert("NextAvailableID", 1);
        root.values.insert("ShapeCookie", controls.len() as i64);
        root.values.insert("DrawBuffer", 32000);
        root.sizes.insert("DisplayedSize", size);
        root.sizes.insert("LogicalSize", (0, 0));
        if depth > 0 {
            root.mask |= 1 << 6;
            root.values.insert("BooleanProperties", 0x8004);
            if class == 14 {
                root.mask |= 1 << 17;
                root.values.insert("SpecialEffect", 3);
            }
            if class == 57 {
                root.values.insert("BooleanProperties", 0xc004);
            }
        }
        root.set("Caption", &PValue::Str(caption.to_string()))?;
        for (name, value) in properties {
            if skipped_property(name, class) {
                continue;
            }
            root.set(name, value)?;
        }
        let class_table_raw =
            if root.values.get("BooleanProperties").copied().unwrap_or(0) & 0x8000 != 0 {
                Vec::new()
            } else {
                vec![0, 0]
            };
        let mut trailing_raw = Vec::new();
        let mut multipage_disabled = false;
        if class == 57 {
            let tabs = tabs.unwrap_or(TabStripState {
                tabs: Vec::new(),
                selected_index: if controls.is_empty() { -1 } else { 0 },
            });
            if tabs.tabs.len() != controls.len() {
                return invalid("MultiPage のタブとページの数が一致しません");
            }
            let orientation = match properties.get("TabOrientation") {
                Some(PValue::Int(n)) => *n,
                _ => 0,
            };
            let (position, page_size) = generated_page_box(size, orientation);
            for (i, control) in controls.iter_mut().enumerate() {
                if control.class != 7 {
                    return invalid("MultiPage には Page しか入れられません");
                }
                control.position = position;
                control.size = page_size;
                control.visible = i as i32 == tabs.selected_index;
            }
            controls.insert(
                0,
                ControlDefinition {
                    name: String::new(),
                    class: 18,
                    size,
                    position: (0, 0),
                    tab_index: 0,
                    visible: true,
                    // The tab settings of the MultiPage are the TabStrip's.
                    properties: properties
                        .iter()
                        .filter(|(name, _)| crate::scalars::TAB_STRIP_FIELDS.contains(&name.as_str()))
                        .map(|(name, value)| (name.clone(), value.clone()))
                        .collect(),
                    controls: Vec::new(),
                    tabs: Some(tabs),
                    internal: true,
                },
            );
            trailing_raw = vec![0, 2, 12, 0, 25, 0, 0, 0, 252, 143, 0, 0, 255, 1, 0, 0];
            multipage_disabled = matches!(
                properties.get("MultiPageEnabled"),
                Some(PValue::Bool(false))
            );
        }
        let index = self.levels.len();
        let (compobj, clsid) = if depth > 0 {
            let (clsid, prog_id) = match class {
                7 => (PAGE_CLSID, "Forms.Form.1"),
                57 => (MULTI_PAGE_CLSID, "Forms.MultiPage.1"),
                _ => (FRAME_CLSID, "Forms.Frame.1"),
            };
            (container_compobj(&clsid, prog_id), clsid)
        } else {
            (root_compobj(), [0; 16])
        };
        self.levels.push(Level {
            path: path.clone(),
            record: root,
            font_raw: Vec::new(),
            class_table_raw,
            controls: Vec::new(),
            depths_raw: Vec::new(),
            trailing_raw,
            multipage: None,
            compobj,
            vbframe: None,
            clsid,
        });

        let mut built = Vec::with_capacity(controls.len());
        for control in controls {
            let lower = control.name.to_lowercase();
            if (!control.internal && (control.name.is_empty() || self.names.contains(&lower)))
                || control.name.contains(['\\', '/', '\0', '\r', '\n', '"'])
                || control.tab_index < 0
                || control.size.0 < 0
                || control.size.1 < 0
            {
                return invalid(format!("コントロール {:?} が不正です", control.name));
            }
            if !control.internal {
                self.names.insert(lower);
            }
            let id = self.next_id;
            self.next_id += 1;
            let container = is_container_class(control.class);
            let spec = if container {
                Some(&FORM)
            } else {
                spec_for_class(control.class)
            };
            let Some(spec) = spec else {
                return unsupported(format!("クラス {}", control.class));
            };
            if control.class == 15 || (!container && !control.controls.is_empty()) {
                return unsupported(format!("クラス {}", control.class));
            }
            let mut record = Record::new(spec);
            if spec.extra.iter().any(|x| x.name == "Size") {
                record.set_size("Size", control.size);
            }
            if std::ptr::eq(spec, &MORPH_DATA) {
                record.mask |= 1 << 31;
                let style = match control.class {
                    23 => 1,
                    24 => 2,
                    25 => 3,
                    26 => 4,
                    27 => 5,
                    28 => 6,
                    _ => 0,
                };
                record.set_int("DisplayStyle", style)?;
                // What Excel stores for a new control; the record's own
                // defaults are no drop button and no auto-completion.
                if control.class == 25 {
                    record.set_int("ShowDropButtonWhen", 2)?;
                    record.set_int("MatchEntry", 1)?;
                }
                if control.class == 24 {
                    record.set_int("MatchEntry", 0)?;
                    // MS-OFORMS: MUST be fmScrollBarsBoth for a ListBox.
                    record.set_int("ScrollBars", 3)?;
                }
            }
            if spec.text_props {
                record.text_props = Some(Box::new(Record::new(&TEXT_PROPS)));
            }
            if control.class == 18 {
                record.mask |= 1 << 19;
                let tabs = control.tabs.clone().unwrap_or(TabStripState {
                    tabs: Vec::new(),
                    selected_index: -1,
                });
                set_tab_strip(&mut record, &tabs)?;
                if tabs.tabs.is_empty() {
                    record.mask |= 1;
                    record.values.insert("ListIndex", -1);
                }
                let font = record.text_props.as_mut().expect("TabStrip has TextProps");
                font.set("FontName", &PValue::Str("Tahoma".into()))?;
                font.set_int("FontHeight", 165)?;
            }
            for (name, value) in &control.properties {
                if skipped_property(name, control.class) {
                    continue;
                }
                record.set(name, value).map_err(|e| match e {
                    PersistError::Invalid(m) => {
                        PersistError::Invalid(format!("コントロール {:?}: {m}", control.name))
                    }
                    PersistError::Unsupported(m) => {
                        PersistError::Unsupported(format!("コントロール {:?}: {m}", control.name))
                    }
                })?;
            }
            if let (Some(text_props), Some(PValue::Int(bits))) =
                (record.text_props.as_mut(), control.properties.get("VariousPropertyBits"))
            {
                if bits & 2 == 0 && DISABLED_TEXT_CLASSES.contains(&control.class) {
                    // A disabled control: Excel keeps its text grayed in the font effects.
                    let effects = text_props.values.get("FontEffects").copied().unwrap_or(0);
                    text_props.set_int("FontEffects", effects | DISABLED_TEXT_EFFECTS)?;
                }
            }
            if let Some(PValue::Str(text)) = control.properties.get("ColumnWidths") {
                // The widths go to the MorphDataColumnInfo array after the TextProps.
                let widths = crate::scalars::parse_column_widths(text).map_err(|e| {
                    PersistError::Invalid(format!("コントロール {:?}: columnWidths: {e}", control.name))
                })?;
                if !widths.is_empty() {
                    record.set_int("cColumnInfo", widths.len() as i64)?;
                    record.tail = crate::scalars::encode_column_infos(&widths);
                }
            }
            record.encode()?;
            let mut child = None;
            if container {
                let child_path = if path.is_empty() {
                    format!("i{id:02}")
                } else {
                    format!("{path}/i{id:02}")
                };
                child = Some(self.level(
                    child_path,
                    "",
                    control.size,
                    &control.properties,
                    control.controls.clone(),
                    depth + 1,
                    control.class,
                    control.tabs.clone(),
                )?);
            }
            let site_name = edited_string(&control.name)?;
            let mut flags = crate::scalars::fresh_site_flags(control.class);
            if !control.visible {
                flags &= !2;
            }
            // TabStop, Default, Cancel.
            for bit in [0u8, 2, 3] {
                if let Some(PValue::Bool(on)) = control.properties.get(&format!("SiteFlag{bit}")) {
                    if *on {
                        flags |= 1 << bit;
                    } else {
                        flags &= !(1 << bit);
                    }
                }
            }
            let mut values = HashMap::new();
            values.insert("NameData", site_name.packed());
            values.insert("ID", id);
            values.insert("BitFlags", flags);
            values.insert("TabIndex", control.tab_index as i64);
            values.insert("ClsidCacheIndex", control.class as i64);
            let mut site = Site {
                mask: 1 << 0 | 1 << 2 | 1 << 4 | 1 << 5 | 1 << 6 | 1 << 7 | 1 << 8,
                values,
                name: site_name,
                tag: None,
                tip: None,
                sources: [None, None],
                position: control.position,
            };
            if container {
                site.mask &= !(1 << 5);
            }
            for (property, length, bit) in [
                ("Tag", "TagData", 1u32 << 1),
                ("ControlTipText", "ControlTipTextData", 1 << 11),
            ] {
                if let Some(value) = control.properties.get(property) {
                    let PValue::Str(text) = value else {
                        return invalid(format!("{property} は文字列が必要です"));
                    };
                    let stored = edited_string(text)?;
                    site.mask |= bit;
                    site.values.insert(length, stored.packed());
                    if property == "Tag" {
                        site.tag = Some(stored);
                    } else {
                        site.tip = Some(stored);
                    }
                }
            }
            // The cell and the range a control is bound to (scalars.rs, `SiteText`).
            for (index, (property, length, bit)) in [
                ("SiteControlSource", "ControlSourceData", 1u32 << 13),
                ("SiteRowSource", "RowSourceData", 1 << 14),
            ]
            .into_iter()
            .enumerate()
            {
                if let Some(PValue::Str(text)) = control.properties.get(property) {
                    let stored = edited_string(text)?;
                    site.mask |= bit;
                    site.values.insert(length, stored.packed());
                    site.sources[index] = Some(stored);
                }
            }
            built.push(Control {
                name: control.name.clone(),
                class: control.class,
                id,
                site,
                record: (!container).then_some(record),
                child,
            });
        }
        let mut depths = Vec::with_capacity(built.len() * 2);
        for _ in &built {
            depths.extend_from_slice(&[0, 1]);
        }
        pad(&mut depths, 4);
        if self.next_id > MAX_SITES_PER_FORM + 1 {
            return invalid("コントロールの総数が上限を超えています");
        }
        if class == 57 {
            self.levels[index].multipage = Some(MultiPageState {
                hidden_id: built[0].id,
                page_ids: built[1..].iter().map(|c| c.id).collect(),
                disabled: multipage_disabled,
            });
        }
        let level = &mut self.levels[index];
        level.record.values.insert("NextAvailableID", self.next_id);
        level
            .record
            .values
            .insert("ShapeCookie", self.next_id - first_id);
        level.depths_raw = depths;
        level.controls = built;
        Ok(index)
    }
}

fn skipped_property(name: &str, class: u16) -> bool {
    matches!(
        name,
        "Tag" | "ControlTipText" | "Accelerator" | "MultiPageEnabled" | "ColumnWidths"
    ) || name.starts_with("Site")
        || (class == 7 && (name == "Caption" || name == "BooleanProperties"))
        // Held by the hidden TabStrip of a MultiPage.
        || (class == 57 && crate::scalars::TAB_STRIP_FIELDS.contains(&name))
}

// ---------------------------------------------------------------------------
// The generated form and its typed edits (mutate.go, font.go)

/// A partial font edit; `None` members are left untouched.
#[derive(Clone, Debug, Default, PartialEq)]
pub(crate) struct FontEdit {
    pub name: Option<String>,
    pub size: Option<f64>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub strike_through: Option<bool>,
    pub charset: Option<i64>,
    pub pitch_and_family: Option<i64>,
}

impl FontEdit {
    pub(crate) fn is_empty(&self) -> bool {
        *self == FontEdit::default()
    }
}

/// A typed appearance edit applied after generation.
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum Edit {
    Font(FontEdit),
    Int(&'static str, i64),
}

pub(crate) struct Form {
    name: String,
    levels: Vec<Level>,
}

fn edit_text_font(record: &mut Record, font: &FontEdit) -> Result<()> {
    if let Some(name) = &font.name {
        if name.is_empty() {
            return invalid("フォント名が空です");
        }
        record.set("FontName", &PValue::Str(name.clone()))?;
    }
    if let Some(size) = font.size {
        let v = size * 20.0;
        if !v.is_finite() || v <= 0.0 || v > 4_294_967.0 || (v - v.round()).abs() > 1e-7 {
            return invalid("フォントサイズは 0.05 ポイント単位の正の値にしてください");
        }
        record.set_int("FontHeight", v.round() as i64)?;
    }
    if let Some(charset) = font.charset {
        if !(0..=255).contains(&charset) {
            return invalid("charset は 0〜255 にしてください");
        }
        record.set_int("FontCharSet", charset)?;
    }
    if let Some(pitch) = font.pitch_and_family {
        if !(0..=255).contains(&pitch) {
            return invalid("pitchAndFamily は 0〜255 にしてください");
        }
        record.set_int("FontPitchAndFamily", pitch)?;
    }
    for (bit, value) in [font.bold, font.italic, font.underline, font.strike_through]
        .into_iter()
        .enumerate()
    {
        let Some(value) = value else { continue };
        let mut flags = record.values.get("FontEffects").copied().unwrap_or(0);
        if value {
            flags |= 1 << bit;
        } else {
            flags &= !(1 << bit);
        }
        record.set_int("FontEffects", flags)?;
    }
    if let Some(bold) = font.bold {
        record.set_int("FontWeight", if bold { 700 } else { 400 })?;
    }
    Ok(())
}

impl Form {
    fn find(&self, name: &str) -> Result<(usize, usize)> {
        let mut found = None;
        for (l, level) in self.levels.iter().enumerate() {
            for (c, control) in level.controls.iter().enumerate() {
                if control.name == name {
                    if found.is_some() {
                        return invalid("コントロール名があいまいです");
                    }
                    found = Some((l, c));
                }
            }
        }
        found
            .ok_or_else(|| PersistError::Invalid(format!("コントロール {name:?} が見つかりません")))
    }

    fn edit_level_font(&mut self, level: usize, font: &FontEdit) -> Result<()> {
        let level = &mut self.levels[level];
        let mut text_props = Record::new(&TEXT_PROPS);
        if !level.font_raw.is_empty() {
            // Only fonts written by this module are edited (one edit per level).
            return unsupported("フォントの再編集");
        }
        edit_text_font(&mut text_props, font)?;
        let mut raw = TEXT_PROPS_CLSID.to_vec();
        raw.extend_from_slice(&text_props.encode()?);
        level.font_raw = raw;
        level.record.mask |= FORM_MASK_FONT;
        level.record.values.insert("Font", 0xFFFF);
        Ok(())
    }

    /// `ApplyEdits` restricted to the typed appearance properties used by
    /// generation. `control` `None` addresses the form itself.
    pub(crate) fn apply(
        &mut self,
        control: Option<&str>,
        property: &'static str,
        edit: &Edit,
    ) -> Result<()> {
        let Some(name) = control else {
            return match edit {
                Edit::Font(font) => self.edit_level_font(0, font),
                Edit::Int(p, value) => {
                    if !matches!(*p, "BackColor" | "ForeColor" | "BorderStyle") {
                        return unsupported(property);
                    }
                    set_checked(&mut self.levels[0].record, p, *value, None)
                }
            };
        };
        let (l, c) = self.find(name)?;
        let class = self.levels[l].controls[c].class;
        let child = self.levels[l].controls[c].child;
        // A MultiPage exposes the font and colors of its hidden TabStrip.
        let hidden = child.filter(|_| class == 57).map(|level| (level, 0usize));
        match edit {
            Edit::Font(font) => {
                if let Some((hl, hc)) = hidden {
                    let record = self.levels[hl].controls[hc]
                        .record
                        .as_mut()
                        .expect("hidden TabStrip record");
                    return edit_text_font(record.text_props.as_mut().expect("TextProps"), font);
                }
                if let Some(level) = child {
                    return self.edit_level_font(level, font);
                }
                let Some(text_props) = self.levels[l].controls[c]
                    .record
                    .as_mut()
                    .and_then(|r| r.text_props.as_mut())
                else {
                    return unsupported(property);
                };
                edit_text_font(text_props, font)
            }
            Edit::Int("TextAlign", value) => {
                let Some(text_props) = self.levels[l].controls[c]
                    .record
                    .as_mut()
                    .and_then(|r| r.text_props.as_mut())
                else {
                    return unsupported(property);
                };
                if !(1..=3).contains(value) {
                    return invalid("textAlign は 1〜3 にしてください");
                }
                // PARAFORMAT uses right = 2 / center = 3; MSForms the reverse.
                let align = match value {
                    2 => 3,
                    3 => 2,
                    v => *v,
                };
                text_props.set_int("ParagraphAlign", align)
            }
            Edit::Int(p, value) => {
                if (*p == "BackColor" || *p == "ForeColor") && class == 57 {
                    let (hl, hc) = hidden.expect("MultiPage level");
                    let record = self.levels[hl].controls[hc]
                        .record
                        .as_mut()
                        .expect("hidden TabStrip record");
                    return set_checked(record, p, *value, Some(18));
                }
                if let Some(level) = child {
                    if class != 14 && class != 7 {
                        return unsupported(property);
                    }
                    return set_checked(&mut self.levels[level].record, p, *value, Some(class));
                }
                let record = self.levels[l].controls[c]
                    .record
                    .as_mut()
                    .expect("control record");
                set_checked(record, p, *value, Some(class))
            }
        }
    }

    /// Serialises the form (`SerializeForm`).
    pub(crate) fn serialize(&self) -> Result<Designer> {
        let mut streams = BTreeMap::new();
        let mut storages = BTreeMap::new();
        for level in &self.levels {
            let prefix = if level.path.is_empty() {
                String::new()
            } else {
                format!("{}/", level.path)
            };
            let mut object = Vec::new();
            let mut sites = Vec::with_capacity(level.controls.len());
            for control in &level.controls {
                let mut site_values = control.site.values.clone();
                if let Some(record) = &control.record {
                    let raw = record.encode()?;
                    site_values.insert("ObjectStreamSize", raw.len() as i64);
                    object.extend_from_slice(&raw);
                }
                let site = Site {
                    values: site_values,
                    ..clone_site(&control.site)
                };
                sites.push(site.encode()?);
            }
            if object.len() > MAX_DESIGNER_STREAM_SIZE {
                return invalid("o ストリームが大きすぎます");
            }
            let mut f = level.record.encode()?;
            // FormStreamData: MouseIcon, Font, Picture.
            if let Some(icon) = level.record.pictures.get("MouseIcon") {
                f.extend_from_slice(icon);
            }
            f.extend_from_slice(&level.font_raw);
            if let Some(picture) = level.record.pictures.get("Picture") {
                f.extend_from_slice(picture);
            }
            f.extend_from_slice(&level.class_table_raw);
            f.extend_from_slice(&(sites.len() as u32).to_le_bytes());
            let count_at = f.len();
            f.extend_from_slice(&[0; 4]);
            let start = f.len();
            f.extend_from_slice(&level.depths_raw);
            for site in &sites {
                f.extend_from_slice(site);
            }
            let site_bytes = (f.len() - start) as u32;
            f[count_at..count_at + 4].copy_from_slice(&site_bytes.to_le_bytes());
            f.extend_from_slice(&level.trailing_raw);
            if f.len() > MAX_DESIGNER_STREAM_SIZE {
                return invalid("f ストリームが大きすぎます");
            }
            streams.insert(format!("{prefix}f"), f);
            streams.insert(format!("{prefix}o"), object);
            streams.insert(format!("{prefix}\u{1}CompObj"), level.compobj.clone());
            if let Some(frame) = &level.vbframe {
                streams.insert(format!("{prefix}\u{3}VBFrame"), frame.clone());
            }
            if let Some(state) = &level.multipage {
                streams.insert(format!("{prefix}x"), encode_multipage_x(state)?);
            }
            storages.insert(level.path.clone(), level.clsid);
        }
        Ok(Designer {
            name: self.name.clone(),
            streams,
            storages,
        })
    }
}

fn clone_site(site: &Site) -> Site {
    Site {
        mask: site.mask,
        values: site.values.clone(),
        name: site.name.clone(),
        tag: site.tag.clone(),
        tip: site.tip.clone(),
        sources: site.sources.clone(),
        position: site.position,
    }
}

fn set_checked(record: &mut Record, property: &str, value: i64, class: Option<u16>) -> Result<()> {
    let Some(field) = record.spec.data.iter().find(|f| f.name == property) else {
        return unsupported(property.to_string());
    };
    if value < 0 || (value as u64) >= (1u64 << (8 * field.size)) {
        return invalid(format!("{property} の値 {value} が範囲外です"));
    }
    match property {
        "BorderStyle" if value > 1 => return invalid("borderStyle は 0 か 1 にしてください"),
        // An Image's picture, or the background of a Frame or Page.
        "PictureAlignment" if !matches!(class, Some(7 | 12 | 14)) || value > 4 => {
            return invalid("pictureAlignment が不正です")
        }
        "PictureSizeMode" if !matches!(class, Some(7 | 12 | 14)) || !matches!(value, 0 | 1 | 3) => {
            return invalid("pictureSizeMode が不正です")
        }
        _ => {}
    }
    record.set_int(property, value)
}

fn encode_multipage_x(state: &MultiPageState) -> Result<Vec<u8>> {
    let mut b = vec![0u8, 2, 4, 0, 0, 0, 0, 0];
    for _ in &state.page_ids {
        b.extend_from_slice(&[0, 2, 4, 0, 0, 0, 0, 0]);
    }
    let mut props = Record::new(&MULTI_PAGE_PROPERTIES);
    if state.disabled {
        props.mask |= 1 << 3;
    }
    props.set_int("PageCount", state.page_ids.len() as i64)?;
    props.set_int("ID", state.hidden_id)?;
    b.extend_from_slice(&props.encode()?);
    for id in &state.page_ids {
        b.extend_from_slice(&(*id as u32).to_le_bytes());
    }
    Ok(b)
}

#[cfg(test)]
mod tests {
    use super::generated_page_box;

    #[test]
    fn the_page_box_leaves_the_band_of_tabs_on_their_side() {
        // 240 × 180 points.
        let size = (8467, 6350);
        let top = generated_page_box(size, 0);
        let bottom = generated_page_box(size, 1);
        let left = generated_page_box(size, 2);
        let right = generated_page_box(size, 3);
        // 2px and 21px at 96 DPI, in HIMETRIC.
        let (two, band) = (53, 556);
        assert_eq!(top.0, (two, band));
        assert_eq!(bottom.0, (two, two));
        assert_eq!(left.0, (band, two));
        assert_eq!(right.0, (two, two));
        // The band takes the same room whichever side it is on.
        assert_eq!(top.1, bottom.1);
        assert_eq!(left.1, right.1);
        assert_eq!(top.1 .0 - left.1 .0, band - two);
        assert_eq!(left.1 .1 - top.1 .1, band - two);
    }
}
