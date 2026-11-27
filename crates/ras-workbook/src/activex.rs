//! The persisted data of an ActiveX control of Microsoft Forms 2.0 on a sheet (MS-OFORMS):
//! read, changed and written again. Excel keeps it in `xl/activeX/activeX<n>.bin` (the control's
//! CLSID, then its stream: `ax:persistence="persistStream"` or `"persistStreamInit"`).
//!
//! A control's stream starts with its record: a mask of the properties it has, a `DataBlock`
//! (each property aligned to its size) and an `ExtraDataBlock` (strings and the size, aligned
//! to 4). What follows the record (its pictures, its font) is kept as it is. The fields are
//! those umya-spreadsheet's `drawingml::controls` reads the stream by.

use std::collections::HashMap;

/// A field of a `DataBlock`: (mask bit, name, size).
type Field = (u8, &'static str, usize);

const COMMAND_BUTTON: &[Field] = &[
    (0, "ForeColor", 4),
    (1, "BackColor", 4),
    (2, "VariousPropertyBits", 4),
    (3, "Caption", 4),
    (4, "PicturePosition", 4),
    (6, "MousePointer", 1),
    (7, "Picture", 2),
    (8, "Accelerator", 2),
    (10, "MouseIcon", 2),
];
const LABEL: &[Field] = &[
    (0, "ForeColor", 4),
    (1, "BackColor", 4),
    (2, "VariousPropertyBits", 4),
    (3, "Caption", 4),
    (4, "PicturePosition", 4),
    (6, "MousePointer", 1),
    (7, "BorderColor", 4),
    (8, "BorderStyle", 2),
    (9, "SpecialEffect", 2),
    (10, "Picture", 2),
    (11, "Accelerator", 2),
    (12, "MouseIcon", 2),
];
const MORPH_DATA: &[Field] = &[
    (0, "VariousPropertyBits", 4),
    (1, "BackColor", 4),
    (2, "ForeColor", 4),
    (3, "MaxLength", 4),
    (4, "BorderStyle", 1),
    (5, "ScrollBars", 1),
    (6, "DisplayStyle", 1),
    (7, "MousePointer", 1),
    (9, "PasswordChar", 2),
    (10, "ListWidth", 4),
    (11, "BoundColumn", 2),
    (12, "TextColumn", 2),
    (13, "ColumnCount", 2),
    (14, "ListRows", 2),
    (15, "cColumnInfo", 2),
    (16, "MatchEntry", 1),
    (17, "ListStyle", 1),
    (18, "ShowDropButtonWhen", 1),
    (20, "DropButtonStyle", 1),
    (21, "MultiSelect", 1),
    (22, "Value", 4),
    (23, "Caption", 4),
    (24, "PicturePosition", 4),
    (25, "BorderColor", 4),
    (26, "SpecialEffect", 4),
    (27, "MouseIcon", 2),
    (28, "Picture", 2),
    (29, "Accelerator", 2),
    (32, "GroupName", 4),
];
const SPIN_BUTTON: &[Field] = &[
    (0, "ForeColor", 4),
    (1, "BackColor", 4),
    (2, "VariousPropertyBits", 4),
    (5, "Min", 4),
    (6, "Max", 4),
    (7, "Position", 4),
    (8, "PrevEnabled", 4),
    (9, "NextEnabled", 4),
    (10, "SmallChange", 4),
    (11, "Orientation", 4),
];
const SCROLL_BAR: &[Field] = &[
    (0, "ForeColor", 4),
    (1, "BackColor", 4),
    (2, "VariousPropertyBits", 4),
    (4, "MousePointer", 1),
    (5, "Min", 4),
    (6, "Max", 4),
    (7, "Position", 4),
    (9, "PrevEnabled", 4),
    (10, "NextEnabled", 4),
    (11, "SmallChange", 4),
    (12, "LargeChange", 4),
    (13, "Orientation", 4),
];

/// The font of a control (what follows the record of a `CommandButton`, `Label` or
/// `MorphData`, after its pictures).
const TEXT_PROPS: &[Field] = &[
    (0, "FontName", 4),
    (1, "FontEffects", 4),
    (2, "FontHeight", 4),
    (4, "FontCharSet", 1),
    (5, "FontPitchAndFamily", 1),
    (6, "ParagraphAlign", 1),
    (7, "FontWeight", 2),
];

/// The size in the `ExtraDataBlock`.
const SIZE: &str = "#Size";

/// A class's record: its fields, a 64-bit mask, its `ExtraDataBlock`.
type Spec = (&'static [Field], bool, &'static [(u8, &'static str)]);

/// How a class's record is laid out: its fields, whether its mask has 64 bits, its
/// `ExtraDataBlock` (mask bit, name: a string, or the size).
fn spec(class: &str) -> Option<Spec> {
    Some(match class {
        "CommandButton" => (COMMAND_BUTTON, false, &[(3, "Caption"), (5, SIZE)]),
        "Label" => (LABEL, false, &[(3, "Caption"), (5, SIZE)]),
        "MorphData" => (
            MORPH_DATA,
            true,
            &[(8, SIZE), (22, "Value"), (23, "Caption"), (32, "GroupName")],
        ),
        "SpinButton" => (SPIN_BUTTON, false, &[(3, SIZE)]),
        "ScrollBar" => (SCROLL_BAR, false, &[(3, SIZE)]),
        "TextProps" => (TEXT_PROPS, false, &[(0, "FontName")]),
        _ => return None,
    })
}

/// The class of an ActiveX control by its CLSID, with `MorphData`'s `DisplayStyle` (TextBox
/// 1, ListBox 2, ComboBox 3, CheckBox 4, OptionButton 5, ToggleButton 6).
pub(crate) fn class_of(clsid: &str) -> (&'static str, u8) {
    match clsid.trim_matches(['{', '}']).to_ascii_uppercase().as_str() {
        "D7053240-CE69-11CD-A777-00DD01143C57" => ("CommandButton", 0),
        "978C9E23-D4B0-11CE-BF2D-00AA003F40D0" => ("Label", 0),
        "8BD21D10-EC42-11CE-9E0D-00AA006002F3" => ("MorphData", 1),
        "8BD21D20-EC42-11CE-9E0D-00AA006002F3" => ("MorphData", 2),
        "8BD21D30-EC42-11CE-9E0D-00AA006002F3" => ("MorphData", 3),
        "8BD21D40-EC42-11CE-9E0D-00AA006002F3" => ("MorphData", 4),
        "8BD21D50-EC42-11CE-9E0D-00AA006002F3" => ("MorphData", 5),
        "8BD21D60-EC42-11CE-9E0D-00AA006002F3" => ("MorphData", 6),
        "79176FB0-B7F2-11CE-97EF-00AA006D2776" => ("SpinButton", 0),
        "DFD181E0-5E2F-11CE-A449-00AA004A803D" => ("ScrollBar", 0),
        _ => ("", 0),
    }
}

/// The Microsoft Forms name of a class (`MorphData` by its display style).
pub(crate) fn forms_type(class: &str, display: u8) -> &'static str {
    match (class, display) {
        ("MorphData", 1) => "TextBox",
        ("MorphData", 2) => "ListBox",
        ("MorphData", 3) => "ComboBox",
        ("MorphData", 4) => "CheckBox",
        ("MorphData", 5) => "OptionButton",
        ("MorphData", 6) => "ToggleButton",
        ("CommandButton", _) => "CommandButton",
        ("Label", _) => "Label",
        ("SpinButton", _) => "SpinButton",
        ("ScrollBar", _) => "ScrollBar",
        _ => "",
    }
}

/// The CLSID of a Microsoft Forms type (`CommandButton`, `TextBox`, ...).
pub(crate) fn clsid_of(forms: &str) -> Option<&'static str> {
    Some(match forms {
        "CommandButton" => "D7053240-CE69-11CD-A777-00DD01143C57",
        "Label" => "978C9E23-D4B0-11CE-BF2D-00AA003F40D0",
        "TextBox" => "8BD21D10-EC42-11CE-9E0D-00AA006002F3",
        "ListBox" => "8BD21D20-EC42-11CE-9E0D-00AA006002F3",
        "ComboBox" => "8BD21D30-EC42-11CE-9E0D-00AA006002F3",
        "CheckBox" => "8BD21D40-EC42-11CE-9E0D-00AA006002F3",
        "OptionButton" => "8BD21D50-EC42-11CE-9E0D-00AA006002F3",
        "ToggleButton" => "8BD21D60-EC42-11CE-9E0D-00AA006002F3",
        "SpinButton" => "79176FB0-B7F2-11CE-97EF-00AA006D2776",
        "ScrollBar" => "DFD181E0-5E2F-11CE-A449-00AA004A803D",
        _ => return None,
    })
}

/// A CLSID as it is stored (a GUID: its first three parts little-endian).
fn clsid_bytes(clsid: &str) -> Option<[u8; 16]> {
    let hex: Vec<u8> = clsid
        .chars()
        .filter_map(|c| c.to_digit(16).map(|d| d as u8))
        .collect();
    if hex.len() != 32 {
        return None;
    }
    let mut out = [0u8; 16];
    for (i, o) in out.iter_mut().enumerate() {
        *o = hex[i * 2] << 4 | hex[i * 2 + 1];
    }
    out[0..4].reverse();
    out[4..6].reverse();
    out[6..8].reverse();
    Some(out)
}

/// The font the VBA editor (in Japanese) stores in every control it makes, as `ras-form`
/// writes a new control's: MS UI Gothic 9pt (in twips), `SHIFTJIS_CHARSET`, `VARIABLE_PITCH`.
const FRESH_FONT: (&str, i64, i64, i64) = ("MS UI Gothic", 180, 128, 2);

/// The data (a `.bin` part: the CLSID and the stream) of a new control of a Microsoft Forms
/// type (`CommandButton`, `TextBox`, ...), as a control just put on a sheet: the
/// `VariousPropertyBits`, colors and list settings of a new control of its kind, its caption
/// and value, its size (HIMETRIC), and the font of the classes that have one. None for a type
/// that is not one of these.
pub(crate) fn fresh(
    forms: &str,
    caption: Option<&str>,
    value: Option<&str>,
    size: (i32, i32),
) -> Option<Vec<u8>> {
    let clsid = clsid_of(forms)?;
    let (class, display) = class_of(clsid);
    let empty = |class: &'static str, head: Vec<u8>| Stream {
        class,
        head,
        mask: 0,
        values: HashMap::new(),
        strings: HashMap::new(),
        size: None,
        rest: Vec::new(),
    };
    let mut s = empty(spec_class(class)?, clsid_bytes(clsid)?.to_vec());
    // As Excel writes them (`ras-form`'s `fresh_bits`): enabled, opaque, ...; a text box and a
    // combo box editable, a label word-wrapped.
    let bits: i64 = match forms {
        "TextBox" | "ComboBox" => 0x2c80_481b,
        "Label" => 0x0080_001b,
        "ListBox" | "CheckBox" | "OptionButton" | "ToggleButton" => 0x2c80_081b,
        _ => 0x1b,
    };
    s.set_number("VariousPropertyBits", bits);
    if class == "MorphData" {
        // (A bit without data that MS-OFORMS reserves; a new control has it.)
        s.mask |= 1 << 31;
        s.set_number("DisplayStyle", i64::from(display));
        match forms {
            // The drop button always shown, the entry completed.
            "ComboBox" => {
                s.set_number("ShowDropButtonWhen", 2);
                s.set_number("MatchEntry", 1);
            }
            // (MS-OFORMS: a list box's scroll bars are fmScrollBarsBoth.)
            "ListBox" => {
                s.set_number("MatchEntry", 0);
                s.set_number("ScrollBars", 3);
            }
            // The button face and text colors (the record's defaults are a text box's).
            "CheckBox" | "OptionButton" | "ToggleButton" => {
                s.set_number("BackColor", 0x8000_000F);
                s.set_number("ForeColor", 0x8000_0012);
            }
            _ => {}
        }
    }
    if let Some(c) = caption {
        s.set_string("Caption", c);
    }
    if let Some(v) = value {
        s.set_string("Value", v);
    }
    s.set_size(size.0, size.1);
    // The font (TextProps), right after the record (there are no pictures).
    if matches!(class, "CommandButton" | "Label" | "MorphData") {
        let mut font = empty("TextProps", Vec::new());
        let (name, height, charset, pitch) = FRESH_FONT;
        font.set_string("FontName", name);
        font.set_number("FontHeight", height);
        font.set_number("FontCharSet", charset);
        font.set_number("FontPitchAndFamily", pitch);
        // A button's caption centred (PARAFORMAT's center).
        if matches!(forms, "CommandButton" | "ToggleButton") {
            font.set_number("ParagraphAlign", 3);
        }
        s.rest = font.write();
    }
    Some(s.write())
}

/// The persisted data of a control: its record read, and what is around it.
#[derive(Clone, Debug)]
pub(crate) struct Stream {
    class: &'static str,
    /// The CLSID before the stream (of a `.bin` part), kept.
    head: Vec<u8>,
    mask: u64,
    values: HashMap<&'static str, u32>,
    /// The strings of the `ExtraDataBlock` as they are stored.
    strings: HashMap<&'static str, Vec<u8>>,
    size: Option<(i32, i32)>,
    /// What follows the record.
    rest: Vec<u8>,
}

fn u16_at(d: &[u8], i: usize) -> Option<u16> {
    d.get(i..i + 2).map(|b| u16::from_le_bytes([b[0], b[1]]))
}

fn u32_at(d: &[u8], i: usize) -> Option<u32> {
    d.get(i..i + 4)
        .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

impl Stream {
    /// The data of a control of `class` (a `.bin` part with its CLSID, or the stream alone).
    /// None when its record is not read whole as the class lays it out.
    pub(crate) fn read(class: &str, data: &[u8]) -> Option<Stream> {
        let (fields, mask64, extra) = spec(class)?;
        let class = spec_class(class)?;
        let (head, d) = if data.len() > 16 && data[16..].starts_with(&[0, 2]) {
            (data[..16].to_vec(), &data[16..])
        } else {
            (Vec::new(), data)
        };
        if d.len() < 8 || d[0] != 0 || d[1] != 2 {
            return None;
        }
        let end = 4 + u16_at(d, 2)? as usize;
        let mut mask = u64::from(u32_at(d, 4)?);
        let mut pos: usize = 8;
        if mask64 {
            mask |= u64::from(u32_at(d, 8)?) << 32;
            pos = 12;
        }
        let mut values = HashMap::new();
        for &(bit, name, size) in fields {
            if mask & (1 << bit) == 0 {
                continue;
            }
            pos = pos.next_multiple_of(size);
            let v = match size {
                1 => u32::from(*d.get(pos)?),
                2 => u32::from(u16_at(d, pos)?),
                _ => u32_at(d, pos)?,
            };
            values.insert(name, v);
            pos += size;
        }
        pos = pos.next_multiple_of(4);
        let mut strings = HashMap::new();
        let mut size = None;
        for &(bit, name) in extra {
            if mask & (1 << bit) == 0 {
                continue;
            }
            if name == SIZE {
                size = Some((u32_at(d, pos)? as i32, u32_at(d, pos + 4)? as i32));
                pos += 8;
                continue;
            }
            let len = (values.get(name).copied().unwrap_or(0) & 0x7fff_ffff) as usize;
            strings.insert(name, d.get(pos..pos.checked_add(len)?)?.to_vec());
            pos = (pos + len).next_multiple_of(4);
        }
        // Read whole, or not at all (a property the fields do not know).
        if pos != end {
            return None;
        }
        Some(Stream {
            class,
            head,
            mask,
            values,
            strings,
            size,
            rest: d[end..].to_vec(),
        })
    }

    /// The data written again: the CLSID, the record, what followed it.
    pub(crate) fn write(&self) -> Vec<u8> {
        let (fields, mask64, extra) = spec(self.class).unwrap_or((&[], false, &[]));
        let mut r: Vec<u8> = vec![0, 2, 0, 0];
        r.extend((self.mask as u32).to_le_bytes());
        if mask64 {
            r.extend(((self.mask >> 32) as u32).to_le_bytes());
        }
        let pad = |r: &mut Vec<u8>, n: usize| {
            while r.len() % n != 0 {
                r.push(0);
            }
        };
        for &(bit, name, size) in fields {
            if self.mask & (1 << bit) == 0 {
                continue;
            }
            pad(&mut r, size);
            let v = self.values.get(name).copied().unwrap_or(0);
            r.extend(&v.to_le_bytes()[..size]);
        }
        pad(&mut r, 4);
        for &(bit, name) in extra {
            if self.mask & (1 << bit) == 0 {
                continue;
            }
            if name == SIZE {
                let (w, h) = self.size.unwrap_or((0, 0));
                r.extend(w.to_le_bytes());
                r.extend(h.to_le_bytes());
                continue;
            }
            r.extend(self.strings.get(name).map_or(&[][..], Vec::as_slice));
            pad(&mut r, 4);
        }
        let cb = (r.len() - 4) as u16;
        r[2..4].copy_from_slice(&cb.to_le_bytes());
        let mut out = self.head.clone();
        out.extend(r);
        out.extend(&self.rest);
        out
    }

    fn bit_of(&self, name: &str) -> Option<u8> {
        let (fields, _, extra) = spec(self.class)?;
        fields
            .iter()
            .find(|f| f.1 == name)
            .map(|f| f.0)
            .or_else(|| extra.iter().find(|e| e.1 == name).map(|e| e.0))
    }

    /// A number of the `DataBlock` (signed).
    pub(crate) fn number(&self, name: &str) -> Option<i64> {
        self.values.get(name).map(|v| i64::from(*v as i32))
    }

    pub(crate) fn set_number(&mut self, name: &str, v: i64) {
        if let Some(bit) = self.bit_of(name) {
            self.mask |= 1 << bit;
            let key = self.key(name);
            self.values.insert(key, v as i32 as u32);
        }
    }

    /// The `&'static` name of a field.
    fn key(&self, name: &str) -> &'static str {
        let (fields, _, extra) = spec(self.class).unwrap_or((&[], false, &[]));
        fields
            .iter()
            .map(|f| f.1)
            .chain(extra.iter().map(|e| e.1))
            .find(|n| *n == name)
            .unwrap_or("")
    }

    /// A string of the `ExtraDataBlock` (`Caption`, `Value`).
    #[cfg_attr(not(test), allow(dead_code))]
    pub(crate) fn string(&self, name: &str) -> Option<String> {
        let raw = self.strings.get(name)?;
        let packed = self.values.get(name).copied().unwrap_or(0);
        Some(if packed & 0x8000_0000 != 0 {
            raw.iter().map(|b| char::from(*b)).collect()
        } else {
            let units: Vec<u16> = raw
                .chunks_exact(2)
                .map(|c| u16::from_le_bytes([c[0], c[1]]))
                .collect();
            String::from_utf16_lossy(&units)
        })
    }

    /// A string set: stored compressed (a byte a character) when it is ASCII, else as UTF-16.
    pub(crate) fn set_string(&mut self, name: &str, s: &str) {
        let Some(bit) = self.bit_of(name) else {
            return;
        };
        let key = self.key(name);
        let (raw, flag): (Vec<u8>, u32) = if s.is_ascii() {
            (s.as_bytes().to_vec(), 0x8000_0000)
        } else {
            (s.encode_utf16().flat_map(u16::to_le_bytes).collect(), 0)
        };
        self.mask |= 1 << bit;
        self.values.insert(key, raw.len() as u32 | flag);
        self.strings.insert(key, raw);
    }

    /// The size (HIMETRIC: hundredths of a millimeter).
    #[cfg_attr(not(test), allow(dead_code))]
    pub(crate) fn size(&self) -> Option<(i32, i32)> {
        self.size
    }

    /// The size set.
    pub(crate) fn set_size(&mut self, w: i32, h: i32) {
        if let Some(bit) = self.bit_of(SIZE) {
            self.mask |= 1 << bit;
            self.size = Some((w, h));
        }
    }
}

/// The mask bits of a class's pictures (`Picture`, `MouseIcon`): its `StreamData` has them
/// before its font.
fn picture_bits(class: &str) -> &'static [u8] {
    match class {
        "CommandButton" => &[7, 10],
        "Label" => &[10, 12],
        "MorphData" => &[27, 28],
        _ => &[],
    }
}

/// `VariousPropertyBits` (MS-OFORMS): 有効 (`Enabled`), 背景が不透明 (`BackStyle`), 折り返し
/// (`WordWrap`), 自動サイズ (`AutoSize`).
pub(crate) const ENABLED: i64 = 1 << 1;
pub(crate) const BACK_STYLE: i64 = 1 << 3;
pub(crate) const WORD_WRAP: i64 = 1 << 23;
pub(crate) const AUTO_SIZE: i64 = 1 << 28;

/// `FontEffects` of a font (MS-OFORMS `TextProps`): 太字, 斜体, 下線, 取り消し線.
pub(crate) const BOLD: i64 = 1;
pub(crate) const ITALIC: i64 = 2;
pub(crate) const UNDERLINE: i64 = 4;
pub(crate) const STRIKE: i64 = 8;

/// The `VariousPropertyBits` a class has when its record leaves them out (MS-OFORMS).
fn default_bits(class: &str) -> i64 {
    match class {
        "Label" => 0x0080_001b,
        "MorphData" => 0x2c80_481b,
        _ => 0x1b,
    }
}

/// The colors (OLE_COLOR) a Microsoft Forms type has as Excel puts it on a sheet: (ForeColor,
/// BackColor). A text, list or combo box: the window's; the others: the button face's.
pub(crate) fn default_colors(forms: &str) -> (u32, u32) {
    match forms {
        "TextBox" | "ListBox" | "ComboBox" => (0x8000_0008, 0x8000_0005),
        _ => (0x8000_0012, 0x8000_000F),
    }
}

/// An OLE_COLOR of `#RRGGBB` (`0x00BBGGRR`).
pub(crate) fn ole_of(hex: &str) -> Option<u32> {
    let h = hex.trim().trim_start_matches('#');
    if h.len() != 6 {
        return None;
    }
    let v = u32::from_str_radix(h, 16).ok()?;
    Some(((v & 0xff) << 16) | (v & 0xff00) | ((v >> 16) & 0xff))
}

impl Stream {
    /// Where the font starts in what follows the record (after its pictures).
    fn font_at(&self) -> Option<usize> {
        let mut pos = 0usize;
        for bit in picture_bits(self.class) {
            if self.mask & (1 << bit) == 0 {
                continue;
            }
            let size = u32_at(&self.rest, pos + 20)? as usize;
            pos = pos.checked_add(24 + size)?;
            if pos > self.rest.len() {
                return None;
            }
        }
        Some(pos)
    }

    /// The font (`TextProps`) of a control that has one.
    pub(crate) fn font(&self) -> Option<Stream> {
        Stream::read("TextProps", &self.rest[self.font_at()?..])
    }

    /// The font changed by `f` (one made as a new control's when it has none). False for a
    /// class without a font or one whose font is not read.
    pub(crate) fn change_font(&mut self, f: impl FnOnce(&mut Stream)) -> bool {
        if !matches!(self.class, "CommandButton" | "Label" | "MorphData") {
            return false;
        }
        let Some(at) = self.font_at() else {
            return false;
        };
        let tail = &self.rest[at..];
        let (mut font, after) = if tail.is_empty() {
            let mut font = Stream {
                class: "TextProps",
                head: Vec::new(),
                mask: 0,
                values: HashMap::new(),
                strings: HashMap::new(),
                size: None,
                rest: Vec::new(),
            };
            let (name, height, charset, pitch) = FRESH_FONT;
            font.set_string("FontName", name);
            font.set_number("FontHeight", height);
            font.set_number("FontCharSet", charset);
            font.set_number("FontPitchAndFamily", pitch);
            (font, Vec::new())
        } else {
            let Some(mut font) = Stream::read("TextProps", tail) else {
                return false;
            };
            let after = std::mem::take(&mut font.rest);
            (font, after)
        };
        f(&mut font);
        let mut rest = self.rest[..at].to_vec();
        rest.extend(font.write());
        rest.extend(after);
        self.rest = rest;
        true
    }

    /// `VariousPropertyBits` (the class's default when the record leaves them out).
    pub(crate) fn bits(&self) -> i64 {
        self.number("VariousPropertyBits")
            .map_or_else(|| default_bits(self.class), |b| b & 0xffff_ffff)
    }

    /// A bit of `VariousPropertyBits` set or cleared.
    pub(crate) fn set_bit(&mut self, bit: i64, on: bool) {
        let b = self.bits();
        self.set_number("VariousPropertyBits", if on { b | bit } else { b & !bit });
    }

    /// A number stored as it is (unsigned: a color).
    pub(crate) fn set_raw(&mut self, name: &str, v: u32) {
        self.set_number(name, i64::from(v as i32));
    }
}

fn spec_class(class: &str) -> Option<&'static str> {
    [
        "CommandButton",
        "Label",
        "MorphData",
        "SpinButton",
        "ScrollBar",
        "TextProps",
    ]
    .into_iter()
    .find(|c| *c == class)
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;

    /// A `MorphData` stream: a check box with a caption and a value.
    pub(crate) fn check_box(caption: &str, value: &str) -> Vec<u8> {
        let mut s = Stream {
            class: "MorphData",
            head: Vec::new(),
            mask: 1 << 6,
            values: HashMap::from([("DisplayStyle", 4)]),
            strings: HashMap::new(),
            size: None,
            rest: vec![9, 9, 9],
        };
        s.set_size(2000, 500);
        s.set_string("Caption", caption);
        s.set_string("Value", value);
        s.write()
    }

    #[test]
    fn the_data_of_a_new_control_is_read_back() {
        let data = fresh("CommandButton", Some("実行"), None, (2540, 847)).unwrap();
        // The CLSID as it is stored.
        assert_eq!(
            data[..16],
            [
                0x40, 0x32, 0x05, 0xD7, 0x69, 0xCE, 0xCD, 0x11, 0xA7, 0x77, 0x00, 0xDD, 0x01, 0x14,
                0x3C, 0x57
            ]
        );
        let s = Stream::read("CommandButton", &data).unwrap();
        assert_eq!(s.write(), data);
        assert_eq!(s.string("Caption").as_deref(), Some("実行"));
        assert_eq!(s.number("VariousPropertyBits"), Some(0x1b));
        assert_eq!(s.size(), Some((2540, 847)));
        // The font after the record.
        let font = Stream::read("TextProps", &s.rest).unwrap();
        assert_eq!(font.string("FontName").as_deref(), Some("MS UI Gothic"));
        assert_eq!(
            (font.number("FontHeight"), font.number("ParagraphAlign")),
            (Some(180), Some(3))
        );
        // A combo box: its display style and settings; a scroll bar: no font.
        let s = Stream::read(
            "MorphData",
            &fresh("ComboBox", None, Some("A"), (1, 1)).unwrap(),
        )
        .unwrap();
        assert_eq!(
            (s.number("DisplayStyle"), s.number("ShowDropButtonWhen")),
            (Some(3), Some(2))
        );
        assert_eq!(s.string("Value").as_deref(), Some("A"));
        let s = Stream::read(
            "ScrollBar",
            &fresh("ScrollBar", None, None, (1, 1)).unwrap(),
        )
        .unwrap();
        assert!(s.rest.is_empty());
        assert!(fresh("Image", None, None, (1, 1)).is_none());
    }

    #[test]
    fn a_stream_is_read_changed_and_written_again() {
        let data = check_box("確認", "0");
        let mut s = Stream::read("MorphData", &data).unwrap();
        assert_eq!(s.string("Caption").as_deref(), Some("確認"));
        assert_eq!(s.string("Value").as_deref(), Some("0"));
        assert_eq!(s.number("DisplayStyle"), Some(4));
        // Written again unchanged: the same bytes.
        assert_eq!(s.write(), data);
        s.set_string("Caption", "OK");
        s.set_string("Value", "1");
        let again = Stream::read("MorphData", &s.write()).unwrap();
        assert_eq!(again.string("Caption").as_deref(), Some("OK"));
        assert_eq!(again.string("Value").as_deref(), Some("1"));
        assert_eq!(again.size, Some((2000, 500)));
        assert_eq!(again.rest, [9, 9, 9]);
        // umya-spreadsheet reads what is written.
        let mut ax = umya_spreadsheet::drawingml::controls::ActiveX {
            class: "MorphData",
            ..Default::default()
        };
        umya_spreadsheet::drawingml::controls::read_activex_stream(&mut ax, &s.write());
        assert_eq!(
            (ax.caption.as_deref(), ax.value.as_deref()),
            (Some("OK"), Some("1"))
        );
    }
}
