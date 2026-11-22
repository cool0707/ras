//! The persisted data of the ActiveX controls of a sheet (MS-OFORMS: the MSForms controls
//! CommandButton, Label, TextBox, ListBox, ComboBox, CheckBox, OptionButton, ToggleButton,
//! SpinButton and ScrollBar): read for the grid (`view::controls`), and changed and written anew
//! for the edits of the grid (`patch::objects`).
//!
//! A control's stream is a record (a mask of the properties it holds, their values, its strings
//! and its size), its pictures, and for the controls with a text the font (`TextProps`). The
//! layouts are those of [MS-OFORMS] 2.2 (as `crates/ras-form` writes them for user forms). A
//! record read is written back as it was but for what was changed: the pictures and whatever
//! follows the font are kept as bytes.

use std::collections::HashMap;

/// How the record of a class is laid out.
pub(crate) struct Layout {
    /// The fields of the DataBlock: (mask bit, name, bytes).
    fields: &'static [(u8, &'static str, usize)],
    /// The mask has 64 bits.
    mask64: bool,
    /// The ExtraDataBlock in its order: the strings by the field that holds their length, and
    /// the size (`SIZE`).
    extra: &'static [(u8, &'static str)],
    /// The mask bits of the pictures that follow the record.
    pictures: &'static [u8],
    /// A `TextProps` record follows the pictures.
    text: bool,
}

const SIZE: &str = "#Size";

static COMMAND_BUTTON: Layout = Layout {
    fields: &[
        (0, "ForeColor", 4),
        (1, "BackColor", 4),
        (2, "VariousPropertyBits", 4),
        (3, "Caption", 4),
        (4, "PicturePosition", 4),
        (6, "MousePointer", 1),
        (7, "Picture", 2),
        (8, "Accelerator", 2),
        (10, "MouseIcon", 2),
    ],
    mask64: false,
    extra: &[(3, "Caption"), (5, SIZE)],
    pictures: &[7, 10],
    text: true,
};

static LABEL: Layout = Layout {
    fields: &[
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
    ],
    mask64: false,
    extra: &[(3, "Caption"), (5, SIZE)],
    pictures: &[10, 12],
    text: true,
};

static MORPH_DATA: Layout = Layout {
    fields: &[
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
    ],
    mask64: true,
    extra: &[(8, SIZE), (22, "Value"), (23, "Caption"), (32, "GroupName")],
    pictures: &[27, 28],
    text: true,
};

static SPIN_BUTTON: Layout = Layout {
    fields: &[
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
        (12, "Delay", 4),
        (13, "MouseIcon", 2),
        (14, "MousePointer", 1),
    ],
    mask64: false,
    extra: &[(3, SIZE)],
    pictures: &[13],
    text: false,
};

static SCROLL_BAR: Layout = Layout {
    fields: &[
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
        (14, "ProportionalThumb", 2),
        (15, "Delay", 4),
        (16, "MouseIcon", 2),
    ],
    mask64: false,
    extra: &[(3, SIZE)],
    pictures: &[16],
    text: false,
};

static TEXT_PROPS: Layout = Layout {
    fields: &[
        (0, "FontName", 4),
        (1, "FontEffects", 4),
        (2, "FontHeight", 4),
        (4, "FontCharSet", 1),
        (5, "FontPitchAndFamily", 1),
        (6, "ParagraphAlign", 1),
        (7, "FontWeight", 2),
    ],
    mask64: false,
    extra: &[(0, "FontName")],
    pictures: &[],
    text: false,
};

/// The classes of MSForms controls a sheet holds: (class id, name as VBA names its type, the
/// record's name, MorphData's DisplayStyle).
pub(crate) const CLASSES: &[(&str, &str, &str, u8)] = &[
    (
        "D7053240-CE69-11CD-A777-00DD01143C57",
        "CommandButton",
        "CommandButton",
        0,
    ),
    ("978C9E23-D4B0-11CE-BF2D-00AA003F40D0", "Label", "Label", 0),
    (
        "8BD21D10-EC42-11CE-9E0D-00AA006002F3",
        "TextBox",
        "MorphData",
        1,
    ),
    (
        "8BD21D20-EC42-11CE-9E0D-00AA006002F3",
        "ListBox",
        "MorphData",
        2,
    ),
    (
        "8BD21D30-EC42-11CE-9E0D-00AA006002F3",
        "ComboBox",
        "MorphData",
        3,
    ),
    (
        "8BD21D40-EC42-11CE-9E0D-00AA006002F3",
        "CheckBox",
        "MorphData",
        4,
    ),
    (
        "8BD21D50-EC42-11CE-9E0D-00AA006002F3",
        "OptionButton",
        "MorphData",
        5,
    ),
    (
        "8BD21D60-EC42-11CE-9E0D-00AA006002F3",
        "ToggleButton",
        "MorphData",
        6,
    ),
    (
        "79176FB0-B7F2-11CE-97EF-00AA006D2776",
        "SpinButton",
        "SpinButton",
        0,
    ),
    (
        "DFD181E0-5E2F-11CE-A449-00AA004A803D",
        "ScrollBar",
        "ScrollBar",
        0,
    ),
];

/// The class of a class id (`{D7053240-...}`): (name, record, DisplayStyle).
pub(crate) fn class_of(clsid: &str) -> Option<(&'static str, &'static str, u8)> {
    let id = clsid.trim_matches(['{', '}']);
    CLASSES
        .iter()
        .find(|c| c.0.eq_ignore_ascii_case(id))
        .map(|c| (c.1, c.2, c.3))
}

/// The class id of a class by its name (`CommandButton`; case does not matter).
pub(crate) fn clsid_of(
    name: &str,
) -> Option<&'static (&'static str, &'static str, &'static str, u8)> {
    CLASSES.iter().find(|c| c.1.eq_ignore_ascii_case(name))
}

/// A class id as the bytes that start the persisted stream of a control.
pub(crate) fn clsid_bytes(clsid: &str) -> Option<[u8; 16]> {
    let hex: Vec<u8> = clsid
        .trim_matches(['{', '}'])
        .bytes()
        .filter(|b| *b != b'-')
        .collect();
    if hex.len() != 32 {
        return None;
    }
    let mut raw = [0u8; 16];
    for (i, pair) in hex.chunks(2).enumerate() {
        raw[i] = u8::from_str_radix(std::str::from_utf8(pair).ok()?, 16).ok()?;
    }
    // The first three groups are stored little-endian.
    raw[0..4].reverse();
    raw[4..6].reverse();
    raw[6..8].reverse();
    Some(raw)
}

fn layout_of(record: &str) -> Option<&'static Layout> {
    Some(match record {
        "CommandButton" => &COMMAND_BUTTON,
        "Label" => &LABEL,
        "MorphData" => &MORPH_DATA,
        "SpinButton" => &SPIN_BUTTON,
        "ScrollBar" => &SCROLL_BAR,
        _ => return None,
    })
}

/// A record: the properties it holds, and what follows it in the stream.
#[derive(Clone)]
pub(crate) struct Record {
    layout: &'static Layout,
    mask: u64,
    values: HashMap<&'static str, u32>,
    strings: HashMap<&'static str, String>,
    /// Width and height (HIMETRIC).
    size: (i32, i32),
    /// The pictures after the record, as they are.
    pictures: Vec<u8>,
    /// The font of the control's text.
    pub text: Option<Box<Record>>,
    /// What follows (the column widths of a list).
    tail: Vec<u8>,
}

fn u16_at(data: &[u8], i: usize) -> Option<u16> {
    data.get(i..i + 2).map(|b| u16::from_le_bytes([b[0], b[1]]))
}

fn u32_at(data: &[u8], i: usize) -> Option<u32> {
    data.get(i..i + 4)
        .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
}

fn pad(out: &mut Vec<u8>, size: usize) {
    while out.len() % size != 0 {
        out.push(0);
    }
}

impl Record {
    /// An empty record of a class (`CommandButton`, `MorphData`, ...), with an empty font where
    /// the class has one.
    pub(crate) fn new(record: &str) -> Option<Record> {
        let layout = layout_of(record)?;
        let empty = |layout: &'static Layout| Record {
            layout,
            mask: 0,
            values: HashMap::new(),
            strings: HashMap::new(),
            size: (0, 0),
            pictures: Vec::new(),
            text: None,
            tail: Vec::new(),
        };
        let mut r = empty(layout);
        if layout.text {
            r.text = Some(Box::new(empty(&TEXT_PROPS)));
        }
        // (Excel always sets this reserved bit of a MorphData.)
        if record == "MorphData" {
            r.mask |= 1 << 31;
        }
        Some(r)
    }

    /// Reads the stream of a control of a class. None when it is not laid out as the class's
    /// record is (another version, a damaged stream).
    pub(crate) fn parse(record: &str, data: &[u8]) -> Option<Record> {
        let layout = layout_of(record)?;
        let (mut r, end) = Record::read(layout, data)?;
        let count = layout
            .pictures
            .iter()
            .filter(|b| r.mask & (1u64 << **b) != 0)
            .count();
        // The pictures (GuidAndPicture: a class id, a stamp, the size of the data, the data).
        let mut at = end;
        for _ in 0..count {
            let size = u32_at(data, at + 20)? as usize;
            at = at.checked_add(24 + size)?;
            if at > data.len() {
                return None;
            }
        }
        r.pictures = data[end..at].to_vec();
        if layout.text {
            let (text, used) = Record::read(&TEXT_PROPS, &data[at..])?;
            r.text = Some(Box::new(text));
            at += used;
        }
        r.tail = data[at..].to_vec();
        Some(r)
    }

    /// One record of `layout` at the start of `data`, and where it ends.
    fn read(layout: &'static Layout, data: &[u8]) -> Option<(Record, usize)> {
        if data.len() < 8 || data[0] != 0 || data[1] != 2 {
            return None;
        }
        let end = 4 + u16_at(data, 2)? as usize;
        if end > data.len() {
            return None;
        }
        let mut mask = u32_at(data, 4)? as u64;
        let mut pos: usize = 8;
        if layout.mask64 {
            mask |= (u32_at(data, 8)? as u64) << 32;
            pos = 12;
        }
        let mut values = HashMap::new();
        for &(bit, name, size) in layout.fields {
            if mask & (1u64 << bit) == 0 {
                continue;
            }
            pos = pos.next_multiple_of(size);
            let v = match size {
                1 => *data.get(pos)? as u32,
                2 => u16_at(data, pos)? as u32,
                _ => u32_at(data, pos)?,
            };
            values.insert(name, v);
            pos += size;
        }
        pos = pos.next_multiple_of(4);
        let mut strings = HashMap::new();
        let mut size = (0, 0);
        for &(bit, name) in layout.extra {
            if mask & (1u64 << bit) == 0 {
                continue;
            }
            if name == SIZE {
                size = (u32_at(data, pos)? as i32, u32_at(data, pos + 4)? as i32);
                pos += 8;
                continue;
            }
            let packed = values.get(name).copied().unwrap_or(0);
            let len = (packed & 0x7fff_ffff) as usize;
            let raw = data.get(pos..pos.checked_add(len)?)?;
            let text = if packed & 0x8000_0000 != 0 {
                raw.iter().map(|b| *b as char).collect()
            } else {
                let units: Vec<u16> = raw
                    .chunks_exact(2)
                    .map(|c| u16::from_le_bytes([c[0], c[1]]))
                    .collect();
                String::from_utf16_lossy(&units)
            };
            strings.insert(name, text);
            pos = (pos + len).next_multiple_of(4);
        }
        if pos > end {
            return None;
        }
        Some((
            Record {
                layout,
                mask,
                values,
                strings,
                size,
                pictures: Vec::new(),
                text: None,
                tail: Vec::new(),
            },
            end,
        ))
    }

    /// The stream of the control.
    pub(crate) fn encode(&self) -> Vec<u8> {
        let layout = self.layout;
        // A string as it is stored: a byte for each character where that holds them all.
        let stored = |text: &str| -> (Vec<u8>, u32) {
            let units: Vec<u16> = text.encode_utf16().collect();
            if units.iter().all(|u| *u <= 0xff) {
                let raw: Vec<u8> = units.iter().map(|u| *u as u8).collect();
                let n = raw.len() as u32;
                (raw, if n == 0 { 0 } else { n | 0x8000_0000 })
            } else {
                let raw: Vec<u8> = units.iter().flat_map(|u| u.to_le_bytes()).collect();
                let n = raw.len() as u32;
                (raw, n)
            }
        };
        let mut out = vec![0u8, 2, 0, 0];
        out.extend((self.mask as u32).to_le_bytes());
        if layout.mask64 {
            out.extend(((self.mask >> 32) as u32).to_le_bytes());
        }
        for &(bit, name, size) in layout.fields {
            if self.mask & (1u64 << bit) == 0 {
                continue;
            }
            pad(&mut out, size);
            let value = match self.strings.get(name) {
                Some(text) => stored(text).1,
                None => self.values.get(name).copied().unwrap_or(0),
            };
            match size {
                1 => out.push(value as u8),
                2 => out.extend((value as u16).to_le_bytes()),
                _ => out.extend(value.to_le_bytes()),
            }
        }
        pad(&mut out, 4);
        for &(bit, name) in layout.extra {
            if self.mask & (1u64 << bit) == 0 {
                continue;
            }
            if name == SIZE {
                out.extend(self.size.0.to_le_bytes());
                out.extend(self.size.1.to_le_bytes());
            } else if let Some(text) = self.strings.get(name) {
                out.extend(stored(text).0);
                pad(&mut out, 4);
            }
        }
        let cb = (out.len() - 4).min(u16::MAX as usize) as u16;
        out[2..4].copy_from_slice(&cb.to_le_bytes());
        out.extend(&self.pictures);
        if let Some(text) = &self.text {
            out.extend(text.encode());
        }
        out.extend(&self.tail);
        out
    }

    fn field(&self, name: &str) -> Option<&'static (u8, &'static str, usize)> {
        self.layout.fields.iter().find(|f| f.1 == name)
    }

    /// The class's record has the property.
    pub(crate) fn has(&self, name: &str) -> bool {
        self.field(name).is_some()
    }

    /// The value of a property the record holds.
    pub(crate) fn get(&self, name: &str) -> Option<u32> {
        self.values.get(name).copied()
    }

    pub(crate) fn signed(&self, name: &str) -> Option<i64> {
        self.get(name).map(|v| v as i32 as i64)
    }

    pub(crate) fn string(&self, name: &str) -> Option<&str> {
        self.strings.get(name).map(String::as_str)
    }

    /// Sets a number. False when the class has no such property.
    pub(crate) fn set(&mut self, name: &str, value: u32) -> bool {
        let Some(&(bit, name, _)) = self.field(name) else {
            return false;
        };
        self.mask |= 1u64 << bit;
        self.values.insert(name, value);
        true
    }

    /// Sets a string. False when the class has no such property.
    pub(crate) fn set_string(&mut self, name: &str, text: &str) -> bool {
        let Some(&(bit, name, _)) = self.field(name) else {
            return false;
        };
        if !self.layout.extra.iter().any(|x| x.1 == name) {
            return false;
        }
        self.mask |= 1u64 << bit;
        self.strings.insert(name, text.to_string());
        true
    }

    /// The property takes its default: the record no longer holds it.
    pub(crate) fn clear(&mut self, name: &str) {
        if let Some(&(bit, name, _)) = self.field(name) {
            self.mask &= !(1u64 << bit);
            self.values.remove(name);
            self.strings.remove(name);
        }
    }

    #[cfg(test)]
    pub(crate) fn size(&self) -> (i32, i32) {
        self.size
    }

    pub(crate) fn set_size(&mut self, size: (i32, i32)) {
        if let Some(&(bit, _)) = self.layout.extra.iter().find(|x| x.1 == SIZE) {
            self.mask |= 1u64 << bit;
            self.size = size;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn class_ids() {
        assert_eq!(
            class_of("{d7053240-ce69-11cd-a777-00dd01143c57}"),
            Some(("CommandButton", "CommandButton", 0))
        );
        assert_eq!(clsid_of("checkbox").map(|c| c.3), Some(4));
        assert_eq!(
            clsid_bytes("{D7053240-CE69-11CD-A777-00DD01143C57}").unwrap(),
            [
                0x40, 0x32, 0x05, 0xD7, 0x69, 0xCE, 0xCD, 0x11, 0xA7, 0x77, 0x00, 0xDD, 0x01, 0x14,
                0x3C, 0x57
            ]
        );
    }

    /// A CommandButton as Excel writes a new one: its caption, its size and the font.
    #[test]
    fn a_record_written_and_read_back() {
        let mut r = Record::new("CommandButton").unwrap();
        r.set_string("Caption", "CommandButton1");
        r.set_size((2540, 847));
        let font = r.text.as_mut().unwrap();
        font.set_string("FontName", "Tahoma");
        font.set("FontHeight", 165);
        font.set("FontCharSet", 0);
        font.set("FontPitchAndFamily", 2);
        let data = r.encode();
        // Caption and Size; the caption a byte a character; 32 bytes after the header.
        assert_eq!(
            &data[..12],
            [0, 2, 0x20, 0, 0x28, 0, 0, 0, 0x0E, 0, 0, 0x80]
        );
        assert_eq!(&data[12..26], b"CommandButton1");
        assert_eq!(&data[36..48], [0, 2, 0x18, 0, 0x35, 0, 0, 0, 6, 0, 0, 0x80]);
        let back = Record::parse("CommandButton", &data).unwrap();
        assert_eq!(back.string("Caption"), Some("CommandButton1"));
        assert_eq!(back.size(), (2540, 847));
        assert_eq!(
            back.text.as_ref().unwrap().string("FontName"),
            Some("Tahoma")
        );
        assert_eq!(back.encode(), data);
    }

    /// What a record holds besides what is changed stays: a picture, the bits of the mask
    /// that carry no value, what follows the font.
    #[test]
    fn a_changed_record_keeps_the_rest() {
        let mut r = Record::new("MorphData").unwrap();
        r.set("DisplayStyle", 4);
        r.set("Picture", 0xffff);
        r.set_string("Caption", "チェック");
        r.set_string("Value", "1");
        r.set_size((1000, 500));
        r.pictures = {
            let mut p = vec![7u8; 20];
            p.extend(3u32.to_le_bytes());
            p.extend([1, 2, 3]);
            p
        };
        r.tail = vec![9, 9];
        let data = r.encode();
        let mut back = Record::parse("MorphData", &data).unwrap();
        assert_eq!(back.string("Caption"), Some("チェック"));
        assert_eq!(back.get("DisplayStyle"), Some(4));
        assert_eq!(back.mask & (1 << 31), 1 << 31);
        back.set_string("Caption", "A longer caption");
        back.set_string("Value", "0");
        let again = Record::parse("MorphData", &back.encode()).unwrap();
        assert_eq!(again.string("Caption"), Some("A longer caption"));
        assert_eq!(again.string("Value"), Some("0"));
        assert_eq!(again.pictures, r.pictures);
        assert_eq!(again.tail, [9, 9]);
        assert!(Record::parse("MorphData", &data[..10]).is_none());
        assert!(!back.set("Min", 1));
    }
}
