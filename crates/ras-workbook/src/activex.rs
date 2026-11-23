//! The persisted data of an ActiveX control of Microsoft Forms 2.0 on a sheet (MS-OFORMS):
//! read, changed and written again. Excel keeps it in `xl/activeX/activeX<n>.bin` (the control's
//! CLSID, then its stream: `ax:persistence="persistStream"`).
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
    pub(crate) fn set_size(&mut self, w: i32, h: i32) {
        if let Some(bit) = self.bit_of(SIZE) {
            self.mask |= 1 << bit;
            self.size = Some((w, h));
        }
    }
}

fn spec_class(class: &str) -> Option<&'static str> {
    [
        "CommandButton",
        "Label",
        "MorphData",
        "SpinButton",
        "ScrollBar",
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
