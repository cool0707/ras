//! Reading an MS-OFORMS designer storage back into a model.
//!
//! Port of the read path of xlflow's `internal/vba/userforms/oforms`
//! (`reader.go`, `records.go`, the parse side of `tabstrip.go`,
//! `multipage.go` and `font.go`). Copyright (c) 2026 harumiWeb, MIT License;
//! see `NOTICE.md`.
//!
//! Only what the projection to a FormSpec needs is kept; the bytes are not
//! retained for writing back.

use std::collections::{BTreeMap, HashMap};

use crate::oforms::{
    array_size_field, is_container_class, spec_for_class, ExtraKind, FieldKind, RecordSpec, Tab,
    TabStripState, FORM, MAX_NESTING_DEPTH, MULTI_PAGE_PROPERTIES, PAGE_PROPERTIES, TEXT_PROPS,
};

const MAX_STREAM_SIZE: usize = 64 << 20;
const MAX_SITES: usize = 65_535;
const FORM_MASK_MOUSE_ICON: u64 = 1 << 15;
const FORM_MASK_FONT: u64 = 1 << 20;
const FORM_MASK_PICTURE: u64 = 1 << 21;
const GUID_STD_FONT: u32 = 0x0be3_5203;
const GUID_TEXT_PROPS: u32 = 0xafc2_0920;

/// Why a designer storage could not be read.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ReadError {
    /// Storage path below the form storage (`""` for the form itself).
    pub storage: String,
    /// Stream concerned (may be empty).
    pub stream: String,
    /// What is wrong.
    pub detail: String,
}

impl std::fmt::Display for ReadError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let place = match (self.storage.is_empty(), self.stream.is_empty()) {
            (true, true) => String::new(),
            (true, false) => format!("{}: ", printable(&self.stream)),
            (false, true) => format!("{}: ", self.storage),
            (false, false) => format!("{}/{}: ", self.storage, printable(&self.stream)),
        };
        write!(
            f,
            "デザイナストレージを読めません（{place}{}）",
            self.detail
        )
    }
}

impl std::error::Error for ReadError {}

fn printable(name: &str) -> String {
    name.chars().filter(|c| !c.is_control()).collect()
}

type Result<T> = std::result::Result<T, String>;

struct Reader<'a> {
    data: &'a [u8],
    pos: usize,
    end: usize,
}

impl<'a> Reader<'a> {
    fn new(data: &'a [u8]) -> Reader<'a> {
        Reader {
            data,
            pos: 0,
            end: data.len(),
        }
    }

    fn remaining(&self) -> usize {
        self.end - self.pos
    }

    fn take(&mut self, n: usize) -> Result<&'a [u8]> {
        if n > self.remaining() {
            return Err(format!(
                "{} バイト目から {n} バイトを読もうとして、範囲（{} バイト）を超えました",
                self.pos, self.end
            ));
        }
        let bytes = &self.data[self.pos..self.pos + n];
        self.pos += n;
        Ok(bytes)
    }

    fn u8(&mut self) -> Result<u8> {
        Ok(self.take(1)?[0])
    }

    fn u16(&mut self) -> Result<u16> {
        let b = self.take(2)?;
        Ok(u16::from_le_bytes([b[0], b[1]]))
    }

    fn u32(&mut self) -> Result<u32> {
        let b = self.take(4)?;
        Ok(u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
    }

    fn i32(&mut self) -> Result<i32> {
        Ok(self.u32()? as i32)
    }

    fn align(&mut self, base: usize, size: usize) -> Result<()> {
        let over = (self.pos - base) % size;
        if over != 0 {
            self.take(size - over)?;
        }
        Ok(())
    }

    fn integer(&mut self, size: usize, signed: bool) -> Result<i64> {
        Ok(match (size, signed) {
            (1, false) => self.u8()? as i64,
            (1, true) => self.u8()? as i8 as i64,
            (2, false) => self.u16()? as i64,
            (2, true) => self.u16()? as i16 as i64,
            (_, false) => self.u32()? as i64,
            (_, true) => self.u32()? as i32 as i64,
        })
    }

    /// A `GuidAndPicture` record, returned whole.
    fn guid_and_picture(&mut self) -> Result<Vec<u8>> {
        let start = self.pos;
        self.take(16)?;
        self.u32()?;
        let size = self.u32()? as usize;
        if size > self.remaining() {
            return Err(format!(
                "画像のサイズ {size} が残り {} バイトを超えています",
                self.remaining()
            ));
        }
        self.take(size)?;
        Ok(self.data[start..self.pos].to_vec())
    }
}

/// A string as stored: the count says whether one byte or two are used per
/// UTF-16 unit (independent of the project code page).
fn stored_text(raw: &[u8], compressed: bool) -> Result<String> {
    if compressed {
        return Ok(raw.iter().map(|b| *b as char).collect());
    }
    if raw.len() % 2 != 0 {
        return Err(format!("UTF-16 文字列の長さ {} が奇数です", raw.len()));
    }
    let units: Vec<u16> = raw
        .chunks_exact(2)
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .collect();
    Ok(String::from_utf16_lossy(&units))
}

fn stored_string(r: &mut Reader, packed: u32) -> Result<String> {
    let raw = r.take((packed & 0x7fff_ffff) as usize)?;
    stored_text(raw, packed & 0x8000_0000 != 0)
}

/// A decoded property record.
#[derive(Clone, Debug)]
pub(crate) struct Record {
    pub spec: &'static RecordSpec,
    pub mask: u64,
    /// Scalar fields and string lengths that are present, plus flag bits.
    pub values: HashMap<&'static str, i64>,
    pub strings: HashMap<&'static str, String>,
    pub sizes: HashMap<&'static str, (i32, i32)>,
    pub arrays: HashMap<&'static str, Vec<u8>>,
    /// `GuidAndPicture` records that follow the record.
    pub pictures: HashMap<&'static str, Vec<u8>>,
    pub text_props: Option<Box<Record>>,
    pub tail: Vec<u8>,
}

impl Record {
    pub(crate) fn value(&self, name: &str) -> Option<i64> {
        self.values.get(name).copied()
    }

    pub(crate) fn text(&self, name: &str) -> Option<&str> {
        self.strings.get(name).map(String::as_str)
    }
}

/// Parses one record that ends at `limit`; returns it and the bytes consumed.
fn parse_record(data: &[u8], spec: &'static RecordSpec, limit: usize) -> Result<(Record, usize)> {
    if limit > data.len() {
        return Err(format!(
            "レコードの境界 {limit} がデータ（{} バイト）を超えています",
            data.len()
        ));
    }
    let mut r = Reader::new(data);
    r.end = limit;
    let minor = r.u8()?;
    let major = r.u8()?;
    if minor != 0 || major != spec.major {
        return Err(format!(
            "{} のバージョン {major}.{minor} には対応していません",
            spec.type_name
        ));
    }
    let cb = r.u16()? as usize;
    let after_extra = 4 + cb;
    if after_extra > limit {
        return Err(format!(
            "{} の長さ {cb} が境界 {limit} を超えています",
            spec.type_name
        ));
    }
    let mut mask = r.u32()? as u64;
    if spec.mask64 {
        mask |= (r.u32()? as u64) << 32;
    }
    let mut record = Record {
        spec,
        mask,
        values: HashMap::new(),
        strings: HashMap::new(),
        sizes: HashMap::new(),
        arrays: HashMap::new(),
        pictures: HashMap::new(),
        text_props: None,
        tail: Vec::new(),
    };
    for (bit, name, value) in spec.flags {
        if mask & (1u64 << bit) != 0 {
            record.values.insert(name, *value);
        }
    }
    for field in spec.data {
        if mask & (1u64 << field.bit) == 0 {
            continue;
        }
        r.align(0, field.size)?;
        let value = r.integer(field.size, field.kind == FieldKind::Signed)?;
        record.values.insert(field.name, value);
    }
    r.align(0, 4)?;
    for extra in spec.extra {
        if mask & (1u64 << extra.bit) == 0 {
            continue;
        }
        match extra.kind {
            ExtraKind::Size | ExtraKind::Position => {
                let width = r.i32()?;
                let height = r.i32()?;
                record.sizes.insert(extra.name, (width, height));
            }
            ExtraKind::Str => {
                let packed = record.value(extra.name).unwrap_or(0) as u32;
                let text =
                    stored_string(&mut r, packed).map_err(|e| format!("{}: {e}", extra.name))?;
                record.strings.insert(extra.name, text);
                r.align(0, 4)?;
            }
            ExtraKind::Array => {
                let size = record.value(array_size_field(extra.name)).unwrap_or(0);
                if size < 0 {
                    return Err(format!("{} のサイズ {size} が不正です", extra.name));
                }
                record
                    .arrays
                    .insert(extra.name, r.take(size as usize)?.to_vec());
            }
        }
    }
    if r.pos != after_extra {
        return Err(format!(
            "{} の長さが合いません（読んだのは {} バイト、記録は {cb} バイト）",
            spec.type_name,
            r.pos - 4
        ));
    }
    if spec.stop_after_extra {
        return Ok((record, after_extra));
    }
    for (bit, name) in spec.stream {
        if mask & (1u64 << bit) == 0 {
            continue;
        }
        let picture = r.guid_and_picture().map_err(|e| format!("{name}: {e}"))?;
        record.pictures.insert(name, picture);
    }
    if spec.text_props {
        if r.remaining() < 4 {
            return Err(format!("{} に TextProps がありません", spec.type_name));
        }
        let nested = 4 + u16::from_le_bytes([data[r.pos + 2], data[r.pos + 3]]) as usize;
        if nested > r.remaining() {
            return Err(format!(
                "TextProps の長さ {nested} が残り {} バイトを超えています",
                r.remaining()
            ));
        }
        let (text_props, consumed) = parse_record(&data[r.pos..r.end], &TEXT_PROPS, nested)?;
        record.text_props = Some(Box::new(text_props));
        r.pos += consumed;
    }
    if r.pos < limit {
        if !spec.raw_tail {
            return Err(format!(
                "{} の後ろに余分な {} バイトがあります",
                spec.type_name,
                limit - r.pos
            ));
        }
        record.tail = data[r.pos..limit].to_vec();
        r.pos = limit;
    }
    Ok((record, r.pos))
}

/// An `OleSiteConcrete` record: where a control sits in its parent.
#[derive(Clone, Debug, Default)]
pub(crate) struct Site {
    pub values: HashMap<&'static str, i64>,
    pub strings: HashMap<&'static str, String>,
    pub position: Option<(i32, i32)>,
}

/// (bit, name, size, signed)
const SITE_FIELDS: &[(u8, &str, usize, bool)] = &[
    (0, "NameData", 4, false),
    (1, "TagData", 4, false),
    (2, "ID", 4, true),
    (3, "HelpContextID", 4, true),
    (4, "BitFlags", 4, false),
    (5, "ObjectStreamSize", 4, false),
    (6, "TabIndex", 2, true),
    (7, "ClsidCacheIndex", 2, false),
    (9, "GroupID", 2, false),
    (11, "ControlTipTextData", 4, false),
    (12, "RuntimeLicKeyData", 4, false),
    (13, "ControlSourceData", 4, false),
    (14, "RowSourceData", 4, false),
];

const SITE_STRINGS_BEFORE: &[(&str, &str)] = &[("NameData", "Name"), ("TagData", "Tag")];
const SITE_STRINGS_AFTER: &[(&str, &str)] = &[
    ("ControlTipTextData", "ControlTipText"),
    ("RuntimeLicKeyData", "RuntimeLicKey"),
    ("ControlSourceData", "ControlSource"),
    ("RowSourceData", "RowSource"),
];

fn parse_site(r: &mut Reader) -> Result<Site> {
    let start = r.pos;
    let version = r.u16()?;
    if version != 0 {
        return Err(format!(
            "OleSiteConcrete のバージョン {version} には対応していません"
        ));
    }
    let cb = r.u16()? as usize;
    let end = start + 4 + cb;
    if end > r.end {
        return Err(format!("サイトの長さ {cb} が境界を超えています"));
    }
    let mask = r.u32()?;
    let mut site = Site::default();
    for (bit, name, size, signed) in SITE_FIELDS {
        if mask & (1u32 << bit) == 0 {
            continue;
        }
        r.align(start, *size)?;
        let value = r.integer(*size, *signed)?;
        site.values.insert(name, value);
    }
    r.align(start, 4)?;
    let read_strings = |r: &mut Reader, site: &mut Site, names: &[(&'static str, &'static str)]| {
        for (length, name) in names {
            let Some(packed) = site.values.get(length).copied() else {
                continue;
            };
            let text = stored_string(r, packed as u32)?;
            site.strings.insert(name, text);
            r.align(start, 4)?;
        }
        Ok::<(), String>(())
    };
    read_strings(r, &mut site, SITE_STRINGS_BEFORE)?;
    if mask & (1 << 8) != 0 {
        let left = r.i32()?;
        let top = r.i32()?;
        site.position = Some((left, top));
    }
    read_strings(r, &mut site, SITE_STRINGS_AFTER)?;
    if r.pos != end {
        return Err(format!(
            "サイト {:?} の長さが合いません（{} バイト目まで読み、境界は {end}）",
            site.strings.get("Name").cloned().unwrap_or_default(),
            r.pos
        ));
    }
    Ok(site)
}

fn is_trailing_record(data: &[u8], offset: usize) -> bool {
    offset + 4 <= data.len()
        && data[offset] == 0
        && data[offset + 1] == 2
        && offset + 4 + u16::from_le_bytes([data[offset + 2], data[offset + 3]]) as usize
            == data.len()
}

/// `FormSiteData`: the class table (when stored), depths and types, sites.
fn parse_site_data(data: &[u8], start: usize, class_table_stored: bool) -> Result<Vec<Site>> {
    let mut r = Reader::new(data);
    r.pos = start;
    if class_table_stored {
        let count = r.u16()? as usize;
        if count > r.remaining() / 4 {
            return Err(format!(
                "クラス情報の数 {count} が残りのバイト数を超えています"
            ));
        }
        for _ in 0..count {
            let version = r.u16()?;
            if version != 0 {
                return Err(format!(
                    "SiteClassInfo のバージョン {version} には対応していません"
                ));
            }
            let cb = r.u16()? as usize;
            r.take(cb)?;
        }
    }
    let count = r.u32()? as usize;
    if count > MAX_SITES {
        return Err(format!(
            "サイト数 {count} が上限 {MAX_SITES} を超えています"
        ));
    }
    let count_of_bytes = r.u32()? as usize;
    if count_of_bytes > r.remaining() {
        return Err(format!(
            "CountOfBytes {count_of_bytes} が残り {} バイトを超えています",
            r.remaining()
        ));
    }
    let end = r.pos + count_of_bytes;
    if end != data.len() && !is_trailing_record(data, end) {
        return Err(format!(
            "CountOfBytes の終端 {end} がストリームの終端 {} と合いません",
            data.len()
        ));
    }
    let depth_start = r.pos;
    let mut seen = 0usize;
    while seen < count {
        let _depth = r.u8()?;
        let type_or_count = r.u8()?;
        let mut run = 1usize;
        if type_or_count & 0x80 != 0 {
            run = (type_or_count & 0x7f) as usize;
            if run == 0 {
                return Err("SiteDepthsAndTypes に長さ 0 の連続があります".into());
            }
            r.u8()?;
        }
        if seen + run > count {
            return Err(format!(
                "SiteDepthsAndTypes が {count} 個を超えるサイトを数えています"
            ));
        }
        seen += run;
    }
    r.align(depth_start, 4)?;
    let mut sites = Vec::with_capacity(count);
    for index in 0..count {
        sites.push(parse_site(&mut r).map_err(|e| format!("サイト {index}: {e}"))?);
    }
    if r.pos != end {
        return Err(format!(
            "サイトの終端 {} が CountOfBytes の境界 {end} と合いません",
            r.pos
        ));
    }
    Ok(sites)
}

fn read_font_blob(r: &mut Reader) -> Result<Vec<u8>> {
    let start = r.pos;
    let guid = r.take(16)?;
    match u32::from_le_bytes([guid[0], guid[1], guid[2], guid[3]]) {
        GUID_STD_FONT => {
            let version = r.u8()?;
            if version != 1 {
                return Err(format!(
                    "StdFont のバージョン {version} には対応していません"
                ));
            }
            r.take(9)?;
            let name_length = r.u8()? as usize;
            r.take(name_length)?;
        }
        GUID_TEXT_PROPS => {
            r.take(2)?;
            let cb = r.u16()? as usize;
            r.take(cb)?;
        }
        _ => return Err("フォントの GUID が不明です".into()),
    }
    Ok(r.data[start..r.pos].to_vec())
}

/// One parent-control storage: the form itself or a Frame, MultiPage, Page.
#[derive(Clone, Debug)]
pub(crate) struct Level {
    pub record: Record,
    /// The `GuidAndPicture` of the level's MouseIcon, if it has one.
    pub mouse_icon: Option<Vec<u8>>,
    pub font_raw: Vec<u8>,
    /// The `GuidAndPicture` of the level's Picture, if it has one.
    pub picture: Option<Vec<u8>>,
    /// Streams of the storage that are not part of the designer format.
    pub extra_streams: Vec<String>,
}

/// A font as far as it is stored; `None` members were omitted.
#[derive(Clone, Debug, Default, PartialEq)]
pub(crate) struct Font {
    pub name: Option<String>,
    pub size: Option<f64>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub strike_through: Option<bool>,
    pub charset: Option<i64>,
    pub pitch_and_family: Option<i64>,
    /// Whether the weight is stored (Bold is false without it).
    pub weight_stored: bool,
}

/// `ReadTextFont`: omitted members stay omitted.
pub(crate) fn text_font(record: Option<&Record>) -> Option<Font> {
    let record = record?;
    let mut font = Font {
        name: record.text("FontName").map(str::to_string),
        size: record.value("FontHeight").map(|v| v as f64 / 20.0),
        charset: record.value("FontCharSet"),
        pitch_and_family: record.value("FontPitchAndFamily"),
        weight_stored: record.value("FontWeight").is_some(),
        ..Font::default()
    };
    if let Some(effects) = record.value("FontEffects") {
        // The stored weight decides Bold; the effects bit alone does not.
        font.bold = Some(false);
        font.italic = Some(effects & 2 != 0);
        font.underline = Some(effects & 4 != 0);
        font.strike_through = Some(effects & 8 != 0);
    }
    if let Some(weight) = record.value("FontWeight") {
        font.bold = Some(weight > 550);
    }
    if font.name.is_none() && font.size.is_none() && font.charset.is_none() && font.bold.is_none() {
        return None;
    }
    Some(font)
}

/// `ReadLevelFont`: `Ok(None)` without a font, `Err(())` for a font this
/// module cannot interpret.
pub(crate) fn level_font(level: &Level) -> std::result::Result<Option<Font>, ()> {
    let raw = &level.font_raw;
    if raw.is_empty() {
        return Ok(None);
    }
    if raw.len() < 16 {
        return Err(());
    }
    match u32::from_le_bytes([raw[0], raw[1], raw[2], raw[3]]) {
        GUID_STD_FONT => {
            if raw.len() < 27
                || raw[16] != 1
                || raw.len() != 27 + raw[26] as usize
                || raw[27..].iter().any(|b| *b > 127)
            {
                return Err(());
            }
            let flags = raw[19];
            Ok(Some(Font {
                name: Some(raw[27..].iter().map(|b| *b as char).collect()),
                size: Some(
                    u32::from_le_bytes([raw[22], raw[23], raw[24], raw[25]]) as f64 / 10000.0,
                ),
                charset: Some(i16::from_le_bytes([raw[17], raw[18]]) as i64),
                bold: Some(u16::from_le_bytes([raw[20], raw[21]]) > 550),
                italic: Some(flags & 2 != 0),
                underline: Some(flags & 4 != 0),
                strike_through: Some(flags & 8 != 0),
                pitch_and_family: None,
                weight_stored: true,
            }))
        }
        GUID_TEXT_PROPS => match parse_record(&raw[16..], &TEXT_PROPS, raw.len() - 16) {
            Ok((record, consumed)) if consumed == raw.len() - 16 => Ok(text_font(Some(&record))),
            _ => Err(()),
        },
        _ => Err(()),
    }
}

/// A MultiPage: its hidden TabStrip and its pages in tab order.
#[derive(Clone, Debug)]
pub(crate) struct MultiPage {
    pub hidden: Box<Control>,
    pub pages: Vec<Control>,
    /// The `MultiPageProperties` record of the `x` stream.
    pub properties: Record,
}

/// A control of the form.
#[derive(Clone, Debug)]
pub(crate) struct Control {
    pub name: String,
    /// `MSForms.TextBox`, ..., `MSForms.Control` or `ActiveX.Control`.
    pub kind: &'static str,
    pub id: i32,
    pub class: u16,
    pub tab_index: Option<i16>,
    pub site: Site,
    /// Property record of a non-container control.
    pub record: Option<Record>,
    /// The control's data could not be interpreted.
    pub opaque: bool,
    /// Child level of a container control.
    pub level: Option<Level>,
    /// Children of a Frame or Page (empty for a MultiPage, see `multipage`).
    pub children: Vec<Control>,
    pub multipage: Option<MultiPage>,
    pub tab_strip: Option<TabStripState>,
}

impl Control {
    /// The record that holds the control's properties: its own, or the
    /// form record of its child level.
    pub(crate) fn properties(&self) -> Option<&Record> {
        self.record
            .as_ref()
            .or(self.level.as_ref().map(|l| &l.record))
    }
}

/// A whole designer storage.
#[derive(Clone, Debug)]
pub(crate) struct Form {
    pub root: Level,
    pub controls: Vec<Control>,
}

fn control_kind(class: u16, record: Option<&Record>) -> &'static str {
    if class >= 0x8000 {
        return "ActiveX.Control";
    }
    match class {
        7 => "MSForms.Form",
        12 => "MSForms.Image",
        14 => "MSForms.Frame",
        15 => match record.map(|r| r.value("DisplayStyle").unwrap_or(0)) {
            None => "MSForms.MorphData",
            Some(0 | 1) => "MSForms.TextBox",
            Some(2) => "MSForms.ListBox",
            Some(3 | 7) => "MSForms.ComboBox",
            Some(4) => "MSForms.CheckBox",
            Some(5) => "MSForms.OptionButton",
            Some(6) => "MSForms.ToggleButton",
            Some(_) => "MSForms.Control",
        },
        16 => "MSForms.SpinButton",
        17 => "MSForms.CommandButton",
        18 => "MSForms.TabStrip",
        21 => "MSForms.Label",
        23 => "MSForms.TextBox",
        24 => "MSForms.ListBox",
        25 => "MSForms.ComboBox",
        26 => "MSForms.CheckBox",
        27 => "MSForms.OptionButton",
        28 => "MSForms.ToggleButton",
        47 => "MSForms.ScrollBar",
        57 => "MSForms.MultiPage",
        _ => "MSForms.Control",
    }
}

fn decode_array(raw: &[u8]) -> Result<Vec<String>> {
    let mut values = Vec::new();
    let mut offset = 0;
    while offset < raw.len() {
        let start = offset;
        if raw.len() - offset < 4 {
            return Err(format!("{offset} バイト目の文字数が途中で切れています"));
        }
        let packed = u32::from_le_bytes([
            raw[offset],
            raw[offset + 1],
            raw[offset + 2],
            raw[offset + 3],
        ]);
        offset += 4;
        let compressed = packed & 0x8000_0000 != 0;
        let count = (packed & 0x7fff_ffff) as usize;
        let length = if compressed { count } else { count * 2 };
        if length > raw.len() - offset {
            return Err(format!(
                "文字列の {length} バイトが残り {} バイトを超えています",
                raw.len() - offset
            ));
        }
        values.push(stored_text(&raw[offset..offset + length], compressed)?);
        offset += length;
        let padding = (4 - (offset - start) % 4) % 4;
        if padding > raw.len() - offset {
            return Err(format!("{offset} バイト目の詰め物が途中で切れています"));
        }
        offset += padding;
    }
    Ok(values)
}

/// `ParseTabStrip`.
fn parse_tab_strip(record: &Record) -> Result<TabStripState> {
    type Setter = fn(&mut Tab, String);
    let arrays: [(u8, &str, Setter); 5] = [
        (5, "Items", |t, v| t.caption = v),
        (15, "TipStrings", |t, v| t.control_tip_text = v),
        (17, "TabNames", |t, v| t.name = v),
        (21, "Tags", |t, v| t.tag = v),
        (23, "Accelerators", |t, v| t.accelerator = v),
    ];
    let mut decoded: Vec<Option<Vec<String>>> = Vec::new();
    for (bit, name, _) in &arrays {
        let present = record.mask & (1u64 << bit) != 0;
        let raw = record.arrays.get(name);
        if present != raw.is_some() {
            return Err(format!("TabStrip: {name} の有無がマスクと合いません"));
        }
        let size = record.value(array_size_field(name));
        match raw {
            None => {
                if size.is_some_and(|s| s != 0) {
                    return Err(format!("TabStrip: {name} が無いのにサイズがあります"));
                }
                decoded.push(None);
            }
            Some(raw) => {
                if size != Some(raw.len() as i64) {
                    return Err(format!("TabStrip: {name} のサイズが合いません"));
                }
                decoded.push(Some(
                    decode_array(raw).map_err(|e| format!("TabStrip {name}: {e}"))?,
                ));
            }
        }
    }
    let count = decoded[0].as_ref().map_or(0, Vec::len);
    for (index, (_, name, _)) in arrays.iter().enumerate().skip(1) {
        if decoded[index].as_ref().is_some_and(|v| v.len() != count) {
            return Err(format!(
                "TabStrip: {name} の数がタブの数 {count} と合いません"
            ));
        }
    }
    let mut tabs: Vec<Tab> = (0..count)
        .map(|_| Tab {
            enabled: true,
            visible: true,
            ..Tab::default()
        })
        .collect();
    for (index, (_, _, set)) in arrays.iter().enumerate() {
        if let Some(values) = decoded[index].take() {
            for (tab, value) in tabs.iter_mut().zip(values) {
                set(tab, value);
            }
        }
    }
    let selected_index = record.value("ListIndex").unwrap_or(-1);
    if selected_index < -1 || (!tabs.is_empty() && selected_index >= tabs.len() as i64) {
        return Err(format!(
            "TabStrip: ListIndex {selected_index} がタブの範囲外です"
        ));
    }
    let data_count = if record.mask & (1 << 22) != 0 {
        record.value("TabData").unwrap_or(0).max(0) as usize
    } else {
        0
    };
    if data_count > tabs.len() || data_count * 4 > record.tail.len() {
        return Err(format!(
            "TabStrip: TabData {data_count} がタブまたはデータの範囲を超えています"
        ));
    }
    for (index, tab) in tabs.iter_mut().enumerate().take(data_count) {
        let word = u32::from_le_bytes([
            record.tail[index * 4],
            record.tail[index * 4 + 1],
            record.tail[index * 4 + 2],
            record.tail[index * 4 + 3],
        ]);
        tab.visible = word & 1 != 0;
        tab.enabled = word & 2 != 0;
    }
    Ok(TabStripState {
        tabs,
        selected_index: selected_index as i32,
    })
}

struct State<'a> {
    streams: &'a BTreeMap<String, Vec<u8>>,
    storages: &'a BTreeMap<String, [u8; 16]>,
    total_sites: usize,
}

fn join(storage: &str, name: &str) -> String {
    if storage.is_empty() {
        name.to_string()
    } else {
        format!("{storage}/{name}")
    }
}

/// The part of `path` directly below `storage`, when it is a direct child.
fn direct_child<'p>(storage: &str, path: &'p str) -> Option<&'p str> {
    let rest = if storage.is_empty() {
        path
    } else {
        path.strip_prefix(storage)?.strip_prefix('/')?
    };
    (!rest.is_empty() && !rest.contains('/')).then_some(rest)
}

impl<'a> State<'a> {
    fn stream(&self, storage: &str, name: &str) -> Option<&'a [u8]> {
        if let Some(data) = self.streams.get(&join(storage, name)) {
            return Some(data);
        }
        // Compound file names are compared without regard to case.
        self.streams
            .iter()
            .find(|(path, _)| {
                direct_child(storage, path).is_some_and(|n| n.to_uppercase() == name.to_uppercase())
            })
            .map(|(_, data)| data.as_slice())
    }

    fn fail<T>(
        &self,
        storage: &str,
        stream: &str,
        detail: String,
    ) -> std::result::Result<T, ReadError> {
        Err(ReadError {
            storage: storage.to_string(),
            stream: stream.to_string(),
            detail,
        })
    }

    /// Child storages `i<ID>` of `storage`, keyed by the site ID.
    fn child_storages(
        &self,
        storage: &str,
    ) -> std::result::Result<HashMap<i32, String>, ReadError> {
        let mut children = HashMap::new();
        for path in self.storages.keys() {
            let Some(name) = direct_child(storage, path) else {
                continue;
            };
            let id = name
                .strip_prefix(['i', 'I'])
                .filter(|digits| !digits.is_empty() && digits.bytes().all(|b| b.is_ascii_digit()))
                .and_then(|digits| digits.parse::<i32>().ok());
            let Some(id) = id else {
                return self.fail(
                    storage,
                    "",
                    format!("見慣れない子ストレージ {name:?} があります"),
                );
            };
            if children.insert(id, path.clone()).is_some() {
                return self.fail(
                    storage,
                    "",
                    format!("ID {id} の子ストレージが重複しています"),
                );
            }
        }
        Ok(children)
    }

    fn level(
        &mut self,
        storage: &str,
        depth: usize,
        root: bool,
    ) -> std::result::Result<(Level, Vec<Control>), ReadError> {
        if depth >= MAX_NESTING_DEPTH {
            return self.fail(
                storage,
                "",
                format!("入れ子が {MAX_NESTING_DEPTH} 段を超えています"),
            );
        }
        let Some(f) = self.stream(storage, "f") else {
            return self.fail(storage, "f", "ストリームがありません".into());
        };
        let Some(o) = self.stream(storage, "o") else {
            return self.fail(storage, "o", "ストリームがありません".into());
        };
        if f.len() > MAX_STREAM_SIZE || o.len() > MAX_STREAM_SIZE {
            return self.fail(storage, "", "ストリームが大きすぎます".into());
        }
        for required in ["\u{1}CompObj", "\u{3}VBFrame"] {
            if root && self.stream(storage, required).is_none() {
                return self.fail(storage, required, "ストリームがありません".into());
            }
        }
        if let Some(compobj) = self.stream(storage, "\u{1}CompObj") {
            let ok = compobj.len() >= 28
                && compobj[..4] == 0xfffe_0001u32.to_le_bytes()
                && compobj[4..8] == 0x0a03u32.to_le_bytes();
            if !ok {
                return self.fail(storage, "\u{1}CompObj", "ヘッダーが不正です".into());
            }
        }

        let (level, sites) = match parse_form_stream(f) {
            Ok(parsed) => parsed,
            Err(detail) => return self.fail(storage, "f", detail),
        };
        let mut level = level;
        level.extra_streams = self
            .streams
            .keys()
            .filter_map(|path| direct_child(storage, path))
            .filter(|name| {
                !["F", "O", "X", "\u{1}COMPOBJ", "\u{3}VBFRAME"]
                    .contains(&name.to_uppercase().as_str())
            })
            .map(str::to_string)
            .collect();

        self.total_sites += sites.len();
        if self.total_sites > MAX_SITES {
            return self.fail(
                storage,
                "f",
                format!("サイトが {MAX_SITES} 個を超えています"),
            );
        }
        let mut child_paths = self.child_storages(storage)?;
        let mut offset = 0usize;
        let mut claimed = std::collections::HashSet::new();
        let mut controls = Vec::with_capacity(sites.len());
        for site in sites {
            let name = site.strings.get("Name").cloned().unwrap_or_default();
            let size = site.values.get("ObjectStreamSize").copied().unwrap_or(0) as u32 as usize;
            if offset + size > o.len() {
                return self.fail(
                    storage,
                    "o",
                    format!("サイト {name:?} のサイズ {size} がストリームを超えています"),
                );
            }
            let id = site.values.get("ID").copied().unwrap_or(0) as i32;
            if !claimed.insert(id) {
                return self.fail(storage, "f", format!("サイトの ID {id} が重複しています"));
            }
            let class = site.values.get("ClsidCacheIndex").copied().unwrap_or(0) as u16;
            let mut control = Control {
                name: name.clone(),
                kind: control_kind(class, None),
                id,
                class,
                tab_index: site.values.get("TabIndex").map(|v| *v as i16),
                site,
                record: None,
                opaque: false,
                level: None,
                children: Vec::new(),
                multipage: None,
                tab_strip: None,
            };
            let child = child_paths.remove(&id);
            let container = is_container_class(class);
            match child {
                Some(path) if !container => {
                    return self.fail(storage, "f", format!("コンテナーではないサイト {name:?} が子ストレージ {path:?} を持っています"));
                }
                None if container => {
                    return self.fail(
                        storage,
                        "f",
                        format!("コンテナーのサイト {name:?}（ID {id}）に子ストレージがありません"),
                    );
                }
                Some(path) => {
                    if size != 0 {
                        return self.fail(
                            storage,
                            "o",
                            format!("コンテナーのサイト {name:?} にオブジェクト データがあります"),
                        );
                    }
                    let (child_level, children) = self.level(&path, depth + 1, false)?;
                    control.level = Some(child_level);
                    control.children = children;
                    if class == 57 {
                        let x = self.stream(&path, "x").map(<[u8]>::to_vec);
                        if let Err(detail) = bind_multi_page(&mut control, x.as_deref()) {
                            return self.fail(&path, "x", detail);
                        }
                    }
                }
                None => {
                    let payload = &o[offset..offset + size];
                    match spec_for_class(class).filter(|_| !payload.is_empty()) {
                        Some(spec) => {
                            let record = match parse_record(payload, spec, payload.len()) {
                                Ok((record, _)) => record,
                                Err(detail) => {
                                    return self.fail(storage, "o", format!("{name}: {detail}"))
                                }
                            };
                            control.kind = control_kind(class, Some(&record));
                            if class == 18 {
                                match parse_tab_strip(&record) {
                                    Ok(strip) => control.tab_strip = Some(strip),
                                    Err(detail) => return self.fail(storage, "o", detail),
                                }
                            }
                            control.record = Some(record);
                        }
                        None => control.opaque = !payload.is_empty(),
                    }
                    offset += size;
                }
            }
            controls.push(control);
        }
        if offset != o.len() {
            return self.fail(
                storage,
                "o",
                format!(
                    "サイトの合計 {offset} バイトがストリームの {} バイトと合いません",
                    o.len()
                ),
            );
        }
        if let Some(path) = child_paths.values().min() {
            return self.fail(
                storage,
                "",
                format!("どのサイトにも属さない子ストレージ {path:?} があります"),
            );
        }
        Ok((level, controls))
    }
}

fn parse_form_stream(data: &[u8]) -> Result<(Level, Vec<Site>)> {
    if data.len() < 4 {
        return Err("ヘッダーより短いストリームです".into());
    }
    let form_end = 4 + u16::from_le_bytes([data[2], data[3]]) as usize;
    if form_end > data.len() {
        return Err(format!(
            "フォーム レコードの長さがストリーム（{} バイト）を超えています",
            data.len()
        ));
    }
    let (record, _) = parse_record(data, &FORM, form_end)?;
    let mut r = Reader::new(data);
    r.pos = form_end;
    let mut level = Level {
        mouse_icon: None,
        font_raw: Vec::new(),
        picture: None,
        extra_streams: Vec::new(),
        record,
    };
    if level.record.mask & FORM_MASK_MOUSE_ICON != 0 {
        level.mouse_icon = Some(
            r.guid_and_picture()
                .map_err(|e| format!("MouseIcon: {e}"))?,
        );
    }
    if level.record.mask & FORM_MASK_FONT != 0 {
        level.font_raw = read_font_blob(&mut r).map_err(|e| format!("Font: {e}"))?;
    }
    if level.record.mask & FORM_MASK_PICTURE != 0 {
        level.picture = Some(r.guid_and_picture().map_err(|e| format!("Picture: {e}"))?);
    }
    let start = r.pos;
    let sites = match parse_site_data(data, start, true) {
        Ok(sites) => sites,
        Err(stored) => parse_site_data(data, start, false).map_err(|omitted| {
            format!("サイトの並びを読めません（クラス表あり: {stored} / クラス表なし: {omitted}）")
        })?,
    };
    Ok((level, sites))
}

/// `bindMultiPage`: orders the pages by the `x` stream and separates the
/// hidden TabStrip.
fn bind_multi_page(control: &mut Control, x: Option<&[u8]>) -> Result<()> {
    let Some(raw) = x else {
        return Err("MultiPage に x ストリームがありません".into());
    };
    let mut hidden = None;
    let mut pages: Vec<Control> = Vec::new();
    for child in std::mem::take(&mut control.children) {
        match child.class {
            18 => {
                if hidden.is_some() || !child.name.is_empty() {
                    return Err("MultiPage の内部 TabStrip が不正です".into());
                }
                hidden = Some(child);
            }
            7 => pages.push(child),
            _ => return Err("MultiPage に Page 以外のコントロールがあります".into()),
        }
    }
    let Some(hidden) = hidden else {
        return Err("MultiPage の内部 TabStrip がありません".into());
    };
    let mut offset = 0usize;
    let mut next_record = |spec: &'static RecordSpec| -> Result<Record> {
        if raw.len() - offset < 8 {
            return Err(format!("{} が途中で切れています", spec.type_name));
        }
        let end = offset + 4 + u16::from_le_bytes([raw[offset + 2], raw[offset + 3]]) as usize;
        if end > raw.len() {
            return Err(format!("{} の境界が不正です", spec.type_name));
        }
        let (record, _) = parse_record(&raw[offset..end], spec, end - offset)?;
        offset = end;
        Ok(record)
    };
    for _ in 0..pages.len() + 1 {
        next_record(&PAGE_PROPERTIES)?;
    }
    let properties = next_record(&MULTI_PAGE_PROPERTIES)?;
    if properties.value("PageCount") != Some(pages.len() as i64)
        || properties.value("ID") != Some(hidden.id as i64)
        || raw.len() - offset != 4 * pages.len()
    {
        return Err("MultiPage のページ数・内部 TabStrip の ID・長さが合いません".into());
    }
    let mut ordered = Vec::with_capacity(pages.len());
    for index in 0..pages.len() {
        let at = offset + 4 * index;
        let id = i32::from_le_bytes([raw[at], raw[at + 1], raw[at + 2], raw[at + 3]]);
        let Some(position) = pages.iter().position(|p| p.id == id) else {
            return Err(format!(
                "MultiPage のページ ID {id} が重複しているか、ありません"
            ));
        };
        ordered.push(pages.remove(position));
    }
    let tabs = hidden.tab_strip.as_ref().map(|s| s.tabs.len());
    if tabs.is_none() || (!ordered.is_empty() && tabs != Some(ordered.len())) {
        return Err("MultiPage のタブとページの数が合いません".into());
    }
    control.multipage = Some(MultiPage {
        hidden: Box::new(hidden),
        pages: ordered,
        properties,
    });
    Ok(())
}

/// Reads a designer storage (`ReadForm`). `streams` and `storages` are keyed
/// by their path below the form storage, as in [`crate::Designer`].
pub(crate) fn read_form(
    streams: &BTreeMap<String, Vec<u8>>,
    storages: &BTreeMap<String, [u8; 16]>,
) -> std::result::Result<Form, ReadError> {
    let mut state = State {
        streams,
        storages,
        total_sites: 0,
    };
    let (root, controls) = state.level("", 0, true)?;
    Ok(Form { root, controls })
}
