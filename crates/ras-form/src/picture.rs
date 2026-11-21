//! Pictures of controls and forms: reading the source file and the
//! `StdPicture` payload MS-OFORMS stores for a picture property.
//!
//! A `Picture` is always stored as a BMP (40-byte `BITMAPINFOHEADER`, 24 or
//! 32 bit, uncompressed) or a JPEG; these two are the payloads known to load
//! in MS-Forms. A source file in one of these forms is stored as it is.
//! Any other picture — PNG, GIF, WebP, TIFF, ICO, and BMP files of another
//! flavour — is decoded and stored as a 24-bit BMP ([`prepare`]).
//!
//! A `MouseIcon` is an icon or a cursor file (ICO / CUR), stored as it is
//! ([`prepare_icon`]); `StdPicture` holds icons ([MS-OFORMS] 2.4.13).
//!
//! The checks of stored pictures are a port of xlflow's
//! `internal/vba/userforms/picture` (Copyright (c) 2026 harumiWeb, MIT
//! License; see `NOTICE.md`). JPEG files are checked structurally (marker
//! segments and frame header) instead of being decoded.

use std::path::{Component, Path, PathBuf};

/// Largest picture file accepted.
pub const MAX_BYTES: usize = 16 << 20;
/// Largest picture accepted, in pixels.
pub const MAX_PIXELS: u64 = 16_000_000;

const STD_PICTURE_TAG: u32 = 0x746c;
const STD_PICTURE_CLSID: [u8; 16] = [
    0x04, 0x52, 0xe3, 0x0b, 0x91, 0x8f, 0xce, 0x11, 0x9d, 0xe3, 0x00, 0xaa, 0x00, 0x4b, 0xb8, 0x51,
];

/// What transparent pixels are blended onto when no color is given: the
/// default face color of a form (`buttonFace`).
pub const DEFAULT_BACKGROUND: [u8; 3] = [0xf0, 0xf0, 0xf0];

/// Default colors of the Windows system colors (index of `COLOR_*`), as
/// red, green, blue. Used where a color has to be known at build time.
const SYSTEM_COLOR_DEFAULTS: &[(u32, [u8; 3])] = &[
    (0, [0xc8, 0xc8, 0xc8]),
    (1, [0x00, 0x00, 0x00]),
    (2, [0x99, 0xb4, 0xd1]),
    (3, [0xbf, 0xcd, 0xdb]),
    (4, [0xf0, 0xf0, 0xf0]),
    (5, [0xff, 0xff, 0xff]),
    (6, [0x64, 0x64, 0x64]),
    (7, [0x00, 0x00, 0x00]),
    (8, [0x00, 0x00, 0x00]),
    (9, [0x00, 0x00, 0x00]),
    (10, [0xb4, 0xb4, 0xb4]),
    (11, [0xf4, 0xf7, 0xfc]),
    (12, [0xab, 0xab, 0xab]),
    (13, [0x00, 0x78, 0xd7]),
    (14, [0xff, 0xff, 0xff]),
    (15, [0xf0, 0xf0, 0xf0]),
    (16, [0xa0, 0xa0, 0xa0]),
    (17, [0x6d, 0x6d, 0x6d]),
    (18, [0x00, 0x00, 0x00]),
    (19, [0x00, 0x00, 0x00]),
    (20, [0xff, 0xff, 0xff]),
    (21, [0x69, 0x69, 0x69]),
    (22, [0xe3, 0xe3, 0xe3]),
    (23, [0x00, 0x00, 0x00]),
    (24, [0xff, 0xff, 0xe1]),
    (26, [0x00, 0x66, 0xcc]),
    (27, [0xb9, 0xd1, 0xea]),
    (28, [0xd7, 0xe4, 0xf2]),
    (29, [0x33, 0x99, 0xff]),
    (30, [0xf0, 0xf0, 0xf0]),
];

/// The color transparent pixels of a converted picture are blended onto:
/// the `backColor` of the Image control. A system color counts with its
/// Windows default; without a usable color it is [`DEFAULT_BACKGROUND`].
pub fn background_of(back_color: Option<&crate::FormSpecColor>) -> [u8; 3] {
    let Some(ole) = back_color.and_then(|c| c.ole().ok()) else {
        return DEFAULT_BACKGROUND;
    };
    if ole & 0xff00_0000 == 0 {
        // OLE colors are 0x00BBGGRR.
        return [ole as u8, (ole >> 8) as u8, (ole >> 16) as u8];
    }
    if ole & 0xffff_ff00 == 0x8000_0000 {
        if let Some((_, rgb)) = SYSTEM_COLOR_DEFAULTS.iter().find(|(i, _)| *i == ole & 0xff) {
            return *rgb;
        }
    }
    DEFAULT_BACKGROUND
}

/// Format of a picture as it is stored in a form.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Format {
    /// Windows bitmap (24 or 32 bit, uncompressed).
    Bmp,
    /// JPEG.
    Jpeg,
    /// Icon (ICO): a `MouseIcon`.
    Icon,
    /// Cursor (CUR): a `MouseIcon`.
    Cursor,
}

impl Format {
    /// File name extension (`bmp` / `jpg` / `ico` / `cur`).
    pub fn extension(self) -> &'static str {
        match self {
            Format::Bmp => "bmp",
            Format::Jpeg => "jpg",
            Format::Icon => "ico",
            Format::Cursor => "cur",
        }
    }
}

/// A checked picture.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Picture {
    /// The file contents.
    pub data: Vec<u8>,
    /// What kind of file it is.
    pub format: Format,
    /// Width in pixels.
    pub width: u32,
    /// Height in pixels.
    pub height: u32,
    /// Format of the source file when `data` is a conversion of it
    /// (`"PNG"`, `"GIF"`, ...); `None` when `data` is the file itself.
    pub converted_from: Option<&'static str>,
}

fn u16le(data: &[u8], at: usize) -> u32 {
    u16::from_le_bytes([data[at], data[at + 1]]) as u32
}

fn u32le(data: &[u8], at: usize) -> u32 {
    u32::from_le_bytes([data[at], data[at + 1], data[at + 2], data[at + 3]])
}

fn check_dimensions(width: u64, height: u64) -> Result<(), String> {
    if width == 0 || height == 0 || width * height > MAX_PIXELS {
        return Err(format!(
            "画像の大きさ {width}x{height} が上限（{MAX_PIXELS} ピクセル）を超えているか、0 です"
        ));
    }
    Ok(())
}

fn inspect_bmp(data: &[u8]) -> Result<(u32, u32), String> {
    if data.len() < 54 {
        return Err("BMP のヘッダーが途中で切れています".into());
    }
    let declared = u32le(data, 2) as u64;
    let pixel_offset = u32le(data, 10) as u64;
    if u32le(data, 14) != 40 {
        return Err(
            "BMP は 40 バイトの BITMAPINFOHEADER 形式だけに対応しています（ペイントなどで 24 ビット ビットマップとして保存し直してください）"
                .into(),
        );
    }
    let width = u32le(data, 18) as i32 as i64;
    let height = u32le(data, 22) as i32 as i64;
    if width <= 0 || height == 0 || height == i32::MIN as i64 {
        return Err("BMP の幅・高さが不正です".into());
    }
    let (width, height) = (width as u64, height.unsigned_abs());
    check_dimensions(width, height)?;
    if u16le(data, 26) != 1 {
        return Err("BMP のプレーン数が 1 ではありません".into());
    }
    let bits = u16le(data, 28) as u64;
    if bits != 24 && bits != 32 {
        return Err(format!(
            "BMP は 24 ビットか 32 ビットだけに対応しています（このファイルは {bits} ビット）"
        ));
    }
    let compression = u32le(data, 30);
    if compression != 0 {
        return Err(format!(
            "BMP は無圧縮（BI_RGB）だけに対応しています（このファイルの圧縮形式は {compression}）"
        ));
    }
    if pixel_offset < 54 || pixel_offset > data.len() as u64 {
        return Err("BMP の画素データの位置がファイルの外を指しています".into());
    }
    let row_bytes = (width * bits).div_ceil(32) * 4;
    let pixel_bytes = row_bytes * height;
    if pixel_offset + pixel_bytes > data.len() as u64 {
        return Err("BMP の画素データが途中で切れています".into());
    }
    if declared != 0 && (declared > data.len() as u64 || declared < pixel_offset + pixel_bytes) {
        return Err("BMP のヘッダーにあるファイル サイズが画素データと合いません".into());
    }
    let image_size = u32le(data, 34) as u64;
    if image_size != 0 && image_size < pixel_bytes {
        return Err("BMP のヘッダーにある画像サイズが画素データより小さくなっています".into());
    }
    Ok((width as u32, height as u32))
}

/// Walks the marker segments of a JPEG file up to the start of the scan.
fn inspect_jpeg(data: &[u8]) -> Result<(u32, u32), String> {
    let bad = |what: &str| format!("JPEG として読めません（{what}）");
    let mut at = 2;
    let mut frame = None;
    loop {
        // Fill bytes (0xFF) may precede a marker.
        while at < data.len() && data[at] == 0xff && data.get(at + 1) == Some(&0xff) {
            at += 1;
        }
        if at + 4 > data.len() || data[at] != 0xff {
            return Err(bad("マーカーが見つかりません"));
        }
        let marker = data[at + 1];
        let length = u16::from_be_bytes([data[at + 2], data[at + 3]]) as usize;
        if length < 2 || at + 2 + length > data.len() {
            return Err(bad("セグメントが途中で切れています"));
        }
        let body = &data[at + 4..at + 2 + length];
        match marker {
            // SOF0–SOF15 without DHT (C4), JPG (C8) and DAC (CC).
            0xc0..=0xcf if !matches!(marker, 0xc4 | 0xc8 | 0xcc) => {
                if body.len() < 6 {
                    return Err(bad("フレーム ヘッダーが短すぎます"));
                }
                let height = u16::from_be_bytes([body[1], body[2]]) as u32;
                let width = u16::from_be_bytes([body[3], body[4]]) as u32;
                frame = Some((width, height));
            }
            // SOS: the entropy-coded data follows.
            0xda => break,
            0xd8 | 0xd9 => return Err(bad("マーカーの並びが不正です")),
            _ => {}
        }
        at += 2 + length;
    }
    let Some((width, height)) = frame else {
        return Err(bad("フレーム ヘッダーがありません"));
    };
    if !data[at..].windows(2).any(|w| w == [0xff, 0xd9]) {
        return Err(bad("終了マーカーがありません"));
    }
    check_dimensions(width as u64, height as u64)?;
    Ok((width, height))
}

/// Checks an icon or cursor file (`ICONDIR` and its entries); returns its kind and the size
/// of its first image.
fn inspect_icon(data: &[u8]) -> Result<(Format, u32, u32), String> {
    let bad = |what: &str| format!("アイコン（ICO / CUR）として読めません（{what}）");
    if data.len() < 6 || u16le(data, 0) != 0 {
        return Err(bad("ヘッダーが不正です"));
    }
    let format = match u16le(data, 2) {
        1 => Format::Icon,
        2 => Format::Cursor,
        _ => return Err(bad("種類がアイコンでもカーソルでもありません")),
    };
    let count = u16le(data, 4) as usize;
    if count == 0 || data.len() < 6 + count * 16 {
        return Err(bad("画像の一覧が途中で切れています"));
    }
    for entry in 0..count {
        let at = 6 + entry * 16;
        let size = u32le(data, at + 8) as u64;
        let offset = u32le(data, at + 12) as u64;
        if size == 0 || offset < (6 + count * 16) as u64 || offset + size > data.len() as u64 {
            return Err(bad("画像のデータがファイルの外を指しています"));
        }
    }
    let side = |b: u8| if b == 0 { 256 } else { b as u32 };
    Ok((format, side(data[6]), side(data[7])))
}

/// Checks that `data` is a picture that can be stored in a form as it is: a BMP or JPEG
/// picture, or an icon or cursor.
pub fn inspect(data: &[u8]) -> Result<(Format, u32, u32), String> {
    if data.len() >= 4 && data[..2] == [0, 0] && matches!(data[2..4], [1, 0] | [2, 0]) {
        if data.len() > MAX_BYTES {
            return Err(format!("画像は {MAX_BYTES} バイト以下にしてください"));
        }
        return inspect_icon(data);
    }
    inspect_picture(data)
}

/// Checks that `data` is a BMP or JPEG picture that can be stored in a form as it is.
fn inspect_picture(data: &[u8]) -> Result<(Format, u32, u32), String> {
    if data.is_empty() || data.len() > MAX_BYTES {
        return Err(format!(
            "画像は 1 バイト以上 {MAX_BYTES} バイト以下にしてください"
        ));
    }
    if data.starts_with(b"BM") {
        let (width, height) = inspect_bmp(data)?;
        return Ok((Format::Bmp, width, height));
    }
    if data.starts_with(&[0xff, 0xd8]) {
        let (width, height) = inspect_jpeg(data)?;
        return Ok((Format::Jpeg, width, height));
    }
    Err("画像の形式は BMP か JPEG にしてください".into())
}

/// Resolves a project-root-relative picture path. The path must stay inside
/// `root`: no absolute path, no drive, no `..`.
pub fn resolve(root: &Path, path: &str) -> Result<PathBuf, String> {
    let normalized = path.replace('\\', "/");
    let relative = Path::new(&normalized);
    let local = !path.trim().is_empty()
        && !normalized.starts_with('/')
        && !normalized.contains(':')
        && relative
            .components()
            .all(|c| matches!(c, Component::Normal(_) | Component::CurDir))
        && relative
            .components()
            .any(|c| matches!(c, Component::Normal(_)));
    if !local {
        return Err(
            "画像のパスは、プロジェクトのルートからの相対パスにしてください（`..` や絶対パスは使えません）"
                .into(),
        );
    }
    Ok(root.join(relative))
}

/// Size of a 24-bit BMP file of the given dimensions.
fn bmp_size(width: u64, height: u64) -> u64 {
    54 + (width * 3).div_ceil(4) * 4 * height
}

/// Writes pixels as a bottom-up 24-bit BMP. `rgba` holds the rows top
/// down; transparent pixels are blended onto `background`.
fn encode_bmp(width: u32, height: u32, rgba: &[u8], background: [u8; 3]) -> Vec<u8> {
    let row = (width as usize * 3).div_ceil(4) * 4;
    let size = 54 + row * height as usize;
    let mut data = vec![0u8; size];
    data[..2].copy_from_slice(b"BM");
    data[2..6].copy_from_slice(&(size as u32).to_le_bytes());
    data[10..14].copy_from_slice(&54u32.to_le_bytes());
    data[14..18].copy_from_slice(&40u32.to_le_bytes());
    data[18..22].copy_from_slice(&(width as i32).to_le_bytes());
    data[22..26].copy_from_slice(&(height as i32).to_le_bytes());
    data[26..28].copy_from_slice(&1u16.to_le_bytes());
    data[28..30].copy_from_slice(&24u16.to_le_bytes());
    data[34..38].copy_from_slice(&((size - 54) as u32).to_le_bytes());
    // 96 DPI, like the bitmaps Windows tools write.
    data[38..42].copy_from_slice(&3780u32.to_le_bytes());
    data[42..46].copy_from_slice(&3780u32.to_le_bytes());
    let blend = |value: u8, back: u8, alpha: u32| -> u8 {
        ((value as u32 * alpha + back as u32 * (255 - alpha) + 127) / 255) as u8
    };
    for y in 0..height as usize {
        let source = &rgba[y * width as usize * 4..(y + 1) * width as usize * 4];
        let start = 54 + (height as usize - 1 - y) * row;
        for (x, pixel) in source.chunks_exact(4).enumerate() {
            let alpha = pixel[3] as u32;
            // BMP stores blue, green, red.
            data[start + x * 3] = blend(pixel[2], background[2], alpha);
            data[start + x * 3 + 1] = blend(pixel[1], background[1], alpha);
            data[start + x * 3 + 2] = blend(pixel[0], background[0], alpha);
        }
    }
    data
}

/// The formats that are converted, with the name used in messages.
fn convertible(format: image::ImageFormat) -> Option<&'static str> {
    Some(match format {
        image::ImageFormat::Png => "PNG",
        image::ImageFormat::Gif => "GIF",
        image::ImageFormat::WebP => "WebP",
        image::ImageFormat::Tiff => "TIFF",
        image::ImageFormat::Ico => "ICO",
        image::ImageFormat::Bmp => "BMP",
        _ => return None,
    })
}

/// Decodes a picture MS-Forms cannot store and writes it as a 24-bit BMP.
fn convert(data: &[u8], background: [u8; 3]) -> Result<Picture, String> {
    let unknown = || {
        "画像の形式を判別できません（BMP / JPEG / PNG / GIF / WebP / TIFF / ICO に対応しています）"
            .to_string()
    };
    let format = image::guess_format(data).map_err(|_| unknown())?;
    let name = convertible(format).ok_or_else(unknown)?;
    let reader = || {
        let mut reader = image::ImageReader::with_format(std::io::Cursor::new(data), format);
        let mut limits = image::Limits::default();
        limits.max_alloc = Some(512 << 20);
        reader.limits(limits);
        reader
    };
    let unreadable = |e: image::ImageError| format!("{name} として読めません（{e}）");
    let (width, height) = reader().into_dimensions().map_err(unreadable)?;
    check_dimensions(width as u64, height as u64)?;
    if bmp_size(width as u64, height as u64) > MAX_BYTES as u64 {
        return Err(format!(
            "{name} の画像（{width}x{height}）は、BMP に変換すると {MAX_BYTES} バイトを超えます。画像を小さくするか、JPEG にしてください"
        ));
    }
    // An animated GIF gives its first frame.
    let pixels = reader().decode().map_err(unreadable)?.into_rgba8();
    let bmp = encode_bmp(width, height, pixels.as_raw(), background);
    debug_assert_eq!(inspect(&bmp), Ok((Format::Bmp, width, height)));
    Ok(Picture {
        data: bmp,
        format: Format::Bmp,
        width,
        height,
        converted_from: Some(name),
    })
}

/// Turns the contents of a picture file into what is stored in a form: the
/// file itself when it is a BMP or JPEG MS-Forms can hold, otherwise a
/// 24-bit BMP made from it. Transparent pixels of a converted picture are
/// blended onto `background` (red, green, blue).
pub fn prepare(data: &[u8], background: [u8; 3]) -> Result<Picture, String> {
    if data.is_empty() || data.len() > MAX_BYTES {
        return Err(format!(
            "画像は 1 バイト以上 {MAX_BYTES} バイト以下にしてください"
        ));
    }
    if data.starts_with(&[0xff, 0xd8]) {
        let (width, height) = inspect_jpeg(data)?;
        return Ok(Picture {
            data: data.to_vec(),
            format: Format::Jpeg,
            width,
            height,
            converted_from: None,
        });
    }
    if data.starts_with(b"BM") {
        if let Ok((width, height)) = inspect_bmp(data) {
            return Ok(Picture {
                data: data.to_vec(),
                format: Format::Bmp,
                width,
                height,
                converted_from: None,
            });
        }
    }
    convert(data, background)
}

/// Reads the picture at the project-root-relative `path` and prepares it
/// for a form (see [`prepare`]) with the default background.
pub fn load(root: &Path, path: &str) -> Result<Picture, String> {
    load_with(root, path, DEFAULT_BACKGROUND)
}

/// Like [`load`], blending transparent pixels onto `background`.
pub fn load_with(root: &Path, path: &str, background: [u8; 3]) -> Result<Picture, String> {
    let data = read_file(root, path)?;
    prepare(&data, background).map_err(|e| format!("画像 {path:?}: {e}"))
}

/// The contents of the picture file at the project-root-relative `path`.
fn read_file(root: &Path, path: &str) -> Result<Vec<u8>, String> {
    let file = resolve(root, path)?;
    // A link must not lead out of the project.
    if let (Ok(real_root), Ok(real_file)) = (root.canonicalize(), file.canonicalize()) {
        if !real_file.starts_with(&real_root) {
            return Err(format!(
                "画像 {path:?} はプロジェクトの外を指しています"
            ));
        }
    }
    let meta =
        std::fs::metadata(&file).map_err(|e| format!("画像 {path:?} を開けません: {e}"))?;
    if !meta.is_file() {
        return Err(format!("画像 {path:?} は通常のファイルではありません"));
    }
    if meta.len() == 0 || meta.len() > MAX_BYTES as u64 {
        return Err(format!(
            "画像 {path:?} は 1 バイト以上 {MAX_BYTES} バイト以下にしてください"
        ));
    }
    std::fs::read(&file).map_err(|e| format!("画像 {path:?} を読めません: {e}"))
}

/// Turns the contents of an icon or cursor file into what is stored for a `MouseIcon`: the
/// file itself. A BMP or JPEG that MS-Forms holds is taken as it is too (Excel may have
/// stored one), but nothing is converted.
pub fn prepare_icon(data: &[u8]) -> Result<Picture, String> {
    let (format, width, height) = inspect(data).map_err(|_| {
        "MouseIcon にはアイコン（.ico）かカーソル（.cur）のファイルを指定してください".to_string()
    })?;
    Ok(Picture {
        data: data.to_vec(),
        format,
        width,
        height,
        converted_from: None,
    })
}

/// Reads the icon or cursor at the project-root-relative `path` for a `MouseIcon`.
pub fn load_icon(root: &Path, path: &str) -> Result<Picture, String> {
    let data = read_file(root, path)?;
    prepare_icon(&data).map_err(|e| format!("画像 {path:?}: {e}"))
}

/// Wraps a stored picture (one that passes [`inspect`]) in the
/// `StdPicture` form of a `GuidAndPicture` record.
pub fn encode(data: &[u8]) -> Result<Vec<u8>, String> {
    inspect(data)?;
    let mut raw = Vec::with_capacity(24 + data.len());
    raw.extend_from_slice(&STD_PICTURE_CLSID);
    raw.extend_from_slice(&STD_PICTURE_TAG.to_le_bytes());
    raw.extend_from_slice(&(data.len() as u32).to_le_bytes());
    raw.extend_from_slice(data);
    Ok(raw)
}

/// Unwraps a stored `StdPicture` payload.
pub fn decode(raw: &[u8]) -> Result<Picture, String> {
    if raw.len() < 24 {
        return Err("画像のデータが途中で切れています".into());
    }
    if raw[..16] != STD_PICTURE_CLSID {
        return Err("StdPicture ではない画像です".into());
    }
    if u32le(raw, 16) != STD_PICTURE_TAG {
        return Err("StdPicture の先頭が不正です".into());
    }
    if u32le(raw, 20) as usize != raw.len() - 24 {
        return Err("StdPicture の長さがデータと合いません".into());
    }
    let (format, width, height) = inspect(&raw[24..])?;
    Ok(Picture {
        data: raw[24..].to_vec(),
        format,
        width,
        height,
        converted_from: None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn bmp(width: i32, height: i32, bits: u16) -> Vec<u8> {
        let row = ((width.unsigned_abs() as usize * bits as usize).div_ceil(32)) * 4;
        let pixels = row * height.unsigned_abs() as usize;
        let mut data = vec![0u8; 54 + pixels];
        data[..2].copy_from_slice(b"BM");
        data[2..6].copy_from_slice(&((54 + pixels) as u32).to_le_bytes());
        data[10..14].copy_from_slice(&54u32.to_le_bytes());
        data[14..18].copy_from_slice(&40u32.to_le_bytes());
        data[18..22].copy_from_slice(&width.to_le_bytes());
        data[22..26].copy_from_slice(&height.to_le_bytes());
        data[26..28].copy_from_slice(&1u16.to_le_bytes());
        data[28..30].copy_from_slice(&bits.to_le_bytes());
        data
    }

    fn jpeg(width: u16, height: u16) -> Vec<u8> {
        let mut data = vec![0xff, 0xd8, 0xff, 0xe0, 0, 4, 0, 0];
        data.extend_from_slice(&[0xff, 0xc0, 0, 11, 8]);
        data.extend_from_slice(&height.to_be_bytes());
        data.extend_from_slice(&width.to_be_bytes());
        data.extend_from_slice(&[1, 1, 0x11, 0]);
        data.extend_from_slice(&[0xff, 0xda, 0, 2, 0x12, 0x34, 0xff, 0xd9]);
        data
    }

    #[test]
    fn bitmaps() {
        assert_eq!(inspect(&bmp(3, 2, 24)), Ok((Format::Bmp, 3, 2)));
        // Top-down bitmaps have a negative height.
        assert_eq!(inspect(&bmp(3, -2, 32)), Ok((Format::Bmp, 3, 2)));
        assert!(inspect(&bmp(3, 2, 8)).is_err());
        assert!(inspect(&bmp(0, 2, 24)).is_err());
        let mut v5 = bmp(3, 2, 24);
        v5[14] = 124;
        assert!(inspect(&v5).unwrap_err().contains("BITMAPINFOHEADER"));
        let mut cut = bmp(3, 2, 24);
        cut.truncate(60);
        assert!(inspect(&cut).is_err());
        assert!(inspect(&bmp(5000, 5000, 24)).is_err(), "pixel limit");
    }

    #[test]
    fn jpegs() {
        assert_eq!(inspect(&jpeg(24, 16)), Ok((Format::Jpeg, 24, 16)));
        assert!(inspect(&jpeg(0, 16)).is_err());
        assert!(inspect(&[0xff, 0xd8, 0xff]).is_err());
        let mut no_frame = vec![0xff, 0xd8];
        no_frame.extend_from_slice(&[0xff, 0xda, 0, 2, 0xff, 0xd9]);
        assert!(inspect(&no_frame).is_err());
        assert!(inspect(b"\x89PNG\r\n\x1a\n").unwrap_err().contains("BMP か JPEG"));
        assert!(inspect(b"").is_err());
    }

    fn png(width: u32, height: u32, pixel: [u8; 4]) -> Vec<u8> {
        let image = image::RgbaImage::from_pixel(width, height, image::Rgba(pixel));
        let mut data = std::io::Cursor::new(Vec::new());
        image.write_to(&mut data, image::ImageFormat::Png).unwrap();
        data.into_inner()
    }

    #[test]
    fn stored_formats_are_kept_as_they_are() {
        for data in [bmp(3, 2, 24), bmp(3, -2, 32), jpeg(24, 16)] {
            let picture = prepare(&data, DEFAULT_BACKGROUND).unwrap();
            assert_eq!(picture.data, data);
            assert_eq!(picture.converted_from, None);
        }
        // A broken JPEG is not something to convert.
        assert!(prepare(&[0xff, 0xd8, 0xff], DEFAULT_BACKGROUND).is_err());
    }

    #[test]
    fn other_formats_become_24_bit_bitmaps() {
        // 3x2, opaque red: rows are padded to 4 bytes, stored bottom up as B, G, R.
        let picture = prepare(&png(3, 2, [255, 0, 0, 255]), DEFAULT_BACKGROUND).unwrap();
        assert_eq!(picture.converted_from, Some("PNG"));
        assert_eq!((picture.format, picture.width, picture.height), (Format::Bmp, 3, 2));
        assert_eq!(picture.data.len(), 54 + 12 * 2);
        assert_eq!(inspect(&picture.data), Ok((Format::Bmp, 3, 2)));
        assert_eq!(&picture.data[54..66], [0, 0, 255, 0, 0, 255, 0, 0, 255, 0, 0, 0]);
        assert!(encode(&picture.data).is_ok());

        // Transparency is blended onto the background.
        let clear = prepare(&png(1, 1, [255, 0, 0, 0]), [10, 20, 30]).unwrap();
        assert_eq!(&clear.data[54..57], [30, 20, 10]);
        let half = prepare(&png(1, 1, [255, 255, 255, 128]), [0, 0, 0]).unwrap();
        assert_eq!(&half.data[54..57], [128, 128, 128]);

        // The same picture gives the same bytes every time.
        assert_eq!(prepare(&png(3, 2, [1, 2, 3, 255]), [0; 3]), prepare(&png(3, 2, [1, 2, 3, 255]), [0; 3]));

        // The orientation survives: top-left pixel of the source is in the last stored row.
        let mut image = image::RgbaImage::from_pixel(2, 2, image::Rgba([0, 0, 0, 255]));
        image.put_pixel(0, 0, image::Rgba([255, 255, 255, 255]));
        let mut data = std::io::Cursor::new(Vec::new());
        image.write_to(&mut data, image::ImageFormat::Png).unwrap();
        let picture = prepare(&data.into_inner(), [0; 3]).unwrap();
        assert_eq!(&picture.data[54..62], [0, 0, 0, 0, 0, 0, 0, 0]);
        assert_eq!(&picture.data[62..70], [255, 255, 255, 0, 0, 0, 0, 0]);
    }

    #[test]
    fn every_convertible_format_is_read() {
        let source = image::RgbaImage::from_fn(4, 3, |x, y| image::Rgba([x as u8 * 60, y as u8 * 80, 7, 255]));
        let expected = prepare(&{
            let mut data = std::io::Cursor::new(Vec::new());
            source.write_to(&mut data, image::ImageFormat::Png).unwrap();
            data.into_inner()
        }, [0; 3])
        .unwrap();
        // Lossless formats give the same bitmap as the PNG.
        for (format, name) in [
            (image::ImageFormat::Gif, "GIF"),
            (image::ImageFormat::Tiff, "TIFF"),
            (image::ImageFormat::Ico, "ICO"),
            (image::ImageFormat::WebP, "WebP"),
        ] {
            let mut data = std::io::Cursor::new(Vec::new());
            source.write_to(&mut data, format).unwrap();
            let picture = prepare(&data.into_inner(), [0; 3]).unwrap_or_else(|e| panic!("{name}: {e}"));
            assert_eq!(picture.converted_from, Some(name));
            assert_eq!(picture.data, expected.data, "{name}");
        }
        // A bitmap MS-Forms cannot hold (8 bits per pixel with a palette).
        let mut indexed = vec![0u8; 54 + 8 + 4];
        indexed[..2].copy_from_slice(b"BM");
        indexed[10..14].copy_from_slice(&62u32.to_le_bytes());
        indexed[14..18].copy_from_slice(&40u32.to_le_bytes());
        indexed[18..22].copy_from_slice(&1i32.to_le_bytes());
        indexed[22..26].copy_from_slice(&1i32.to_le_bytes());
        indexed[26..28].copy_from_slice(&1u16.to_le_bytes());
        indexed[28..30].copy_from_slice(&8u16.to_le_bytes());
        indexed[46..50].copy_from_slice(&2u32.to_le_bytes());
        indexed[58..62].copy_from_slice(&[10, 20, 30, 0]);
        indexed[62] = 1;
        assert!(inspect(&indexed).is_err());
        let picture = prepare(&indexed, [0; 3]).unwrap();
        assert_eq!(picture.converted_from, Some("BMP"));
        assert_eq!(&picture.data[54..57], [10, 20, 30]);
    }

    #[test]
    fn pictures_that_cannot_be_converted() {
        assert!(prepare(b"plain text, not a picture", DEFAULT_BACKGROUND).unwrap_err().contains("判別できません"));
        assert!(prepare(b"", DEFAULT_BACKGROUND).is_err());
        // A PNG that is cut off.
        let mut cut = png(8, 8, [1, 2, 3, 255]);
        cut.truncate(cut.len() - 20);
        assert!(prepare(&cut, DEFAULT_BACKGROUND).unwrap_err().contains("PNG として読めません"));
        // 3000x2000 pixels would be an 18 MB bitmap.
        let large = png(3000, 2000, [0, 0, 0, 255]);
        assert!(large.len() < MAX_BYTES);
        assert!(prepare(&large, DEFAULT_BACKGROUND).unwrap_err().contains("BMP に変換すると"));
    }

    #[test]
    fn backgrounds() {
        let color = |kind: &str, value: &str| crate::FormSpecColor { kind: kind.into(), value: value.into() };
        assert_eq!(background_of(None), DEFAULT_BACKGROUND);
        assert_eq!(background_of(Some(&color("rgb", "#102030"))), [0x10, 0x20, 0x30]);
        assert_eq!(background_of(Some(&color("system", "window"))), [255, 255, 255]);
        assert_eq!(background_of(Some(&color("system", "buttonFace"))), [0xf0, 0xf0, 0xf0]);
        assert_eq!(background_of(Some(&color("ole", "0x00302010"))), [0x10, 0x20, 0x30]);
        assert_eq!(background_of(Some(&color("ole", "0x02000000"))), DEFAULT_BACKGROUND);
    }

    #[test]
    fn std_picture_round_trip() {
        let data = bmp(2, 2, 24);
        let raw = encode(&data).unwrap();
        assert_eq!(&raw[16..20], &[0x6c, 0x74, 0, 0]);
        assert_eq!(decode(&raw).unwrap().data, data);
        assert!(decode(&raw[..raw.len() - 1]).is_err());
        assert!(encode(b"GIF89a").is_err());
    }

    #[test]
    fn paths_stay_inside_the_project() {
        let root = Path::new("/project");
        assert_eq!(
            resolve(root, "forms/assets/a.bmp").unwrap(),
            root.join("forms/assets/a.bmp")
        );
        assert_eq!(
            resolve(root, "forms\\assets\\a.bmp").unwrap(),
            root.join("forms/assets/a.bmp")
        );
        for bad in ["", " ", "/etc/a.bmp", "C:/a.bmp", "../a.bmp", "forms/../../a.bmp", "."] {
            assert!(resolve(root, bad).is_err(), "{bad:?}");
        }
    }
}
