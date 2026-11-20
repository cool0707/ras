//! Pictures of `Image` controls: reading a BMP or JPEG source file and the
//! `StdPicture` payload MS-OFORMS stores for a picture property.
//!
//! Port of xlflow's `internal/vba/userforms/picture` (Copyright (c) 2026
//! harumiWeb, MIT License; see `NOTICE.md`). JPEG files are checked
//! structurally (marker segments and frame header) instead of being decoded.

use std::path::{Component, Path, PathBuf};

/// Largest picture file accepted.
pub const MAX_BYTES: usize = 16 << 20;
/// Largest picture accepted, in pixels.
pub const MAX_PIXELS: u64 = 16_000_000;

const STD_PICTURE_TAG: u32 = 0x746c;
const STD_PICTURE_CLSID: [u8; 16] = [
    0x04, 0x52, 0xe3, 0x0b, 0x91, 0x8f, 0xce, 0x11, 0x9d, 0xe3, 0x00, 0xaa, 0x00, 0x4b, 0xb8, 0x51,
];

/// File format of a picture.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Format {
    /// Windows bitmap (24 or 32 bit, uncompressed).
    Bmp,
    /// JPEG.
    Jpeg,
}

impl Format {
    /// File name extension (`bmp` / `jpg`).
    pub fn extension(self) -> &'static str {
        match self {
            Format::Bmp => "bmp",
            Format::Jpeg => "jpg",
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

/// Checks that `data` is a picture MS-OFORMS can hold.
pub fn inspect(data: &[u8]) -> Result<(Format, u32, u32), String> {
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

/// Reads and checks the picture at the project-root-relative `path`.
pub fn load(root: &Path, path: &str) -> Result<Picture, String> {
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
    let data = std::fs::read(&file).map_err(|e| format!("画像 {path:?} を読めません: {e}"))?;
    let (format, width, height) = inspect(&data).map_err(|e| format!("画像 {path:?}: {e}"))?;
    Ok(Picture {
        data,
        format,
        width,
        height,
    })
}

/// Wraps a picture file in the `StdPicture` form of a `GuidAndPicture`
/// record.
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
