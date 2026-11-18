//! Helpers shared by the integration tests.
#![allow(dead_code)]

use std::fs;
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};

use ras_ovba::{ModuleKind, VbaModule};

/// Path of a vendored fixture (see `tests/fixtures/README.md`).
pub fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests").join("fixtures").join(name)
}

pub fn fixture_bytes(name: &str) -> Vec<u8> {
    fs::read(fixture(name)).expect("fixture exists")
}

/// A fresh, empty directory below Cargo's temporary directory for
/// integration tests. Left in place after the test so that the produced
/// workbooks can be inspected with external tools.
pub fn fresh_dir(name: &str) -> PathBuf {
    let dir = Path::new(env!("CARGO_TARGET_TMPDIR")).join("ras-ovba").join(name);
    let _ = fs::remove_dir_all(&dir);
    fs::create_dir_all(&dir).expect("create test directory");
    dir
}

pub fn module(kind: ModuleKind, name: &str, source: &str) -> VbaModule {
    VbaModule { name: name.to_string(), kind, source: source.to_string() }
}

pub fn standard(name: &str, source: &str) -> VbaModule {
    module(ModuleKind::Standard, name, source)
}

pub fn class(name: &str, source: &str) -> VbaModule {
    module(ModuleKind::Class, name, source)
}

/// Names of all entries of a zip file, in order.
pub fn zip_names(path: &Path) -> Vec<String> {
    let archive = zip::ZipArchive::new(Cursor::new(fs::read(path).unwrap())).unwrap();
    (0..archive.len()).map(|i| archive.name_for_index(i).unwrap().to_string()).collect()
}

/// Decompressed contents of a zip entry.
pub fn zip_entry(path: &Path, name: &str) -> Option<Vec<u8>> {
    let mut archive = zip::ZipArchive::new(Cursor::new(fs::read(path).unwrap())).unwrap();
    let mut file = archive.by_name(name).ok()?;
    let mut data = Vec::new();
    file.read_to_end(&mut data).unwrap();
    Some(data)
}

pub fn zip_text(path: &Path, name: &str) -> String {
    String::from_utf8(zip_entry(path, name).unwrap_or_else(|| panic!("entry {name} exists"))).unwrap()
}

/// `(crc32, compressed bytes exactly as stored)` of a zip entry.
pub fn zip_raw(path: &Path, name: &str) -> Option<(u32, Vec<u8>)> {
    let mut archive = zip::ZipArchive::new(Cursor::new(fs::read(path).unwrap())).unwrap();
    let index = archive.index_for_name(name)?;
    let mut file = archive.by_index_raw(index).unwrap();
    let crc = file.crc32();
    let mut data = Vec::new();
    file.read_to_end(&mut data).unwrap();
    Some((crc, data))
}

/// Builds a zip file from `(name, contents)` pairs.
pub fn build_zip(entries: &[(&str, Vec<u8>)]) -> Vec<u8> {
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    for (name, data) in entries {
        writer.start_file(*name, options).unwrap();
        writer.write_all(data).unwrap();
    }
    writer.finish().unwrap().into_inner()
}

/// Rebuilds a zip, replacing / adding / removing entries.
pub fn patch_zip(bytes: &[u8], changes: &[(&str, Option<Vec<u8>>)]) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes.to_vec())).unwrap();
    let mut entries: Vec<(String, Vec<u8>)> = Vec::new();
    for i in 0..archive.len() {
        let mut file = archive.by_index(i).unwrap();
        let mut data = Vec::new();
        file.read_to_end(&mut data).unwrap();
        entries.push((file.name().to_string(), data));
    }
    for (name, change) in changes {
        entries.retain(|(n, _)| n != name);
        if let Some(data) = change {
            entries.push((name.to_string(), data.clone()));
        }
    }
    let refs: Vec<(&str, Vec<u8>)> = entries.iter().map(|(n, d)| (n.as_str(), d.clone())).collect();
    build_zip(&refs)
}

/// `(module name, source)` of every module, read with the independent
/// `ovba` crate.
pub fn ovba_modules(bin: &[u8]) -> Vec<(String, String)> {
    let project = ovba::open_project(bin.to_vec()).expect("ovba opens the project");
    project
        .modules
        .iter()
        .map(|m| {
            let source = project.module_source(&m.name).expect("ovba reads the module source");
            (m.name.clone(), source)
        })
        .collect()
}

/// All streams of a compound file as `(path, contents)`, sorted by path.
pub fn cfb_streams(bin: &[u8]) -> Vec<(String, Vec<u8>)> {
    let mut file = cfb::CompoundFile::open(Cursor::new(bin.to_vec())).expect("valid compound file");
    let entries: Vec<cfb::Entry> = file.walk().filter(|e| e.is_stream()).collect();
    let mut result = Vec::new();
    for entry in entries {
        let mut data = Vec::new();
        file.open_stream(entry.path()).unwrap().read_to_end(&mut data).unwrap();
        result.push((entry.path().to_string_lossy().replace('\\', "/"), data));
    }
    result.sort();
    result
}

/// All storages of a compound file as `(path, clsid text)`, sorted by path.
pub fn cfb_storages(bin: &[u8]) -> Vec<(String, String)> {
    let file = cfb::CompoundFile::open(Cursor::new(bin.to_vec())).expect("valid compound file");
    let mut result: Vec<(String, String)> = file
        .walk()
        .filter(|e| e.is_storage())
        .map(|e| (e.path().to_string_lossy().replace('\\', "/"), e.clsid().to_string()))
        .collect();
    result.sort();
    result
}

pub fn cfb_stream(bin: &[u8], path: &str) -> Option<Vec<u8>> {
    cfb_streams(bin).into_iter().find(|(p, _)| p == path).map(|(_, d)| d)
}
