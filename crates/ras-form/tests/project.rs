//! Storing generated forms in a `vbaProject.bin` with `ras-ovba`.

use std::collections::BTreeMap;
use std::io::Read;
use std::path::Path;

use ras_ovba::{DocumentSpec, ModuleKind, ProjectEdit, UserForm};

fn fixture(name: &str) -> ras_form::Designer {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures")
        .join(name);
    let parsed = ras_form::parse(&std::fs::read_to_string(path).unwrap());
    ras_form::compile(parsed.spec.as_ref().expect("valid fixture"), 932).unwrap()
}

fn user_form(designer: &ras_form::Designer, code: &str) -> UserForm {
    UserForm {
        name: designer.name.clone(),
        code: code.into(),
        streams: designer.streams.clone(),
        storages: designer.storages.clone(),
    }
}

/// Every stream below the storage `top`, keyed by its path below it.
fn streams_under(bin: &[u8], top: &str) -> BTreeMap<String, Vec<u8>> {
    let mut file = cfb::CompoundFile::open(std::io::Cursor::new(bin.to_vec())).unwrap();
    let entries: Vec<_> = file
        .walk()
        .filter(|e| e.is_stream())
        .map(|e| e.path().to_path_buf())
        .collect();
    let mut out = BTreeMap::new();
    for path in entries {
        let parts: Vec<String> = path
            .iter()
            .skip(1)
            .map(|p| p.to_string_lossy().to_string())
            .collect();
        if parts.first().map(String::as_str) != Some(top) {
            continue;
        }
        let mut data = Vec::new();
        file.open_stream(&path)
            .unwrap()
            .read_to_end(&mut data)
            .unwrap();
        out.insert(parts[1..].join("/"), data);
    }
    out
}

fn storage_clsid(bin: &[u8], path: &str) -> [u8; 16] {
    let file = cfb::CompoundFile::open(std::io::Cursor::new(bin.to_vec())).unwrap();
    file.entry(path).unwrap().clsid().to_bytes_le()
}

fn dir_text(bin: &[u8]) -> Vec<u8> {
    let mut file = cfb::CompoundFile::open(std::io::Cursor::new(bin.to_vec())).unwrap();
    let mut data = Vec::new();
    file.open_stream("/VBA/dir")
        .unwrap()
        .read_to_end(&mut data)
        .unwrap();
    ras_ovba::compression::decompress(&data).unwrap()
}

fn count(haystack: &[u8], needle: &[u8]) -> usize {
    haystack
        .windows(needle.len())
        .filter(|w| *w == needle)
        .count()
}

fn source_of(bin: &[u8], name: &str) -> (ModuleKind, String) {
    let info = ras_ovba::read_project_bin(bin).unwrap();
    let module = info.modules.into_iter().find(|m| m.name == name).unwrap();
    (module.kind, module.source)
}

#[test]
fn new_project_with_a_generated_form() {
    let designer = fixture("05-multipage.form.json");
    let edit = ProjectEdit {
        forms: vec![user_form(
            &designer,
            "Private Sub UserForm_Initialize()\nEnd Sub",
        )],
        ..ProjectEdit::default()
    };
    let documents = [DocumentSpec {
        code_name: "ThisWorkbook".into(),
        is_workbook: true,
    }];
    let bin = ras_ovba::write_project_bin(None, &documents, &edit).unwrap();

    let (kind, source) = source_of(&bin, "Pages");
    assert_eq!(kind, ModuleKind::Form);
    let lines: Vec<&str> = source.split("\r\n").collect();
    assert_eq!(lines[0], "Attribute VB_Name = \"Pages\"");
    assert!(
        lines[1].starts_with("Attribute VB_Base = \"0{")
            && lines[1].len() == "Attribute VB_Base = \"0{}{}\"".len() + 72,
        "{}",
        lines[1]
    );
    assert_eq!(lines[4], "Attribute VB_PredeclaredId = True");
    assert_eq!(
        &lines[8..10],
        ["Private Sub UserForm_Initialize()", "End Sub"]
    );

    // The designer storage is stored exactly as generated.
    let stored = streams_under(&bin, "Pages");
    assert_eq!(stored, designer.streams);
    assert_eq!(storage_clsid(&bin, "/Pages/i01"), designer.storages["i01"]);
    assert_eq!(
        storage_clsid(&bin, "/Pages/i01/i05/i06"),
        designer.storages["i01/i05/i06"]
    );

    // PROJECT names it as a designer; the MSForms reference was added once.
    let mut file = cfb::CompoundFile::open(std::io::Cursor::new(bin.clone())).unwrap();
    let mut project = String::new();
    file.open_stream("/PROJECT")
        .unwrap()
        .read_to_string(&mut project)
        .unwrap();
    assert!(project.contains("BaseClass=Pages\r\n"), "{project}");
    assert_eq!(
        count(&dir_text(&bin), b"{0D452EE1-E08F-101A-852E-02608C4D0BB4}"),
        1
    );

    // Writing the same form again replaces it in place and keeps its identity.
    let again = ProjectEdit {
        forms: vec![user_form(&fixture("05-multipage.form.json"), "")],
        ..ProjectEdit::default()
    };
    let bin2 = ras_ovba::write_project_bin(Some(&bin), &documents, &again).unwrap();
    let (_, source2) = source_of(&bin2, "Pages");
    assert_eq!(
        source2.split("\r\n").take(8).collect::<Vec<_>>(),
        lines[..8]
    );
    assert!(!source2.contains("UserForm_Initialize"));
    assert_eq!(streams_under(&bin2, "Pages"), designer.streams);
    assert_eq!(
        count(&dir_text(&bin2), b"{0D452EE1-E08F-101A-852E-02608C4D0BB4}"),
        1
    );
}

#[test]
fn excel_project_gets_a_new_form_and_keeps_its_own() {
    // xl/vbaProject.bin saved by Excel 16 with the hand-made form
    // `MutationForm` (from xlflow's test data).
    let original = std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/data/mutation.bin"),
    )
    .unwrap();
    let before = streams_under(&original, "MutationForm");
    assert!(before.contains_key("f"));
    let msforms_before = count(
        &dir_text(&original),
        b"{0D452EE1-E08F-101A-852E-02608C4D0BB4}",
    );

    let designer = fixture("03-all-controls.form.json");
    let edit = ProjectEdit {
        forms: vec![user_form(&designer, "")],
        ..ProjectEdit::default()
    };
    let bin = ras_ovba::write_project_bin(Some(&original), &[], &edit).unwrap();
    assert_eq!(streams_under(&bin, "MutationForm"), before);
    assert_eq!(streams_under(&bin, "AllControls"), designer.streams);
    assert_eq!(source_of(&bin, "AllControls").0, ModuleKind::Form);
    assert_eq!(
        count(&dir_text(&bin), b"{0D452EE1-E08F-101A-852E-02608C4D0BB4}"),
        msforms_before
    );

    // Replacing the designer of the Excel form keeps the module header.
    let (_, old_source) = source_of(&original, "MutationForm");
    let mut spec = ras_form::FormSpec::new_form("MutationForm");
    spec.form.caption = Some("Replaced".into());
    let replacement = ras_form::compile(&spec, 932).unwrap();
    let edit = ProjectEdit {
        forms: vec![user_form(&replacement, "Sub X()\nEnd Sub\n")],
        ..ProjectEdit::default()
    };
    let bin = ras_ovba::write_project_bin(Some(&original), &[], &edit).unwrap();
    assert_eq!(streams_under(&bin, "MutationForm"), replacement.streams);
    let (_, new_source) = source_of(&bin, "MutationForm");
    let header = |s: &str| {
        s.split("\r\n")
            .take_while(|l| l.starts_with("Attribute "))
            .map(str::to_string)
            .collect::<Vec<_>>()
    };
    assert_eq!(header(&new_source), header(&old_source));
    assert!(new_source.ends_with("Sub X()\r\nEnd Sub\r\n"));
}

#[test]
fn bad_requests_are_refused() {
    let designer = fixture("01-minimal.form.json");
    let mut form = user_form(&designer, "Attribute VB_Exposed = True\n");
    let refused = ras_ovba::write_project_bin(
        None,
        &[],
        &ProjectEdit {
            forms: vec![form.clone()],
            ..ProjectEdit::default()
        },
    );
    assert!(refused.is_err());
    form.code.clear();
    form.streams.remove("o");
    assert!(ras_ovba::write_project_bin(
        None,
        &[],
        &ProjectEdit {
            forms: vec![form],
            ..ProjectEdit::default()
        }
    )
    .is_err());
    // A standard module of the same name is not silently turned into a form.
    let module = ras_ovba::VbaModule {
        name: "MinimalForm".into(),
        kind: ModuleKind::Standard,
        source: String::new(),
    };
    let bin = ras_ovba::write_project_bin(
        None,
        &[],
        &ProjectEdit {
            upsert: vec![module],
            ..ProjectEdit::default()
        },
    )
    .unwrap();
    let edit = ProjectEdit {
        forms: vec![user_form(&designer, "")],
        ..ProjectEdit::default()
    };
    assert!(ras_ovba::write_project_bin(Some(&bin), &[], &edit).is_err());
}

#[test]
fn forms_are_read_back_from_a_project() {
    let designer = fixture("05-multipage.form.json");
    let plain = fixture("01-minimal.form.json");
    let edit = ProjectEdit {
        forms: vec![
            user_form(&designer, "Private Sub UserForm_Initialize()\nEnd Sub"),
            user_form(&plain, ""),
        ],
        ..ProjectEdit::default()
    };
    let documents = [DocumentSpec {
        code_name: "ThisWorkbook".into(),
        is_workbook: true,
    }];
    let bin = ras_ovba::write_project_bin(None, &documents, &edit).unwrap();

    let forms = ras_ovba::read_user_forms(&bin).unwrap();
    assert_eq!(
        forms.iter().map(|f| f.name.as_str()).collect::<Vec<_>>(),
        ["Pages", "MinimalForm"]
    );
    assert_eq!(forms[0].streams, designer.streams);
    assert_eq!(forms[0].storages, designer.storages);
    assert_eq!(
        forms[0].code,
        "Private Sub UserForm_Initialize()\nEnd Sub\n"
    );
    assert_eq!(forms[1].code, "");
    assert!(!ras_ovba::is_generated_form(&forms[0].code));

    // The forms of a project saved by Excel.
    let excel = std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/data/generation.bin"),
    )
    .unwrap();
    let forms = ras_ovba::read_user_forms(&excel).unwrap();
    let names: Vec<&str> = forms.iter().map(|f| f.name.as_str()).collect();
    assert_eq!(names, ["GenerationBaseline", "EmptyForm"]);
    for form in &forms {
        assert!(form.streams.contains_key("f") && form.streams.contains_key("\u{3}VBFrame"));
        assert_eq!(form.storages[""], [0; 16]);
    }
}

/// `CompObj` of the root storage of a `.frx` file exported by Excel 16
/// (identical in every export looked at).
const FRX_COMPOBJ: &str = "0100feff030a0000fffffffff0692ac6dc16ce119e9800aa00574a4f190000004d6963726f736f667420466f726d7320322e3020466f726d0010000000456d626564646564204f626a656374000d000000466f726d732e466f726d2e3100f439b271000000000000000000000000";

#[test]
fn a_form_is_exported_as_frm_and_frx() {
    let designer = fixture("04-frames.form.json");
    let form = user_form(&designer, "Option Explicit\n\nPrivate Sub UserForm_Click()\nEnd Sub\n");
    let files = ras_ovba::export_user_form(&form, 932).unwrap();

    // The text part: designer block, attributes, code.
    let lines: Vec<&str> = files.frm.split("\r\n").collect();
    assert_eq!(lines[0], "VERSION 5.00");
    assert_eq!(
        lines[1],
        format!("Begin {{C62A69F0-16DC-11CE-9E98-00AA00574A4F}} {}", form.name)
    );
    let end = lines.iter().position(|l| *l == "End").unwrap();
    let keys: Vec<&str> = lines[2..end]
        .iter()
        .map(|l| l.split('=').next().unwrap().trim())
        .collect();
    assert_eq!(
        keys,
        ["Caption", "ClientHeight", "ClientLeft", "ClientTop", "ClientWidth", "OleObjectBlob", "StartUpPosition"]
    );
    assert!(lines[2..end]
        .contains(&format!("   OleObjectBlob   =   \"{}.frx\":0000", form.name).as_str()));
    assert_eq!(
        &lines[end + 1..end + 6],
        [
            format!("Attribute VB_Name = \"{}\"", form.name).as_str(),
            "Attribute VB_GlobalNameSpace = False",
            "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = True",
            "Attribute VB_Exposed = False",
        ]
    );
    assert_eq!(lines[end + 6], "Option Explicit");
    assert!(files.frm.ends_with("End Sub\r\n"));

    // The binary part: a 24-byte header, then a compound file.
    let frx = &files.frx;
    assert_eq!(&frx[..4], b"LB\x08\x00");
    assert_eq!(
        u32::from_le_bytes(frx[4..8].try_into().unwrap()) as usize,
        frx.len() - 24
    );
    let int = |at: usize| i32::from_le_bytes(frx[at..at + 4].try_into().unwrap());
    let property = |name: &str| -> i32 {
        lines[2..end]
            .iter()
            .find(|l| l.trim_start().starts_with(name))
            .and_then(|l| l.split('=').nth(1))
            .and_then(|v| v.trim().parse().ok())
            .unwrap()
    };
    assert_eq!((int(8), int(12)), (0, 0));
    assert_eq!(int(16), property("ClientWidth") + 2 * property("ClientLeft"));
    assert_eq!(
        int(20),
        property("ClientHeight") + property("ClientTop") + property("ClientLeft")
    );
    let compound = frx[24..].to_vec();
    assert_eq!(&compound[..8], [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    let file = cfb::CompoundFile::open(std::io::Cursor::new(compound.clone())).unwrap();
    assert_eq!(file.version(), cfb::Version::V3);
    assert_eq!(
        file.root_entry().clsid().to_bytes_le(),
        [0xf0, 0x69, 0x2a, 0xc6, 0xdc, 0x16, 0xce, 0x11, 0x9e, 0x98, 0, 0xaa, 0, 0x57, 0x4a, 0x4f]
    );

    // Everything of the designer storage except VBFrame, whose properties
    // are in the text part; CompObj is the one of an exported form.
    let mut stored = BTreeMap::new();
    let mut file = cfb::CompoundFile::open(std::io::Cursor::new(compound)).unwrap();
    let paths: Vec<_> = file
        .walk()
        .filter(|e| e.is_stream())
        .map(|e| e.path().to_path_buf())
        .collect();
    for path in paths {
        let key: Vec<String> = path
            .iter()
            .skip(1)
            .map(|p| p.to_string_lossy().to_string())
            .collect();
        let mut data = Vec::new();
        file.open_stream(&path).unwrap().read_to_end(&mut data).unwrap();
        stored.insert(key.join("/"), data);
    }
    let mut expected = designer.streams.clone();
    expected.remove("\u{3}VBFrame");
    let compobj: Vec<u8> = (0..FRX_COMPOBJ.len() / 2)
        .map(|i| u8::from_str_radix(&FRX_COMPOBJ[i * 2..i * 2 + 2], 16).unwrap())
        .collect();
    expected.insert("\u{1}CompObj".into(), compobj);
    assert_eq!(stored, expected);
    for (path, clsid) in designer.storages.iter().filter(|(p, _)| !p.is_empty()) {
        assert_eq!(
            file.entry(format!("/{path}")).unwrap().clsid().to_bytes_le(),
            *clsid,
            "{path}"
        );
    }

    // A form without a designer block cannot be exported.
    let mut broken = form.clone();
    broken.streams.remove("\u{3}VBFrame");
    assert!(ras_ovba::export_user_form(&broken, 932).is_err());
}

/// `tests/data/excel-export.frx` was written by Excel 16 (`VBComponent.Export`)
/// for a form whose `.frm` says `ClientHeight = 7032`, `ClientLeft = 108`,
/// `ClientTop = 456`, `ClientWidth = 6180`. Exporting the same designer
/// storage gives the same header and the same streams.
#[test]
fn the_frx_layout_is_the_one_excel_writes() {
    let original = std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/data/excel-export.frx"),
    )
    .unwrap();
    let read = |frx: &[u8]| {
        let mut file = cfb::CompoundFile::open(std::io::Cursor::new(frx[24..].to_vec())).unwrap();
        let clsid = file.root_entry().clsid().to_bytes_le();
        let mut streams = BTreeMap::new();
        for name in ["f", "o", "\u{1}CompObj"] {
            let mut data = Vec::new();
            file.open_stream(format!("/{name}")).unwrap().read_to_end(&mut data).unwrap();
            streams.insert(name.to_string(), data);
        }
        let count = file.walk().count();
        (clsid, streams, count)
    };
    let (clsid, mut streams, count) = read(&original);
    assert_eq!(count, 4, "root, f, o, CompObj");

    // The same designer storage as it would lie in a project.
    let frame = "VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MazeChaseForm \r\n   Caption         =   \"Maze Chase\"\r\n   ClientHeight    =   7032\r\n   ClientLeft      =   108\r\n   ClientTop       =   456\r\n   ClientWidth     =   6180\r\n   StartUpPosition =   1\r\nEnd\r\n";
    let expected = streams.clone();
    streams.insert("\u{3}VBFrame".into(), frame.as_bytes().to_vec());
    let form = UserForm {
        name: "MazeChaseForm".into(),
        code: String::new(),
        streams,
        storages: [(String::new(), [0; 16])].into(),
    };
    let files = ras_ovba::export_user_form(&form, 932).unwrap();
    assert_eq!(files.frx[..24], original[..24], "the 24-byte header");
    assert_eq!(files.frx.len(), original.len());
    let (new_clsid, new_streams, new_count) = read(&files.frx);
    assert_eq!((new_clsid, new_count), (clsid, count));
    assert_eq!(new_streams, expected);
    assert!(files.frm.contains("   OleObjectBlob   =   \"MazeChaseForm.frx\":0000\r\n"));
}

/// What `VBComponent.Export` writes reads back as the same form: the forms of
/// projects saved by Excel, written as `.frm` / `.frx` and read again, have the
/// same designer streams and nested storages and give the same FormSpec.
#[test]
fn an_exported_form_reads_back() {
    for file in ["generation.bin", "frame.bin", "multipage.bin", "pictures.bin", "typed.bin"] {
        let bin = std::fs::read(Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/data").join(file)).unwrap();
        for form in ras_ovba::read_user_forms(&bin).unwrap() {
            let files = ras_ovba::export_user_form(&form, 932).unwrap();
            let back = ras_ovba::import_user_form(&files.frm, &files.frx, 932).unwrap();
            assert_eq!(back.name, form.name);
            assert_eq!(back.code, form.code, "{file} {}", form.name);
            // The root storage: Excel stores it with or without the form's CLSID; the `.frx` has it.
            let nested = |s: &BTreeMap<String, [u8; 16]>| s.iter().filter(|(p, _)| !p.is_empty()).map(|(p, c)| (p.clone(), *c)).collect::<Vec<_>>();
            assert_eq!(nested(&back.storages), nested(&form.storages), "{file} {}", form.name);
            assert!(back.storages.contains_key(""));
            for (path, data) in &form.streams {
                if path != "\u{1}CompObj" && path != "\u{3}VBFrame" {
                    assert_eq!(back.streams.get(path), Some(data), "{file} {} {path:?}", form.name);
                }
            }
            let original = ras_form::snapshot(&form.name, &form.streams, &form.storages, 932).unwrap();
            let read = ras_form::snapshot(&back.name, &back.streams, &back.storages, 932).unwrap();
            assert_eq!(read.spec.to_json(), original.spec.to_json(), "{file} {}", form.name);
        }
    }
    // Not an exported form.
    assert!(ras_ovba::import_user_form("Attribute VB_Name = \"X\"\r\n", b"", 932).is_err());
}
