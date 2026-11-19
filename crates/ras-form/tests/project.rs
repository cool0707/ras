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
        Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/data/excel-mutationform.bin"),
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
