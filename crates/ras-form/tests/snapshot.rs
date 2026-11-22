//! Reading existing forms: forms saved by Excel, the round trip through
//! generation, and `workspace::pull`.

use std::collections::BTreeMap;
use std::collections::HashMap;
use std::path::{Path, PathBuf};

use ras_form::workspace::{self, PullOutcome};
use ras_ovba::UserForm;
use serde_json::Value;

fn data_dir() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/data")
}

fn forms_of(label: &str) -> Vec<UserForm> {
    let bin = std::fs::read(data_dir().join(format!("{label}.bin"))).unwrap();
    ras_ovba::read_user_forms(&bin).unwrap()
}

fn form(label: &str, name: &str) -> UserForm {
    forms_of(label)
        .into_iter()
        .find(|f| f.name == name)
        .unwrap_or_else(|| panic!("{label}.bin has no form {name}"))
}

/// `(label, form name)` of every form of `tests/data/*.bin` with a `<label>.<Form>.expected.json`.
fn excel_forms() -> Vec<(String, String)> {
    let mut out = Vec::new();
    for entry in std::fs::read_dir(data_dir()).unwrap() {
        let file = entry.unwrap().file_name().to_string_lossy().to_string();
        if let Some(stem) = file.strip_suffix(".expected.json") {
            let (label, name) = stem.split_once('.').unwrap();
            out.push((label.to_string(), name.to_string()));
        }
    }
    out.sort();
    assert!(out.len() >= 9, "expected readings missing");
    out
}

/// First place where two JSON values differ; numbers are compared as `f64`.
fn difference(path: &str, expected: &Value, actual: &Value) -> Option<String> {
    match (expected, actual) {
        (Value::Number(a), Value::Number(b)) if a.as_f64() == b.as_f64() => None,
        (Value::Array(a), Value::Array(b)) => {
            if a.len() != b.len() {
                return Some(format!(
                    "{path}: {} elements, expected {}",
                    b.len(),
                    a.len()
                ));
            }
            a.iter()
                .zip(b)
                .enumerate()
                .find_map(|(i, (x, y))| difference(&format!("{path}[{i}]"), x, y))
        }
        (Value::Object(a), Value::Object(b)) => {
            let keys: std::collections::BTreeSet<&String> = a.keys().chain(b.keys()).collect();
            keys.into_iter()
                .find_map(|key| match (a.get(key), b.get(key)) {
                    (Some(x), Some(y)) => difference(&format!("{path}.{key}"), x, y),
                    (Some(x), None) => Some(format!("{path}.{key}: missing, expected {x}")),
                    (None, Some(y)) => Some(format!("{path}.{key}: unexpected {y}")),
                    (None, None) => None,
                })
        }
        _ if expected == actual => None,
        _ => Some(format!("{path}: {actual}, expected {expected}")),
    }
}

/// What `snapshot` makes of a form saved by Excel: the FormSpec, with its warnings, and the pictures.
fn reading(label: &str, name: &str) -> Value {
    let source = form(label, name);
    let snapshot = ras_form::snapshot(name, &source.streams, &source.storages, 932)
        .unwrap_or_else(|e| panic!("{label}.{name}: {e}"));
    let pictures: Vec<Value> = snapshot
        .pictures
        .iter()
        .map(|p| {
            let control = p
                .control_index
                .map(|i| snapshot.spec.controls[i].name.clone())
                .unwrap_or_default();
            serde_json::json!({
                "control": control, "property": p.property, "format": p.picture.format.extension(),
                "size": p.picture.data.len(), "width": p.picture.width, "height": p.picture.height,
            })
        })
        .collect();
    serde_json::json!({ "spec": serde_json::to_value(&snapshot.spec).unwrap(), "pictures": pictures })
}

/// The forms saved by Excel (`tests/data/*.bin`) are read as `<label>.<Form>.expected.json` says.
/// With `RAS_UPDATE_EXPECTED=1` the files are written instead (check the change with `git diff`).
#[test]
fn forms_saved_by_excel_are_read_as_expected() {
    let update = std::env::var_os("RAS_UPDATE_EXPECTED").is_some();
    let mut failures = Vec::new();
    for (label, name) in excel_forms() {
        let actual = reading(&label, &name);
        let path = data_dir().join(format!("{label}.{name}.expected.json"));
        if update {
            std::fs::write(&path, serde_json::to_string_pretty(&actual).unwrap() + "\n").unwrap();
            continue;
        }
        let expected: Value =
            serde_json::from_str(&std::fs::read_to_string(&path).unwrap()).unwrap();
        if let Some(diff) = difference(&format!("{label}.{name}"), &expected, &actual) {
            failures.push(diff);
        }
        // The spec read is a valid FormSpec.
        let spec: ras_form::FormSpec = serde_json::from_value(actual["spec"].clone()).unwrap();
        let parsed = ras_form::parse(&spec.to_json());
        if parsed.has_errors() && !spec.controls.iter().any(|c| c.control_type == "Control") {
            failures.push(format!("{label}.{name}: not valid: {:?}", parsed.issues));
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

#[test]
fn the_pictures_of_forms_saved_by_excel_are_extracted() {
    let mut seen = 0;
    for (label, name) in excel_forms() {
        let source = form(&label, &name);
        let snapshot = ras_form::snapshot(&name, &source.streams, &source.storages, 932).unwrap();
        for found in &snapshot.pictures {
            let Some(index) = found.control_index else {
                continue;
            };
            // An extracted picture is no longer reported as lost.
            let control = &snapshot.spec.controls[index];
            assert!(
                !control
                    .unsupported
                    .iter()
                    .flatten()
                    .any(|u| u == found.property),
                "{label}.{name}"
            );
            seen += 1;
        }
    }
    assert!(seen >= 3, "no pictures in the test data");
}

fn temp_root(tag: &str) -> PathBuf {
    let dir = std::env::temp_dir().join(format!("ras-form-{tag}-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(&dir).unwrap();
    dir
}

/// What a round trip through `forms/` has to preserve of a control.
fn essence(spec: &ras_form::FormSpec) -> Vec<Value> {
    let name_of = |id: &str| {
        spec.controls
            .iter()
            .find(|c| c.id == id)
            .map(|c| c.name.clone())
            .unwrap_or_default()
    };
    let round = |v: Option<f64>| v.map(|v| (v * 20.0).round() / 20.0);
    spec.controls
        .iter()
        .map(|c| {
            // Generation stores the default value where Excel stores none.
            let value = c.value.clone().or_else(|| match c.control_type.as_str() {
                "TextBox" | "ComboBox" => Some(Value::from("")),
                "CheckBox" | "OptionButton" | "ToggleButton" => Some(Value::from("False")),
                "SpinButton" | "ScrollBar" => Some(Value::from(0)),
                _ => None,
            });
            let text = c.text.clone().or_else(|| (c.control_type == "TextBox").then(String::new));
            serde_json::json!({
                "name": c.name, "type": c.control_type, "parent": name_of(&c.parent_id),
                "left": round(c.left), "top": round(c.top), "width": round(c.width), "height": round(c.height),
                "caption": c.caption, "text": text, "value": value,
                "enabled": c.enabled, "visible": c.visible, "tabIndex": c.tab_index,
                "tag": c.tag, "tip": c.control_tip_text, "selectedIndex": c.selected_index,
                "tabs": c.tabs, "foreColor": c.fore_color, "backColor": c.back_color,
                "borderStyle": c.border_style, "textAlign": c.text_align,
                "pictureAlignment": c.picture_alignment, "pictureSizeMode": c.picture_size_mode,
                "properties": c.properties, "zIndex": c.z_index, "font": c.font,
            })
        })
        .collect()
}

#[test]
fn a_pulled_form_is_generated_again_with_the_same_design() {
    let root = temp_root("roundtrip");
    let mut checked = 0;
    for (label, name) in excel_forms() {
        let source = form(&label, &name);
        let dir = root.join(format!("{label}.{name}"));
        std::fs::create_dir_all(&dir).unwrap();
        let outcome = workspace::pull(&dir, &source, 932, false).unwrap();
        let PullOutcome::Written(pulled) = outcome else {
            panic!("{label}.{name}: nothing written");
        };
        assert!(pulled.spec_path.is_file() && pulled.code_path.is_file());

        let loaded = workspace::load(&dir).unwrap();
        assert_eq!(loaded.len(), 1);
        let first = loaded[0].parsed.spec.clone().unwrap_or_else(|| {
            panic!(
                "{label}.{name}: pulled spec is not valid: {:?}",
                loaded[0].parsed.issues
            )
        });
        let built = workspace::build(&loaded, &dir, 932, &HashMap::new());
        assert_eq!(built.errors(), 0, "{label}.{name}: {:?}", built.problems);
        let generated = &built.forms[0];
        assert!(ras_ovba::is_generated_form(&generated.code));

        // Reading the generated form gives the same design again.
        let again =
            ras_form::snapshot(&name, &generated.streams, &generated.storages, 932).unwrap();
        let (before, after) = (essence(&first), essence(&again.spec));
        assert_eq!(before.len(), after.len(), "{label}.{name}");
        for (a, b) in before.iter().zip(&after) {
            assert_eq!(difference(&format!("{label}.{name}"), a, b), None);
        }
        assert_eq!(first.form.caption, again.spec.form.caption);
        assert_eq!(first.form.width, again.spec.form.width);
        assert_eq!(first.form.height, again.spec.form.height);
        assert_eq!(first.form.font, again.spec.form.font, "{label}.{name}");
        assert_eq!(
            first.form.back_color, again.spec.form.back_color,
            "{label}.{name}"
        );
        assert_eq!(
            first.form.fore_color, again.spec.form.fore_color,
            "{label}.{name}"
        );
        assert_eq!(
            first.form.border_style, again.spec.form.border_style,
            "{label}.{name}"
        );
        assert_eq!(
            again.pictures.len(),
            pulled.pictures.len(),
            "{label}.{name}"
        );
        checked += 1;
    }
    assert!(checked >= 9);
    let _ = std::fs::remove_dir_all(&root);
}

#[test]
fn pull_writes_the_files_of_a_form() {
    let root = temp_root("pull");
    let source = form("pictures", "PictureForm");
    let PullOutcome::Written(pulled) = workspace::pull(&root, &source, 932, false).unwrap() else {
        panic!("nothing written");
    };
    assert_eq!(pulled.spec_path, root.join("forms/PictureForm.form.json"));
    assert_eq!(pulled.code_path, root.join("forms/PictureForm.ras"));
    assert_eq!(
        pulled.pictures,
        [
            "forms/assets/PictureForm.ImageBmp.bmp",
            "forms/assets/PictureForm.ImageJpeg.bmp"
        ]
    );
    for picture in &pulled.pictures {
        assert!(std::fs::read(root.join(picture))
            .unwrap()
            .starts_with(b"BM"));
    }
    let text = std::fs::read_to_string(&pulled.spec_path).unwrap();
    let spec: Value = serde_json::from_str(&text).unwrap();
    assert_eq!(spec["kind"], "ras.userform");
    assert!(
        spec.get("warnings").is_none(),
        "warnings are reported, not stored"
    );
    assert_eq!(
        spec["controls"][0]["picture"]["path"],
        "forms/assets/PictureForm.ImageBmp.bmp"
    );
    // The code-behind is the code pane converted to RAS.
    let code = std::fs::read_to_string(&pulled.code_path).unwrap();
    assert!(
        !code.contains("Attribute VB_") && !code.contains('\r'),
        "{code}"
    );

    // Existing files are left alone …
    std::fs::write(&pulled.code_path, "' edited\n").unwrap();
    match workspace::pull(&root, &source, 932, false).unwrap() {
        PullOutcome::Exists(files) => assert_eq!(files.len(), 2),
        PullOutcome::Written(_) => panic!("existing files were replaced"),
    }
    assert_eq!(
        std::fs::read_to_string(&pulled.code_path).unwrap(),
        "' edited\n"
    );
    // … unless asked, and the same pictures are not written twice.
    let PullOutcome::Written(again) = workspace::pull(&root, &source, 932, true).unwrap() else {
        panic!("nothing written");
    };
    assert_eq!(again.pictures, pulled.pictures);
    assert_eq!(std::fs::read_to_string(&pulled.code_path).unwrap(), code);
    let _ = std::fs::remove_dir_all(&root);
}

/// The code the RAS compiler generates for the `.ras` code-behinds of the project at `root`.
fn compile_forms(root: &Path) -> HashMap<String, String> {
    let (sources, decls) = ras_compiler::load_forms(root).unwrap();
    let config = ras_compiler::Config {
        forms: decls,
        ..Default::default()
    };
    let prog = ras_compiler::analyze(sources, config);
    let out = ras_compiler::build(
        &prog,
        &ras_compiler::BuildOptions {
            no_attributes: true,
            ..Default::default()
        },
    );
    assert!(
        prog.diagnostics
            .iter()
            .chain(&out.diagnostics)
            .all(|d| !d.is_error()),
        "{:?}",
        prog.diagnostics
    );
    out.files
        .iter()
        .map(|f| (f.name.clone(), f.code.clone()))
        .collect()
}

#[test]
fn pull_keeps_what_only_the_spec_knows() {
    let root = temp_root("pull-lists");
    let dir = root.join("forms");
    std::fs::create_dir_all(dir.join("assets")).unwrap();
    let logo = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/assets/logo.bmp");
    std::fs::copy(&logo, dir.join("assets/logo.bmp")).unwrap();
    let spec = r#"{
  "schemaVersion": 1, "kind": "ras.userform", "basis": "designer",
  "form": { "name": "frmList", "caption": "List" },
  "controls": [
    { "id": "c", "name": "cboKind", "type": "ComboBox", "list": ["A", "B"], "selectedIndex": 1 },
    { "id": "i", "name": "imgLogo", "type": "Image", "picture": { "path": "forms/assets/logo.bmp" } },
    { "id": "t", "name": "txtName", "type": "TextBox", "tag": "name", "controlTipText": "Your name" }
  ]
}"#;
    std::fs::write(dir.join("frmList.form.json"), spec).unwrap();
    let ras = "@on(\"UserForm.Initialize\")\nfunction userFormInitialize() {\n    this.caption = \"ready\";\n}\n";
    std::fs::write(dir.join("frmList.ras"), ras).unwrap();

    let built = workspace::build(
        &workspace::load(&root).unwrap(),
        &root,
        932,
        &compile_forms(&root),
    );
    assert_eq!(built.errors(), 0, "{:?}", built.problems);
    let generated = built.forms[0].clone();
    let code = &generated.code;
    assert!(code.contains(".AddItem \"A\""), "{code}");
    assert!(
        code.contains("Private Sub UserForm_Initialize()\n    RAS_FormInit '--- RAS ---\n"),
        "{code}"
    );
    assert!(code.contains("Me.Caption = \"ready\""), "{code}");

    // As if the form came back from a workbook.
    let PullOutcome::Written(pulled) = workspace::pull(&root, &generated, 932, true).unwrap()
    else {
        panic!("nothing written");
    };
    assert_eq!(pulled.warnings, [], "nothing of a generated form is lost");
    assert!(pulled.kept_code, "the code was compiled from the .ras file");
    assert_eq!(std::fs::read_to_string(&pulled.code_path).unwrap(), ras);
    assert_eq!(
        pulled.pictures,
        ["forms/assets/logo.bmp"],
        "the existing file is reused"
    );
    let text = std::fs::read_to_string(&pulled.spec_path).unwrap();
    let spec: Value = serde_json::from_str(&text).unwrap();
    assert_eq!(spec["controls"][0]["list"], serde_json::json!(["A", "B"]));
    assert_eq!(spec["controls"][0]["selectedIndex"], 1);
    assert_eq!(spec["controls"][2]["tag"], "name");
    assert_eq!(spec["controls"][2]["controlTipText"], "Your name");

    // Building the pulled files gives the same form again.
    let rebuilt = workspace::build(
        &workspace::load(&root).unwrap(),
        &root,
        932,
        &compile_forms(&root),
    );
    assert_eq!(rebuilt.errors(), 0, "{:?}", rebuilt.problems);
    assert_eq!(rebuilt.forms[0].streams, generated.streams);
    assert_eq!(rebuilt.forms[0].code, generated.code);

    // In a project without the form, its VBA comes back as RAS.
    let other = temp_root("pull-lists-other");
    let PullOutcome::Written(fresh) = workspace::pull(&other, &generated, 932, false).unwrap()
    else {
        panic!("nothing written");
    };
    assert!(!fresh.kept_code);
    assert_eq!(fresh.todos, []);
    assert_eq!(std::fs::read_to_string(&fresh.code_path).unwrap(), ras);
    let _ = std::fs::remove_dir_all(&root);
    let _ = std::fs::remove_dir_all(&other);
}

#[test]
fn typed_tag_fields_equal_the_property_bag() {
    let spec = |fields: &str| {
        let text = format!(
            r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"F"}},"controls":[{{"id":"t","name":"txt","type":"TextBox",{fields}}}]}}"#
        );
        let parsed = ras_form::parse(&text);
        assert!(!parsed.has_errors(), "{:?}", parsed.issues);
        ras_form::compile(parsed.spec.as_ref().unwrap(), 932).unwrap()
    };
    // The same written in the property bag (`properties`).
    let bag = spec(r#""properties":{"tag":"t1","controlTipText":"tip"}"#);
    let typed = spec(r#""tag":"t1","controlTipText":"tip""#);
    assert_eq!(bag.streams, typed.streams);
    assert_ne!(spec(r#""tag":"other""#).streams, typed.streams);
}

#[test]
fn list_items_do_not_change_the_designer() {
    let spec = |fields: &str| {
        let text = format!(
            r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"F"}},"controls":[{{"id":"c","name":"cbo","type":"ComboBox"{fields}}}]}}"#
        );
        ras_form::parse(&text)
    };
    let plain = ras_form::compile(spec("").spec.as_ref().unwrap(), 932).unwrap();
    let listed = spec(r#","list":["a","b"],"selectedIndex":1"#);
    assert!(listed.issues.is_empty(), "{:?}", listed.issues);
    let listed = ras_form::compile(listed.spec.as_ref().unwrap(), 932).unwrap();
    assert_eq!(plain.streams, listed.streams);

    // A selection has to name an item.
    for bad in [
        r#","list":["a"],"selectedIndex":1"#,
        r#","selectedIndex":0"#,
        r#","list":[],"selectedIndex":-2"#,
    ] {
        let parsed = spec(bad);
        assert!(
            parsed.issues.iter().any(|i| i.code == "UFV020"),
            "{bad}: {:?}",
            parsed.issues
        );
    }
    assert!(!spec(r#","selectedIndex":-1"#).has_errors());
}

#[test]
fn list_properties_go_round_and_default_to_what_excel_writes() {
    let project = |controls: &str| {
        let text = format!(
            r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"F"}},"controls":[{controls}]}}"#
        );
        let parsed = ras_form::parse(&text);
        assert!(!parsed.has_errors(), "{:?}", parsed.issues);
        let designer = ras_form::compile(parsed.spec.as_ref().unwrap(), 932).unwrap();
        let snapshot = ras_form::snapshot("F", &designer.streams, &designer.storages, 932).unwrap();
        snapshot.spec.controls
    };
    // A new ComboBox has a drop button and completes entries, as one Excel makes;
    // none of these is projected, nor left unsupported.
    let plain = project(
        r#"{"id":"c","name":"cbo","type":"ComboBox"},{"id":"l","name":"lst","type":"ListBox","top":30}"#,
    );
    let keys = [
        "style",
        "multiSelect",
        "listStyle",
        "matchEntry",
        "showDropButtonWhen",
        "listRows",
    ];
    for control in &plain {
        let json = serde_json::to_value(control).unwrap();
        for key in keys {
            assert!(
                json.get(key).is_none(),
                "{}: {key} = {}",
                control.name,
                json[key]
            );
        }
        let unsupported = control.unsupported.clone().unwrap_or_default();
        assert!(
            !unsupported.iter().any(|u| keys.contains(&u.as_str())),
            "{unsupported:?}"
        );
    }
    let set = project(
        r#"{"id":"c","name":"cbo","type":"ComboBox","style":2,"listStyle":1,"matchEntry":0,"showDropButtonWhen":1,"listRows":4},
           {"id":"l","name":"lst","type":"ListBox","top":30,"multiSelect":2,"listStyle":1,"matchEntry":2}"#,
    );
    let (combo, list) = (&set[0], &set[1]);
    let values = |c: &ras_form::FormSpecControl, names: &[&str]| -> Vec<Value> {
        names
            .iter()
            .map(|n| c.scalar(n).cloned().unwrap_or(Value::Null))
            .collect()
    };
    assert_eq!(
        values(
            combo,
            &[
                "style",
                "listStyle",
                "matchEntry",
                "showDropButtonWhen",
                "listRows"
            ]
        ),
        [2, 1, 0, 1, 4].map(Value::from)
    );
    assert_eq!(
        values(list, &["multiSelect", "listStyle", "matchEntry"]),
        [2, 1, 2].map(Value::from)
    );

    // Each belongs to its control type only.
    for bad in [
        r#"{"id":"t","name":"txt","type":"TextBox","style":2}"#,
        r#"{"id":"c","name":"cbo","type":"ComboBox","multiSelect":1}"#,
        r#"{"id":"l","name":"lst","type":"ListBox","showDropButtonWhen":0}"#,
        r#"{"id":"c","name":"cbo","type":"ComboBox","style":1}"#,
        r#"{"id":"c","name":"cbo","type":"ComboBox","listRows":-1}"#,
    ] {
        let text = format!(
            r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"F"}},"controls":[{bad}]}}"#
        );
        assert!(ras_form::parse(&text).has_errors(), "{bad}");
    }
}

#[test]
fn a_storage_that_is_not_a_designer_is_refused() {
    let mut source = form("mutation", "MutationForm");
    let f = source.streams.get_mut("f").unwrap();
    f.truncate(f.len() - 7);
    let error =
        ras_form::snapshot("MutationForm", &source.streams, &source.storages, 932).unwrap_err();
    assert!(
        error.to_string().contains("デザイナストレージを読めません"),
        "{error}"
    );

    let empty: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    assert!(ras_form::snapshot("X", &empty, &BTreeMap::new(), 932).is_err());
}

/// A picture MS-Forms cannot store (here a PNG with transparency) is
/// converted when the form is generated; the spec keeps naming the PNG.
#[test]
fn other_picture_formats_are_converted_when_the_form_is_written() {
    let root = temp_root("png");
    let assets = root.join("forms/assets");
    std::fs::create_dir_all(&assets).unwrap();
    let png =
        std::fs::read(Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/assets/logo.png"))
            .unwrap();
    std::fs::write(assets.join("logo.png"), &png).unwrap();
    let spec = |path: &str, back: &str| {
        format!(
            r##"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"frmPng"}},
            "controls":[{{"id":"i","name":"imgLogo","type":"Image","picture":{{"path":"{path}"}}{back}}}]}}"##
        )
    };
    let build = |text: &str| {
        std::fs::write(root.join("forms/frmPng.form.json"), text).unwrap();
        let built = workspace::build(
            &workspace::load(&root).unwrap(),
            &root,
            932,
            &HashMap::new(),
        );
        assert_eq!(built.errors(), 0, "{:?}", built.problems);
        built.forms[0].clone()
    };

    // The form holds the bitmap the PNG converts to …
    let from_png = build(&spec("forms/assets/logo.png", ""));
    let converted =
        ras_form::picture::prepare(&png, ras_form::picture::DEFAULT_BACKGROUND).unwrap();
    assert_eq!(converted.converted_from, Some("PNG"));
    assert_eq!((converted.width, converted.height), (24, 16));
    // The transparent frame shows the background, the opaque centre the picture.
    assert_eq!(&converted.data[54..57], [0xf0, 0xf0, 0xf0]);
    let stored = ras_form::snapshot("frmPng", &from_png.streams, &from_png.storages, 932).unwrap();
    assert_eq!(stored.pictures[0].picture.data, converted.data);

    // … exactly as if that bitmap had been named in the spec.
    std::fs::write(assets.join("converted.bmp"), &converted.data).unwrap();
    let from_bmp = build(&spec("forms/assets/converted.bmp", ""));
    assert_eq!(from_png.streams, from_bmp.streams);

    // Transparent pixels take the control's background color.
    let on_red = build(&spec(
        "forms/assets/logo.png",
        r##","backColor":{"kind":"rgb","value":"#FF0000"}"##,
    ));
    let stored = ras_form::snapshot("frmPng", &on_red.streams, &on_red.storages, 932).unwrap();
    assert_eq!(
        &stored.pictures[0].picture.data[54..57],
        [0, 0, 255],
        "blue, green, red"
    );

    // Pulling the form back keeps the PNG the spec names.
    let PullOutcome::Written(pulled) = workspace::pull(&root, &on_red, 932, true).unwrap() else {
        panic!("nothing written");
    };
    assert_eq!(pulled.pictures, ["forms/assets/logo.png"]);
    assert!(!assets.join("frmPng.imgLogo.bmp").exists());
    // Into a project without the PNG, the stored bitmap comes out.
    let other = temp_root("png-other");
    let PullOutcome::Written(pulled) = workspace::pull(&other, &on_red, 932, false).unwrap() else {
        panic!("nothing written");
    };
    assert_eq!(pulled.pictures, ["forms/assets/frmPng.imgLogo.bmp"]);
    assert!(std::fs::read(other.join(&pulled.pictures[0]))
        .unwrap()
        .starts_with(b"BM"));

    // A file that is no picture is still an error that names the file.
    std::fs::write(assets.join("notes.png"), "not a picture").unwrap();
    std::fs::write(
        root.join("forms/frmPng.form.json"),
        spec("forms/assets/notes.png", ""),
    )
    .unwrap();
    let built = workspace::build(
        &workspace::load(&root).unwrap(),
        &root,
        932,
        &HashMap::new(),
    );
    assert_eq!(built.errors(), 1);
    assert!(
        built.problems[0].message.contains("notes.png"),
        "{:?}",
        built.problems
    );
    let _ = std::fs::remove_dir_all(&root);
    let _ = std::fs::remove_dir_all(&other);
}

/// The designer storage of a form saved by Excel (`data/triple-state/`): two
/// OptionButtons and two CheckBoxes, the first of each with TripleState True.
fn triple_state_form() -> (BTreeMap<String, Vec<u8>>, BTreeMap<String, [u8; 16]>) {
    let dir = data_dir().join("triple-state");
    let streams = [
        ("f", "f"),
        ("o", "o"),
        ("\u{1}CompObj", "CompObj"),
        ("\u{3}VBFrame", "VBFrame"),
    ]
    .into_iter()
    .map(|(stream, file)| (stream.to_string(), std::fs::read(dir.join(file)).unwrap()))
    .collect();
    (streams, BTreeMap::from([(String::new(), [0u8; 16])]))
}

#[test]
fn triple_state_is_read_from_excel_and_written_back() {
    let (streams, storages) = triple_state_form();
    let snapshot = ras_form::snapshot("UserForm1", &streams, &storages, 932).unwrap();
    let controls = &snapshot.spec.controls;
    let triple: Vec<(&str, Option<&Value>)> = controls
        .iter()
        .map(|c| (c.name.as_str(), c.scalar("tripleState")))
        .collect();
    assert_eq!(
        triple,
        [
            ("OptionButton1", Some(&Value::Bool(true))),
            ("OptionButton2", None),
            ("CheckBox1", Some(&Value::Bool(true))),
            ("CheckBox2", None),
        ]
    );
    for control in controls {
        let unsupported = control.unsupported.clone().unwrap_or_default();
        assert!(
            !unsupported.contains(&"multiSelect".to_string()),
            "{}: {unsupported:?}",
            control.name
        );
        assert_eq!(
            control.value,
            Some(Value::from("False")),
            "{}",
            control.name
        );
    }
    // Written again, the toggles hold the same records as Excel's.
    let designer = ras_form::compile(&snapshot.spec, 932).unwrap();
    let again =
        ras_form::snapshot("UserForm1", &designer.streams, &designer.storages, 932).unwrap();
    for (a, b) in controls.iter().zip(&again.spec.controls) {
        assert_eq!(
            a.scalar("tripleState"),
            b.scalar("tripleState"),
            "{}",
            a.name
        );
    }
}

#[test]
fn a_toggle_can_be_null() {
    let spec = |value: &str| {
        let text = format!(
            r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"F"}},"controls":[
                {{"id":"c","name":"chk","type":"CheckBox","tripleState":true,"value":{value}}},
                {{"id":"t","name":"tgl","type":"ToggleButton","top":30,"value":{value}}}]}}"#
        );
        let parsed = ras_form::parse(&text);
        assert!(!parsed.has_errors(), "{:?}", parsed.issues);
        parsed.spec.unwrap()
    };
    for (written, read) in [
        ("\"Null\"", "Null"),
        ("\"null\"", "Null"),
        ("true", "True"),
        ("0", "False"),
    ] {
        let designer = ras_form::compile(&spec(written), 932).unwrap();
        let snapshot = ras_form::snapshot("F", &designer.streams, &designer.storages, 932).unwrap();
        for control in &snapshot.spec.controls {
            assert_eq!(
                control.value,
                Some(Value::from(read)),
                "{written}: {}",
                control.name
            );
        }
    }
    let error = ras_form::compile(&spec("\"maybe\""), 932).unwrap_err();
    assert!(error.reason.contains("Null"), "{error}");
}

#[test]
fn controls_are_bound_to_cells() {
    let text = r#"{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{"name":"F"},"controls":[
        {"id":"t","name":"txt","type":"TextBox","controlSource":"Sheet1!B2","tag":"t","controlTipText":"tip"},
        {"id":"l","name":"lst","type":"ListBox","top":30,"rowSource":"'売上 2026'!A2:B9","controlSource":"Sheet1!C2"}]}"#;
    let parsed = ras_form::parse(text);
    assert!(!parsed.has_errors(), "{:?}", parsed.issues);
    let designer = ras_form::compile(parsed.spec.as_ref().unwrap(), 932).unwrap();
    let snapshot = ras_form::snapshot("F", &designer.streams, &designer.storages, 932).unwrap();
    let (txt, lst) = (&snapshot.spec.controls[0], &snapshot.spec.controls[1]);
    assert_eq!(txt.scalar("controlSource"), Some(&Value::from("Sheet1!B2")));
    assert_eq!(
        (txt.tag.as_deref(), txt.control_tip_text.as_deref()),
        (Some("t"), Some("tip"))
    );
    assert_eq!(
        lst.scalar("rowSource"),
        Some(&Value::from("'売上 2026'!A2:B9"))
    );
    assert_eq!(lst.scalar("controlSource"), Some(&Value::from("Sheet1!C2")));
    for control in [txt, lst] {
        assert_eq!(control.unsupported, None, "{}", control.name);
    }

    // A list from cells cannot be filled by AddItem as well; a Label has no ControlSource.
    for bad in [
        r#"{"id":"l","name":"lst","type":"ListBox","rowSource":"A1:A3","list":["a"]}"#,
        r#"{"id":"b","name":"lbl","type":"Label","controlSource":"A1"}"#,
        r#"{"id":"t","name":"txt","type":"TextBox","rowSource":"A1:A3"}"#,
    ] {
        let text = format!(
            r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"F"}},"controls":[{bad}]}}"#
        );
        assert!(ras_form::parse(&text).has_errors(), "{bad}");
    }
}

#[test]
fn columns_of_a_list_go_round() {
    let text = r#"{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{"name":"F"},"controls":[
        {"id":"l","name":"lst","type":"ListBox","columnCount":3,"columnWidths":"36;;1 in","boundColumn":2,"textColumn":3,"columnHeads":true,"list":[["a","b","c"]]},
        {"id":"c","name":"cbo","type":"ComboBox","top":80,"columnCount":-1,"columnWidths":"0 pt;72 pt","boundColumn":0}]}"#;
    let parsed = ras_form::parse(text);
    assert!(!parsed.has_errors(), "{:?}", parsed.issues);
    let designer = ras_form::compile(parsed.spec.as_ref().unwrap(), 932).unwrap();
    let snapshot = ras_form::snapshot("F", &designer.streams, &designer.storages, 932).unwrap();
    let (lst, cbo) = (&snapshot.spec.controls[0], &snapshot.spec.controls[1]);
    let values = |c: &ras_form::FormSpecControl, names: &[&str]| -> Vec<Value> {
        names
            .iter()
            .map(|n| c.scalar(n).cloned().unwrap_or(Value::Null))
            .collect()
    };
    let names = [
        "columnCount",
        "columnWidths",
        "boundColumn",
        "textColumn",
        "columnHeads",
    ];
    assert_eq!(
        values(lst, &names),
        [
            Value::from(3),
            Value::from("36 pt;;72 pt"),
            Value::from(2),
            Value::from(3),
            Value::from(true)
        ]
    );
    assert_eq!(
        values(cbo, &names),
        [
            Value::from(-1),
            Value::from("0 pt;72 pt"),
            Value::from(0),
            Value::Null,
            Value::Null
        ]
    );
    for control in [lst, cbo] {
        assert_eq!(control.unsupported, None, "{}", control.name);
    }
    // Neither the items nor the columns change the design of a list without them.
    let bare = ras_form::parse(
        r#"{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{"name":"F"},"controls":[{"id":"l","name":"lst","type":"ListBox","list":["a"]}]}"#,
    );
    let bare = ras_form::compile(bare.spec.as_ref().unwrap(), 932).unwrap();
    let again = ras_form::snapshot("F", &bare.streams, &bare.storages, 932).unwrap();
    assert!(names
        .iter()
        .all(|n| again.spec.controls[0].scalar(n).is_none()));

    for bad in [
        r#""columnWidths":"x""#,
        r#""columnWidths":"-5""#,
        r#""columnCount":-2"#,
        r#""textColumn":-2"#,
    ] {
        let text = format!(
            r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"F"}},"controls":[{{"id":"l","name":"lst","type":"ListBox",{bad}}}]}}"#
        );
        assert!(ras_form::parse(&text).has_errors(), "{bad}");
    }
}

#[test]
fn properties_of_the_form_go_round() {
    let form = r##""startUpPosition":2,"showModal":false,"tag":"メモ \"1\"","helpContextId":5,"rightToLeft":true,
        "whatsThisButton":true,"whatsThisHelp":true,"enabled":false,"borderColor":{"kind":"rgb","value":"#102030"},
        "cycle":2,"mousePointer":6,"specialEffect":1,"zoom":150,"drawBuffer":64000,"scrollBars":3,
        "keepScrollBarsVisible":0,"scrollWidth":400,"scrollHeight":300,"scrollLeft":10,"scrollTop":20"##;
    let text = format!(
        r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"F","caption":"C",{form}}},"controls":[]}}"#
    );
    let parsed = ras_form::parse(&text);
    assert!(!parsed.has_errors(), "{:?}", parsed.issues);
    let spec = parsed.spec.unwrap();
    let designer = ras_form::compile(&spec, 932).unwrap();
    // The designer block in the order of [MS-OVBA], booleans as VBABOOL.
    let block = ras_ovba::decode_text(932, &designer.streams["\u{3}VBFrame"]).unwrap();
    let keys: Vec<&str> = block
        .lines()
        .filter_map(|l| l.split_once(" = ").map(|(k, _)| k.trim()))
        .collect();
    assert_eq!(
        keys,
        [
            "Caption",
            "ClientHeight",
            "ClientWidth",
            "Enabled",
            "HelpContextID",
            "RightToLeft",
            "ShowModal",
            "StartUpPosition",
            "Tag",
            "WhatsThisButton",
            "WhatsThisHelp"
        ]
    );
    assert!(
        block.contains(
            "   ShowModal = 0\r\n   StartUpPosition = 2\r\n   Tag = \"メモ \"\"1\"\"\"\r\n"
        ),
        "{block}"
    );
    let snapshot = ras_form::snapshot("F", &designer.streams, &designer.storages, 932).unwrap();
    let back = &snapshot.spec.form.scalars;
    assert_eq!(
        back.keys().collect::<Vec<_>>(),
        spec.form.scalars.keys().collect::<Vec<_>>()
    );
    for (name, value) in &spec.form.scalars {
        match (value.as_f64(), back[name].as_f64()) {
            // Lengths are stored in HIMETRIC.
            (Some(a), Some(b)) => assert!((a - b).abs() < 0.02, "{name}: {b}, expected {a}"),
            _ => assert_eq!(&back[name], value, "{name}"),
        }
    }
    assert!(snapshot.warnings.is_empty(), "{:?}", snapshot.warnings);

    // A new form, and one the VBA editor made, have none of them.
    let plain = ras_form::parse(
        r#"{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{"name":"F"},"controls":[]}"#,
    );
    let plain = ras_form::compile(plain.spec.as_ref().unwrap(), 932).unwrap();
    let again = ras_form::snapshot("F", &plain.streams, &plain.storages, 932).unwrap();
    assert!(
        again.spec.form.scalars.is_empty(),
        "{:?}",
        again.spec.form.scalars
    );
    let (streams, storages) = triple_state_form();
    let excel = ras_form::snapshot("UserForm1", &streams, &storages, 932).unwrap();
    assert!(
        excel.spec.form.scalars.is_empty(),
        "{:?}",
        excel.spec.form.scalars
    );
    assert!(
        !excel.warnings.iter().any(|w| w.control.is_empty()),
        "{:?}",
        excel.warnings
    );

    for bad in [
        r#""zoom":5"#,
        r#""startUpPosition":4"#,
        r#""tag":"a\nb""#,
        r#""mousePointer":4"#,
        r#""showModal":0"#,
    ] {
        let text = format!(
            r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"F",{bad}}},"controls":[]}}"#
        );
        assert!(ras_form::parse(&text).has_errors(), "{bad}");
    }
}

/// A cursor file (CUR) of one 2×2 image.
fn cursor_file() -> Vec<u8> {
    let image: Vec<u8> = {
        // BITMAPINFOHEADER (height doubled for the mask), 2×2 at 32 bits, the AND mask.
        let mut b = Vec::new();
        for v in [40u32, 2, 4] {
            b.extend_from_slice(&v.to_le_bytes());
        }
        b.extend_from_slice(&1u16.to_le_bytes());
        b.extend_from_slice(&32u16.to_le_bytes());
        b.extend_from_slice(&[0; 24]);
        b.extend_from_slice(&[0xff; 16]);
        b.extend_from_slice(&[0; 8]);
        b
    };
    let mut file = vec![0, 0, 2, 0, 1, 0];
    file.extend_from_slice(&[2, 2, 0, 0, 0, 0, 0, 0]);
    file.extend_from_slice(&(image.len() as u32).to_le_bytes());
    file.extend_from_slice(&22u32.to_le_bytes());
    file.extend_from_slice(&image);
    file
}

#[test]
fn pictures_and_mouse_icons_of_controls_and_the_form_go_round() {
    let root = temp_root("pictures-all");
    let assets = root.join("forms/assets");
    std::fs::create_dir_all(&assets).unwrap();
    let logo =
        std::fs::read(Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/assets/logo.bmp"))
            .unwrap();
    std::fs::write(assets.join("logo.bmp"), &logo).unwrap();
    std::fs::write(assets.join("hand.cur"), cursor_file()).unwrap();
    let text = r#"{"schemaVersion":1,"kind":"ras.userform","basis":"designer",
        "form":{"name":"Pics","picture":{"path":"forms/assets/logo.bmp"},"pictureAlignment":0,"pictureTiling":true,
                "mouseIcon":{"path":"forms/assets/hand.cur"},"mousePointer":99},
        "controls":[
          {"id":"b","name":"btn","type":"CommandButton","picture":{"path":"forms/assets/logo.bmp"},"picturePosition":1},
          {"id":"c","name":"chk","type":"CheckBox","top":30,"picture":{"path":"forms/assets/logo.bmp"},"mouseIcon":{"path":"forms/assets/hand.cur"},"mousePointer":99},
          {"id":"l","name":"lbl","type":"Label","top":60,"picture":{"path":"forms/assets/logo.bmp"},"picturePosition":12},
          {"id":"f","name":"fra","type":"Frame","left":100,"picture":{"path":"forms/assets/logo.bmp"},"pictureSizeMode":1,"pictureTiling":true,"mouseIcon":{"path":"forms/assets/hand.cur"}},
          {"id":"m","name":"mp","type":"MultiPage","top":120},
          {"id":"p","parentId":"m","name":"pg","type":"Page","picture":{"path":"forms/assets/logo.bmp"},"pictureAlignment":4}]}"#;
    let parsed = ras_form::parse(text);
    assert!(!parsed.has_errors(), "{:?}", parsed.issues);
    let designer = ras_form::compile_in(parsed.spec.as_ref().unwrap(), 932, &root).unwrap();
    let snapshot = ras_form::snapshot("Pics", &designer.streams, &designer.storages, 932).unwrap();
    assert!(snapshot.warnings.is_empty(), "{:?}", snapshot.warnings);
    let found: Vec<(String, &str, &str)> = snapshot
        .pictures
        .iter()
        .map(|p| {
            let owner = p
                .control_index
                .map_or(String::new(), |i| snapshot.spec.controls[i].name.clone());
            (owner, p.property, p.picture.format.extension())
        })
        .collect();
    let expected = [
        ("", "picture", "bmp"),
        ("", "mouseIcon", "cur"),
        ("btn", "picture", "bmp"),
        ("chk", "picture", "bmp"),
        ("chk", "mouseIcon", "cur"),
        ("lbl", "picture", "bmp"),
        ("fra", "picture", "bmp"),
        ("fra", "mouseIcon", "cur"),
        ("pg", "picture", "bmp"),
    ];
    let mut found_sorted = found.clone();
    found_sorted.sort();
    let mut expected_sorted: Vec<(String, &str, &str)> = expected
        .iter()
        .map(|(o, p, e)| (o.to_string(), *p, *e))
        .collect();
    expected_sorted.sort();
    assert_eq!(found_sorted, expected_sorted);
    let control = |name: &str| {
        snapshot
            .spec
            .controls
            .iter()
            .find(|c| c.name == name)
            .unwrap()
    };
    assert_eq!(
        control("btn").scalar("picturePosition"),
        Some(&Value::from(1))
    );
    assert_eq!(
        control("lbl").scalar("picturePosition"),
        Some(&Value::from(12))
    );
    assert_eq!(
        control("fra").scalar("pictureTiling"),
        Some(&Value::from(true))
    );
    assert_eq!(control("fra").picture_size_mode, Some(1));
    assert_eq!(control("pg").picture_alignment, Some(4));
    assert_eq!(
        control("chk").scalar("mousePointer"),
        Some(&Value::from(99))
    );
    let form = &snapshot.spec.form.scalars;
    assert_eq!(
        (form.get("pictureAlignment"), form.get("pictureTiling")),
        (Some(&Value::from(0)), Some(&Value::from(true)))
    );
    for c in &snapshot.spec.controls {
        assert_eq!(c.unsupported, None, "{}", c.name);
    }

    // Pulled, every picture is a file under forms/assets; written again, the design is the same.
    let pulled_root = temp_root("pictures-all-pull");
    let source = ras_ovba::UserForm {
        name: "Pics".into(),
        code: String::new(),
        streams: designer.streams.clone(),
        storages: designer.storages.clone(),
    };
    let PullOutcome::Written(pulled) = workspace::pull(&pulled_root, &source, 932, false).unwrap()
    else {
        panic!("nothing written");
    };
    let mut files = pulled.pictures.clone();
    files.sort();
    assert_eq!(
        files,
        [
            "forms/assets/Pics.bmp",
            "forms/assets/Pics.btn.bmp",
            "forms/assets/Pics.chk.bmp",
            "forms/assets/Pics.chk.mouseIcon.cur",
            "forms/assets/Pics.fra.bmp",
            "forms/assets/Pics.fra.mouseIcon.cur",
            "forms/assets/Pics.lbl.bmp",
            "forms/assets/Pics.mouseIcon.cur",
            "forms/assets/Pics.pg.bmp",
        ]
    );
    let text = std::fs::read_to_string(&pulled.spec_path).unwrap();
    let respec = ras_form::parse(&text);
    assert!(!respec.has_errors(), "{:?}", respec.issues);
    let again = ras_form::compile_in(respec.spec.as_ref().unwrap(), 932, &pulled_root).unwrap();
    let reread = ras_form::snapshot("Pics", &again.streams, &again.storages, 932).unwrap();
    let pictures = |s: &ras_form::Snapshot| -> Vec<(Option<String>, &'static str, Vec<u8>)> {
        let mut out: Vec<_> = s
            .pictures
            .iter()
            .map(|p| {
                (
                    p.control_index.map(|i| s.spec.controls[i].name.clone()),
                    p.property,
                    p.picture.data.clone(),
                )
            })
            .collect();
        out.sort();
        out
    };
    assert_eq!(pictures(&reread), pictures(&snapshot));
    assert_eq!(reread.spec.form.scalars, snapshot.spec.form.scalars);
    for (a, b) in reread.spec.controls.iter().zip(&snapshot.spec.controls) {
        assert_eq!(a.scalars, b.scalars, "{}", a.name);
        assert_eq!(
            (a.picture_alignment, a.picture_size_mode),
            (b.picture_alignment, b.picture_size_mode),
            "{}",
            a.name
        );
    }

    // A mouse icon has to be an icon or a cursor; a TextBox has no picture.
    for bad in [
        r#""controls":[{"id":"t","name":"txt","type":"TextBox","picture":{"path":"forms/assets/logo.bmp"}}]"#,
        r#""controls":[{"id":"m","name":"mp","type":"MultiPage","mouseIcon":{"path":"forms/assets/hand.cur"}}]"#,
    ] {
        let text = format!(
            r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"F"}},{bad}}}"#
        );
        assert!(ras_form::parse(&text).has_errors(), "{bad}");
    }
    std::fs::write(assets.join("not-an-icon.png"), b"\x89PNG\r\n\x1a\n").unwrap();
    let text = r#"{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{"name":"F","mouseIcon":{"path":"forms/assets/not-an-icon.png"}},"controls":[]}"#;
    let error =
        ras_form::compile_in(ras_form::parse(text).spec.as_ref().unwrap(), 932, &root).unwrap_err();
    assert!(error.reason.contains("MouseIcon"), "{error}");
    let _ = std::fs::remove_dir_all(&root);
    let _ = std::fs::remove_dir_all(&pulled_root);
}
