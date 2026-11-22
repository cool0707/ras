//! Diagnostics of `ras_form::parse`.

fn codes(text: &str) -> Vec<String> {
    ras_form::parse(text)
        .issues
        .into_iter()
        .filter(|i| i.severity == ras_form::Severity::Error)
        .map(|i| i.code)
        .collect()
}

fn doc(form: &str, controls: &str) -> String {
    format!(
        r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{form},"controls":{controls}}}"#
    )
}

#[test]
fn minimal_and_new_form_are_valid() {
    let parsed = ras_form::parse(&ras_form::FormSpec::new_form("UserForm1").to_json());
    assert!(parsed.issues.is_empty(), "{:?}", parsed.issues);
    let spec = parsed.spec.unwrap();
    assert_eq!(spec.form.width, Some(240.0));
}

#[test]
fn error_codes() {
    assert_eq!(codes("{"), ["UFY001"]);
    assert_eq!(codes("[]"), ["UFV002"]);
    assert_eq!(codes(&doc(r#"{"name":"F","extra":1}"#, "[]")), ["UFV001"]);
    assert_eq!(codes(&doc(r#"{"name":5}"#, "[]")), ["UFV002"]);
    assert_eq!(
        codes(&doc(r#"{"name":"F"}"#, "[]").replace("\"designer\"", "\"runtime\"")),
        ["UFV003"]
    );
    assert_eq!(codes(&doc(r#"{"caption":"x"}"#, "[]")), ["UFV004"]);
    assert_eq!(
        codes(&doc(
            r#"{"name":"F"}"#,
            r#"[{"id":"a","name":"a","type":"TextBox","caption":"x"}]"#
        )),
        ["UFV005"]
    );
    assert_eq!(
        codes(&doc(
            r#"{"name":"F"}"#,
            r#"[{"id":"a","name":"a","type":"Gizmo"}]"#
        )),
        ["UFV006"]
    );
    assert_eq!(
        codes(&doc(
            r#"{"name":"F"}"#,
            r#"[{"id":"a","name":"a","type":"Label"},{"id":"a","name":"b","type":"Label"}]"#
        )),
        ["UFV007"]
    );
    assert_eq!(
        codes(&doc(
            r#"{"name":"F"}"#,
            r#"[{"id":"a","parentId":"z","name":"a","type":"Label"}]"#
        )),
        ["UFV008"]
    );
    assert_eq!(
        codes(&doc(
            r#"{"name":"F"}"#,
            r#"[{"id":"a","parentId":"a","name":"a","type":"Frame"}]"#
        )),
        ["UFV009"]
    );
    assert_eq!(
        codes(&doc(
            r#"{"name":"F"}"#,
            r#"[{"id":"a","parentId":"b","name":"a","type":"Frame"},{"id":"b","parentId":"a","name":"b","type":"Frame"}]"#
        )),
        ["UFV010"]
    );
    assert_eq!(
        codes(&doc(
            r#"{"name":"F"}"#,
            r#"[{"id":"p","name":"p","type":"Page"}]"#
        )),
        ["UFV011"]
    );
    assert_eq!(
        codes(&doc(
            r#"{"name":"F"}"#,
            r#"[{"id":"a","name":"a","type":"Label","progId":"Forms.TextBox.1"}]"#
        )),
        ["UFV012"]
    );
    assert_eq!(codes(&doc(r#"{"name":"F","width":-1}"#, "[]")), ["UFV016"]);
    assert_eq!(
        codes(&doc(
            r#"{"name":"F"}"#,
            r#"[{"id":"t","name":"t","type":"TabStrip","tabs":[{"name":"x"},{"name":"X"}]}]"#
        )),
        ["UFV019"]
    );
    assert_eq!(
        codes(&doc(
            r#"{"name":"F"}"#,
            r#"[{"id":"t","name":"t","type":"TabStrip","selectedIndex":3,"tabs":[{"name":"x"}]}]"#
        )),
        ["UFV020"]
    );
    assert_eq!(
        codes(&doc(
            r#"{"name":"F","foreColor":{"kind":"rgb","value":"red"}}"#,
            "[]"
        )),
        ["UFV021"]
    );
}

#[test]
fn warnings_do_not_block_generation() {
    let text = doc(
        r#"{"name":"F"}"#,
        r#"[{"id":"a","name":"a","type":"Label","properties":{"tag":"x"}},{"id":"c","name":"c","type":"Calendar","progId":"MSCAL.Calendar.7"}]"#,
    );
    let parsed = ras_form::parse(&text);
    let warnings: Vec<_> = parsed
        .issues
        .iter()
        .map(|i| (i.code.as_str(), i.severity))
        .collect();
    assert_eq!(
        warnings,
        [
            ("UFV013", ras_form::Severity::Warning),
            ("UFV014", ras_form::Severity::Warning)
        ]
    );
    // ...but a custom control cannot be generated.
    let error = ras_form::compile(parsed.spec.as_ref().unwrap(), 932).unwrap_err();
    assert_eq!(error.code, "userform_generation_unsupported");
    assert_eq!(error.property, "controls[1].type");
}

#[test]
fn locate_points_at_the_field() {
    let text = "{\n  \"form\": { \"name\": \"F\" },\n  \"controls\": [\n    { \"name\": \"a\" },\n    { \"name\": \"b\", \"caption\": 1 }\n  ]\n}\n";
    assert_eq!(ras_form::locate(text, "controls[1].caption"), Some((5, 20)));
    assert_eq!(ras_form::locate(text, "controls[1].name"), Some((5, 7)));
    assert_eq!(ras_form::locate(text, "form.name"), Some((2, 13)));
    assert_eq!(ras_form::locate(text, ""), None);
}

#[test]
fn round_trip_keeps_unknown_free_json() {
    let parsed = ras_form::parse(
        &std::fs::read_to_string(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/07-appearance.form.json"
        ))
        .unwrap(),
    );
    let spec = parsed.spec.unwrap();
    let again = ras_form::parse(&spec.to_json()).spec.unwrap();
    assert_eq!(
        ras_form::compile(&spec, 932).unwrap(),
        ras_form::compile(&again, 932).unwrap()
    );
}

#[test]
fn an_edit_field_lower_than_its_font_is_warned_about() {
    let warnings = |control: &str| -> Vec<String> {
        ras_form::parse(&doc(r#"{"name":"F"}"#, &format!("[{control}]")))
            .issues
            .into_iter()
            .filter(|i| i.severity == ras_form::Severity::Warning && i.code == "UFV022")
            .map(|i| i.field)
            .collect()
    };
    // 9pt MS UI Gothic (no font written) needs 13.5pt.
    assert!(warnings(r#"{"id":"t","name":"t","type":"TextBox","height":18}"#).is_empty());
    assert_eq!(
        warnings(r#"{"id":"t","name":"t","type":"TextBox","height":12}"#),
        ["controls[0].height"]
    );
    // 16pt Meiryo UI needs 26.25pt, as the Designer computes.
    let meiryo = r#""font":{"name":"Meiryo UI","size":16}"#;
    assert_eq!(
        warnings(&format!(
            r#"{{"id":"c","name":"c","type":"ComboBox","height":26,{meiryo}}}"#
        )),
        ["controls[0].height"]
    );
    assert!(warnings(&format!(
        r#"{{"id":"c","name":"c","type":"ComboBox","height":26.25,{meiryo}}}"#
    ))
    .is_empty());
    // Not an edit field, or several lines: no warning.
    assert!(warnings(&format!(
        r#"{{"id":"l","name":"l","type":"Label","height":5,{meiryo}}}"#
    ))
    .is_empty());
    assert!(warnings(
        r#"{"id":"t","name":"t","type":"TextBox","height":5,"properties":{"multiLine":true}}"#
    )
    .is_empty());
}

#[test]
fn a_single_border_with_an_effect_is_warned_about_and_written_flat() {
    let parsed = |form: &str, control: &str| ras_form::parse(&doc(form, &format!("[{control}]")));
    let warnings = |form: &str, control: &str| -> Vec<String> {
        parsed(form, control)
            .issues
            .into_iter()
            .filter(|i| i.severity == ras_form::Severity::Warning && i.code == "UFV023")
            .map(|i| i.field)
            .collect()
    };
    let both = r#"{"id":"t","name":"t","type":"TextBox","borderStyle":1,"specialEffect":3}"#;
    assert_eq!(
        warnings(r#"{"name":"F"}"#, both),
        ["controls[0].specialEffect"]
    );
    assert_eq!(
        warnings(
            r#"{"name":"F","borderStyle":1,"specialEffect":2}"#,
            r#"{"id":"l","name":"l","type":"Label"}"#
        ),
        ["form.specialEffect"]
    );
    // One of them, or a flat effect: no warning.
    assert!(warnings(
        r#"{"name":"F"}"#,
        r#"{"id":"t","name":"t","type":"TextBox","borderStyle":1}"#
    )
    .is_empty());
    assert!(warnings(
        r#"{"name":"F"}"#,
        r#"{"id":"t","name":"t","type":"TextBox","borderStyle":1,"specialEffect":0}"#
    )
    .is_empty());

    // Written with one edge, as the VBA editor leaves it: the border wins.
    let read_back = |control: &str| -> ras_form::FormSpecControl {
        let spec = parsed(r#"{"name":"F"}"#, control).spec.unwrap();
        let designer = ras_form::compile(&spec, 932).unwrap();
        let back = ras_form::snapshot("F", &designer.streams, &designer.storages, 932)
            .unwrap()
            .spec;
        back.controls.into_iter().next().unwrap()
    };
    let effect = |c: &ras_form::FormSpecControl| c.scalar("specialEffect").and_then(|v| v.as_i64());
    for control in [
        both,
        r#"{"id":"t","name":"t","type":"TextBox","borderStyle":1}"#,
        r#"{"id":"f","name":"f","type":"Frame","borderStyle":1}"#,
    ] {
        let back = read_back(control);
        assert_eq!(
            (back.border_style, effect(&back)),
            (Some(1), Some(0)),
            "{control}"
        );
    }
    // An Image is made with a Single border, which an effect takes away.
    let back = read_back(r#"{"id":"i","name":"i","type":"Image","specialEffect":2}"#);
    assert_eq!((back.border_style, effect(&back)), (Some(0), Some(2)));
}

#[test]
fn a_form_size_in_whole_points_is_stored_in_whole_pixels() {
    let read_back = |form: &str| -> (f64, f64) {
        let spec = ras_form::parse(&doc(form, "[]")).spec.unwrap();
        let designer = ras_form::compile(&spec, 932).unwrap();
        let back = ras_form::snapshot("F", &designer.streams, &designer.storages, 932)
            .unwrap()
            .spec
            .form;
        (back.width.unwrap(), back.height.unwrap())
    };
    let pixels = |points: f64| points / 0.75;
    // 340 x 250 pt is 453⅓ x 333⅓ px at 96 DPI: the VBA editor keeps 453 x 333 px.
    let (width, height) = read_back(r#"{"name":"F","width":340,"height":250}"#);
    assert!(
        (pixels(width) - 453.0).abs() < 0.02 && (pixels(height) - 333.0).abs() < 0.02,
        "{width} x {height}"
    );
    // A multiple of 3 pt is whole pixels already; a size read from a workbook saved at
    // 120 DPI (349.2 pt = 582 px) is kept.
    let (width, height) = read_back(r#"{"name":"F","width":330,"height":349.2}"#);
    assert!(
        (width - 330.0).abs() < 0.02 && (height - 349.2).abs() < 0.02,
        "{width} x {height}"
    );
}
