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
    // xlflow's kind is accepted as well.
    let xlflow = doc(r#"{"name":"F"}"#, "[]").replace("ras.userform", "xlflow.userform");
    assert!(codes(&xlflow).is_empty());
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
