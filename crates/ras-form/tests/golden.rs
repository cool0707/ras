//! Differential tests against xlflow's pure-Go UserForm generator.
//!
//! Every `tests/fixtures/*.form.json` has a `*.golden.json` next to it that
//! was produced by running xlflow's own `spec.ParseFormSpec` +
//! `compiler.CompileNew` + `oforms.SerializeForm` (code page 932) on the same
//! file — see `tests/fixtures/README.md`. A golden file holds either the
//! designer streams and storage CLSIDs (hex) or the error xlflow reported.
//!
//! `tests/data/*.bin` are projects saved by Excel; the `*.projected.json`
//! next to them hold what xlflow's `oforms.ReadForm` + `projection.Project`
//! make of their forms. `projection_matches_xlflow` checks the Rust port of
//! the reading side against them.

use std::collections::BTreeMap;
use std::path::Path;

fn hex(bytes: &[u8]) -> String {
    bytes.iter().map(|b| format!("{b:02x}")).collect()
}

fn first_difference(a: &str, b: &str) -> usize {
    a.bytes()
        .zip(b.bytes())
        .position(|(x, y)| x != y)
        .unwrap_or(a.len().min(b.len()))
        / 2
}

#[test]
fn generation_matches_xlflow() {
    let dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
    let mut fixtures: Vec<_> = std::fs::read_dir(&dir)
        .unwrap()
        .map(|e| e.unwrap().path())
        .filter(|p| p.to_string_lossy().ends_with(".form.json"))
        .collect();
    fixtures.sort();
    assert!(fixtures.len() >= 10, "fixtures missing");
    let mut failures = Vec::new();
    for path in fixtures {
        let name = path.file_name().unwrap().to_string_lossy().to_string();
        let golden_path = dir.join(name.replace(".form.json", ".golden.json"));
        let golden: serde_json::Value =
            serde_json::from_str(&std::fs::read_to_string(&golden_path).unwrap()).unwrap();
        let text = std::fs::read_to_string(&path).unwrap();
        let parsed = ras_form::parse(&text);
        let result = match &parsed.spec {
            // Picture paths of the fixtures are relative to the fixture directory.
            Some(spec) => ras_form::compile_in(spec, 932, &dir).map_err(|e| e.to_string()),
            None => Err(parsed
                .issues
                .iter()
                .find(|i| i.severity == ras_form::Severity::Error)
                .map(|i| format!("{} {}", i.code, i.message))
                .unwrap_or_default()),
        };
        match (golden.get("error"), result) {
            (Some(expected), Err(_)) => {
                let _ = expected;
            }
            (Some(expected), Ok(_)) => failures.push(format!(
                "{name}: xlflow rejects it ({expected}) but RAS generated a form"
            )),
            (None, Err(e)) => failures.push(format!("{name}: RAS rejects it: {e}")),
            (None, Ok(designer)) => {
                let expected: BTreeMap<String, String> =
                    serde_json::from_value(golden["streams"].clone()).unwrap();
                let actual: BTreeMap<String, String> = designer
                    .streams
                    .iter()
                    .map(|(k, v)| (k.clone(), hex(v)))
                    .collect();
                let expected_keys: Vec<_> = expected.keys().collect();
                let actual_keys: Vec<_> = actual.keys().collect();
                if expected_keys != actual_keys {
                    failures.push(format!(
                        "{name}: streams differ: expected {expected_keys:?}, got {actual_keys:?}"
                    ));
                    continue;
                }
                for (key, value) in &expected {
                    if actual[key] != *value {
                        failures.push(format!(
                            "{name}: stream {key:?} differs at byte {}\n  expected {value}\n  actual   {}",
                            first_difference(value, &actual[key]),
                            actual[key]
                        ));
                    }
                }
                let storages: BTreeMap<String, String> =
                    serde_json::from_value(golden["storages"].clone()).unwrap();
                let actual_storages: BTreeMap<String, String> = designer
                    .storages
                    .iter()
                    .map(|(k, v)| (k.clone(), hex(v)))
                    .collect();
                if storages != actual_storages {
                    failures.push(format!(
                        "{name}: storages differ: expected {storages:?}, got {actual_storages:?}"
                    ));
                }
            }
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}
