//! The FormSpecs of `tests/fixtures`: the valid ones survive generation and
//! reading (a form written from one, read back with `workspace::pull` and
//! written again, gives the same designer storage), and the invalid ones
//! (`9x-error-*`) are rejected with the right diagnostic.

use std::collections::HashMap;
use std::path::{Path, PathBuf};

use ras_form::workspace::{self, PullOutcome};

fn fixtures_dir() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures")
}

/// `(file name, text)` of every fixture.
fn fixtures() -> Vec<(String, String)> {
    let mut out: Vec<_> = std::fs::read_dir(fixtures_dir())
        .unwrap()
        .map(|e| e.unwrap().path())
        .filter(|p| p.to_string_lossy().ends_with(".form.json"))
        .map(|p| (p.file_name().unwrap().to_string_lossy().to_string(), std::fs::read_to_string(&p).unwrap()))
        .collect();
    out.sort();
    assert!(out.len() >= 20, "fixtures missing");
    out
}

fn temp_root(tag: &str) -> PathBuf {
    let dir = std::env::temp_dir().join(format!("ras-form-fixtures-{tag}-{}", std::process::id()));
    let _ = std::fs::remove_dir_all(&dir);
    std::fs::create_dir_all(dir.join("forms")).unwrap();
    dir
}

fn copy_dir(from: &Path, to: &Path) {
    std::fs::create_dir_all(to).unwrap();
    for entry in std::fs::read_dir(from).unwrap() {
        let path = entry.unwrap().path();
        std::fs::copy(&path, to.join(path.file_name().unwrap())).unwrap();
    }
}

#[test]
fn valid_fixtures_survive_generation_and_reading() {
    let mut failures = Vec::new();
    for (file, text) in fixtures().into_iter().filter(|(f, _)| !f.contains("-error-")) {
        let parsed = ras_form::parse(&text);
        let Some(spec) = parsed.spec.as_ref().filter(|_| !parsed.has_errors()) else {
            failures.push(format!("{file}: rejected: {:?}", parsed.issues));
            continue;
        };
        let name = spec.form.name.clone();
        // A project with the fixture as its form; picture paths are relative to the fixture directory.
        let first = temp_root(&format!("{name}-a"));
        copy_dir(&fixtures_dir().join("assets"), &first.join("assets"));
        std::fs::write(first.join(format!("forms/{name}.form.json")), &text).unwrap();
        let built = workspace::build(&workspace::load(&first).unwrap(), &first, 932, &HashMap::new());
        if built.errors() > 0 {
            failures.push(format!("{file}: {:?}", built.problems));
            continue;
        }
        let generated = built.forms[0].clone();

        // Read back into another project, and written again.
        let second = temp_root(&format!("{name}-b"));
        let PullOutcome::Written(pulled) = workspace::pull(&second, &generated, 932, true).unwrap() else {
            panic!("{file}: nothing written");
        };
        if !pulled.warnings.is_empty() {
            failures.push(format!("{file}: lost in reading: {:?}", pulled.warnings));
        }
        let again = workspace::build(&workspace::load(&second).unwrap(), &second, 932, &HashMap::new());
        if again.errors() > 0 {
            failures.push(format!("{file}: the pulled form is not valid: {:?}", again.problems));
        } else if again.forms[0].streams != generated.streams {
            let differ: Vec<_> = generated
                .streams
                .iter()
                .filter(|(k, v)| again.forms[0].streams.get(*k) != Some(v))
                .map(|(k, _)| k.clone())
                .collect();
            failures.push(format!("{file}: written again, the streams {differ:?} differ"));
        } else if again.forms[0].storages != generated.storages {
            failures.push(format!("{file}: written again, the storages differ"));
        }
        let _ = std::fs::remove_dir_all(&first);
        let _ = std::fs::remove_dir_all(&second);
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
}

/// The diagnostic each `9x-error-*` fixture is rejected with: code and part of the message.
const ERRORS: &[(&str, &str, &str)] = &[
    ("90-error-picture-format.form.json", "userform_generation_invalid", "画像の形式を判別できません"),
    ("91-error-parent.form.json", "UFV011", "コンテナーではないコントロール \"txtA\""),
    ("92-error-duplicate-name.form.json", "userform_generation_invalid", "\"TXTA\""),
    ("93-error-page-geometry.form.json", "UFV005", "controls[1].left"),
    ("94-error-custom.form.json", "userform_generation_unsupported", "生成できません"),
    ("95-error-tab-index.form.json", "UFV019", "タブ名 \"a\" と重複"),
    ("96-error-unknown-field.form.json", "UFV001", "form.build は廃止"),
    ("97-error-picture-missing.form.json", "userform_generation_invalid", "を開けません"),
    ("98-error-picture-outside.form.json", "userform_generation_invalid", "相対パスにしてください"),
];

#[test]
fn invalid_fixtures_are_rejected() {
    let mut failures = Vec::new();
    let mut seen = 0;
    for (file, text) in fixtures().into_iter().filter(|(f, _)| f.contains("-error-")) {
        let parsed = ras_form::parse(&text);
        let (code, message) = match &parsed.spec {
            Some(spec) if !parsed.has_errors() => match ras_form::compile_in(spec, 932, &fixtures_dir()) {
                Ok(_) => {
                    failures.push(format!("{file}: generated"));
                    continue;
                }
                Err(e) => (e.code.to_string(), e.to_string()),
            },
            _ => parsed
                .issues
                .iter()
                .find(|i| i.severity == ras_form::Severity::Error)
                .map(|i| (i.code.clone(), i.message.clone()))
                .unwrap_or_default(),
        };
        match ERRORS.iter().find(|(f, ..)| *f == file) {
            Some((_, c, m)) if *c == code && message.contains(m) => seen += 1,
            Some((_, c, m)) => failures.push(format!("{file}: {code} {message:?}, expected {c} with {m:?}")),
            None => failures.push(format!("{file}: {code} {message:?}, not listed in ERRORS")),
        }
    }
    assert!(failures.is_empty(), "{}", failures.join("\n"));
    assert_eq!(seen, ERRORS.len());
}
