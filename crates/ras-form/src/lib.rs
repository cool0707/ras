//! User forms described as data (RAS language specification §14.11).
//!
//! A user form lives next to `ras.toml` as `forms/<Name>.form.json` — a
//! FormSpec version 1 document (`kind` `ras.userform`; the format of
//! [xlflow](https://github.com/harumiWeb/xlflow)'s FormSpec, extended) — with
//! its code-behind in `forms/<Name>.ras` (RAS, compiled by `ras-compiler`).
//! This crate
//!
//! * parses and validates a FormSpec ([`parse`]; diagnostics `UFY001`,
//!   `UFV001`–`UFV027`),
//! * generates the MS-OFORMS designer storage of the form from scratch
//!   ([`compile`], [`compile_in`]), so that `ras pack` can store the form in
//!   a workbook without Excel,
//! * reads the designer storage of an existing form back into a FormSpec
//!   ([`snapshot`]), and
//! * handles the files of a project ([`workspace`]) and the code module of a
//!   generated form ([`codebehind`]), including the code that fills
//!   ComboBox / ListBox lists, and
//! * takes over the hand-written VBA of a workbook when it is synced
//!   ([`import`]: forms, standard modules and the code of sheets).
//!
//! The specification model, validation, generation, reading and projection
//! are a port of xlflow's `internal/vba/userforms` packages (Copyright (c)
//! 2026 harumiWeb, MIT License; see `NOTICE.md`). Editing an existing
//! binary form in place is not ported: a form is always generated whole.

#![forbid(unsafe_code)]
#![warn(missing_docs)]

pub mod anchor;
pub mod codebehind;
mod compile;
pub mod contract;
pub mod import;
mod oforms;
pub mod picker;
pub mod picture;
pub mod progress;
mod reader;
mod scalars;
pub mod sizing;
mod snapshot;
pub mod spec;
mod validate;
pub mod vbdiff;
pub mod workspace;

pub use compile::CompileError;
pub use oforms::Designer;
pub use reader::ReadError;
pub use snapshot::{
    snapshot, Snapshot, SnapshotError, SnapshotPicture, SnapshotWarning, PLACEHOLDER_TYPE,
};
pub use spec::{
    FormSpec, FormSpecColor, FormSpecControl, FormSpecFont, FormSpecForm, FormSpecListItem,
    FormSpecTab,
};
pub use validate::{Issue, Severity};

/// Result of [`parse`].
#[derive(Clone, Debug)]
pub struct Parsed {
    /// The normalised spec (legacy nested controls flattened, ids and
    /// `zIndex` filled in); `None` when the text is not a usable FormSpec.
    pub spec: Option<FormSpec>,
    /// Errors and warnings, in document order of discovery.
    pub issues: Vec<Issue>,
}

impl Parsed {
    /// Whether any issue is an error.
    pub fn has_errors(&self) -> bool {
        self.issues.iter().any(|i| i.severity == Severity::Error)
    }
}

/// Line and column (1-based) of a JSON path in `text`, best effort: the
/// position of the last property name of the path. Used to point
/// diagnostics into the source.
pub fn locate(text: &str, field: &str) -> Option<(usize, usize)> {
    let key = field.rsplit('.').next()?.split('[').next()?;
    if key.is_empty() {
        return None;
    }
    let needle = format!("\"{key}\"");
    // Walk the path's property names in order so that `controls[3].name`
    // finds the fourth control's `name`, not the form's.
    let mut from = 0usize;
    for segment in field.split('.') {
        let name = segment.split('[').next().unwrap_or("");
        if name.is_empty() {
            continue;
        }
        let quoted = format!("\"{name}\"");
        let Some(at) = text[from..].find(&quoted) else {
            break;
        };
        from += at;
        if let Some(index) = segment
            .split('[')
            .nth(1)
            .and_then(|s| s.trim_end_matches(']').parse::<usize>().ok())
        {
            // Skip to the index-th object of the array that follows.
            let mut depth = 0i32;
            let mut seen = 0usize;
            let mut in_string = false;
            let mut escaped = false;
            for (offset, ch) in text[from..].char_indices() {
                if in_string {
                    match ch {
                        _ if escaped => escaped = false,
                        '\\' => escaped = true,
                        '"' => in_string = false,
                        _ => {}
                    }
                    continue;
                }
                match ch {
                    '"' => in_string = true,
                    '[' | '{' => {
                        depth += 1;
                        if ch == '{' && depth == 2 {
                            if seen == index {
                                from += offset;
                                break;
                            }
                            seen += 1;
                        }
                    }
                    ']' | '}' => {
                        depth -= 1;
                        if depth <= 0 {
                            break;
                        }
                    }
                    _ => {}
                }
            }
        }
    }
    let at = text[from..]
        .find(&needle)
        .map(|i| i + from)
        .or_else(|| text.find(&needle))?;
    let line = text[..at].matches('\n').count() + 1;
    let column = at - text[..at].rfind('\n').map_or(0, |i| i + 1) + 1;
    Some((line, column))
}

/// Parses and validates the text of a `.form.json` file.
pub fn parse(text: &str) -> Parsed {
    let value: serde_json::Value = match serde_json::from_str(text) {
        Ok(value) => value,
        Err(e) => {
            return Parsed {
                spec: None,
                issues: vec![Issue {
                    code: "UFY001".into(),
                    severity: Severity::Error,
                    message: format!(
                        "JSON として読めません（{} 行 {} 列）: {e}",
                        e.line(),
                        e.column()
                    ),
                    field: String::new(),
                }],
            }
        }
    };
    let mut issues = validate::validate_raw(&value);
    if issues.iter().any(|i| i.severity == Severity::Error) {
        return Parsed { spec: None, issues };
    }
    let mut spec: FormSpec = match serde_json::from_value(value) {
        Ok(spec) => spec,
        Err(e) => {
            issues.push(Issue {
                code: "UFV002".into(),
                severity: Severity::Error,
                message: format!("FormSpec として読めません: {e}"),
                field: String::new(),
            });
            return Parsed { spec: None, issues };
        }
    };
    spec::normalize(&mut spec);
    let strict = validate::validate_strict(&spec);
    for issue in strict {
        if !issues.contains(&issue) {
            issues.push(issue);
        }
    }
    let ok = !issues.iter().any(|i| i.severity == Severity::Error);
    Parsed {
        spec: ok.then_some(spec),
        issues,
    }
}

/// Generates the designer storage of a form. `spec` should come from
/// [`parse`] (it is normalised and validated again here); `codepage` is the
/// project code page (`target.codepage`, normally 932), used for the
/// `VBFrame` stream.
///
/// A form with an `Image` picture (`picture.path`) needs the project root:
/// use [`compile_in`].
pub fn compile(spec: &FormSpec, codepage: u16) -> Result<Designer, CompileError> {
    let mut spec = spec.clone();
    spec::normalize(&mut spec);
    compile::compile_new(&spec, codepage, None)
}

/// Like [`compile`], reading the pictures of `Image` controls from the
/// files their `picture.path` names relative to `root`, the project root.
/// A picture MS-Forms cannot store as it is (PNG, GIF, ...) is converted to
/// a 24-bit BMP (see [`picture::prepare`]).
pub fn compile_in(
    spec: &FormSpec,
    codepage: u16,
    root: &std::path::Path,
) -> Result<Designer, CompileError> {
    let mut spec = spec.clone();
    spec::normalize(&mut spec);
    compile::compile_new(&spec, codepage, Some(root))
}
