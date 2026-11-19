//! Diagnostics.

use ras_syntax::ast::Span;

use crate::model::ModuleId;

#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum Severity {
    Error,
    Warning,
    Info,
    Hint,
}

#[derive(Clone, Debug)]
pub struct Diagnostic {
    pub severity: Severity,
    /// Stable identifier such as `E0102` / `W0203`.
    pub code: &'static str,
    pub message: String,
    pub module: Option<ModuleId>,
    pub span: Span,
}

/// A mechanical correction of the source (`ras check --fix`, editor quick fix).
#[derive(Clone, Debug)]
pub struct Fix {
    pub module: ModuleId,
    pub span: Span,
    pub replacement: String,
    /// Code of the diagnostic this corrects.
    pub code: &'static str,
}

/// `text` with the replacements applied. Overlapping replacements are skipped.
pub fn apply_fixes(text: &str, fixes: &[&Fix]) -> String {
    let mut sorted: Vec<&Fix> = fixes.to_vec();
    sorted.sort_by_key(|f| u32::from(f.span.start()));
    let mut out = String::with_capacity(text.len());
    let mut pos = 0usize;
    for f in sorted {
        let (start, end) = (u32::from(f.span.start()) as usize, u32::from(f.span.end()) as usize);
        if start < pos || end > text.len() {
            continue;
        }
        out.push_str(&text[pos..start]);
        out.push_str(&f.replacement);
        pos = end;
    }
    out.push_str(&text[pos..]);
    out
}

impl Diagnostic {
    pub fn is_error(&self) -> bool {
        self.severity == Severity::Error
    }
}
