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

impl Diagnostic {
    pub fn is_error(&self) -> bool {
        self.severity == Severity::Error
    }
}
