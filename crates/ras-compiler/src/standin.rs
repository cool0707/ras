//! What a program uses that the stand-in for Excel does not have (`ras check --target stand-in`).
//!
//! Without Excel (`ras run`, `ras test`, the debug run on macOS and Linux) the program works on
//! the stand-in of the debug run time (spec §14.10.8). A member it does not implement raises
//! error 438 there, an operation that needs Excel itself 445, a COM object it cannot make 429 —
//! found only when that line runs. These warnings tell before.
//!
//! What the stand-in has is `stdlib/standin.json`, listed from its implementation by
//! `tools/standin/members.js`.

use std::collections::{HashMap, HashSet};
use std::sync::OnceLock;

use ras_syntax::ast::{Expr, ExprKind, Span};

use crate::astutil::{visit_module, Visit};
use crate::diag::{Diagnostic, Severity};
use crate::model::*;

const MEMBERS_JSON: &str = include_str!("../stdlib/standin.json");

struct Members {
    members: HashSet<String>,
    excel: HashSet<String>,
}

/// The objects of the stand-in by the name of their Excel class, names in lower case.
fn table() -> &'static HashMap<String, Members> {
    static TABLE: OnceLock<HashMap<String, Members>> = OnceLock::new();
    TABLE.get_or_init(|| {
        let v: serde_json::Value = serde_json::from_str(MEMBERS_JSON).unwrap_or_default();
        let set = |x: &serde_json::Value| x.as_array().map(|a| a.iter().filter_map(|s| s.as_str().map(str::to_lowercase)).collect()).unwrap_or_default();
        v.as_object().map(|o| o.iter().map(|(k, m)| (k.clone(), Members { members: set(&m["members"]), excel: set(&m["excel"]) })).collect()).unwrap_or_default()
    })
}

/// The object of the stand-in for a host class (`Excel.Worksheets` are its `Sheets`).
fn stand_in_class(def: &ClassDef) -> Option<&'static str> {
    if def.kind != ClassKind::Host || def.namespace.as_deref() != Some("Excel") {
        return None;
    }
    Some(match def.name.as_str() {
        "Application" => "Application",
        "Workbooks" => "Workbooks",
        "Workbook" => "Workbook",
        "Sheets" | "Worksheets" => "Sheets",
        "Worksheet" => "Worksheet",
        "Range" => "Range",
        "Names" => "Names",
        "Name" => "Name",
        "Sort" => "Sort",
        _ => return None,
    })
}

/// The ProgIDs the run time makes without COM.
const CREATABLE: &[&str] = &["scripting.filesystemobject", "scripting.dictionary", "vbscript.regexp", "excel.application"];

struct Walker<'p> {
    prog: &'p Program,
    module: ModuleId,
    /// The functions of `Excel.WorksheetFunction`: `Application` has them too.
    functions: &'p HashSet<String>,
    seen: HashSet<(u32, &'static str)>,
    out: Vec<Diagnostic>,
}

impl Walker<'_> {
    fn warn(&mut self, span: Span, code: &'static str, message: String) {
        if self.seen.insert((u32::from(span.start()), code)) {
            self.out.push(Diagnostic { severity: Severity::Warning, code, message, module: Some(self.module), span });
        }
    }

    /// A member `host` (the type library's spelling) of the stand-in's object `class`, written `shown`.
    fn member(&mut self, span: Span, class: &str, host: &str, shown: &str) {
        let Some(m) = table().get(class) else { return };
        let lower = host.to_lowercase();
        if class == "Application" && lower == "run" {
            return self.warn(span, "W0511", format!("`{shown}`（手書き VBA の実行）は Excel が必要です。ブックの代役では実行時エラー 445 になります（テストでは `mockRun`、`ras run` では `--replace-run` で置き換えられます）"));
        }
        if m.excel.contains(&lower) {
            return self.warn(span, "W0511", format!("`{shown}` は Excel が必要です。ブックの代役では実行時エラー 445 になります"));
        }
        if !m.members.contains(&lower) && !(class == "Application" && self.functions.contains(&lower)) {
            let what = if class == "VBA" { "デバッグ実行のランタイムにありません（実行時エラー 445 になります）".to_string() } else { format!("ブックの代役の `{class}` にありません（実行時エラー 438 になります）") };
            self.warn(span, "W0510", format!("`{shown}` は{what}"));
        }
    }

    fn host_fn(&mut self, span: Span, f: FnId) {
        let prog = self.prog;
        let d = &prog.fns[f];
        match d.kind {
            FnKind::Host { class: Some(c) } => {
                if let Some(class) = stand_in_class(&prog.classes[c]) {
                    self.member(span, class, d.vba_name(), &d.name);
                }
            }
            FnKind::Host { class: None } => match prog.modules[d.module].key.as_str() {
                "excel.d" => self.member(span, "Application", d.vba_name(), &d.name),
                "vba.d" => self.member(span, "VBA", d.vba_name(), &d.name),
                // Declared in the project's own type definitions: a procedure of the workbook.
                _ if prog.modules[d.module].kind == ModuleKind::TypeDef && !prog.modules[d.module].rel_path.starts_with("<stdlib>/") => {
                    self.warn(span, "W0511", format!("`{}`（手書き VBA）は Excel が必要です。ブックの代役では実行時エラー 445 になります（テストでは `mockRun`、`ras run` では `--replace-run` で置き換えられます）", d.name));
                }
                _ => {}
            },
            _ => {}
        }
    }
}

impl<'a> Visit<'a> for Walker<'_> {
    fn expr(&mut self, e: &'a Expr) {
        let prog = self.prog;
        let info = &prog.modules[self.module].info;
        match &e.kind {
            ExprKind::Member { name, .. } => match info.res.get(&e.id) {
                Some(Res::Field(c, i)) => {
                    if let Some(class) = stand_in_class(&prog.classes[*c]) {
                        let f = &prog.classes[*c].fields[*i];
                        self.member(name.span, class, f.vba_name(), &f.name);
                    }
                }
                Some(Res::Method(f) | Res::Fn(f)) => self.host_fn(name.span, *f),
                _ => {}
            },
            ExprKind::Name(id) => match info.res.get(&e.id) {
                Some(Res::Fn(f)) => self.host_fn(id.span, *f),
                Some(Res::Global(v)) => {
                    let d = &prog.vars[*v];
                    // (The sheets' code names and the forms are globals too, of the project.)
                    let own = prog.config.sheet_code_names.iter().any(|n| n.eq_ignore_ascii_case(&d.name)) || prog.config.forms.iter().any(|f| f.name.eq_ignore_ascii_case(&d.name));
                    if d.declared && !own && prog.modules[d.module].key == "excel.d" && !matches!(d.vba_name(), "Application" | "ThisWorkbook") {
                        self.member(id.span, "Application", d.vba_name(), &d.name);
                    }
                }
                _ => {}
            },
            ExprKind::Call { callee, args, .. } => {
                // createObject("ADODB.Connection")
                let is_create = matches!(info.calls.get(&e.id).map(|c| &c.target), Some(CallTarget::Host(f)) if prog.fns[*f].vba_name() == "CreateObject" && prog.modules[prog.fns[*f].module].key == "vba.d");
                if let (true, Some(ExprKind::Str(id))) = (is_create, args.first().map(|a| &a.peel().kind)) {
                    if !CREATABLE.contains(&id.trim().to_lowercase().as_str()) {
                        self.warn(callee.span, "W0512", format!("`createObject(\"{id}\")` は COM を使わない実行では作成できません（実行時エラー 429。テストでは `mockCreateObject`、`ras run` では `--replace-object` で置き換えられます）"));
                    }
                }
            }
            ExprKind::New { .. } => {
                if let Type::Class(c, _) = info.ty(e.id) {
                    let d = &prog.classes[c];
                    if d.kind == ClassKind::Host {
                        let id = format!("{}.{}", d.namespace.as_deref().unwrap_or_default(), d.name);
                        if !CREATABLE.contains(&id.to_lowercase().as_str()) && id != "VBA.Collection" {
                            self.warn(e.span, "W0512", format!("`new {id}()` は COM を使わない実行では作成できません（実行時エラー 429。テストでは `mockCreateObject`、`ras run` では `--replace-object` で置き換えられます）"));
                        }
                    }
                }
            }
            _ => {}
        }
    }
}

/// The warnings for the modules of the project (its tests included, when they are loaded).
pub fn check(prog: &Program) -> Vec<Diagnostic> {
    let functions: HashSet<String> = prog
        .namespaces
        .get("Excel")
        .and_then(|ns| match ns.get("WorksheetFunction") {
            Some(Sym::Class(c)) => Some(prog.classes[*c].methods.iter().map(|&f| prog.fns[f].vba_name().to_lowercase()).collect()),
            _ => None,
        })
        .unwrap_or_default();
    let mut out = Vec::new();
    for m in 0..prog.modules.len() {
        if prog.modules[m].kind != ModuleKind::User {
            continue;
        }
        let mut w = Walker { prog, module: m, functions: &functions, seen: HashSet::new(), out: Vec::new() };
        visit_module(&prog.modules[m].ast, &mut w);
        out.append(&mut w.out);
    }
    out
}
