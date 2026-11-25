//! JavaScript back end (spec §14.10): typed AST → CommonJS modules with source maps.
//!
//! The output is not meant to be fast or small. It exists so that a RAS project can be
//! stepped through in a JavaScript debugger: one JS statement per RAS statement, the same
//! local names, and every Excel access routed through the run-time library (`$rt`), which
//! talks to the real Excel over COM.
//!
//! None of the VBA-side optimisations (inlining, loop fusion, bulk transfer, `DoEvents`
//! insertion) are applied here: the code runs the program as written.

mod expr;
mod stmt;

use std::collections::{BTreeMap, HashMap, HashSet};

use ras_syntax::ast::{self, PassMode, Span};

use crate::diag::{Diagnostic, Severity};
use crate::model::*;

const RUNTIME_JS: &str = include_str!("runtime.js");
const FORMS_JS: &str = include_str!("forms.js");
const EXCEL_JS: &str = include_str!("excel.js");
const SCRIPTING_JS: &str = include_str!("scripting.js");
const MAIN_JS: &str = include_str!("main.js");
const TEST_JS: &str = include_str!("test.js");
const TEST_MAIN_JS: &str = include_str!("test_main.js");

#[derive(Debug, Clone)]
pub struct JsFile {
    /// Path relative to the output directory, `/`-separated (`report.js`, `report.js.map`).
    pub rel_path: String,
    pub text: String,
}

#[derive(Debug, Default)]
pub struct JsOutput {
    pub files: Vec<JsFile>,
    pub diagnostics: Vec<Diagnostic>,
    /// Functions that can be started from the debugger (`Module.proc`), macro entries first,
    /// then the forms (`frmInput.show`).
    pub entries: Vec<String>,
}

/// Generates the debug build of an analysed program. Returns no files when the program has errors.
///
/// `source_prefix` is the path from the output directory back to the project root
/// (`../../` for `build/debug`), used for the `sources` of the source maps.
pub fn build_js(prog: &Program, source_prefix: &str) -> JsOutput {
    let mut out = JsOutput::default();
    if prog.diagnostics.iter().any(|d| d.is_error()) {
        return out;
    }
    let mut module_files = Vec::new();
    let mut macros = Vec::new();
    let mut others = Vec::new();
    let mut forms = Vec::new();
    let mut tests = Vec::new();
    let mut module_names = serde_json::Map::new();
    for m in 0..prog.modules.len() {
        // Bundled library modules with code of their own (the struct map) run as RAS code too.
        let std_code =
            prog.modules[m].kind == ModuleKind::Std && !prog.modules[m].classes.is_empty();
        if prog.modules[m].kind != ModuleKind::User && !std_code {
            continue;
        }
        // A form's code-behind is skipped only when its design is missing (an error already).
        if prog.modules[m]
            .form
            .as_ref()
            .is_some_and(|f| f.decl.is_none())
        {
            continue;
        }
        let mut g = Gen::new(prog, m);
        g.emit_module();
        let md = &prog.modules[m];
        let key = js_key(md);
        let depth = key.matches('/').count();
        let up = "../".repeat(depth);
        let file_name = format!("{key}.js");
        let base_name = file_name
            .rsplit('/')
            .next()
            .unwrap_or(&file_name)
            .to_string();
        let source = format!("{up}{source_prefix}{}", md.rel_path);
        let (text, map) = finish(&g.assemble(&up), md, &base_name, &source);
        out.files.push(JsFile {
            rel_path: file_name.clone(),
            text: format!("{text}//# sourceMappingURL={base_name}.map\n"),
        });
        out.files.push(JsFile {
            rel_path: format!("{file_name}.map"),
            text: map,
        });
        if md.is_test() {
            // Run by `ras test` (ras-test-main.js), which loads the test file it is given.
            if prog.is_test_file(m) {
                tests.push(serde_json::json!({ "file": md.rel_path, "module": md.vba_name, "js": format!("./{file_name}") }));
            }
            out.diagnostics.append(&mut g.diags);
            continue;
        }
        module_names.insert(
            md.vba_name.clone(),
            serde_json::Value::String(format!("./{file_name}")),
        );
        module_files.push(file_name);
        if md.is_form() {
            // A form is started by showing it (spec §14.10.7).
            forms.push(format!("{}.show", md.vba_name));
        } else {
            macros.extend(g.macros.iter().map(|n| format!("{}.{n}", md.vba_name)));
            others.extend(
                g.entries
                    .iter()
                    .filter(|n| !g.macros.contains(n))
                    .map(|n| format!("{}.{n}", md.vba_name)),
            );
        }
        out.diagnostics.append(&mut g.diags);
    }
    out.entries = macros;
    out.entries.extend(others);
    out.entries.extend(forms);
    let list = module_files
        .iter()
        .map(|f| format!("    {},\n", js_str(&format!("./{f}"))))
        .collect::<String>();
    let form_names: Vec<&str> = prog
        .modules
        .iter()
        .filter(|m| m.is_form() && m.form.as_ref().is_some_and(|f| f.decl.is_some()))
        .map(|m| m.vba_name.as_str())
        .collect();
    let manifest = serde_json::json!({
        "project": prog.config.project.name,
        "workbook": prog.config.project.workbook,
        "entries": out.entries,
        "forms": form_names,
        "modules": module_names,
    });
    if prog.modules.iter().any(|m| m.is_test()) {
        let mut manifest = manifest.clone();
        manifest["tests"] = serde_json::Value::Array(tests);
        manifest["eventErrors"] = serde_json::json!(prog.config.test.event_errors);
        let main = TEST_MAIN_JS
            .replace("/*@MODULES@*/", list.trim())
            .replace("/*@MANIFEST@*/", &manifest.to_string())
            .replace(
                "/*@ROOT@*/",
                &js_str(if source_prefix.is_empty() {
                    "."
                } else {
                    source_prefix
                }),
            );
        out.files.push(JsFile {
            rel_path: "ras-test-main.js".into(),
            text: main,
        });
        out.files.push(JsFile {
            rel_path: "ras-test.js".into(),
            text: TEST_JS.to_string(),
        });
    }
    let main = MAIN_JS
        .replace("/*@MODULES@*/", list.trim())
        .replace("/*@MANIFEST@*/", &manifest.to_string())
        .replace(
            "/*@ROOT@*/",
            &js_str(if source_prefix.is_empty() {
                "."
            } else {
                source_prefix
            }),
        );
    out.files.push(JsFile {
        rel_path: "ras-main.js".into(),
        text: main,
    });
    out.files.push(JsFile {
        rel_path: "ras-runtime.js".into(),
        text: RUNTIME_JS.to_string(),
    });
    out.files.push(JsFile {
        rel_path: "ras-forms.js".into(),
        text: FORMS_JS.to_string(),
    });
    out.files.push(JsFile {
        rel_path: "ras-excel.js".into(),
        text: EXCEL_JS.to_string(),
    });
    out.files.push(JsFile {
        rel_path: "ras-scripting.js".into(),
        text: SCRIPTING_JS.to_string(),
    });
    out.files.push(JsFile {
        rel_path: "ras-debug.json".into(),
        text: format!(
            "{}\n",
            serde_json::to_string_pretty(&manifest).unwrap_or_default()
        ),
    });
    if out.diagnostics.iter().any(|d| d.is_error()) {
        out.files.clear();
    }
    out
}

/// Path of a module's JavaScript file without `.js`: the import key, except for a form's
/// code-behind (`@forms/frmInput`: its key `<forms>/frmInput` is not a valid file name).
fn js_key(md: &ModuleData) -> String {
    match (
        md.key.strip_prefix(crate::collect::FORM_KEY_PREFIX),
        md.key.strip_prefix(crate::collect::OBJECT_KEY_PREFIX),
    ) {
        (Some(name), _) => format!("@forms/{name}"),
        (_, Some(name)) => format!("@objects/{name}"),
        _ => md.key.clone(),
    }
}

/// A JavaScript string literal (JSON strings are valid JavaScript).
pub(crate) fn js_str(s: &str) -> String {
    serde_json::to_string(s).unwrap_or_else(|_| "\"\"".into())
}

/// Names that cannot be used as (or would shadow something the generated code needs as) a JS identifier.
const JS_RESERVED: &[&str] = &[
    "arguments",
    "await",
    "break",
    "case",
    "catch",
    "class",
    "const",
    "continue",
    "debugger",
    "default",
    "delete",
    "do",
    "else",
    "enum",
    "eval",
    "export",
    "extends",
    "false",
    "finally",
    "for",
    "function",
    "if",
    "implements",
    "import",
    "in",
    "instanceof",
    "interface",
    "let",
    "new",
    "null",
    "package",
    "private",
    "protected",
    "public",
    "return",
    "static",
    "super",
    "switch",
    "this",
    "throw",
    "true",
    "try",
    "typeof",
    "var",
    "void",
    "while",
    "with",
    "yield",
    "undefined",
    "NaN",
    "Infinity",
    "Math",
    "console",
    "Map",
    "String",
    "Number",
    "Object",
    "Array",
    "Error",
    "Symbol",
    "require",
    "exports",
    "module",
    "process",
    "globalThis",
    "Date",
    "JSON",
    "Boolean",
    "Function",
    "Promise",
    "Proxy",
    "Reflect",
    "Set",
];

pub(crate) fn js_ident(name: &str) -> String {
    if JS_RESERVED.contains(&name) {
        format!("{name}$")
    } else {
        name.to_string()
    }
}

const MARK_START: char = '\u{1}';
const MARK_END: char = '\u{2}';

/// Strips the position markers from generated code and builds the source map (version 3).
fn finish(code: &str, md: &ModuleData, file: &str, source: &str) -> (String, String) {
    let mut text = String::with_capacity(code.len());
    let mut mappings = String::new();
    let (mut gen_col, mut prev_gen_col) = (0u32, 0u32);
    let (mut prev_line, mut prev_col) = (0i64, 0i64);
    let mut first_in_line = true;
    let mut chars = code.chars();
    while let Some(c) = chars.next() {
        match c {
            MARK_START => {
                let mut num = String::new();
                for d in chars.by_ref() {
                    if d == MARK_END {
                        break;
                    }
                    num.push(d);
                }
                let offset: u32 = num.parse().unwrap_or(0);
                let (line, col) = md.line_col_utf16(offset);
                if !first_in_line {
                    mappings.push(',');
                }
                vlq(&mut mappings, gen_col as i64 - prev_gen_col as i64);
                vlq(&mut mappings, 0);
                vlq(&mut mappings, line as i64 - prev_line);
                vlq(&mut mappings, col as i64 - prev_col);
                prev_gen_col = gen_col;
                prev_line = line as i64;
                prev_col = col as i64;
                first_in_line = false;
            }
            '\n' => {
                text.push('\n');
                mappings.push(';');
                gen_col = 0;
                prev_gen_col = 0;
                first_in_line = true;
            }
            c => {
                text.push(c);
                gen_col += c.len_utf16() as u32;
            }
        }
    }
    let mut map = serde_json::json!({
        "version": 3,
        "file": file,
        "sources": [source],
        "names": [],
        "mappings": mappings,
    });
    // The standard library is not a file of the project: the debugger shows it from the map.
    if md.rel_path.starts_with("<stdlib>/") {
        map["sourcesContent"] = serde_json::json!([&*md.text]);
    }
    (text, map.to_string())
}

fn vlq(out: &mut String, value: i64) {
    const B64: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut v = if value < 0 {
        ((-value) << 1) | 1
    } else {
        value << 1
    } as u64;
    loop {
        let mut digit = (v & 31) as usize;
        v >>= 5;
        if v != 0 {
            digit |= 32;
        }
        out.push(B64[digit] as char);
        if v == 0 {
            break;
        }
    }
}

pub(crate) struct Gen<'p> {
    pub prog: &'p Program,
    /// Module being emitted (its `TypeInfo` describes every expression).
    pub module: ModuleId,
    /// Function whose locals are in scope.
    pub fn_id: Option<FnId>,
    /// Return type of the function being emitted (`None` inside arrow functions).
    pub ret_ty: Option<Type>,
    bufs: Vec<String>,
    pub indent: usize,
    pub diags: Vec<Diagnostic>,
    /// > 0 inside `try { ... } catch { resume; }`: every simple statement ignores its own errors.
    pub resume: u32,
    /// Receiver that unqualified `Cells` / `Range` inside `ws.Range(...)` refer to (spec §13.4).
    pub qualifier: Option<String>,
    pub tmp_n: u32,
    imports: BTreeMap<ModuleId, String>,
    aliases: HashSet<String>,
    reported: HashSet<&'static str>,
    pub entries: Vec<String>,
    pub macros: Vec<String>,
}

impl<'p> Gen<'p> {
    fn new(prog: &'p Program, module: ModuleId) -> Self {
        Gen {
            prog,
            module,
            fn_id: None,
            ret_ty: None,
            bufs: vec![String::new()],
            indent: 0,
            diags: Vec::new(),
            resume: 0,
            qualifier: None,
            tmp_n: 0,
            imports: BTreeMap::new(),
            aliases: HashSet::new(),
            reported: HashSet::new(),
            entries: Vec::new(),
            macros: Vec::new(),
        }
    }

    // ------------------------------------------------------------------ output helpers

    pub(crate) fn w(&mut self, text: impl AsRef<str>) {
        let pad = "    ".repeat(self.indent);
        let buf = self.bufs.last_mut().unwrap();
        buf.push_str(&pad);
        buf.push_str(text.as_ref());
        buf.push('\n');
    }

    /// Writes a line that starts a RAS statement: breakpoints and stepping map to `span`.
    pub(crate) fn wm(&mut self, span: Span, text: impl AsRef<str>) {
        let marked = format!(
            "{MARK_START}{}{MARK_END}{}",
            u32::from(span.start()),
            text.as_ref()
        );
        self.w(marked);
    }

    /// Runs `f` with a fresh buffer and returns what it wrote.
    pub(crate) fn capture(&mut self, f: impl FnOnce(&mut Self)) -> String {
        self.bufs.push(String::new());
        f(self);
        self.bufs.pop().unwrap()
    }

    pub(crate) fn pad(&self) -> String {
        "    ".repeat(self.indent)
    }

    pub(crate) fn tmp(&mut self, base: &str) -> String {
        self.tmp_n += 1;
        format!("${base}{}", self.tmp_n)
    }

    /// Reports a construct the debug run time cannot execute (once per kind and module) and
    /// returns an expression that raises when it is reached.
    pub(crate) fn unsupported(&mut self, span: Span, what: &'static str) -> String {
        if self.reported.insert(what) {
            self.diags.push(Diagnostic {
                severity: Severity::Warning,
                code: "W1410",
                message: format!("{what} はデバッグ実行（JavaScript ランタイム）では動作しません。実行時にエラーになります"),
                module: Some(self.module),
                span,
            });
        }
        format!("$rt.unsupported({})", js_str(what))
    }

    // ------------------------------------------------------------------ lookups

    pub(crate) fn info(&self) -> &'p TypeInfo {
        &self.prog.modules[self.module].info
    }

    pub(crate) fn ty(&self, e: &ast::Expr) -> Type {
        self.info().ty(e.id)
    }

    pub(crate) fn local(&self, l: LocalId) -> Option<&'p LocalDef> {
        self.fn_id.and_then(|f| self.prog.fns[f].locals.get(l))
    }

    pub(crate) fn local_ty(&self, l: LocalId) -> Type {
        self.local(l).map(|d| d.ty.clone()).unwrap_or(Type::Any)
    }

    /// JS name of a local (`$this` for the receiver of an extension method).
    pub(crate) fn local_name(&self, l: LocalId) -> String {
        match self.local(l) {
            Some(d) if d.kind == LocalKind::This => "$this".into(),
            Some(d) => js_ident(&d.name),
            None => "undefined".into(),
        }
    }

    /// Expression reading a local (by-reference parameters are `{ v }` cells).
    pub(crate) fn local_js(&self, l: LocalId) -> String {
        let name = self.local_name(l);
        match self.local(l) {
            Some(d) if d.byref => format!("{name}.v"),
            _ => name,
        }
    }

    fn module_alias(&mut self, m: ModuleId) -> String {
        if let Some(a) = self.imports.get(&m) {
            return a.clone();
        }
        let key = &self.prog.modules[m].key;
        let mut alias = format!(
            "$m_{}",
            key.chars()
                .map(|c| if c.is_ascii_alphanumeric() { c } else { '_' })
                .collect::<String>()
        );
        if !self.aliases.insert(alias.clone()) {
            alias = format!("{alias}_{m}");
            self.aliases.insert(alias.clone());
        }
        self.imports.insert(m, alias.clone());
        alias
    }

    /// A top-level name of module `m`, as seen from the module being emitted.
    pub(crate) fn module_member(&mut self, m: ModuleId, name: &str) -> String {
        if m == self.module {
            name.to_string()
        } else {
            format!("{}.{name}", self.module_alias(m))
        }
    }

    pub(crate) fn class_ref(&mut self, c: ClassId) -> String {
        let d = &self.prog.classes[c];
        self.module_member(d.module, &js_ident(&d.name))
    }

    pub(crate) fn struct_ref(&mut self, s: StructId) -> String {
        let d = &self.prog.structs[s];
        self.module_member(d.module, &js_ident(&d.name))
    }

    /// Reference to a RAS function (not a call).
    pub(crate) fn fn_ref(&mut self, f: FnId) -> String {
        let d = &self.prog.fns[f];
        match &d.kind {
            FnKind::Method { class, .. }
            | FnKind::Getter { class, .. }
            | FnKind::Setter { class, .. } => format!("{}.{}", self.class_ref(*class), d.name),
            FnKind::Declare { lib, alias } => format!(
                "$rt.declare({}, {})",
                js_str(lib),
                js_str(alias.as_deref().unwrap_or(&d.name))
            ),
            _ => self.module_member(d.module, &js_ident(&d.name)),
        }
    }

    pub(crate) fn global_ref(&mut self, v: VarId) -> String {
        let d = &self.prog.vars[v];
        if d.declared {
            return match self.prog.modules[d.module].key.as_str() {
                "vba.d" => format!("$rt.vba.{}", d.vba_name()),
                "msforms.d" => format!("$rt.msforms.{}", d.vba_name()),
                "test.d" => format!("$rt.test.{}", d.vba_name()),
                _ => format!("$rt.global({})", js_str(d.vba_name())),
            };
        }
        self.module_member(d.module, &js_ident(&d.name))
    }

    /// Initial value of a variable of type `t`.
    pub(crate) fn zero(&mut self, t: &Type) -> String {
        match t {
            Type::Byte
            | Type::Short
            | Type::Int
            | Type::Long
            | Type::Float
            | Type::Double
            | Type::Currency
            | Type::Decimal
            | Type::Ptr
            | Type::Enum(_) => "0".into(),
            Type::Str | Type::WStr | Type::AStr => "\"\"".into(),
            Type::Bool => "false".into(),
            Type::Date => "$rt.date(0)".into(),
            Type::Array(_) => "[]".into(),
            Type::Struct(s) => format!("new {}()", self.struct_ref(*s)),
            // Type arguments are erased; an instance of a generic class created with a struct
            // argument carries the constructors of its zero values (see `new_expr`).
            Type::Param(n) => format!("(this?.$zero?.{}?.() ?? null)", js_ident(n)),
            Type::Void | Type::Ns(_) => "undefined".into(),
            _ => "null".into(),
        }
    }

    // ------------------------------------------------------------------ module

    fn emit_module(&mut self) {
        let prog = self.prog;
        let md = &prog.modules[self.module];
        // Module-level variables.
        for &v in &md.vars {
            let d = &prog.vars[v];
            if d.declared {
                continue;
            }
            let init = match &d.const_value {
                Some(ConstVal::Int(i)) => i.to_string(),
                Some(ConstVal::Float(f)) => format!("{f:?}"),
                Some(ConstVal::Str(s)) => js_str(s),
                Some(ConstVal::Bool(b)) => b.to_string(),
                None => self.zero(&d.ty),
            };
            self.w(format!("let {} = {init};", js_ident(&d.name)));
        }
        for &s in &md.structs {
            let d = &prog.structs[s];
            let name = js_ident(&d.name);
            self.wm(d.name_span, format!("class {name} extends $rt.Struct {{"));
            self.indent += 1;
            self.w("constructor() {");
            self.indent += 1;
            self.w("super();");
            for (fname, fty) in &d.fields {
                let z = self.zero(fty);
                self.w(format!("this.{fname} = {z};"));
            }
            self.indent -= 1;
            self.w("}");
            self.indent -= 1;
            self.w("}");
        }
        let mut static_inits = String::new();
        for &c in &md.classes {
            if prog.classes[c].kind == ClassKind::Class {
                static_inits.push_str(&self.emit_class(c));
            }
        }
        let mut inits = Vec::new();
        let mut events = Vec::new();
        // Handlers of the objects that module variables hold (`WithEvents`, spec §12.2.2).
        let mut sources: Vec<String> = Vec::new();
        for &f in &md.fns {
            let d = &prog.fns[f];
            if !matches!(d.kind, FnKind::Free | FnKind::Main | FnKind::ModInit) {
                continue;
            }
            let Some(decl) = d.decl.clone() else { continue };
            if decl.body.is_none() {
                continue;
            }
            self.emit_fn(f, "function ", &js_ident(&d.name));
            if matches!(d.kind, FnKind::ModInit) {
                continue;
            }
            if d.attrs.init {
                inits.push(js_ident(&d.name));
            }
            for ev in &d.attrs.on_events {
                match &md.form {
                    // The events of the object a variable holds (spec §12.2.2).
                    Some(_) if prog.event_source(f, &ev.target).is_some() => {
                        if let Some(EventSource::Var(v)) = prog.event_source(f, &ev.target) {
                            let source = format!("() => {}", js_ident(&prog.vars[v].name));
                            let call = js_ident(&d.name);
                            sources.push(self.source_handler(ev, f, &source, &call, None));
                        }
                    }
                    Some(form) if form.kind == ObjectKind::Form => {
                        events.push(self.form_handler(form, ev, f))
                    }
                    Some(form) => events.extend(self.document_handler(form, ev, f)),
                    // An error already (`E0610`): `@on` is written in the file of the sheet.
                    None => {}
                }
            }
            if md.is_form() {
                continue;
            }
            // The bundled library is no part of the project (`z.string()` is no entry point).
            if d.is_debug_entry() && !md.rel_path.starts_with("<stdlib>/") {
                self.entries.push(d.name.clone());
                if d.is_macro_entry() {
                    self.macros.push(d.name.clone());
                }
            }
        }
        if !sources.is_empty() {
            self.w(format!("$rt.withEvents([{}]);", sources.join(", ")));
        }
        if md.is_form() {
            self.register_form(events);
            return;
        }
        // Registration with the run time.
        let init_fn = md.init_fn.map(|f| js_ident(&prog.fns[f].name));
        let init = if static_inits.is_empty() && init_fn.is_none() {
            "null".to_string()
        } else {
            self.w("function $init() {");
            for line in static_inits.lines() {
                self.w(format!("    {line}"));
            }
            if let Some(f) = &init_fn {
                self.w(format!("    {f}();"));
            }
            self.w("}");
            "$init".to_string()
        };
        let entries = self
            .entries
            .iter()
            .map(|n| format!("{}: {}", js_str(n), js_ident(n)))
            .collect::<Vec<_>>()
            .join(", ");
        let macros = self
            .macros
            .iter()
            .map(|n| js_str(n))
            .collect::<Vec<_>>()
            .join(", ");
        self.w(format!("$rt.module({}, {{", js_str(&md.vba_name)));
        self.w(format!("    file: {},", js_str(&md.rel_path)));
        self.w(format!("    init: {init},"));
        self.w(format!("    inits: [{}],", inits.join(", ")));
        self.w(format!("    entries: {{ {entries} }},"));
        self.w(format!("    macros: [{macros}],"));
        self.w(format!("    events: [{}],", events.join(", ")));
        self.w("});");
    }

    /// The entry of `$rt.module` for an `@on` handler of the workbook or a sheet (spec §12.2).
    /// The stand-in for Excel calls it with the arguments of the event as `{ v }` cells, as the
    /// forms do: a `byref` parameter (`cancel`) gets the cell, the others its value. A handler
    /// takes as many of the event's arguments as it declares parameters.
    ///
    /// The target is the sheet's code name, or `Workbook`. Events of controls on a sheet are
    /// raised by the user only: they are not registered.
    fn document_handler(&self, doc: &FormModule, ev: &OnEvent, f: FnId) -> Option<String> {
        let d = &self.prog.fns[f];
        let md = &self.prog.modules[self.module];
        let target = match doc.kind {
            ObjectKind::Workbook
                if ev.target.eq_ignore_ascii_case("Workbook")
                    || ev.target.eq_ignore_ascii_case("ThisWorkbook") =>
            {
                "Workbook".to_string()
            }
            ObjectKind::Sheet
                if ev.target.eq_ignore_ascii_case("Worksheet")
                    || ev.target.to_lowercase() == md.vba_name.to_lowercase() =>
            {
                md.vba_name.clone()
            }
            _ => return None,
        };
        let event = crate::lower::document_event(doc.kind == ObjectKind::Workbook, &ev.event)?;
        let args: Vec<String> = d
            .params
            .iter()
            .enumerate()
            .map(|(i, p)| {
                if p.mode == PassMode::ByRef {
                    format!("$a[{i}]")
                } else {
                    format!("$a[{i}].v")
                }
            })
            .collect();
        Some(format!(
            "{{ target: {}, event: {}, reentrant: {}, fn: ($a) => {}({}) }}",
            js_str(&target),
            js_str(event),
            ev.reentrant,
            js_ident(&d.name),
            args.join(", ")
        ))
    }

    /// The entry of `$rt.withEvents` for an `@on("<variable>.<Event>")` handler (spec §12.2.2):
    /// `source` gives the object the variable holds, `call` is the handler. The stand-in for
    /// Excel calls it with the arguments of the event as `{ v }` cells, as the other handlers.
    fn source_handler(
        &self,
        ev: &OnEvent,
        f: FnId,
        source: &str,
        call: &str,
        alive: Option<&str>,
    ) -> String {
        let d = &self.prog.fns[f];
        let args: Vec<String> = d
            .params
            .iter()
            .enumerate()
            .map(|(i, p)| {
                if p.mode == PassMode::ByRef {
                    format!("$a[{i}]")
                } else {
                    format!("$a[{i}].v")
                }
            })
            .collect();
        format!(
            "{{ target: {}, event: {}, reentrant: {}, source: {source}, {}fn: ($a) => {call}({}) }}",
            js_str(&ev.target),
            js_str(&ev.event),
            ev.reentrant,
            alive.map(|a| format!("alive: {a}, ")).unwrap_or_default(),
            args.join(", ")
        )
    }

    /// The entry of `$rt.form` for an `@on` handler of a form (spec §14.11.2). The run time
    /// passes every argument of the event as a `{ v }` cell; a handler that declares the
    /// parameter `byref` gets the cell, the others its value.
    fn form_handler(&self, form: &FormModule, ev: &OnEvent, f: FnId) -> String {
        let d = &self.prog.fns[f];
        let is_form = ev.target.eq_ignore_ascii_case("UserForm")
            || form
                .decl
                .as_ref()
                .is_some_and(|x| x.name.eq_ignore_ascii_case(&ev.target));
        let target = if is_form {
            "UserForm"
        } else {
            ev.target.as_str()
        };
        let args: Vec<String> = d
            .params
            .iter()
            .enumerate()
            .map(|(i, p)| {
                if p.mode == PassMode::ByRef {
                    format!("$a[{i}]")
                } else {
                    format!("$a[{i}].v")
                }
            })
            .collect();
        format!(
            "{{ target: {}, event: {}, reentrant: {}, fn: ($a) => {}({}) }}",
            js_str(target),
            js_str(&ev.event),
            ev.reentrant,
            js_ident(&d.name),
            args.join(", ")
        )
    }

    /// Registration of a form's code-behind: the run time creates the form (its only, default
    /// instance) and calls `load` with it, which sets `$me` and the module variables.
    fn register_form(&mut self, mut events: Vec<String>) {
        let prog = self.prog;
        let md = &prog.modules[self.module];
        // The form's validators (spec §14.11.12) get the events after the handlers of the
        // code-behind, unguarded, as in the VBA build.
        let validators: Vec<String> = prog
            .form_validators(self.module)
            .into_iter()
            .map(|v| js_ident(&prog.vars[v].name))
            .collect();
        if let Some(decl) = md.form.as_ref().and_then(|f| f.decl.as_ref()) {
            for c in decl.controls.iter().filter(|_| !validators.is_empty()) {
                let hooks: &[(&str, &str)] =
                    if VALIDATED_CONTROLS.contains(&c.control_type.as_str()) {
                        &[("Exit", "commit"), ("Change", "edit")]
                    } else if VALIDATED_CONTAINERS.contains(&c.control_type.as_str()) {
                        &[("Exit", "commitActive")]
                    } else {
                        &[]
                    };
                for (event, method) in hooks {
                    let calls: Vec<String> = validators
                        .iter()
                        .map(|v| {
                            format!(
                                "if ({v} !== null) {v}.{method}($rt.get($me, {}));",
                                js_str(&c.name)
                            )
                        })
                        .collect();
                    events.push(format!(
                        "{{ target: {}, event: {}, reentrant: true, fn: ($a) => {{ {} }} }}",
                        js_str(&c.name),
                        js_str(event),
                        calls.join(" ")
                    ));
                }
            }
        }
        self.w("let $me = null;");
        self.w("function $load($form) {");
        self.indent += 1;
        self.w("$me = $form;");
        for &v in &md.vars {
            let d = &prog.vars[v];
            if d.declared || d.const_value.is_some() {
                continue;
            }
            let z = self.zero(&d.ty);
            self.w(format!("{} = {z};", js_ident(&d.name)));
        }
        if let Some(f) = md.init_fn {
            self.w(format!("{}();", js_ident(&prog.fns[f].name)));
        }
        self.indent -= 1;
        self.w("}");
        // Several instances (`new frmInput()`, spec §14.11.9): the run time swaps their state.
        let state: Vec<String> = md
            .vars
            .iter()
            .map(|&v| &prog.vars[v])
            .filter(|d| !d.declared && d.const_value.is_none())
            .map(|d| js_ident(&d.name))
            .collect();
        self.w(format!(
            "function $save() {{ return {{ {} }}; }}",
            state.join(", ")
        ));
        self.w("function $restore($s, $form) {");
        self.indent += 1;
        self.w("$me = $form;");
        if !state.is_empty() {
            self.w(format!("if ($s) ({{ {} }} = $s);", state.join(", ")));
        }
        self.indent -= 1;
        self.w("}");
        let ctor = md
            .form
            .as_ref()
            .and_then(|f| f.ctor)
            .map(|f| js_ident(&prog.fns[f].name));
        let members: Vec<String> = md
            .fns
            .iter()
            .map(|&f| &prog.fns[f])
            .filter(|d| {
                d.exported
                    && matches!(d.kind, FnKind::Free)
                    && d.decl.as_ref().is_some_and(|x| x.body.is_some())
            })
            .map(|d| format!("{}: {}", js_str(&d.name), js_ident(&d.name)))
            .collect();
        let dir = md
            .rel_path
            .rsplit_once('/')
            .map(|(d, _)| format!("{d}/"))
            .unwrap_or_default();
        self.w(format!("$rt.form({}, {{", js_str(&md.vba_name)));
        self.w(format!("    file: {},", js_str(&md.rel_path)));
        self.w(format!(
            "    spec: {},",
            js_str(&format!("{dir}{}.form.json", md.vba_name))
        ));
        self.w("    load: $load,");
        self.w("    save: $save,");
        self.w("    restore: $restore,");
        if let Some(c) = ctor {
            self.w(format!("    ctor: {c},"));
        }
        self.w(format!("    members: {{ {} }},", members.join(", ")));
        self.w(format!("    events: [{}],", events.join(", ")));
        self.w("});");
    }

    /// Puts the header (exports, imports) in front of the generated body.
    fn assemble(&mut self, up: &str) -> String {
        let prog = self.prog;
        let md = &prog.modules[self.module];
        let mut head = String::new();
        head.push_str(&format!(
            "// Generated by ras from {} — debug build, do not edit.\n",
            md.rel_path
        ));
        head.push_str("\"use strict\";\n");
        head.push_str(&format!(
            "const $rt = require({});\n",
            js_str(&format!("./{up}ras-runtime.js"))
        ));
        // Exports come before the imports so that import cycles between modules work.
        for &f in &md.fns {
            let d = &prog.fns[f];
            if matches!(d.kind, FnKind::Free | FnKind::Main)
                && d.decl.as_ref().is_some_and(|x| x.body.is_some())
            {
                let n = js_ident(&d.name);
                head.push_str(&format!("exports.{n} = {n};\n"));
            }
        }
        for &v in &md.vars {
            let d = &prog.vars[v];
            if !d.declared {
                let n = js_ident(&d.name);
                head.push_str(&format!("Object.defineProperty(exports, {}, {{ get: () => {n}, set: ($v) => {{ {n} = $v; }}, enumerable: true }});\n", js_str(&n)));
            }
        }
        for (m, alias) in &self.imports {
            let key = js_key(&prog.modules[*m]);
            head.push_str(&format!(
                "const {alias} = require({});\n",
                js_str(&format!("./{up}{key}.js"))
            ));
        }
        let mut body = std::mem::take(&mut self.bufs[0]);
        // Classes are not hoisted: export them where they are defined.
        for &c in &md.classes {
            if prog.classes[c].kind == ClassKind::Class {
                let n = js_ident(&prog.classes[c].name);
                body.push_str(&format!("exports.{n} = {n};\n"));
            }
        }
        for &s in &md.structs {
            let n = js_ident(&prog.structs[s].name);
            body.push_str(&format!("exports.{n} = {n};\n"));
        }
        head.push_str(&body);
        head
    }

    // ------------------------------------------------------------------ functions and classes

    /// Parameter list of a function, with default values.
    fn params(&mut self, f: FnId) -> String {
        let prog = self.prog;
        let def = &prog.fns[f];
        let decl = def.decl.clone();
        let mut out = Vec::new();
        for (i, p) in def.params.iter().enumerate() {
            let name = self.local_name(p.local);
            if p.rest {
                out.push(format!("...{name}"));
                continue;
            }
            if p.mode == PassMode::ByRef {
                out.push(name);
                continue;
            }
            let default = decl
                .as_ref()
                .and_then(|d| d.params.get(i))
                .and_then(|ap| ap.default.as_ref());
            match default {
                Some(d) => {
                    let v = self.coerce_store(d, &p.ty);
                    out.push(format!("{name} = {v}"));
                }
                None if p.optional => {
                    let z = self.zero(&p.ty);
                    out.push(format!("{name} = {z}"));
                }
                None => out.push(name),
            }
        }
        out.join(", ")
    }

    /// Emits `prefix name(params) { body }` for function `f`.
    fn emit_fn(&mut self, f: FnId, prefix: &str, name: &str) {
        let prog = self.prog;
        let def = &prog.fns[f];
        let Some(decl) = def.decl.clone() else { return };
        let Some(body) = decl.body.as_ref() else {
            return;
        };
        let (saved_fn, saved_ret, saved_resume) = (self.fn_id, self.ret_ty.take(), self.resume);
        self.fn_id = Some(f);
        self.ret_ty = Some(def.ret_ty());
        self.resume = 0;
        let params = self.params(f);
        self.wm(def.name_span, format!("{prefix}{name}({params}) {{"));
        self.indent += 1;
        if def.attrs.fast {
            self.w("$rt.suspend(63);");
            self.w("try {");
            self.indent += 1;
        }
        self.stmts(&body.stmts);
        if def.attrs.fast {
            self.indent -= 1;
            self.w("} finally {");
            self.w("    $rt.resume(63);");
            self.w("}");
        }
        let ret = def.ret_ty();
        let ends_with_return = matches!(
            body.stmts.last().map(|s| &s.kind),
            Some(ast::StmtKind::Return(_) | ast::StmtKind::Throw(_))
        );
        if ret != Type::Void
            && !ends_with_return
            && !matches!(def.kind, FnKind::Ctor { .. } | FnKind::Setter { .. })
        {
            // A VBA function that falls off its end returns the default value of its type.
            let z = self.zero(&ret);
            self.w(format!("return {z};"));
        }
        self.indent -= 1;
        self.w("}");
        self.fn_id = saved_fn;
        self.ret_ty = saved_ret;
        self.resume = saved_resume;
    }

    /// Emits a class. Returns the statements initialising its static fields (run from `$init`).
    fn emit_class(&mut self, c: ClassId) -> String {
        let prog = self.prog;
        let def = &prog.classes[c];
        let name = js_ident(&def.name);
        let base = def.extends.iter().find_map(|t| match t {
            Type::Class(b, _) if prog.classes[*b].kind == ClassKind::Class => Some(*b),
            _ => None,
        });
        let extends = match base {
            Some(b) => format!(" extends {}", self.class_ref(b)),
            None => String::new(),
        };
        self.wm(def.name_span, format!("class {name}{extends} {{"));
        self.indent += 1;
        // Constructor: field initialisers, then the user's body.
        let ctor_decl = def.ctor.and_then(|f| prog.fns[f].decl.clone());
        let (saved_fn, saved_ret) = (self.fn_id, self.ret_ty.take());
        self.fn_id = def.ctor;
        let params = match def.ctor {
            Some(f) => self.params(f),
            None => String::new(),
        };
        self.w(format!("constructor({params}) {{"));
        self.indent += 1;
        let body = ctor_decl.as_ref().and_then(|d| d.body.as_ref());
        let calls_super = body.is_some_and(|b| {
            b.stmts.first().is_some_and(|s| matches!(&s.kind, ast::StmtKind::Expr(e) if matches!(&e.kind, ast::ExprKind::Call { callee, .. } if matches!(callee.kind, ast::ExprKind::Super))))
        });
        if base.is_some() && !calls_super {
            self.w("super();");
        }
        let mut first_stmt = 0;
        if calls_super {
            if let Some(b) = body {
                self.stmts(&b.stmts[..1]);
                first_stmt = 1;
            }
        }
        for fd in def.fields.iter().filter(|f| !f.is_static) {
            let v = match &fd.init {
                Some(init) => self.coerce_store(init, &fd.ty),
                None => self.zero(&fd.ty),
            };
            self.wm(fd.span, format!("this.{} = {v};", fd.name));
        }
        // The handlers of the objects its fields hold (`WithEvents`, spec §12.2.2). The run
        // time must not keep the instance alive: it is reached through a weak reference.
        let mut sources: Vec<String> = Vec::new();
        for &m in &def.methods {
            for ev in &prog.fns[m].attrs.on_events {
                if let Some(EventSource::Field(_, i)) = prog.event_source(m, &ev.target) {
                    let source = format!("() => $w.deref()?.{}", def.fields[i].name);
                    let call = format!("$w.deref()?.{}", prog.fns[m].name);
                    sources.push(self.source_handler(
                        ev,
                        m,
                        &source,
                        &call,
                        Some("() => $w.deref() !== undefined"),
                    ));
                }
            }
        }
        if !sources.is_empty() {
            self.w("const $w = new WeakRef(this);");
            self.w(format!("$rt.withEvents([{}]);", sources.join(", ")));
        }
        if let Some(b) = body {
            self.ret_ty = Some(Type::Void);
            self.stmts(&b.stmts[first_stmt..]);
        }
        self.indent -= 1;
        self.w("}");
        self.fn_id = saved_fn;
        self.ret_ty = saved_ret;
        for &f in &def.methods {
            let d = &prog.fns[f];
            let prefix = match d.kind {
                FnKind::Method { is_static, .. } => {
                    if is_static { "static " } else { "" }.to_string()
                }
                FnKind::Getter { is_static, .. } => {
                    format!("{}get ", if is_static { "static " } else { "" })
                }
                FnKind::Setter { is_static, .. } => {
                    format!("{}set ", if is_static { "static " } else { "" })
                }
                _ => continue,
            };
            self.emit_fn(f, &prefix, &d.name);
        }
        self.indent -= 1;
        self.w("}");
        let implemented: Vec<String> = self
            .implemented(c)
            .into_iter()
            .map(|n| js_str(&n))
            .collect();
        self.w(format!("{name}.$impl = [{}];", implemented.join(", ")));
        // Static fields get their default now and their initialiser when the project starts.
        let mut inits = String::new();
        for fd in def.fields.iter().filter(|f| f.is_static) {
            let z = self.zero(&fd.ty);
            self.w(format!("{name}.{} = {z};", fd.name));
            if let Some(init) = &fd.init {
                let v = self.coerce_store(init, &fd.ty);
                inits.push_str(&format!("{name}.{} = {v};\n", fd.name));
            }
        }
        inits
    }

    /// Names of every interface a class implements (directly or through its bases).
    fn implemented(&self, c: ClassId) -> Vec<String> {
        let mut out = Vec::new();
        let mut seen = HashMap::new();
        let mut stack = vec![c];
        while let Some(x) = stack.pop() {
            if seen.insert(x, ()).is_some() {
                continue;
            }
            let d = &self.prog.classes[x];
            if d.kind == ClassKind::Interface {
                out.push(d.name.clone());
            }
            for t in d.extends.iter().chain(d.implements.iter()) {
                if let Type::Class(b, _) = t {
                    stack.push(*b);
                }
            }
        }
        out.sort();
        out
    }
}
