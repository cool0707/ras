//! Declaration collection: parses every source file, allocates definitions,
//! resolves imports and signatures. Function bodies are checked in `check.rs`.

use std::collections::HashMap;
use std::path::PathBuf;
use std::rc::Rc;

use ras_syntax::ast::{self, ExprKind, InlineMode, Item, PassMode, Span, TypeRef, Visibility};

use crate::config::Config;
use crate::diag::{Diagnostic, Severity};
use crate::model::*;

#[derive(Clone)]
pub struct SourceInput {
    pub path: PathBuf,
    pub rel_path: String,
    /// Import key (`strutil`, `sub/util`, `win32`).
    pub key: String,
    pub kind: ModuleKind,
    pub text: String,
}

/// Converts a file stem to a VBA module name (`str_util` → `StrUtil`, `report` → `Report`).
pub fn vba_module_name(key: &str) -> String {
    let stem = key.rsplit('/').next().unwrap_or(key);
    let mut out = String::new();
    let mut upper = true;
    for c in stem.chars() {
        if c == '_' || c == '-' || c == '.' || c == ' ' {
            upper = true;
            continue;
        }
        if upper {
            out.extend(c.to_uppercase());
            upper = false;
        } else {
            out.push(c);
        }
    }
    if out.is_empty() || out.chars().next().is_some_and(|c| c.is_ascii_digit()) {
        out.insert_str(0, "Mod");
    }
    let out: String = out.chars().take(31).collect();
    // `circle.ras` must not become a module named `Circle` (a syntax error wherever it is referenced).
    if crate::names::is_reserved(&out) {
        format!("{out}_")
    } else {
        out
    }
}

/// Evaluates a constant integer expression (enum values, array sizes).
pub fn const_int(e: &ast::Expr) -> Option<i128> {
    match &e.kind {
        ExprKind::Int { value, .. } => Some(*value),
        ExprKind::Paren(inner) => const_int(inner),
        ExprKind::Unary {
            op: ast::UnaryOp::Neg,
            operand,
        } => const_int(operand).map(|v| -v),
        ExprKind::Unary {
            op: ast::UnaryOp::Plus,
            operand,
        } => const_int(operand),
        ExprKind::Unary {
            op: ast::UnaryOp::BitNot,
            operand,
        } => const_int(operand).map(|v| !v),
        ExprKind::Binary { op, lhs, rhs } => {
            let (a, b) = (const_int(lhs)?, const_int(rhs)?);
            match op {
                ast::BinaryOp::Add => a.checked_add(b),
                ast::BinaryOp::Sub => a.checked_sub(b),
                ast::BinaryOp::Mul => a.checked_mul(b),
                ast::BinaryOp::Div if b != 0 => Some(a / b),
                ast::BinaryOp::Rem if b != 0 => Some(a % b),
                ast::BinaryOp::BitOr => Some(a | b),
                ast::BinaryOp::BitAnd => Some(a & b),
                ast::BinaryOp::BitXor => Some(a ^ b),
                ast::BinaryOp::Shl if (0..64).contains(&b) => Some(a << b),
                ast::BinaryOp::Shr if (0..64).contains(&b) => Some(a >> b),
                _ => None,
            }
        }
        _ => None,
    }
}

struct Collector {
    prog: Program,
    fn_decls: Vec<(FnId, Option<ClassId>)>,
    class_decls: Vec<(ClassId, ClassAst)>,
    var_decls: Vec<(VarId, Rc<ast::VarDecl>)>,
    struct_decls: Vec<(StructId, Rc<ast::StructDecl>)>,
    alias_decls: Vec<(ModuleId, Rc<ast::TypeAlias>)>,
    /// Namespace currently being declared, and the namespace of each function / variable.
    cur_ns: Option<String>,
    fn_ns: HashMap<FnId, String>,
    var_ns: HashMap<VarId, String>,
}

enum ClassAst {
    Class(Rc<ast::ClassDecl>),
    Interface(Rc<ast::InterfaceDecl>),
}

/// Where a top-level name of a module is registered.
#[derive(Clone)]
enum Target {
    /// Module scope (user / std modules).
    Module(ModuleId),
    /// Global scope (top level of a `.d.ras` file).
    Global,
    /// Host namespace (`declare namespace Excel { ... }`).
    Namespace(String),
}

pub fn collect(inputs: Vec<SourceInput>, config: Config) -> Program {
    let mut c = Collector {
        prog: Program {
            config,
            ..Default::default()
        },
        fn_decls: Vec::new(),
        class_decls: Vec::new(),
        var_decls: Vec::new(),
        struct_decls: Vec::new(),
        alias_decls: Vec::new(),
        cur_ns: None,
        fn_ns: HashMap::new(),
        var_ns: HashMap::new(),
    };
    for input in inputs {
        c.add_module(input);
    }
    c.check_module_names();
    for m in 0..c.prog.modules.len() {
        c.declare_module_items(m);
    }
    // The files of sheets come first: their code names get the type with the file's members.
    c.add_forms();
    c.add_sheet_code_names();
    for m in 0..c.prog.modules.len() {
        c.resolve_imports(m);
    }
    // The API of tests, where the module does not declare or import the name itself.
    for m in 0..c.prog.modules.len() {
        if c.prog.modules[m].is_test() {
            for (name, sym) in c.prog.test_globals.clone() {
                c.prog.modules[m].scope.entry(name).or_insert(sym);
            }
        }
    }
    c.resolve_signatures();
    c.check_document_names();
    c.prog.cur_namespace = None;
    for m in 0..c.prog.modules.len() {
        if !c.is_bundled_typedef(m) {
            continue;
        }
        for &cid in &c.prog.modules[m].classes.clone() {
            let def = &c.prog.classes[cid];
            let mut spelled: Vec<String> = def
                .fields
                .iter()
                .map(|f| f.vba_name().to_string())
                .collect();
            spelled.extend(
                def.methods
                    .iter()
                    .map(|&f| c.prog.fns[f].vba_name().to_string()),
            );
            for s in spelled {
                c.prog
                    .host_member_spelling
                    .entry(s.to_lowercase())
                    .or_insert(s);
            }
        }
    }
    c.prog
}

impl Collector {
    fn err(
        &mut self,
        module: ModuleId,
        span: Span,
        code: &'static str,
        message: impl Into<String>,
    ) {
        self.prog.diagnostics.push(Diagnostic {
            severity: Severity::Error,
            code,
            message: message.into(),
            module: Some(module),
            span,
        });
    }

    fn warn(
        &mut self,
        module: ModuleId,
        span: Span,
        code: &'static str,
        message: impl Into<String>,
    ) {
        self.prog.diagnostics.push(Diagnostic {
            severity: Severity::Warning,
            code,
            message: message.into(),
            module: Some(module),
            span,
        });
    }

    fn add_module(&mut self, input: SourceInput) {
        let id = self.prog.modules.len();
        let parse = ras_syntax::parse(&input.text);
        for e in &parse.errors {
            self.prog.diagnostics.push(Diagnostic {
                severity: Severity::Error,
                code: "E0001",
                message: e.message.clone(),
                module: Some(id),
                span: e.range,
            });
        }
        let line_starts = line_starts(&input.text);
        let mut info = TypeInfo::default();
        info.expr_types = vec![Type::Error; parse.ast.expr_count as usize];
        // A form's code-behind is the form's code module: its VBA name is the form's name.
        // The file of a sheet (of `ThisWorkbook`) is that document module (spec §12.2).
        let form = match (
            input.key.strip_prefix(FORM_KEY_PREFIX),
            input.key.strip_prefix(OBJECT_KEY_PREFIX),
        ) {
            (Some(name), _) if input.kind == ModuleKind::User => Some((
                name.to_string(),
                FormModule {
                    kind: ObjectKind::Form,
                    decl: self
                        .prog
                        .config
                        .forms
                        .iter()
                        .find(|f| f.name.eq_ignore_ascii_case(name))
                        .cloned(),
                    class: None,
                    ctor: None,
                    result: None,
                    modal: false,
                },
            )),
            (_, Some(name)) if input.kind == ModuleKind::User => {
                let cfg = &self.prog.config;
                let (kind, code_name) = if name.eq_ignore_ascii_case("ThisWorkbook") {
                    (ObjectKind::Workbook, "ThisWorkbook".to_string())
                } else {
                    // The code name as the workbook spells it.
                    (
                        ObjectKind::Sheet,
                        cfg.sheet_code_names
                            .iter()
                            .find(|n| n.to_lowercase() == name.to_lowercase())
                            .cloned()
                            .unwrap_or_else(|| name.to_string()),
                    )
                };
                let controls = cfg
                    .sheet_controls
                    .iter()
                    .find(|d| d.name.to_lowercase() == code_name.to_lowercase())
                    .map(|d| d.controls.clone())
                    .unwrap_or_default();
                Some((
                    code_name.clone(),
                    FormModule {
                        kind,
                        decl: Some(crate::config::FormDecl {
                            name: code_name,
                            controls,
                        }),
                        class: None,
                        ctor: None,
                        result: None,
                        modal: false,
                    },
                ))
            }
            _ => None,
        };
        let vba_name = match &form {
            Some((name, f)) => f
                .decl
                .as_ref()
                .map(|d| d.name.clone())
                .unwrap_or_else(|| name.clone()),
            // A test module is no VBA module: its name only tells the modules apart at run time.
            None if input.kind == ModuleKind::User && input.key.starts_with(TEST_KEY_PREFIX) => {
                format!(
                    "T_{}",
                    input.key[TEST_KEY_PREFIX.len()..]
                        .chars()
                        .map(|c| if c.is_alphanumeric() { c } else { '_' })
                        .collect::<String>()
                )
            }
            None => vba_module_name(&input.key),
        };
        self.prog.modules.push(ModuleData {
            id,
            path: input.path,
            rel_path: input.rel_path,
            vba_name,
            key: input.key,
            kind: input.kind,
            text: Rc::from(input.text),
            ast: Rc::new(parse.ast),
            scope: HashMap::new(),
            exports: HashMap::new(),
            ext_methods: Vec::new(),
            info,
            main_fn: None,
            init_fn: None,
            fns: Vec::new(),
            classes: Vec::new(),
            enums: Vec::new(),
            structs: Vec::new(),
            vars: Vec::new(),
            imports: Vec::new(),
            line_starts,
            form: form.map(|f| f.1),
        });
    }

    fn check_module_names(&mut self) {
        let mut seen: HashMap<String, ModuleId> = HashMap::new();
        for m in 0..self.prog.modules.len() {
            if self.prog.modules[m].kind == ModuleKind::TypeDef {
                continue;
            }
            let lower = self.prog.modules[m].vba_name.to_lowercase();
            // A standard module cannot have the name of a sheet or of the workbook: their code
            // is written in `objects/` (spec §12.1.1).
            if self.prog.modules[m].form.is_none()
                && self.prog.modules[m].kind == ModuleKind::User
                && !self.prog.modules[m].is_test()
            {
                let sheet = self
                    .prog
                    .config
                    .sheet_code_names
                    .iter()
                    .find(|n| n.to_lowercase() == lower)
                    .cloned();
                let document = if lower == "thisworkbook" {
                    Some("ThisWorkbook".to_string())
                } else {
                    sheet
                };
                if let Some(doc) = document {
                    let msg = format!(
                        "VBA モジュール名 `{}` が{}と重複します。{}のコードは `objects/{doc}.ras` に書きます（`src/` のファイルは標準モジュールになります）",
                        self.prog.modules[m].vba_name,
                        if doc == "ThisWorkbook" { "ブックのモジュール `ThisWorkbook` ".to_string() } else { format!("シートのコード名 `{doc}` ") },
                        if doc == "ThisWorkbook" { "ブック" } else { "シート" }
                    );
                    self.err(m, Span::default(), "E0102", msg);
                    continue;
                }
            }
            if let Some(&other) = seen.get(&lower) {
                let msg = format!(
                    "VBA モジュール名 `{}` が `{}` と重複します（VBA のモジュール名は大文字小文字を区別せず、プロジェクト内で一意である必要があります）",
                    self.prog.modules[m].vba_name, self.prog.modules[other].rel_path
                );
                self.err(m, Span::default(), "E0102", msg);
            } else {
                seen.insert(lower, m);
            }
        }
    }

    fn register(
        &mut self,
        module: ModuleId,
        target: &Target,
        name: &ast::Ident,
        sym: Sym,
        exported: bool,
    ) {
        if name.text.is_empty() {
            return;
        }
        let dup = match target {
            Target::Module(m) => {
                let md = &mut self.prog.modules[*m];
                let dup = md.scope.insert(name.text.clone(), sym.clone()).is_some();
                if exported {
                    md.exports.insert(name.text.clone(), sym);
                }
                dup
            }
            Target::Global if self.prog.modules[module].key == TEST_DEFS_MODULE => self
                .prog
                .test_globals
                .insert(name.text.clone(), sym)
                .is_some(),
            Target::Global => self.prog.globals.insert(name.text.clone(), sym).is_some(),
            Target::Namespace(ns) => self
                .prog
                .namespaces
                .entry(ns.clone())
                .or_default()
                .insert(name.text.clone(), sym)
                .is_some(),
        };
        if dup {
            self.err(
                module,
                name.span,
                "E0101",
                format!("`{}` は既に宣言されています", name.text),
            );
        }
    }

    /// What a form's code-behind cannot contain: it is the code module of the form (an object
    /// module), not a standard module (spec §14.11.2). The same holds for the file of a sheet or
    /// of `ThisWorkbook` (spec §12.2).
    fn check_form_items(&mut self, m: ModuleId, kind: ObjectKind) {
        let ast = self.prog.modules[m].ast.clone();
        for item in &ast.items {
            let (span, what) = match item {
                Item::Stmt(s) if !matches!(s.kind, ast::StmtKind::Empty | ast::StmtKind::Error) => {
                    (s.span, "トップレベルの文")
                }
                Item::Class(c) => (c.name.span, "クラス"),
                Item::Interface(i) => (i.name.span, "インターフェイス"),
                Item::Struct(s) => (s.name.span, "構造体"),
                Item::DeclareModule(d) => (d.span, "`declare module`"),
                Item::Var(v) if v.exported => (v.name.span, "`export` した変数"),
                Item::Enum(e) if e.exported => (e.name.span, "`export` した列挙型"),
                Item::TypeAlias(a) if a.exported => (a.name.span, "`export` した型エイリアス"),
                _ => continue,
            };
            if kind == ObjectKind::Form {
                self.err(m, span, "E0630", format!("フォームのコードビハインドには{what}を書けません（関数・変数・定数・列挙型だけを書けます。共有するものは src/ のモジュールに置いてください）"));
            } else {
                let file = if kind == ObjectKind::Workbook {
                    "`ThisWorkbook` のファイル"
                } else {
                    "シートのファイル"
                };
                self.err(m, span, "E0650", format!("{file}には{what}を書けません（VBA のドキュメントモジュールに置けないためです。関数・変数・定数・列挙型だけを書けます。共有するものは src/ のモジュールに置いてください）"));
            }
        }
    }

    /// The type of `this` in a form's code-behind: a host class named like the form whose fields
    /// are the controls, extending `MSForms.UserForm`. The form's default instance is a global
    /// of that type, so other modules can write `frmInput.show()`.
    fn add_forms(&mut self) {
        let msforms = self
            .prog
            .namespaces
            .get("MSForms")
            .cloned()
            .unwrap_or_default();
        let host = |name: &str| match msforms.get(name) {
            Some(Sym::Class(c)) => Some(*c),
            _ => None,
        };
        let excel = self
            .prog
            .namespaces
            .get("Excel")
            .cloned()
            .unwrap_or_default();
        for m in 0..self.prog.modules.len() {
            let Some(form) = self.prog.modules[m].form.clone() else {
                continue;
            };
            self.check_form_items(m, form.kind);
            if form.kind != ObjectKind::Form {
                self.add_document(m, &form, &excel, &msforms);
                continue;
            }
            let Some(decl) = form.decl else {
                let msg = format!("フォーム `{}` のデザイン（forms/{}.form.json）が見つかりません。`ras form new` で作るか、ブックから `ras form pull` で取り出してください", self.prog.modules[m].vba_name, self.prog.modules[m].vba_name);
                self.err(m, Span::default(), "E0631", msg);
                continue;
            };
            let ident = ast::Ident {
                text: decl.name.clone(),
                span: Span::default(),
            };
            let id = self.alloc_class(
                m,
                &ident,
                ClassKind::Host,
                true,
                Some(format!(
                    "ユーザーフォーム `{}`（forms/{}.form.json）。",
                    decl.name, decl.name
                )),
                &Target::Global,
            );
            // `alloc_class` lists the class under the module; it is not emitted (a host class).
            self.prog.classes[id].extends = host("UserForm")
                .map(|c| Type::Class(c, Vec::new()))
                .into_iter()
                .collect();
            let mut fields = Vec::new();
            for control in &decl.controls {
                let ty = match control.control_type.as_str() {
                    t if t.eq_ignore_ascii_case("Control") || t.is_empty() => Type::Object,
                    t => match msforms.iter().find(|(k, _)| k.eq_ignore_ascii_case(t)) {
                        Some((_, Sym::Class(c))) => Type::Class(*c, Vec::new()),
                        _ => Type::Object,
                    },
                };
                if fields
                    .iter()
                    .any(|f: &FieldDef| f.name.eq_ignore_ascii_case(&control.name))
                {
                    continue;
                }
                fields.push(FieldDef {
                    name: control.name.clone(),
                    host_name: None,
                    ty,
                    vis: Visibility::Public,
                    is_static: false,
                    readonly: true,
                    init: None,
                    doc: Some(format!("{}（{}）", control.control_type, control.name)),
                    span: Span::default(),
                });
            }
            self.prog.classes[id].fields = fields;
            // Public functions of the code-behind are members of the form.
            let exported: Vec<FnId> = self.prog.modules[m]
                .fns
                .iter()
                .copied()
                .filter(|&f| {
                    self.prog.fns[f].exported && matches!(self.prog.fns[f].kind, FnKind::Free)
                })
                .collect();
            self.prog.classes[id].methods = exported;
            // `function constructor(...)`: the arguments of `new frmInput(...)` (spec §14.11.9).
            let ctor = self.prog.modules[m].fns.iter().copied().find(|&f| {
                let d = &self.prog.fns[f];
                d.name == "constructor" && matches!(d.kind, FnKind::Free) && !d.is_extension
            });
            if let Some(f) = ctor {
                let d = self.prog.fns[f].clone();
                let span = d.name_span;
                if d.exported {
                    self.err(m, span, "E0633", "フォームの `constructor` は `export` せずに書いてください（`new フォーム名(...)` で呼ばれます）");
                }
                if !d.type_params.is_empty() {
                    self.err(
                        m,
                        span,
                        "E0633",
                        "フォームの `constructor` に型パラメーターは付けられません",
                    );
                }
                if d.decl.as_ref().is_some_and(|x| x.ret.is_some()) {
                    self.err(
                        m,
                        span,
                        "E0633",
                        "フォームの `constructor` に戻り値の型は書けません",
                    );
                }
            }
            if let Some(f) = self.prog.modules[m].form.as_mut() {
                f.class = Some(id);
                f.ctor = ctor;
            }
            // The default instance, visible everywhere by the form's name.
            let var = self.prog.vars.len();
            self.prog.vars.push(VarDef {
                id: var,
                module: m,
                name: decl.name.clone(),
                host_name: None,
                ty: Type::Class(id, Vec::new()),
                is_const: true,
                exported: false,
                declared: true,
                init: None,
                const_value: None,
                doc: Some(format!(
                    "ユーザーフォーム `{}` の既定のインスタンス。",
                    decl.name
                )),
                name_span: Span::default(),
                used_outside_main: false,
                reassigned: false,
            });
            if self.prog.globals.contains_key(&decl.name) {
                let msg = format!(
                    "フォーム名 `{}` は既に宣言されている名前と重複します",
                    decl.name
                );
                self.err(m, Span::default(), "E0101", msg);
            } else {
                self.prog.globals.insert(decl.name.clone(), Sym::Var(var));
            }
        }
    }

    /// The type of `this` in the file of a sheet or of `ThisWorkbook` (spec §12.2): a host class
    /// extending `Excel.Worksheet` (`Excel.Workbook`) whose fields are the ActiveX controls on the
    /// sheet (spec §12.2.1) and whose methods are the file's `export function`s. The sheet's code
    /// name (`thisWorkbook`) is a global of that type, so other modules can write
    /// `Sheet1.clearInput()`.
    fn add_document(
        &mut self,
        m: ModuleId,
        form: &FormModule,
        excel: &HashMap<String, Sym>,
        msforms: &HashMap<String, Sym>,
    ) {
        let Some(decl) = form.decl.clone() else {
            return;
        };
        let workbook = form.kind == ObjectKind::Workbook;
        if !workbook {
            let known = &self.prog.config.sheet_code_names;
            if !known.is_empty()
                && !known
                    .iter()
                    .any(|n| n.to_lowercase() == decl.name.to_lowercase())
            {
                let msg = format!(
                    "`{}`: ブックにコード名 `{}` のワークシートがありません。コード名は、シートの見出しの名前ではなく、VBE のプロジェクトに表示される名前です（ブックのシート: {}）。シートのコード名を変えたときは、ファイル名も変えてください",
                    self.prog.modules[m].rel_path,
                    decl.name,
                    known.join(", ")
                );
                self.err(m, Span::default(), "E0610", msg);
            }
        }
        let base = match excel.get(if workbook { "Workbook" } else { "Worksheet" }) {
            Some(Sym::Class(c)) => Some(*c),
            _ => None,
        };
        let doc = if workbook {
            "ブックのモジュール（objects/ThisWorkbook.ras）。".to_string()
        } else {
            format!(
                "シート `{}` のモジュール（objects/{}.ras）。",
                decl.name, decl.name
            )
        };
        let ident = ast::Ident {
            text: decl.name.clone(),
            span: Span::default(),
        };
        let id = self.alloc_class(
            m,
            &ident,
            ClassKind::Host,
            true,
            Some(doc.clone()),
            &Target::Global,
        );
        self.prog.classes[id].extends = base
            .map(|c| Type::Class(c, Vec::new()))
            .into_iter()
            .collect();
        let mut fields: Vec<FieldDef> = Vec::new();
        for control in &decl.controls {
            let ty = match msforms
                .iter()
                .find(|(k, _)| k.eq_ignore_ascii_case(&control.control_type))
            {
                Some((_, Sym::Class(c))) => Type::Class(*c, Vec::new()),
                _ => Type::Object,
            };
            if fields
                .iter()
                .any(|f| f.name.eq_ignore_ascii_case(&control.name))
            {
                continue;
            }
            fields.push(FieldDef {
                name: control.name.clone(),
                host_name: None,
                ty,
                vis: Visibility::Public,
                is_static: false,
                readonly: true,
                init: None,
                doc: Some(format!(
                    "シートの上のコントロール {}（{}）",
                    control.name,
                    if control.control_type.is_empty() {
                        "ActiveX"
                    } else {
                        control.control_type.as_str()
                    }
                )),
                span: Span::default(),
            });
        }
        self.prog.classes[id].fields = fields;
        let exported: Vec<FnId> = self.prog.modules[m]
            .fns
            .iter()
            .copied()
            .filter(|&f| self.prog.fns[f].exported && matches!(self.prog.fns[f].kind, FnKind::Free))
            .collect();
        self.prog.classes[id].methods = exported;
        if let Some(f) = self.prog.modules[m].form.as_mut() {
            f.class = Some(id);
        }
        // The global the sheet (workbook) is referred to by gets the type with these members.
        let global = if workbook {
            "thisWorkbook".to_string()
        } else {
            decl.name.clone()
        };
        match self.prog.globals.get(&global).cloned() {
            Some(Sym::Var(v)) => self.prog.vars[v].ty = Type::Class(id, Vec::new()),
            Some(_) => {
                let msg = format!(
                    "シートのコード名 `{}` は既に宣言されている名前と重複します",
                    decl.name
                );
                self.err(m, Span::default(), "E0101", msg);
            }
            None => {
                let var = self.prog.vars.len();
                self.prog.vars.push(VarDef {
                    id: var,
                    module: m,
                    name: global.clone(),
                    host_name: if workbook {
                        Some("ThisWorkbook".into())
                    } else {
                        None
                    },
                    ty: Type::Class(id, Vec::new()),
                    is_const: true,
                    exported: false,
                    declared: true,
                    init: None,
                    const_value: None,
                    doc: Some(doc),
                    name_span: Span::default(),
                    used_outside_main: false,
                    reassigned: false,
                });
                self.prog.globals.insert(global, Sym::Var(var));
            }
        }
    }

    /// Functions and variables of a sheet's (the workbook's) file named like a member of the
    /// sheet (workbook) or a control on it: `Sheet1.name` would be either (spec §12.2, `E0651`).
    fn check_document_names(&mut self) {
        for m in 0..self.prog.modules.len() {
            let Some(doc) = self.prog.modules[m].document().cloned() else {
                continue;
            };
            let Some(class) = doc.class else { continue };
            // The global keeps the type with the file's members, whatever its declaration says
            // (`declare const ThisWorkbook: Excel.Workbook`).
            let global = if doc.kind == ObjectKind::Workbook {
                "thisWorkbook".to_string()
            } else {
                self.prog.modules[m].vba_name.clone()
            };
            if let Some(Sym::Var(v)) = self.prog.globals.get(&global).cloned() {
                self.prog.vars[v].ty = Type::Class(class, Vec::new());
            }
            // Members of the sheet: those of `Excel.Worksheet` and what it extends, and the controls.
            let mut members: HashMap<String, String> = HashMap::new();
            let mut pending: Vec<ClassId> = self.prog.classes[class]
                .extends
                .iter()
                .filter_map(|t| {
                    if let Type::Class(c, _) = t {
                        Some(*c)
                    } else {
                        None
                    }
                })
                .collect();
            let mut seen: Vec<ClassId> = Vec::new();
            while let Some(c) = pending.pop() {
                if seen.contains(&c) {
                    continue;
                }
                seen.push(c);
                let def = &self.prog.classes[c];
                for f in &def.fields {
                    members
                        .entry(f.name.to_lowercase())
                        .or_insert_with(|| f.name.clone());
                }
                for &f in &def.methods {
                    let name = &self.prog.fns[f].name;
                    members
                        .entry(name.to_lowercase())
                        .or_insert_with(|| name.clone());
                }
                pending.extend(def.extends.iter().filter_map(|t| {
                    if let Type::Class(c, _) = t {
                        Some(*c)
                    } else {
                        None
                    }
                }));
            }
            for f in &self.prog.classes[class].fields {
                members
                    .entry(f.name.to_lowercase())
                    .or_insert_with(|| f.name.clone());
            }
            let what = if doc.kind == ObjectKind::Workbook {
                "ブック"
            } else {
                "シート"
            };
            let mut bad: Vec<(Span, String, String)> = Vec::new();
            for &f in &self.prog.modules[m].fns {
                let d = &self.prog.fns[f];
                if matches!(d.kind, FnKind::Free) {
                    if let Some(member) = members.get(&d.name.to_lowercase()) {
                        bad.push((d.name_span, d.name.clone(), member.clone()));
                    }
                }
            }
            for &v in &self.prog.modules[m].vars {
                let d = &self.prog.vars[v];
                if d.module == m && d.name_span != Span::default() {
                    if let Some(member) = members.get(&d.name.to_lowercase()) {
                        bad.push((d.name_span, d.name.clone(), member.clone()));
                    }
                }
            }
            for (span, name, member) in bad {
                self.err(m, span, "E0651", format!("`{name}` は{what}のメンバー `{member}` と同じ名前です。別の名前にしてください（`{}.{member}` や `this.{member}` がどちらを指すか決まらなくなります）", doc.decl.as_ref().map(|d| d.name.as_str()).unwrap_or("this")));
            }
        }
    }

    fn declare_module_items(&mut self, m: ModuleId) {
        let ast = self.prog.modules[m].ast.clone();
        let is_typedef = self.prog.modules[m].kind == ModuleKind::TypeDef;
        let target = if is_typedef {
            Target::Global
        } else {
            Target::Module(m)
        };
        let mut main_stmts: Vec<ast::Stmt> = Vec::new();
        let mut has_exec = false;
        for item in &ast.items {
            match item {
                Item::Stmt(s) => {
                    if !matches!(s.kind, ast::StmtKind::Empty | ast::StmtKind::Error) {
                        has_exec = true;
                    }
                    main_stmts.push(s.clone());
                }
                Item::Var(v) => {
                    self.declare_var(m, &target, v, is_typedef);
                    if !is_typedef {
                        main_stmts.push(ast::Stmt {
                            kind: ast::StmtKind::Var(v.clone()),
                            span: v.span,
                        });
                    }
                }
                other => self.declare_item(m, &target, other, is_typedef),
            }
        }
        let has_decls = ast
            .items
            .iter()
            .any(|i| matches!(i, Item::Fn(_) | Item::Class(_) | Item::Interface(_)));
        let only_vars = !has_decls
            && ast
                .items
                .iter()
                .any(|i| matches!(i, Item::Var(v) if v.init.is_some() && !v.declared))
            && !ast
                .items
                .iter()
                .any(|i| matches!(i, Item::Var(v) if v.exported));
        let is_form = self.prog.modules[m].form.is_some();
        if (has_exec || (only_vars && !is_form)) && !is_typedef && !is_form {
            // Script style: top-level statements become `Public Sub Main()`.
            let span = main_stmts.first().map(|s| s.span).unwrap_or_default();
            let decl = ast::FnDecl {
                decorators: Vec::new(),
                doc: None,
                exported: true,
                declared: false,
                inline: InlineMode::Default,
                name: ast::Ident {
                    text: "Main".into(),
                    span: Span::empty(span.start()),
                },
                type_params: Vec::new(),
                params: Vec::new(),
                ret: None,
                body: Some(ast::Block {
                    stmts: main_stmts,
                    span,
                }),
                span,
            };
            let name = decl.name.clone();
            let id = self.alloc_fn(m, Rc::new(decl), FnKind::Main, None, Visibility::Public);
            self.prog.fns[id].ret = Some(Type::Void);
            self.prog.modules[m].main_fn = Some(id);
            if self.prog.modules[m].scope.contains_key("Main") {
                self.err(m, span, "E0103", "トップレベルの文は `Main` という Sub にまとめられますが、`Main` は既に宣言されています");
            } else {
                self.register(m, &target, &name, Sym::Fn(id), true);
            }
        } else if !is_typedef
            && main_stmts.iter().any(
                |s| matches!(&s.kind, ast::StmtKind::Var(v) if v.init.is_some() && !v.declared),
            )
        {
            // Top-level variables with initializers: run them once from the module initializer.
            let span = main_stmts.first().map(|s| s.span).unwrap_or_default();
            let decl = ast::FnDecl {
                decorators: Vec::new(),
                doc: None,
                exported: false,
                declared: false,
                inline: InlineMode::NoInline,
                name: ast::Ident {
                    text: "RAS_ModInit".into(),
                    span: Span::empty(span.start()),
                },
                type_params: Vec::new(),
                params: Vec::new(),
                ret: None,
                body: Some(ast::Block {
                    stmts: main_stmts,
                    span,
                }),
                span,
            };
            let id = self.alloc_fn(m, Rc::new(decl), FnKind::ModInit, None, Visibility::Public);
            self.prog.fns[id].ret = Some(Type::Void);
            self.prog.modules[m].init_fn = Some(id);
        }
    }

    /// The bundled type definitions are written with the type libraries' spelling (`Cells`).
    fn is_bundled_typedef(&self, m: ModuleId) -> bool {
        let md = &self.prog.modules[m];
        md.kind == ModuleKind::TypeDef && md.rel_path.starts_with("<stdlib>/")
    }

    /// RAS name of a host member / global declared as `name` in module `m`, and the
    /// type-library spelling to emit when the two differ (spec §15.3.1).
    fn host_names(&self, m: ModuleId, name: &str, global: bool) -> (String, Option<String>) {
        if !self.is_bundled_typedef(m) {
            return (name.to_string(), None);
        }
        let js = crate::names::js_name(name);
        // A global cannot be spelled like a keyword (`Switch`); a member can (`range.delete()`).
        if js == name || (global && crate::names::is_ras_keyword(&js)) {
            return (name.to_string(), None);
        }
        (js, Some(name.to_string()))
    }

    /// `name` as it is registered in a scope: with the RAS spelling.
    fn host_ident(&mut self, m: ModuleId, target: &Target, name: &ast::Ident) -> ast::Ident {
        let (js, _) = self.host_names(m, &name.text, true);
        if matches!(target, Target::Global)
            && self.is_bundled_typedef(m)
            && self.prog.modules[m].key != TEST_DEFS_MODULE
        {
            self.prog
                .host_global_names
                .insert(name.text.to_lowercase(), js.clone());
        }
        ast::Ident {
            text: js,
            span: name.span,
        }
    }

    fn alloc_fn(
        &mut self,
        m: ModuleId,
        decl: Rc<ast::FnDecl>,
        kind: FnKind,
        class: Option<ClassId>,
        vis: Visibility,
    ) -> FnId {
        let id = self.prog.fns.len();
        let (name, host_name) = match kind {
            FnKind::Host { class } => self.host_names(m, &decl.name.text, class.is_none()),
            _ => (decl.name.text.clone(), None),
        };
        self.prog.fns.push(FnDef {
            id,
            module: m,
            name,
            host_name,
            kind,
            exported: decl.exported,
            vis,
            inline: decl.inline,
            type_params: Vec::new(),
            params: Vec::new(),
            is_extension: decl.params.first().is_some_and(|p| p.is_this),
            ret: None,
            attrs: FnAttrs::default(),
            doc: decl.doc.clone(),
            name_span: decl.name.span,
            decl: Some(decl),
            locals: Vec::new(),
            calls: Vec::new(),
            var_refs: Vec::new(),
            checked: false,
            has_loops: false,
            inline_blockers: Vec::new(),
            cost: 0,
        });
        self.prog.modules[m].fns.push(id);
        self.fn_decls.push((id, class));
        if let Some(ns) = &self.cur_ns {
            self.fn_ns.insert(id, ns.clone());
        }
        id
    }

    fn declare_var(&mut self, m: ModuleId, target: &Target, v: &ast::VarDecl, is_typedef: bool) {
        let id = self.prog.vars.len();
        let (name, host_name) = if v.declared || is_typedef {
            self.host_names(m, &v.name.text, true)
        } else {
            (v.name.text.clone(), None)
        };
        self.prog.vars.push(VarDef {
            id,
            module: m,
            name,
            host_name,
            ty: Type::Error,
            is_const: v.kind == ast::VarKind::Const,
            exported: v.exported,
            declared: v.declared || is_typedef,
            init: v.init.clone().map(Rc::new),
            const_value: None,
            doc: v.doc.clone(),
            name_span: v.name.span,
            used_outside_main: false,
            reassigned: false,
        });
        self.prog.modules[m].vars.push(id);
        self.var_decls.push((id, Rc::new(v.clone())));
        if let Some(ns) = &self.cur_ns {
            self.var_ns.insert(id, ns.clone());
        }
        if v.kind == ast::VarKind::Using {
            self.err(m, v.span, "E0104", "`using` は関数の中でのみ使えます");
        }
        let reg_name = self.host_ident(m, target, &v.name);
        self.register(m, target, &reg_name, Sym::Var(id), v.exported);
    }

    fn declare_item(&mut self, m: ModuleId, target: &Target, item: &Item, is_typedef: bool) {
        match item {
            Item::Fn(f) => {
                let kind = if is_typedef || f.declared {
                    FnKind::Host { class: None }
                } else {
                    FnKind::Free
                };
                if f.body.is_none() && !(is_typedef || f.declared) {
                    self.err(m, f.name.span, "E0105", "関数の本体がありません");
                }
                let id = self.alloc_fn(m, Rc::new(f.clone()), kind, None, Visibility::Public);
                let reg_name = self.host_ident(m, target, &f.name);
                self.register(m, target, &reg_name, Sym::Fn(id), f.exported);
            }
            Item::Class(c) => {
                let id = self.alloc_class(
                    m,
                    &c.name,
                    if is_typedef || c.declared {
                        ClassKind::Host
                    } else {
                        ClassKind::Class
                    },
                    c.exported,
                    c.doc.clone(),
                    target,
                );
                self.prog.classes[id].predeclared =
                    c.decorators.iter().any(|d| d.name.text == "predeclared");
                for d in &c.decorators {
                    if d.name.text != "predeclared" {
                        self.warn(
                            m,
                            d.span,
                            "W0110",
                            format!("不明なクラスデコレータ `@{}` は無視されます", d.name.text),
                        );
                    }
                }
                for member in &c.members {
                    match member {
                        ast::ClassMember::Method(md) => {
                            let kind = FnKind::Method {
                                class: id,
                                is_static: md.is_static,
                            };
                            let f =
                                self.alloc_fn(m, Rc::new(md.func.clone()), kind, Some(id), md.vis);
                            self.prog.classes[id].methods.push(f);
                        }
                        ast::ClassMember::Ctor(f) => {
                            let fid = self.alloc_fn(
                                m,
                                Rc::new(f.clone()),
                                FnKind::Ctor { class: id },
                                Some(id),
                                Visibility::Public,
                            );
                            if self.prog.classes[id].ctor.is_some() {
                                self.err(
                                    m,
                                    f.name.span,
                                    "E0106",
                                    "コンストラクタは 1 つだけ宣言できます",
                                );
                            }
                            self.prog.classes[id].ctor = Some(fid);
                        }
                        ast::ClassMember::Accessor(a) => {
                            let kind = match a.kind {
                                ast::AccessorKind::Get => FnKind::Getter {
                                    class: id,
                                    is_static: a.is_static,
                                },
                                ast::AccessorKind::Set => FnKind::Setter {
                                    class: id,
                                    is_static: a.is_static,
                                },
                            };
                            let f =
                                self.alloc_fn(m, Rc::new(a.func.clone()), kind, Some(id), a.vis);
                            self.prog.classes[id].methods.push(f);
                        }
                        ast::ClassMember::Field(_) => {}
                    }
                }
                self.class_decls
                    .push((id, ClassAst::Class(Rc::new(c.clone()))));
                self.register(m, target, &c.name, Sym::Class(id), c.exported);
            }
            Item::Interface(i) => {
                let host = is_typedef || i.declared;
                let id = self.alloc_class(
                    m,
                    &i.name,
                    if host {
                        ClassKind::Host
                    } else {
                        ClassKind::Interface
                    },
                    i.exported,
                    i.doc.clone(),
                    target,
                );
                for member in &i.members {
                    match member {
                        ast::InterfaceMember::Method(f) => {
                            let kind = if host {
                                FnKind::Host { class: Some(id) }
                            } else {
                                FnKind::Abstract { class: id }
                            };
                            let fid = self.alloc_fn(
                                m,
                                Rc::new(f.clone()),
                                kind,
                                Some(id),
                                Visibility::Public,
                            );
                            if f.name.text.is_empty() {
                                self.prog.classes[id].call_sig = Some(fid);
                            } else {
                                self.prog.classes[id].methods.push(fid);
                            }
                        }
                        ast::InterfaceMember::Index { param, ty, span } => {
                            // `[index: K]: V` is sugar for the call signature `(index: K): V`.
                            let decl = ast::FnDecl {
                                decorators: Vec::new(),
                                doc: None,
                                exported: false,
                                declared: true,
                                inline: InlineMode::Default,
                                name: ast::Ident {
                                    text: String::new(),
                                    span: *span,
                                },
                                type_params: Vec::new(),
                                params: vec![param.clone()],
                                ret: Some(ty.clone()),
                                body: None,
                                span: *span,
                            };
                            let kind = if host {
                                FnKind::Host { class: Some(id) }
                            } else {
                                FnKind::Abstract { class: id }
                            };
                            let fid =
                                self.alloc_fn(m, Rc::new(decl), kind, Some(id), Visibility::Public);
                            self.prog.classes[id].call_sig = Some(fid);
                        }
                        ast::InterfaceMember::Prop(_) => {}
                    }
                }
                self.class_decls
                    .push((id, ClassAst::Interface(Rc::new(i.clone()))));
                self.register(m, target, &i.name, Sym::Class(id), i.exported);
            }
            Item::Enum(e) => {
                let id = self.prog.enums.len();
                let host = is_typedef || e.declared;
                let mut members = Vec::new();
                let mut next: i64 = 0;
                for mem in &e.members {
                    let value = match &mem.value {
                        Some(expr) => match const_int(expr).and_then(|v| i64::try_from(v).ok()) {
                            Some(v) if i32::try_from(v).is_ok() => v,
                            Some(_) => {
                                self.err(
                                    m,
                                    expr.span,
                                    "E0107",
                                    "列挙値は int（32bit）の範囲に収まる必要があります",
                                );
                                0
                            }
                            None => {
                                self.err(
                                    m,
                                    expr.span,
                                    "E0107",
                                    "列挙値は整数の定数式である必要があります",
                                );
                                next
                            }
                        },
                        None => next,
                    };
                    next = value.wrapping_add(1);
                    members.push(EnumMemberDef {
                        name: mem.name.text.clone(),
                        value,
                        doc: mem.doc.clone(),
                        span: mem.name.span,
                    });
                }
                self.prog.enums.push(EnumDef {
                    id,
                    module: m,
                    name: e.name.text.clone(),
                    host,
                    late_bound: host && matches!(target, Target::Namespace(ns) if !crate::model::referenced_namespace(Some(ns.as_str()))),
                    exported: e.exported,
                    members,
                    doc: e.doc.clone(),
                    name_span: e.name.span,
                });
                self.prog.modules[m].enums.push(id);
                self.register(m, target, &e.name, Sym::Enum(id), e.exported);
                if host {
                    // Host enum constants are visible by bare name, as in VBA (`xlUp`).
                    for (i, mem) in e.members.iter().enumerate() {
                        if self
                            .prog
                            .globals
                            .insert(mem.name.text.clone(), Sym::EnumMember(id, i))
                            .is_some()
                        {
                            self.err(
                                m,
                                mem.name.span,
                                "E0101",
                                format!("`{}` は既に宣言されています", mem.name.text),
                            );
                        }
                    }
                }
            }
            Item::Struct(s) => {
                let id = self.prog.structs.len();
                self.prog.structs.push(StructDef {
                    id,
                    module: m,
                    name: s.name.text.clone(),
                    kind: StructKind::Named,
                    fields: Vec::new(),
                    optional: s.fields.iter().map(|f| f.optional).collect(),
                    doc: s.doc.clone(),
                    name_span: s.name.span,
                });
                self.prog.modules[m].structs.push(id);
                self.struct_decls.push((id, Rc::new(s.clone())));
                self.register(m, target, &s.name, Sym::Struct(id), s.exported);
            }
            Item::TypeAlias(a) => {
                if !a.type_params.is_empty() {
                    self.err(
                        m,
                        a.name.span,
                        "E0108",
                        "ジェネリックな型エイリアスには対応していません",
                    );
                }
                self.register(m, target, &a.name, Sym::Alias(Type::Error), a.exported);
                self.alias_decls.push((m, Rc::new(a.clone())));
            }
            Item::DeclareModule(d) => {
                for f in &d.functions {
                    let alias = f
                        .decorators
                        .iter()
                        .find(|x| x.name.text == "alias")
                        .and_then(|x| match x.args.first().map(|a| &a.kind) {
                            Some(ExprKind::Str(s)) => Some(s.clone()),
                            _ => None,
                        });
                    let id = self.alloc_fn(
                        m,
                        Rc::new(f.clone()),
                        FnKind::Declare {
                            lib: d.lib.clone(),
                            alias,
                        },
                        None,
                        Visibility::Public,
                    );
                    self.register(m, target, &f.name, Sym::Fn(id), f.exported);
                }
            }
            Item::Namespace(ns) => {
                if !is_typedef {
                    self.err(
                        m,
                        ns.name.span,
                        "E0109",
                        "`namespace` は型定義ファイル（.d.ras）でのみ使えます",
                    );
                    return;
                }
                let inner = Target::Namespace(ns.name.text.clone());
                self.prog
                    .namespaces
                    .entry(ns.name.text.clone())
                    .or_default();
                let saved = self.cur_ns.replace(ns.name.text.clone());
                for it in &ns.items {
                    match it {
                        Item::Var(v) => self.declare_var(m, &inner, v, true),
                        Item::Stmt(_) | Item::Import(_) => {}
                        other => self.declare_item(m, &inner, other, true),
                    }
                }
                self.cur_ns = saved;
            }
            Item::Import(_) | Item::Var(_) | Item::Stmt(_) => {}
        }
    }

    fn alloc_class(
        &mut self,
        m: ModuleId,
        name: &ast::Ident,
        kind: ClassKind,
        exported: bool,
        doc: Option<String>,
        target: &Target,
    ) -> ClassId {
        let id = self.prog.classes.len();
        self.prog.classes.push(ClassDef {
            id,
            module: m,
            name: name.text.clone(),
            kind,
            namespace: match target {
                Target::Namespace(ns) => Some(ns.clone()),
                _ => None,
            },
            exported,
            type_params: Vec::new(),
            extends: Vec::new(),
            implements: Vec::new(),
            fields: Vec::new(),
            methods: Vec::new(),
            ctor: None,
            call_sig: None,
            predeclared: false,
            doc,
            name_span: name.span,
        });
        self.prog.modules[m].classes.push(id);
        id
    }

    fn add_sheet_code_names(&mut self) {
        let Some(Sym::Class(ws)) = self
            .prog
            .namespaces
            .get("Excel")
            .and_then(|ns| ns.get("Worksheet"))
            .cloned()
        else {
            return;
        };
        let module = self.prog.classes[ws].module;
        for name in self.prog.config.sheet_code_names.clone() {
            if self.prog.globals.contains_key(&name) {
                continue;
            }
            let id = self.prog.vars.len();
            self.prog.vars.push(VarDef {
                id,
                module,
                name: name.clone(),
                host_name: None,
                ty: Type::Class(ws, Vec::new()),
                is_const: true,
                exported: false,
                declared: true,
                init: None,
                const_value: None,
                doc: Some("ブックのシート（コード名）。".into()),
                name_span: Span::default(),
                used_outside_main: false,
                reassigned: false,
            });
            self.prog.globals.insert(name, Sym::Var(id));
        }
    }

    fn resolve_imports(&mut self, m: ModuleId) {
        if self.prog.modules[m].kind == ModuleKind::TypeDef {
            return;
        }
        let ast = self.prog.modules[m].ast.clone();
        // Extension methods declared in the module itself are always in scope.
        let own: Vec<FnId> = self.prog.modules[m]
            .fns
            .iter()
            .copied()
            .filter(|&f| self.prog.fns[f].is_extension)
            .collect();
        self.prog.modules[m].ext_methods.extend(own);
        for item in &ast.items {
            let Item::Import(imp) = item else { continue };
            let target = if let Some(rest) = imp.source.strip_prefix("vba:") {
                self.err(
                    m,
                    imp.source_span,
                    "E0120",
                    format!("`vba:{rest}`（非管理 VBA モジュールの参照）にはまだ対応していません"),
                );
                continue;
            } else {
                import_target(&self.prog, m, &imp.source)
            };
            let Some(t) = target else {
                if imp.source.starts_with('.')
                    && import_by_path(&self.prog, m, &imp.source)
                        .is_some_and(|t| self.prog.modules[t].is_test())
                {
                    self.err(m, imp.source_span, "E0701", format!("`{}` はテスト側のファイルです。`src/` / `forms/` のモジュールからテストのファイルは `import` できません", imp.source));
                } else {
                    self.err(
                        m,
                        imp.source_span,
                        "E0121",
                        format!("モジュール `{}` が見つかりません", imp.source),
                    );
                }
                continue;
            };
            if t == m {
                self.err(
                    m,
                    imp.source_span,
                    "E0122",
                    "自分自身はインポートできません",
                );
                continue;
            }
            if !self.prog.modules[m].imports.contains(&t) {
                self.prog.modules[m].imports.push(t);
            }
            if let Some(ns) = &imp.namespace {
                let exts: Vec<FnId> = self.prog.modules[t]
                    .exports
                    .values()
                    .filter_map(|s| match s {
                        Sym::Fn(f) if self.prog.fns[*f].is_extension => Some(*f),
                        _ => None,
                    })
                    .collect();
                for f in exts {
                    if !self.prog.modules[m].ext_methods.contains(&f) {
                        self.prog.modules[m].ext_methods.push(f);
                    }
                }
                self.register(m, &Target::Module(m), ns, Sym::Module(t), false);
            }
            for n in &imp.names {
                match self.prog.modules[t].exports.get(&n.name.text).cloned() {
                    Some(sym) => {
                        if let Sym::Fn(f) = &sym {
                            if self.prog.fns[*f].is_extension
                                && !self.prog.modules[m].ext_methods.contains(f)
                            {
                                self.prog.modules[m].ext_methods.push(*f);
                            }
                        }
                        let local = n.alias.as_ref().unwrap_or(&n.name);
                        self.register(m, &Target::Module(m), local, sym, false);
                    }
                    None => {
                        let hint = if self.prog.modules[t].scope.contains_key(&n.name.text) {
                            "（`export` されていません）"
                        } else {
                            ""
                        };
                        self.err(
                            m,
                            n.name.span,
                            "E0123",
                            format!(
                                "モジュール `{}` に `{}` はありません{hint}",
                                imp.source, n.name.text
                            ),
                        );
                    }
                }
            }
        }
    }

    // ------------------------------------------------------------------ signatures

    fn resolve_signatures(&mut self) {
        // The number of type parameters of every class is known before any type is resolved,
        // so that a declaration can refer to a generic type declared after it.
        for (id, ast) in &self.class_decls {
            let tps = match ast {
                ClassAst::Class(c) => &c.type_params,
                ClassAst::Interface(i) => &i.type_params,
            };
            self.prog.classes[*id].type_params = tps
                .iter()
                .map(|tp| TypeParamDef {
                    name: tp.name.text.clone(),
                    constraint: Constraint::None,
                    default: None,
                })
                .collect();
        }
        // Then the defaults of their type parameters (`interface Sheets<T = Worksheet | Chart>`),
        // so that a type written without arguments gets them wherever it is declared.
        let decls: Vec<(ClassId, Vec<ast::TypeParam>)> = self
            .class_decls
            .iter()
            .map(|(id, ast)| {
                (
                    *id,
                    match ast {
                        ClassAst::Class(c) => c.type_params.clone(),
                        ClassAst::Interface(i) => i.type_params.clone(),
                    },
                )
            })
            .filter(|(_, tps)| tps.iter().any(|tp| tp.default.is_some()))
            .collect();
        for (id, tps) in decls {
            let m = self.prog.classes[id].module;
            self.prog.cur_namespace = self.prog.classes[id].namespace.clone();
            let mut seen_default = false;
            for (i, tp) in tps.iter().enumerate() {
                match &tp.default {
                    Some(t) => {
                        seen_default = true;
                        let scope = self.prog.classes[id].type_params[..i].to_vec();
                        let ty = resolve_type(&mut self.prog, m, &scope, t);
                        self.prog.classes[id].type_params[i].default = Some(ty);
                    }
                    None if seen_default => {
                        self.err(m, tp.name.span, "E0171", format!("既定の型引数を持つ型パラメータの後の `{}` にも既定の型引数が必要です", tp.name.text));
                    }
                    None => {}
                }
            }
            self.prog.cur_namespace = None;
        }
        // Aliases first (in declaration order), then everything else.
        for (m, a) in std::mem::take(&mut self.alias_decls) {
            let ty = resolve_type(&mut self.prog, m, &[], &a.ty);
            let md = &mut self.prog.modules[m];
            md.scope.insert(a.name.text.clone(), Sym::Alias(ty.clone()));
            if a.exported {
                md.exports.insert(a.name.text.clone(), Sym::Alias(ty));
            }
        }
        // Re-bind imported aliases (they were copied before being resolved).
        for m in 0..self.prog.modules.len() {
            let ast = self.prog.modules[m].ast.clone();
            for item in &ast.items {
                let Item::Import(imp) = item else { continue };
                for n in &imp.names {
                    let local = n.alias.as_ref().unwrap_or(&n.name).text.clone();
                    if let Some(Sym::Alias(_)) = self.prog.modules[m].scope.get(&local) {
                        let resolved =
                            self.prog.modules[m].imports.iter().find_map(|&t| {
                                match self.prog.modules[t].exports.get(&n.name.text) {
                                    Some(Sym::Alias(ty)) => Some(ty.clone()),
                                    _ => None,
                                }
                            });
                        if let Some(ty) = resolved {
                            self.prog.modules[m].scope.insert(local, Sym::Alias(ty));
                        }
                    }
                }
            }
        }
        for (id, ast) in std::mem::take(&mut self.class_decls) {
            self.resolve_class(id, &ast);
        }
        for (id, decl) in std::mem::take(&mut self.struct_decls) {
            let m = self.prog.structs[id].module;
            let mut fields = Vec::new();
            for f in &decl.fields {
                let ty = match &f.ty {
                    Some(t) => resolve_type(&mut self.prog, m, &[], t),
                    None => {
                        self.err(m, f.span, "E0130", "構造体のフィールドには型注釈が必要です");
                        Type::Error
                    }
                };
                fields.push((f.name.text.clone(), ty));
            }
            self.prog.structs[id].fields = fields;
        }
        for (id, class) in std::mem::take(&mut self.fn_decls) {
            self.resolve_fn(id, class);
        }
        self.prog.cur_namespace = None;
        for (id, decl) in std::mem::take(&mut self.var_decls) {
            let m = self.prog.vars[id].module;
            self.prog.cur_namespace = self.var_ns.get(&id).cloned();
            if let Some(t) = &decl.ty {
                self.prog.vars[id].ty = resolve_type(&mut self.prog, m, &[], t);
            } else if self.prog.vars[id].declared {
                self.err(
                    m,
                    decl.name.span,
                    "E0131",
                    "`declare` された変数には型注釈が必要です",
                );
            } else if decl.init.is_none() {
                self.err(m, decl.name.span, "E0132", "型注釈か初期化式が必要です");
            } else {
                // Inferred from the initializer by the checker.
                self.prog.vars[id].ty = Type::Void;
            }
            if decl.kind == ast::VarKind::Const
                && decl.init.is_none()
                && !self.prog.vars[id].declared
            {
                self.err(m, decl.name.span, "E0133", "`const` には初期化式が必要です");
            }
        }
    }

    /// The type parameters of a class or interface, keeping the defaults resolved beforehand.
    fn class_type_params(
        &mut self,
        id: ClassId,
        m: ModuleId,
        tps: &[ast::TypeParam],
    ) -> Vec<TypeParamDef> {
        let mut out = self.type_params(m, &[], tps);
        for (i, p) in out.iter_mut().enumerate() {
            p.default = self.prog.classes[id]
                .type_params
                .get(i)
                .and_then(|d| d.default.clone());
        }
        self.prog.classes[id].type_params = out.clone();
        out
    }

    fn type_params(
        &mut self,
        m: ModuleId,
        outer: &[TypeParamDef],
        tps: &[ast::TypeParam],
    ) -> Vec<TypeParamDef> {
        let mut out: Vec<TypeParamDef> = Vec::new();
        for tp in tps {
            let constraint = match &tp.constraint {
                None => Constraint::None,
                Some(TypeRef::Named { path, args, .. }) if path.len() == 1 && args.is_empty() => {
                    match path[0].text.as_str() {
                        "Comparable" => Constraint::Comparable,
                        "Numeric" => Constraint::Numeric,
                        "ObjectType" => Constraint::ObjectType,
                        _ => {
                            let scope: Vec<TypeParamDef> =
                                outer.iter().chain(out.iter()).cloned().collect();
                            Constraint::Type(resolve_type(
                                &mut self.prog,
                                m,
                                &scope,
                                tp.constraint.as_ref().unwrap(),
                            ))
                        }
                    }
                }
                Some(t) => {
                    let scope: Vec<TypeParamDef> =
                        outer.iter().chain(out.iter()).cloned().collect();
                    Constraint::Type(resolve_type(&mut self.prog, m, &scope, t))
                }
            };
            if outer
                .iter()
                .chain(out.iter())
                .any(|p| p.name == tp.name.text)
            {
                self.err(
                    m,
                    tp.name.span,
                    "E0101",
                    format!("型パラメータ `{}` は既に宣言されています", tp.name.text),
                );
            }
            out.push(TypeParamDef {
                name: tp.name.text.clone(),
                constraint,
                default: None,
            });
        }
        out
    }

    fn resolve_class(&mut self, id: ClassId, ast: &ClassAst) {
        let m = self.prog.classes[id].module;
        self.prog.cur_namespace = self.prog.classes[id].namespace.clone();
        match ast {
            ClassAst::Class(c) => {
                let tps = self.class_type_params(id, m, &c.type_params);
                if let Some(ext) = &c.extends {
                    self.err(m, ext.span(), "E0140", "クラスの継承（`extends`）にはまだ対応していません。インターフェイス（`implements`）か委譲を使ってください");
                }
                let mut implements = Vec::new();
                for t in &c.implements {
                    let ty = resolve_type(&mut self.prog, m, &tps, t);
                    match &ty {
                        Type::Class(i, _) if self.prog.classes[*i].kind != ClassKind::Class => {
                            implements.push(ty)
                        }
                        Type::Error => {}
                        _ => self.err(
                            m,
                            t.span(),
                            "E0141",
                            "`implements` にはインターフェイスを指定してください",
                        ),
                    }
                }
                self.prog.classes[id].implements = implements;
                let mut fields = Vec::new();
                for member in &c.members {
                    let ast::ClassMember::Field(f) = member else {
                        continue;
                    };
                    let ty = match &f.ty {
                        Some(t) => resolve_type(&mut self.prog, m, &tps, t),
                        // Inferred from the initializer by the checker.
                        None if f.init.is_some() => Type::Void,
                        None => {
                            self.err(m, f.name.span, "E0132", "型注釈か初期化式が必要です");
                            Type::Error
                        }
                    };
                    if fields.iter().any(|x: &FieldDef| x.name == f.name.text) {
                        self.err(
                            m,
                            f.name.span,
                            "E0101",
                            format!("`{}` は既に宣言されています", f.name.text),
                        );
                    }
                    fields.push(FieldDef {
                        name: f.name.text.clone(),
                        host_name: None,
                        ty,
                        vis: f.vis,
                        is_static: f.is_static,
                        readonly: f.readonly,
                        init: f.init.clone().map(Rc::new),
                        doc: f.doc.clone(),
                        span: f.name.span,
                    });
                }
                self.prog.classes[id].fields = fields;
            }
            ClassAst::Interface(i) => {
                let tps = self.class_type_params(id, m, &i.type_params);
                let mut extends = Vec::new();
                for t in &i.extends {
                    let ty = resolve_type(&mut self.prog, m, &tps, t);
                    if matches!(ty, Type::Class(..)) {
                        extends.push(ty);
                    }
                }
                self.prog.classes[id].extends = extends;
                let mut fields = Vec::new();
                for member in &i.members {
                    let ast::InterfaceMember::Prop(p) = member else {
                        continue;
                    };
                    let ty = match &p.ty {
                        Some(t) => resolve_type(&mut self.prog, m, &tps, t),
                        None => Type::Any,
                    };
                    let (name, host_name) = self.host_names(m, &p.name.text, false);
                    fields.push(FieldDef {
                        name,
                        host_name,
                        ty,
                        vis: Visibility::Public,
                        is_static: false,
                        readonly: p.readonly,
                        init: None,
                        doc: p.doc.clone(),
                        span: p.name.span,
                    });
                }
                self.prog.classes[id].fields = fields;
            }
        }
    }

    fn resolve_fn(&mut self, id: FnId, class: Option<ClassId>) {
        let m = self.prog.fns[id].module;
        self.prog.cur_namespace = self.fn_ns.get(&id).cloned();
        let decl = self.prog.fns[id].decl.clone().unwrap();
        let outer: Vec<TypeParamDef> = class
            .map(|c| self.prog.classes[c].type_params.clone())
            .unwrap_or_default();
        let own = self.type_params(m, &outer, &decl.type_params);
        for tp in decl.type_params.iter().filter(|tp| tp.default.is_some()) {
            self.err(m, tp.name.span, "E0171", "既定の型引数はクラスとインターフェイスの型パラメータにだけ書けます（関数の型引数は引数の型から推論されます）");
        }
        let scope: Vec<TypeParamDef> = outer.iter().chain(own.iter()).cloned().collect();
        let kind = self.prog.fns[id].kind.clone();
        let is_host = matches!(kind, FnKind::Host { .. } | FnKind::Declare { .. });
        let mut params = Vec::new();
        let mut locals = Vec::new();
        for (i, p) in decl.params.iter().enumerate() {
            let ty = match &p.ty {
                Some(t) => resolve_type(&mut self.prog, m, &scope, t),
                None if is_host => Type::Any,
                None => {
                    self.err(
                        m,
                        p.name.span,
                        "E0150",
                        format!("引数 `{}` には型注釈が必要です", p.name.text),
                    );
                    Type::Error
                }
            };
            if p.is_this && i != 0 {
                self.err(m, p.name.span, "E0151", "`this` は第1引数にのみ書けます");
            }
            if p.is_this && class.is_some() {
                self.err(
                    m,
                    p.name.span,
                    "E0151",
                    "クラスのメソッドには `this` 引数を書けません",
                );
            }
            if p.rest {
                self.err(
                    m,
                    p.span,
                    "E0152",
                    "可変長引数（`...`）には対応していません",
                );
            }
            // VBA can only pass arrays by reference (spec §2.4 example).
            let byref = p.mode == PassMode::ByRef;
            locals.push(LocalDef {
                name: p.name.text.clone(),
                ty: ty.clone(),
                kind: if p.is_this {
                    LocalKind::This
                } else {
                    LocalKind::Param
                },
                byref,
                reassigned: false,
                captured: false,
                span: p.name.span,
                used: false,
            });
            params.push(ParamDef {
                name: p.name.text.clone(),
                ty,
                mode: p.mode,
                optional: p.optional,
                default: p.default.clone().map(Rc::new),
                rest: p.rest,
                span: p.name.span,
                local: i,
            });
        }
        let mut seen_optional = false;
        for (p, a) in params.iter().zip(decl.params.iter()) {
            let opt = p.optional || p.default.is_some();
            if seen_optional && !opt {
                self.err(
                    m,
                    a.name.span,
                    "E0153",
                    "省略可能な引数の後に必須の引数は置けません",
                );
            }
            seen_optional |= opt;
            if opt && p.mode == PassMode::ByRef && !is_host {
                self.err(
                    m,
                    a.name.span,
                    "E0154",
                    "`byref` 引数は省略可能にできません",
                );
            }
        }
        let ret = match (&decl.ret, &kind) {
            (Some(t), _) => Some(resolve_type(&mut self.prog, m, &scope, t)),
            (
                None,
                FnKind::Ctor { .. } | FnKind::Setter { .. } | FnKind::Main | FnKind::ModInit,
            ) => Some(Type::Void),
            (None, FnKind::Host { .. } | FnKind::Declare { .. } | FnKind::Abstract { .. }) => {
                Some(Type::Void)
            }
            // Inferred from the body.
            (None, _) => None,
        };
        let mut attrs = self.fn_attrs(m, &decl, &kind);
        if self.prog.modules[m].document().is_some()
            && matches!(kind, FnKind::Free | FnKind::ModInit)
        {
            attrs.doc_member = true;
            let bad = [
                ("init", attrs.init),
                ("udf", attrs.udf),
                ("ribbon", attrs.ribbon),
            ];
            for (name, set) in bad {
                if set {
                    self.err(m, decl.name.span, "E0650", format!("シート・`ThisWorkbook` のファイルの関数に `@{name}` は付けられません（標準モジュールに置かなければ働きません。src/ のモジュールに置いてください）"));
                }
            }
            if decl.params.first().is_some_and(|p| p.is_this) {
                self.err(m, decl.name.span, "E0650", "シート・`ThisWorkbook` のファイルでは拡張メソッドを宣言できません（src/ のモジュールに置いてください）");
            }
        } else if self.prog.modules[m].form.is_some()
            && matches!(kind, FnKind::Free | FnKind::ModInit)
        {
            attrs.form_member = true;
            let bad = [
                ("macro", attrs.macro_entry),
                ("init", attrs.init),
                ("udf", attrs.udf),
                ("ribbon", attrs.ribbon),
            ];
            for (name, set) in bad {
                if set {
                    self.err(m, decl.name.span, "E0632", format!("フォームのコードビハインドの関数に `@{name}` は付けられません（src/ のモジュールに置いてください）"));
                }
            }
            if decl.params.first().is_some_and(|p| p.is_this) {
                self.err(m, decl.name.span, "E0632", "フォームのコードビハインドでは拡張メソッドを宣言できません（src/ のモジュールに置いてください）");
            }
        }
        let f = &mut self.prog.fns[id];
        f.type_params = own;
        f.params = params;
        f.locals = locals;
        f.ret = ret;
        f.attrs = attrs;
        if matches!(kind, FnKind::Getter { .. }) && !decl.params.is_empty() {
            self.err(m, decl.name.span, "E0155", "getter は引数を取れません");
        }
        if matches!(kind, FnKind::Setter { .. }) && decl.params.len() != 1 {
            self.err(m, decl.name.span, "E0155", "setter の引数は 1 つです");
        }
    }

    fn fn_attrs(&mut self, m: ModuleId, decl: &ast::FnDecl, kind: &FnKind) -> FnAttrs {
        let mut a = FnAttrs::default();
        for d in &decl.decorators {
            let str_arg = |i: usize| match d.args.get(i).map(|e| &e.peel().kind) {
                Some(ExprKind::Str(s)) => Some(s.clone()),
                Some(ExprKind::Name(n)) => Some(n.text.clone()),
                _ => None,
            };
            match d.name.text.as_str() {
                "macro" => a.macro_entry = true,
                "init" => a.init = true,
                "udf" => a.udf = true,
                "fast" => a.fast = true,
                "silent" => a.silent = true,
                "reentrant" => a.reentrant = true,
                "nocancel" => a.nocancel = true,
                "nodoevents" => a.nodoevents = true,
                "doevents" => a.doevents = true,
                "ribbon" => a.ribbon = true,
                "alias" => {
                    if !matches!(kind, FnKind::Declare { .. }) {
                        self.err(
                            m,
                            d.span,
                            "E0160",
                            "`@alias` は `declare module` 内の関数にのみ使えます",
                        );
                    }
                }
                "onError" => match str_arg(0) {
                    Some(s) => a.on_error = Some(s),
                    None => self.err(
                        m,
                        d.span,
                        "E0160",
                        "`@onError(handler)` にはハンドラ関数名を指定してください",
                    ),
                },
                "fallback" => match str_arg(0) {
                    Some(s) => a.fallback = Some(s),
                    None => self.err(
                        m,
                        d.span,
                        "E0160",
                        "`@fallback(fn)` には代替関数名を指定してください",
                    ),
                },
                "on" => {
                    let Some(spec) = str_arg(0) else {
                        self.err(
                            m,
                            d.span,
                            "E0161",
                            "`@on(\"対象.イベント\")` の形式で指定してください",
                        );
                        continue;
                    };
                    let Some((target, event)) = spec.split_once('.') else {
                        self.err(m, d.span, "E0161", format!("`@on(\"{spec}\")`: `対象.イベント`（例: `Workbook.Open`, `Worksheet.Change`）の形式で指定してください"));
                        continue;
                    };
                    let mut reentrant = false;
                    if let Some(ExprKind::Object(props)) = d.args.get(1).map(|e| &e.kind) {
                        for p in props {
                            match (p.key.text.as_str(), &p.value.kind) {
                                ("reentrant", ExprKind::Bool(b)) => reentrant = *b,
                                _ => self.warn(
                                    m,
                                    p.key.span,
                                    "W0110",
                                    format!(
                                        "`@on` の不明なオプション `{}` は無視されます",
                                        p.key.text
                                    ),
                                ),
                            }
                        }
                    }
                    a.on_events.push(OnEvent {
                        target: target.to_string(),
                        event: event.to_string(),
                        reentrant,
                        span: d.span,
                    });
                }
                other => self.warn(
                    m,
                    d.span,
                    "W0110",
                    format!("不明なデコレータ `@{other}` は無視されます"),
                ),
            }
        }
        if (a.macro_entry || !a.on_events.is_empty() || a.init) && !decl.type_params.is_empty() {
            self.err(m, decl.name.span, "E0162", "ジェネリック関数はマクロ入口・`@on`・`@init` の対象にできません。具体型の関数で包んでください");
        }
        a
    }
}

/// Resolves a syntactic type to a semantic type, reporting errors.
pub fn resolve_type(
    prog: &mut Program,
    m: ModuleId,
    type_params: &[TypeParamDef],
    t: &TypeRef,
) -> Type {
    let mut errors: Vec<(Span, String)> = Vec::new();
    let ty = resolve_type_inner(prog, m, type_params, t, &mut errors);
    for (span, message) in errors {
        prog.diagnostics.push(Diagnostic {
            severity: Severity::Error,
            code: "E0170",
            message,
            module: Some(m),
            span,
        });
    }
    ty
}

fn primitive(name: &str) -> Option<Type> {
    Some(match name {
        "byte" => Type::Byte,
        "short" => Type::Short,
        "int" => Type::Int,
        "long" => Type::Long,
        "float" => Type::Float,
        "double" => Type::Double,
        "currency" => Type::Currency,
        "decimal" => Type::Decimal,
        "ptr" => Type::Ptr,
        "date" => Type::Date,
        "string" => Type::Str,
        "boolean" => Type::Bool,
        "any" => Type::Any,
        "cellValue" => Type::CellValue,
        "object" => Type::Object,
        "void" => Type::Void,
        "wstr" => Type::WStr,
        "astr" => Type::AStr,
        "null" | "undefined" => Type::Null,
        _ => return None,
    })
}

pub const PRIMITIVE_NAMES: &[&str] = &[
    "byte",
    "short",
    "int",
    "long",
    "float",
    "double",
    "currency",
    "decimal",
    "ptr",
    "date",
    "string",
    "boolean",
    "any",
    "cellValue",
    "object",
    "void",
];

fn resolve_type_inner(
    prog: &mut Program,
    m: ModuleId,
    tps: &[TypeParamDef],
    t: &TypeRef,
    errors: &mut Vec<(Span, String)>,
) -> Type {
    match t {
        TypeRef::Error { .. } => Type::Error,
        TypeRef::Array { elem, span } => {
            let inner = resolve_type_inner(prog, m, tps, elem, errors);
            if inner.is_array() && inner.has_struct() {
                errors.push((
                    *span,
                    "構造体の 2 次元配列には対応していません。1 次元配列（`T[]`）を使ってください"
                        .into(),
                ));
                return Type::Error;
            }
            Type::Array(Box::new(inner))
        }
        TypeRef::Tuple { elems, span } => {
            let tys: Vec<Type> = elems
                .iter()
                .map(|e| resolve_type_inner(prog, m, tps, e, errors))
                .collect();
            if tys.len() < 2 {
                errors.push((
                    *span,
                    "タプル型には 2 つ以上の要素が必要です（`[int, string]`）".into(),
                ));
                return Type::Error;
            }
            anon_struct(
                prog,
                m,
                StructKind::Tuple,
                tys.into_iter()
                    .enumerate()
                    .map(|(i, t)| (format!("item{i}"), t, false))
                    .collect(),
                *span,
                errors,
            )
        }
        TypeRef::Object { fields, span } => {
            let mut out: Vec<(String, Type, bool)> = Vec::new();
            for f in fields {
                let ty = match &f.ty {
                    Some(t) => resolve_type_inner(prog, m, tps, t, errors),
                    None => {
                        errors.push((f.span, "フィールドには型注釈が必要です".into()));
                        Type::Error
                    }
                };
                if out.iter().any(|o| o.0.eq_ignore_ascii_case(&f.name.text)) {
                    errors.push((
                        f.name.span,
                        format!("フィールド `{}` が重複しています", f.name.text),
                    ));
                    continue;
                }
                out.push((f.name.text.clone(), ty, f.optional));
            }
            if out.is_empty() {
                errors.push((*span, "フィールドのないオブジェクト型は使えません".into()));
                return Type::Error;
            }
            anon_struct(prog, m, StructKind::Object, out, *span, errors)
        }
        TypeRef::Nullable { inner, span } => {
            let ty = resolve_type_inner(prog, m, tps, inner, errors);
            if !(ty.is_object()
                || matches!(
                    ty,
                    Type::Any | Type::CellValue | Type::Error | Type::Param(_)
                ))
            {
                errors.push((*span, format!("`{} | null` は使えません。`null` を持てるのはオブジェクト型と `any` のみです", prog.show(&ty))));
            }
            ty
        }
        TypeRef::Union {
            elems,
            nullable,
            span,
        } => {
            let mut members = Vec::new();
            for el in elems {
                let ty = resolve_type_inner(prog, m, tps, el, errors);
                let shown = prog.show(&ty);
                let bad = match &ty {
                    Type::Void => Some("`void` はユニオン型に含められません".to_string()),
                    Type::Ns(_) => Some(format!("`{shown}` は型ではありません")),
                    Type::Param(_) => Some(format!("型パラメータ `{shown}` はユニオン型に含められません")),
                    Type::WStr | Type::AStr => Some(format!("`{shown}` は `declare` の引数にだけ使えます")),
                    t if t.has_struct() => Some(format!("構造体 `{shown}` はユニオン型に含められません（ユニオン型の値は Variant / Object に格納されます）")),
                    _ => None,
                };
                if let Some(msg) = bad {
                    errors.push((el.span(), msg));
                    return Type::Error;
                }
                members.push(ty);
            }
            let ty = Type::union(members);
            if *nullable && !(ty.is_object() || ty.is_mixed_union() || ty.is_any() || ty.is_error())
            {
                errors.push((*span, format!("`{} | null` は使えません。`null` を持てるのはオブジェクト型を含む型と `any` のみです", prog.show(&ty))));
            }
            ty
        }
        TypeRef::Fn { params, ret, .. } => {
            let ps = params
                .iter()
                .map(|p| match &p.ty {
                    Some(t) => resolve_type_inner(prog, m, tps, t, errors),
                    None => {
                        errors.push((p.span, "関数型の引数には型注釈が必要です".into()));
                        Type::Error
                    }
                })
                .collect();
            let ret = resolve_type_inner(prog, m, tps, ret, errors);
            Type::Func(Box::new(FuncType { params: ps, ret }))
        }
        TypeRef::Named { path, args, span } => {
            let mut resolved: Vec<Type> = args
                .iter()
                .map(|a| resolve_type_inner(prog, m, tps, a, errors))
                .collect();
            let sym = if path.len() == 1 {
                let name = path[0].text.as_str();
                if let Some(p) = primitive(name) {
                    if !resolved.is_empty() {
                        errors.push((*span, format!("`{name}` は型引数を取りません")));
                    }
                    return p;
                }
                if tps.iter().any(|p| p.name == name) {
                    return Type::Param(name.to_string());
                }
                match name {
                    "number" => {
                        errors.push((
                            *span,
                            "`number` 型はありません。`int` または `double` を使ってください"
                                .into(),
                        ));
                        return Type::Error;
                    }
                    "Array" => {
                        if resolved.len() != 1 {
                            errors.push((*span, "`Array<T>` の型引数は 1 つです".into()));
                            return Type::Error;
                        }
                        return Type::Array(Box::new(resolved[0].clone()));
                    }
                    "Map" => {
                        if resolved.len() != 2 {
                            errors.push((*span, "`Map<K, V>` の型引数は 2 つです".into()));
                            return Type::Error;
                        }
                        if resolved[1].is_array() && resolved[1].has_struct() {
                            errors.push((*span, "構造体の配列は `Map` の値にできません".into()));
                            return Type::Error;
                        }
                        // Structs cannot live in a Dictionary (a Variant): such a map is a generated
                        // class that keeps the values in a struct array and maps keys to indices.
                        if let (Type::Struct(_), Some(c)) = (&resolved[1], struct_map_class(prog)) {
                            return Type::Class(c, resolved);
                        }
                        return Type::Map(
                            Box::new(resolved[0].clone()),
                            Box::new(resolved[1].clone()),
                        );
                    }
                    "Comparable" | "Numeric" | "ObjectType" => {
                        errors.push((*span, format!("`{name}` は型パラメータの制約（`T extends {name}`）としてのみ使えます")));
                        return Type::Error;
                    }
                    _ => {}
                }
                let in_ns = prog
                    .cur_namespace
                    .as_ref()
                    .and_then(|ns| prog.namespaces.get(ns))
                    .and_then(|n| n.get(name))
                    .filter(|s| {
                        matches!(
                            s,
                            Sym::Class(_) | Sym::Enum(_) | Sym::Struct(_) | Sym::Alias(_)
                        )
                    });
                in_ns
                    .cloned()
                    .or_else(|| prog.modules[m].scope.get(name).cloned())
                    .or_else(|| prog.globals.get(name).cloned())
            } else if path.len() == 2 {
                let (ns, name) = (path[0].text.as_str(), path[1].text.as_str());
                match prog.modules[m].scope.get(ns) {
                    Some(Sym::Module(t)) => prog.modules[*t].exports.get(name).cloned(),
                    _ => prog.namespaces.get(ns).and_then(|n| n.get(name)).cloned(),
                }
            } else {
                None
            };
            let display = path
                .iter()
                .map(|p| p.text.as_str())
                .collect::<Vec<_>>()
                .join(".");
            match sym {
                Some(Sym::Class(c)) => {
                    let expected = prog.classes[c].type_params.len();
                    // Omitted trailing type arguments take their defaults (`interface Sheets<T = Worksheet | Chart>`).
                    if resolved.len() < expected {
                        let defaults: Vec<Option<Type>> = prog.classes[c].type_params
                            [resolved.len()..]
                            .iter()
                            .map(|p| p.default.clone())
                            .collect();
                        if defaults.iter().all(|d| d.is_some()) {
                            resolved.extend(defaults.into_iter().flatten());
                        }
                    }
                    // A host collection without its element type (`Excel.Sheets`) holds objects.
                    if resolved.is_empty() && prog.classes[c].kind == ClassKind::Host {
                        resolved = vec![Type::Object; expected];
                    }
                    if resolved.len() != expected {
                        errors.push((
                            *span,
                            format!(
                                "`{display}` の型引数は {expected} 個です（{} 個指定されています）",
                                resolved.len()
                            ),
                        ));
                        return Type::Error;
                    }
                    Type::Class(c, resolved)
                }
                Some(Sym::Enum(e)) => Type::Enum(e),
                Some(Sym::Struct(s)) => Type::Struct(s),
                Some(Sym::Alias(t)) => t,
                // A user form's name is its default instance and its type (spec §14.11.9).
                Some(Sym::Var(v)) if matches!(&prog.vars[v].ty, Type::Class(c, _) if prog.form_module_of(*c).is_some() && prog.vars[v].declared) => {
                    prog.vars[v].ty.clone()
                }
                Some(_) => {
                    errors.push((*span, format!("`{display}` は型ではありません")));
                    Type::Error
                }
                None => {
                    errors.push((*span, format!("型 `{display}` が見つかりません")));
                    Type::Error
                }
            }
        }
    }
}

/// The module an `import ... from "<source>"` in module `m` names: a project module for a
/// relative path (`./util`, `../lib/str`), a bundled library module otherwise.
pub fn import_target(prog: &Program, m: ModuleId, source: &str) -> Option<ModuleId> {
    if source.starts_with("vba:") {
        return None;
    }
    if !source.starts_with('.') {
        return prog
            .modules
            .iter()
            .position(|x| x.kind == ModuleKind::Std && x.key == source);
    }
    // A test module names what it imports by its path (`../src/report`, `./helper`).
    if prog.modules[m].is_test() {
        return import_by_path(prog, m, source);
    }
    let base = prog.modules[m]
        .key
        .rsplit_once('/')
        .map(|(d, _)| d.to_string())
        .unwrap_or_default();
    let mut parts: Vec<&str> = if base.is_empty() {
        Vec::new()
    } else {
        base.split('/').collect()
    };
    for seg in source.split('/') {
        match seg {
            "." | "" => {}
            ".." => {
                parts.pop();
            }
            s => parts.push(s),
        }
    }
    let key = parts.join("/");
    let key = key.strip_suffix(".ras").unwrap_or(&key);
    prog.modules
        .iter()
        .position(|x| x.kind == ModuleKind::User && x.key == key)
}

/// The bundled class that implements `Map<K, V>` for struct values.
pub fn struct_map_class(prog: &Program) -> Option<ClassId> {
    let m = prog.module_by_key(STRUCT_MAP_MODULE)?;
    match prog.modules[m].exports.get("RAS_StructMap") {
        Some(Sym::Class(c)) => Some(*c),
        _ => None,
    }
}

/// The module a relative import names by the path of its file from the project root
/// (`test/report.test.ras` importing `../src/report` → `src/report.ras`).
fn import_by_path(prog: &Program, m: ModuleId, source: &str) -> Option<ModuleId> {
    let from = &prog.modules[m].rel_path;
    let mut parts: Vec<&str> = from
        .rsplit_once('/')
        .map(|(d, _)| d.split('/').collect())
        .unwrap_or_default();
    for seg in source.split('/') {
        match seg {
            "." | "" => {}
            ".." => {
                parts.pop()?;
            }
            s => parts.push(s),
        }
    }
    let path = parts.join("/");
    let path = path.strip_suffix(".ras").unwrap_or(&path);
    prog.modules
        .iter()
        .position(|x| x.kind == ModuleKind::User && x.rel_path.strip_suffix(".ras") == Some(path))
}

pub const STRUCT_MAP_MODULE: &str = "ras_collections";

/// The bundled type definitions of the test API (`describe`, `expect`, ...; spec §14.12.4).
pub const TEST_DEFS_MODULE: &str = "test.d";

/// Import-key prefix of the test files and their helper modules (`test/a.test.ras` has the
/// key `@test/test/a.test`): they are apart from the modules built into the workbook.
pub const TEST_KEY_PREFIX: &str = "@test/";

/// Import-key prefix of a form's code-behind (`forms/<Name>.ras` has the key `<forms>/<Name>`).
/// It cannot be spelled by a relative import, so a code-behind is never imported.
pub const FORM_KEY_PREFIX: &str = "<forms>/";

/// Import-key prefix of the file of a sheet or of `ThisWorkbook` (`objects/<CodeName>.ras` has
/// the key `<objects>/<CodeName>`, spec §12.2).
pub const OBJECT_KEY_PREFIX: &str = "<objects>/";

fn anon_struct(
    prog: &mut Program,
    m: ModuleId,
    kind: StructKind,
    fields: Vec<(String, Type, bool)>,
    span: Span,
    errors: &mut Vec<(Span, String)>,
) -> Type {
    if fields.iter().any(|f| f.1.is_error()) {
        return Type::Error;
    }
    if fields.iter().any(|f| f.1.contains_param()) {
        errors.push((span, "型パラメータを含むタプル・オブジェクト型は使えません。`struct` を宣言するか、具体的な型で書いてください".into()));
        return Type::Error;
    }
    if let Some(bad) = fields
        .iter()
        .find(|f| matches!(f.1, Type::Void | Type::Null | Type::Ns(_)))
    {
        errors.push((
            span,
            format!("`{}` はフィールドの型にできません", prog.show(&bad.1)),
        ));
        return Type::Error;
    }
    Type::Struct(prog.intern_struct(m, kind, fields))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn module_names() {
        assert_eq!(vba_module_name("strutil"), "Strutil");
        assert_eq!(vba_module_name("str_util"), "StrUtil");
        assert_eq!(vba_module_name("sub/excel-util"), "ExcelUtil");
        assert_eq!(vba_module_name("StrUtil"), "StrUtil");
        assert_eq!(vba_module_name("1st"), "Mod1st");
    }
}
