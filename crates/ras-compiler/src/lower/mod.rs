//! Lowering: typed AST → VBA AST.
//!
//! One `VModule` is produced per RAS module / class instantiation, plus the
//! document-module code for event handlers and the generated init glue.

mod bulk;
mod events;
pub(crate) mod expr;
mod fusion;
mod stmt;

use std::collections::{HashMap, HashSet};

use ras_syntax::ast::{self, Expr, ExprId, InlineMode, PassMode, Span, Visibility};

use crate::check::const_value;
use crate::config::Config;
use crate::diag::{Diagnostic, Severity};
use crate::model::*;
use crate::names;
use crate::vba::name as vname;
use crate::vba::*;

pub struct Lowered {
    pub modules: Vec<VModule>,
    pub diagnostics: Vec<Diagnostic>,
    pub notes: Vec<(ModuleId, OptNote)>,
    /// Names of macro entry points (`Module.proc`).
    pub macros: Vec<String>,
}

#[derive(Clone, Debug)]
pub(crate) enum Binding {
    /// A VBA variable.
    Var(String),
    /// Direct substitution (inlined arguments, fused lambda parameters, captured values).
    Expr(VExpr),
}

pub(crate) struct Frame {
    /// Module whose `TypeInfo` describes the code being lowered.
    pub module: ModuleId,
    pub fn_id: Option<FnId>,
    pub env: HashMap<String, Type>,
    pub locals: HashMap<LocalId, Binding>,
    /// Value of `this` in class methods.
    pub this: Option<VExpr>,
    /// Inline instance number: locals are named `ras_inl{N}_{name}`.
    pub inl: Option<u32>,
}

pub(crate) enum RetMode {
    /// Assign to the procedure name, then `Exit` (or jump to the exit label).
    Proc,
    /// Inlined body / fused lambda block: assign to `var`, jump to `label` unless in tail position.
    Inline {
        var: Option<String>,
        ty: Type,
        label: String,
        used: bool,
        tries_base: usize,
    },
}

#[derive(Clone, PartialEq)]
pub(crate) enum LoopKind {
    For,
    Do,
    Switch,
}

pub(crate) struct Breakable {
    pub kind: LoopKind,
    /// RAS label (`outer: for ...`).
    pub label: Option<String>,
    pub break_label: Option<String>,
    pub continue_label: Option<String>,
    /// `tries.len()` when the loop was entered.
    pub tries_base: usize,
}

pub(crate) enum PendingExit {
    Return,
    Jump(String, usize),
    Rethrow,
}

pub(crate) enum TryKind {
    /// `On Error Resume Next` region.
    ResumeNext,
    /// try/catch without finally (only `ras_tryId` needs restoring).
    Catch,
    /// A region whose exits must run a `finally` block first.
    Finally {
        label: String,
        exit_var: String,
        pending: Vec<PendingExit>,
    },
}

pub(crate) struct TryCtx {
    pub kind: TryKind,
    /// `ras_tryId` value to restore when leaving the region.
    pub outer_id: u32,
}

#[derive(Default)]
pub(crate) struct ProcCtx {
    pub emit_module: ModuleId,
    pub in_class: bool,
    /// Lower-cased names already used in this procedure.
    pub names: HashSet<String>,
    pub locals: Vec<(String, String, bool)>,
    pub frames: Vec<Frame>,
    pub rets: Vec<RetMode>,
    pub blocks: Vec<Vec<VStmt>>,
    pub label_n: u32,
    pub temp_n: u32,
    pub inl_n: u32,
    pub yield_n: u32,
    pub try_n: u32,
    pub buf_n: u32,
    pub breakables: Vec<Breakable>,
    pub tries: Vec<TryCtx>,
    /// Current `ras_tryId` value while lowering.
    pub cur_try_id: u32,
    pub handler_arms: Vec<(u32, Vec<VStmt>)>,
    pub uses_handler: bool,
    pub uses_try_id: bool,
    pub exit_kw: &'static str,
    /// VBA name of the procedure (return value variable of functions).
    pub proc_name: String,
    pub ret_ty: Type,
    /// `Module.proc`, used as error source.
    pub display_name: String,
    pub exit_label: Option<String>,
    pub exit_label_used: bool,
    pub loop_depth: usize,
    /// Receiver to prepend to auto-qualified `Cells` / `Range` (spec §13.4).
    pub qualifier: Option<VExpr>,
    pub bulk: HashMap<ExprId, bulk::BulkSlot>,
    /// `frames.len()` of the code the bulk slots belong to.
    pub bulk_depth: usize,
    /// No yield points while lowering a bulk-transfer loop (spec §10.4.3).
    pub suppress_yield: bool,
    /// Arrays in capacity-doubling mode: VBA array name → (length var, capacity var).
    pub push_caps: HashMap<String, (String, String)>,
    pub inline_stack: Vec<FnId>,
    pub stmt_count: usize,
    /// True while lowering the body of a closure class (`Apply`).
    pub in_closure: bool,
    /// The arrow function being lowered into a closure class.
    pub closure_span: Option<Span>,
    /// Variable being initialised by its declaration, when the declaration runs only once
    /// (omitted fields of a struct literal need no reset then).
    pub fresh_target: Option<String>,
    /// Top-level variables of a script module that became locals of `Main`.
    pub demoted: HashMap<VarId, String>,
}

pub(crate) struct Lowerer<'p> {
    pub prog: &'p Program,
    pub cfg: &'p Config,
    pub opt: u8,
    pub out: Vec<VModule>,
    index: HashMap<String, usize>,
    fn_insts: HashMap<(FnId, String), String>,
    fn_queue: Vec<(FnId, Vec<Type>)>,
    class_insts: HashMap<(ClassId, String), String>,
    class_queue: Vec<(ClassId, Vec<Type>)>,
    used_declares: HashSet<FnId>,
    /// Structs whose arrays are measured (`RAS_Len_<Struct>` helpers to generate).
    pub(crate) struct_len_used: std::collections::BTreeSet<StructId>,
    pub lambda_n: u32,
    pub diags: Vec<Diagnostic>,
    pub notes: Vec<(ModuleId, OptNote)>,
    pub global_names: HashSet<String>,
    no_yield: HashSet<FnId>,
    yield_plans: HashMap<FnId, std::rc::Rc<HashSet<u32>>>,
    pub pc: ProcCtx,
    force_public: HashSet<FnId>,
    force_public_vars: HashSet<VarId>,
    /// Variables and fields declared `WithEvents` (spec §12.2.2).
    with_events: HashSet<EventSource>,
    fn_procs: HashMap<FnId, Vec<(usize, String)>>,
    /// Procedures to call from `RAS_InitAll`, in order.
    init_calls: Vec<String>,
    macros: Vec<String>,
    instance_counts: HashMap<FnId, usize>,
    reported_3d: bool,
    /// Listener classes of the `on` calls (spec §12.2.3).
    listeners: Vec<events::ListenerClass>,
    /// Per event: the classes whose listener is chosen by the type of the target at run time
    /// (type name, listener class), and whether arrays and collections are listened to.
    dispatchers: Vec<events::Dispatcher>,
    /// Forms whose code attaches event handlers: they are detached when the form is unloaded.
    form_listens: HashSet<ModuleId>,
    /// No "this arrow function becomes a class" hint (the handler of an `on` call).
    quiet_closures: bool,
}

/// Workbook / worksheet event signatures: (name, VBA type, by reference).
/// Events of the workbook and of a sheet that `@on` handles (spec §12.2), as `event_params` names them.
const WORKBOOK_EVENTS: &[&str] = &[
    "Open",
    "Activate",
    "Deactivate",
    "AddinInstall",
    "AddinUninstall",
    "AfterXmlExport",
    "BeforeClose",
    "BeforeSave",
    "AfterSave",
    "BeforePrint",
    "NewSheet",
    "SheetActivate",
    "SheetDeactivate",
    "SheetCalculate",
    "SheetChange",
    "SheetSelectionChange",
    "SheetBeforeDoubleClick",
    "SheetBeforeRightClick",
    "SheetFollowHyperlink",
    "WindowActivate",
    "WindowDeactivate",
    "WindowResize",
];
const SHEET_EVENTS: &[&str] = &[
    "Activate",
    "Deactivate",
    "Calculate",
    "Change",
    "SelectionChange",
    "BeforeDoubleClick",
    "BeforeRightClick",
    "FollowHyperlink",
    "BeforeDelete",
];

/// The event `event` (any case) of the workbook or of a sheet, spelled as `@on` takes it;
/// `None` when `@on` does not handle it.
pub(crate) fn document_event(target_is_workbook: bool, event: &str) -> Option<&'static str> {
    let events = if target_is_workbook {
        WORKBOOK_EVENTS
    } else {
        SHEET_EVENTS
    };
    events
        .iter()
        .copied()
        .find(|e| e.eq_ignore_ascii_case(event))
        .filter(|e| event_params(target_is_workbook, e).is_some())
}

/// The events of a control on a sheet (spec §12.2.1): those of the control on a form, without
/// what a form raises (`Enter`, `Exit`, `BeforeUpdate`, `AfterUpdate`) and with what a sheet
/// raises (`GotFocus`, `LostFocus`).
fn sheet_control_events(control_type: &str) -> Vec<String> {
    let mut names: Vec<String> = crate::userform::events_of(control_type)
        .iter()
        .map(|e| e.name.clone())
        .filter(|n| !FORM_ONLY_EVENTS.iter().any(|x| x.eq_ignore_ascii_case(n)))
        .collect();
    if !names.is_empty() {
        names.extend(SHEET_ONLY_EVENTS.iter().map(|s| s.to_string()));
    }
    names
}

const FORM_ONLY_EVENTS: [&str; 4] = ["Enter", "Exit", "BeforeUpdate", "AfterUpdate"];
const SHEET_ONLY_EVENTS: [&str; 2] = ["GotFocus", "LostFocus"];

/// Name and parameters (name, VBA type, by reference) of an event of a control on a sheet.
pub(crate) fn sheet_control_event(
    control_type: &str,
    event: &str,
) -> Option<(String, Vec<(String, String, bool)>)> {
    if FORM_ONLY_EVENTS
        .iter()
        .any(|x| x.eq_ignore_ascii_case(event))
    {
        return None;
    }
    if let Some(name) = SHEET_ONLY_EVENTS
        .iter()
        .find(|x| x.eq_ignore_ascii_case(event))
    {
        return if crate::userform::events_of(control_type).is_empty() {
            None
        } else {
            Some((name.to_string(), Vec::new()))
        };
    }
    let def = crate::userform::find_event(control_type, event)?;
    Some((
        def.name.clone(),
        def.params
            .iter()
            .map(|p| (p.name.clone(), p.vba_type.clone(), p.byref))
            .collect(),
    ))
}

/// The events an object of the host class `class` raises for a `WithEvents` variable
/// (spec §12.2.2): name and parameters (name, VBA type, by reference). `None` for a class
/// whose events are not known.
pub(crate) fn host_events(
    prog: &Program,
    class: ClassId,
) -> Option<Vec<(String, Vec<(String, String, bool)>)>> {
    type Sig = (&'static str, &'static [(&'static str, &'static str, bool)]);
    const WB: (&str, &str, bool) = ("Wb", "Workbook", false);
    const SH: (&str, &str, bool) = ("Sh", "Object", false);
    const TARGET: (&str, &str, bool) = ("Target", "Range", false);
    const CANCEL: (&str, &str, bool) = ("Cancel", "Boolean", true);
    const WN: (&str, &str, bool) = ("Wn", "Window", false);
    const PVW: (&str, &str, bool) = ("Pvw", "ProtectedViewWindow", false);
    const APPLICATION: &[Sig] = &[
        ("NewWorkbook", &[WB]),
        ("SheetActivate", &[SH]),
        ("SheetDeactivate", &[SH]),
        ("SheetCalculate", &[SH]),
        ("SheetBeforeDelete", &[SH]),
        ("SheetChange", &[SH, TARGET]),
        ("SheetSelectionChange", &[SH, TARGET]),
        ("SheetBeforeDoubleClick", &[SH, TARGET, CANCEL]),
        ("SheetBeforeRightClick", &[SH, TARGET, CANCEL]),
        (
            "SheetFollowHyperlink",
            &[SH, ("Target", "Hyperlink", false)],
        ),
        (
            "SheetPivotTableUpdate",
            &[SH, ("Target", "PivotTable", false)],
        ),
        ("WindowActivate", &[WB, WN]),
        ("WindowDeactivate", &[WB, WN]),
        ("WindowResize", &[WB, WN]),
        ("WorkbookOpen", &[WB]),
        ("WorkbookActivate", &[WB]),
        ("WorkbookDeactivate", &[WB]),
        ("WorkbookAddinInstall", &[WB]),
        ("WorkbookAddinUninstall", &[WB]),
        ("WorkbookBeforeClose", &[WB, CANCEL]),
        ("WorkbookBeforePrint", &[WB, CANCEL]),
        (
            "WorkbookBeforeSave",
            &[WB, ("SaveAsUI", "Boolean", false), CANCEL],
        ),
        ("WorkbookAfterSave", &[WB, ("Success", "Boolean", false)]),
        ("WorkbookNewSheet", &[WB, SH]),
        ("WorkbookNewChart", &[WB, ("Ch", "Chart", false)]),
        ("AfterCalculate", &[]),
        ("ProtectedViewWindowOpen", &[PVW]),
        ("ProtectedViewWindowActivate", &[PVW]),
        ("ProtectedViewWindowDeactivate", &[PVW]),
        ("ProtectedViewWindowResize", &[PVW]),
        ("ProtectedViewWindowBeforeEdit", &[PVW, CANCEL]),
    ];
    const ELEMENT: [(&str, &str, bool); 3] = [
        ("ElementID", "Long", false),
        ("Arg1", "Long", false),
        ("Arg2", "Long", false),
    ];
    const MOUSE: &[(&str, &str, bool)] = &[
        ("Button", "Long", false),
        ("Shift", "Long", false),
        ("x", "Long", false),
        ("y", "Long", false),
    ];
    const CHART: &[Sig] = &[
        ("Activate", &[]),
        ("Deactivate", &[]),
        ("Calculate", &[]),
        ("Resize", &[]),
        (
            "BeforeDoubleClick",
            &[ELEMENT[0], ELEMENT[1], ELEMENT[2], CANCEL],
        ),
        ("BeforeRightClick", &[CANCEL]),
        ("Select", &ELEMENT),
        (
            "SeriesChange",
            &[
                ("SeriesIndex", "Long", false),
                ("PointIndex", "Long", false),
            ],
        ),
        ("MouseDown", MOUSE),
        ("MouseMove", MOUSE),
        ("MouseUp", MOUSE),
    ];
    const QUERY_TABLE: &[Sig] = &[
        ("BeforeRefresh", &[CANCEL]),
        ("AfterRefresh", &[("Success", "Boolean", false)]),
    ];
    let own = |list: &[Sig]| {
        list.iter()
            .map(|(name, params)| {
                (
                    name.to_string(),
                    params
                        .iter()
                        .map(|p| (p.0.to_string(), p.1.to_string(), p.2))
                        .collect(),
                )
            })
            .collect()
    };
    let document = |workbook: bool| {
        let names = if workbook {
            WORKBOOK_EVENTS
        } else {
            SHEET_EVENTS
        };
        names
            .iter()
            .filter_map(|name| {
                event_params(workbook, name).map(|sig| {
                    (
                        name.to_string(),
                        sig.into_iter()
                            .map(|p| (p.0.to_string(), p.1.to_string(), p.2))
                            .collect(),
                    )
                })
            })
            .collect()
    };
    let def = &prog.classes[class];
    match (def.namespace.as_deref(), def.name.as_str()) {
        (Some("Excel"), "Application") => Some(own(APPLICATION)),
        (Some("Excel"), "Workbook") => Some(document(true)),
        (Some("Excel"), "Worksheet") => Some(document(false)),
        (Some("Excel"), "Chart") => Some(own(CHART)),
        (Some("Excel"), "QueryTable") => Some(own(QUERY_TABLE)),
        // What a form raises for its controls (`Enter`, `Exit` …) does not reach a variable.
        (Some("MSForms"), name) => {
            let events: Vec<(String, Vec<(String, String, bool)>)> =
                crate::userform::events_of(name)
                    .into_iter()
                    .filter(|e| {
                        !FORM_ONLY_EVENTS
                            .iter()
                            .any(|x| x.eq_ignore_ascii_case(&e.name))
                    })
                    .map(|e| {
                        (
                            e.name.clone(),
                            e.params
                                .iter()
                                .map(|p| (p.name.clone(), p.vba_type.clone(), p.byref))
                                .collect(),
                        )
                    })
                    .collect();
            if events.is_empty() {
                None
            } else {
                Some(events)
            }
        }
        _ => None,
    }
}

/// What `target.on(…)` can attach a handler to on an object of a class (spec §12.2.3).
pub(crate) enum ListenTarget {
    /// An object with events: the host class a `WithEvents` variable is declared as, and its
    /// events (name and parameters).
    Events(ClassId, Vec<(String, Vec<(String, String, bool)>)>),
    /// An object with an `OnAction` property (a shape, a button of the Forms toolbar).
    OnAction,
}

/// How an object of class `class` is listened to: by the events of the class or of the host
/// class it extends (the class of a sheet's file extends `Excel.Worksheet`), or by `OnAction`.
/// `None` for a class without either. A form is none: its events are the form module's.
pub(crate) fn listen_target(prog: &Program, class: ClassId) -> Option<ListenTarget> {
    let mut cur = Some(class);
    for _ in 0..16 {
        let c = cur?;
        let def = &prog.classes[c];
        if def.kind != ClassKind::Host
            || (def.namespace.as_deref() == Some("MSForms") && def.name == "UserForm")
        {
            return None;
        }
        if let Some(events) = host_events(prog, c) {
            return Some(ListenTarget::Events(c, events));
        }
        if def.fields.iter().any(|f| {
            f.vba_name().eq_ignore_ascii_case("OnAction") && f.ty == Type::Str && !f.readonly
        }) {
            return Some(ListenTarget::OnAction);
        }
        cur = def.extends.iter().find_map(|t| match t {
            Type::Class(b, _) => Some(*b),
            _ => None,
        });
    }
    None
}

/// An event of the workbook or of a sheet with its parameters (name, VBA type, by reference).
fn document_event_sig(
    target_is_workbook: bool,
    event: &str,
) -> Option<(&'static str, Vec<(&'static str, &'static str, bool)>)> {
    let name = document_event(target_is_workbook, event)?;
    Some((name, event_params(target_is_workbook, name)?))
}

fn event_params(
    target_is_workbook: bool,
    event: &str,
) -> Option<Vec<(&'static str, &'static str, bool)>> {
    let p = |n, t, r| (n, t, r);
    Some(if target_is_workbook {
        match event {
            "Open" | "Activate" | "Deactivate" | "AddinInstall" | "AddinUninstall"
            | "AfterXmlExport" => vec![],
            "BeforeClose" => vec![p("Cancel", "Boolean", true)],
            "BeforeSave" => vec![
                p("SaveAsUI", "Boolean", false),
                p("Cancel", "Boolean", true),
            ],
            "AfterSave" => vec![p("Success", "Boolean", false)],
            "BeforePrint" => vec![p("Cancel", "Boolean", true)],
            "NewSheet" | "SheetActivate" | "SheetDeactivate" | "SheetCalculate" => {
                vec![p("Sh", "Object", false)]
            }
            "SheetChange" | "SheetSelectionChange" => {
                vec![p("Sh", "Object", false), p("Target", "Range", false)]
            }
            "SheetBeforeDoubleClick" | "SheetBeforeRightClick" => {
                vec![
                    p("Sh", "Object", false),
                    p("Target", "Range", false),
                    p("Cancel", "Boolean", true),
                ]
            }
            "SheetFollowHyperlink" => {
                vec![p("Sh", "Object", false), p("Target", "Hyperlink", false)]
            }
            "WindowActivate" | "WindowDeactivate" | "WindowResize" => {
                vec![p("Wn", "Window", false)]
            }
            _ => return None,
        }
    } else {
        match event {
            "Activate" | "Deactivate" | "Calculate" => vec![],
            "Change" | "SelectionChange" => vec![p("Target", "Range", false)],
            "BeforeDoubleClick" | "BeforeRightClick" => {
                vec![p("Target", "Range", false), p("Cancel", "Boolean", true)]
            }
            "FollowHyperlink" => vec![p("Target", "Hyperlink", false)],
            "BeforeDelete" => vec![],
            _ => return None,
        }
    })
}

pub fn lower(prog: &Program) -> Lowered {
    let mut lw = Lowerer {
        prog,
        cfg: &prog.config,
        opt: prog.config.opt_level(),
        out: Vec::new(),
        index: HashMap::new(),
        fn_insts: HashMap::new(),
        fn_queue: Vec::new(),
        class_insts: HashMap::new(),
        class_queue: Vec::new(),
        used_declares: HashSet::new(),
        struct_len_used: Default::default(),
        lambda_n: 0,
        diags: Vec::new(),
        notes: Vec::new(),
        global_names: HashSet::new(),
        no_yield: HashSet::new(),
        yield_plans: HashMap::new(),
        pc: ProcCtx::default(),
        force_public: HashSet::new(),
        force_public_vars: HashSet::new(),
        with_events: HashSet::new(),
        fn_procs: HashMap::new(),
        init_calls: Vec::new(),
        macros: Vec::new(),
        instance_counts: HashMap::new(),
        reported_3d: false,
        listeners: Vec::new(),
        dispatchers: Vec::new(),
        form_listens: HashSet::new(),
        quiet_closures: false,
    };
    lw.run();
    Lowered {
        modules: lw.out,
        diagnostics: lw.diags,
        notes: lw.notes,
        macros: lw.macros,
    }
}

impl<'p> Lowerer<'p> {
    // ------------------------------------------------------------------ diagnostics

    pub(crate) fn diag(
        &mut self,
        severity: Severity,
        module: ModuleId,
        span: Span,
        code: &'static str,
        message: impl Into<String>,
    ) {
        self.diags.push(Diagnostic {
            severity,
            code,
            message: message.into(),
            module: Some(module),
            span,
        });
    }

    pub(crate) fn err(&mut self, span: Span, code: &'static str, message: impl Into<String>) {
        let m = self.src_module();
        self.diag(Severity::Error, m, span, code, message);
    }

    /// Records an optimisation note (spec §6 rule 13) for hints / CodeLens.
    pub(crate) fn note(
        &mut self,
        span: Span,
        applied: bool,
        kind: &'static str,
        message: impl Into<String>,
    ) {
        let m = self.src_module();
        self.notes.push((
            m,
            OptNote {
                span,
                applied,
                kind,
                message: message.into(),
            },
        ));
    }

    // ------------------------------------------------------------------ driver

    fn run(&mut self) {
        let prog = self.prog;
        // Names that locals must not shadow.
        for m in &prog.modules {
            if m.kind == ModuleKind::TypeDef {
                continue;
            }
            self.global_names.insert(m.vba_name.to_lowercase());
            for &f in &m.fns {
                // Members of bundled classes (`keys`, `values`, `size` …) are not module-level names.
                let member = matches!(
                    prog.fns[f].kind,
                    FnKind::Method { .. }
                        | FnKind::Getter { .. }
                        | FnKind::Setter { .. }
                        | FnKind::Ctor { .. }
                );
                if m.kind == ModuleKind::Std && member {
                    continue;
                }
                self.global_names.insert(prog.fns[f].name.to_lowercase());
            }
            for &v in &m.vars {
                self.global_names.insert(prog.vars[v].name.to_lowercase());
            }
            for &c in &m.classes {
                self.global_names
                    .insert(prog.classes[c].name.to_lowercase());
            }
            for &e in &m.enums {
                self.global_names.insert(prog.enums[e].name.to_lowercase());
            }
        }
        // spec §10.4.3: no DoEvents in UDFs, event handlers, @init and anything they may call.
        let mut stack: Vec<FnId> = prog
            .fns
            .iter()
            .filter(|f| {
                f.attrs.udf
                    || !f.attrs.on_events.is_empty()
                    || f.attrs.init
                    || matches!(f.kind, FnKind::ModInit)
            })
            .map(|f| f.id)
            .collect();
        while let Some(f) = stack.pop() {
            if self.no_yield.insert(f) {
                stack.extend(prog.fns[f].calls.iter().copied());
            }
        }

        self.collect_with_events();
        for m in 0..prog.modules.len() {
            if prog.modules[m].kind == ModuleKind::User {
                self.emit_module(m);
            }
        }
        loop {
            if let Some((f, targs)) = self.fn_queue.pop() {
                self.emit_fn_instance(f, &targs);
            } else if let Some((c, targs)) = self.class_queue.pop() {
                self.emit_class(c, &targs);
            } else {
                break;
            }
        }
        self.emit_dispatchers();
        self.emit_listeners();
        self.emit_struct_helpers();
        self.emit_std_declares();
        self.emit_events();
        self.emit_module_vars();
        self.emit_init_glue();
        // Visibility fix-ups for functions referenced from other VBA modules.
        for f in self.force_public.clone() {
            for (mi, name) in self.fn_procs.get(&f).cloned().unwrap_or_default() {
                for p in self.out[mi].procs.iter_mut().filter(|p| p.name == name) {
                    p.public = true;
                }
            }
        }
        // spec §2.4.1: warn when one generic definition expands into many instances.
        let limit = self.cfg.build.generic_instance_limit;
        for (f, n) in self.instance_counts.clone() {
            if n > limit {
                let d = &prog.fns[f];
                self.diag(
                    Severity::Warning,
                    d.module,
                    d.name_span,
                    "W0501",
                    format!(
                        "ジェネリック関数 `{}` から {n} 個の実装が出力されます（閾値 {limit}）",
                        d.name
                    ),
                );
            }
        }
        for f in &prog.fns {
            if f.is_generic()
                && f.exported
                && !self.instance_counts.contains_key(&f.id)
                && prog.modules[f.module].kind == ModuleKind::User
            {
                self.diag(
                    Severity::Info,
                    f.module,
                    f.name_span,
                    "I0502",
                    format!(
                        "ジェネリック関数 `{}` は一度も呼ばれていないため、何も出力されません",
                        f.name
                    ),
                );
            }
        }
        // A sheet's file without code leaves the module of the sheet without code (spec §12.2).
        self.out.retain(|m| {
            !(m.decls.is_empty()
                && m.procs.is_empty()
                && matches!(m.kind, VModuleKind::Standard | VModuleKind::Document))
        });
    }

    pub(crate) fn vmodule(&mut self, name: &str, kind: VModuleKind) -> usize {
        if let Some(&i) = self.index.get(name) {
            return i;
        }
        let i = self.out.len();
        self.out.push(VModule::new(name, kind));
        self.index.insert(name.to_string(), i);
        i
    }

    fn std_vmodule(&mut self, m: ModuleId) -> usize {
        let md = &self.prog.modules[m];
        if let Some(form) = &md.form {
            // A form's code-behind is the form's code module (spec §14.11); the file of a sheet
            // or of `ThisWorkbook` is that document module (spec §12.2).
            let kind = if form.kind == ObjectKind::Form {
                VModuleKind::Form
            } else {
                VModuleKind::Document
            };
            let i = self.vmodule(&md.vba_name.clone(), kind);
            if self.out[i].sources.is_empty() {
                self.out[i].sources.push(md.rel_path.clone());
            }
            return i;
        }
        let i = self.vmodule(&md.vba_name.clone(), VModuleKind::Standard);
        if self.out[i].sources.is_empty() {
            self.out[i].sources.push(md.rel_path.clone());
            self.out[i].private_module = true;
        }
        i
    }

    fn emit_module(&mut self, m: ModuleId) {
        let prog = self.prog;
        let md = &prog.modules[m];
        let mi = self.std_vmodule(m);
        let has_public = md.fns.iter().any(|&f| {
            let d = &prog.fns[f];
            matches!(d.kind, FnKind::Free | FnKind::Main) && (d.exported || d.attrs.macro_entry)
        });
        self.out[mi].private_module = !has_public;
        for &e in &md.enums {
            let d = &prog.enums[e];
            let vis = if d.exported { "Public" } else { "Private" };
            let mut s = format!("{vis} Enum {}\n", names::member_ident(&d.name));
            for mem in &d.members {
                s.push_str(&format!(
                    "    {} = {}\n",
                    self.enum_member_name(e, &mem.name),
                    mem.value
                ));
            }
            s.push_str("End Enum");
            self.out[mi].decls.push(s);
        }
        // A `Type` must be declared before the types that contain it ("forward reference" error).
        let mut order: Vec<StructId> = Vec::new();
        fn visit(
            prog: &Program,
            home: &[StructId],
            s: StructId,
            order: &mut Vec<StructId>,
            depth: usize,
        ) {
            if order.contains(&s) || !home.contains(&s) || depth > 64 {
                return;
            }
            for (_, fty) in &prog.structs[s].fields {
                let inner = match fty {
                    Type::Array(_) => fty.elem(),
                    other => Some(other),
                };
                if let Some(Type::Struct(d)) = inner {
                    if *d != s {
                        visit(prog, home, *d, order, depth + 1);
                    }
                }
            }
            if !order.contains(&s) {
                order.push(s);
            }
        }
        for &sid in &md.structs {
            visit(prog, &md.structs, sid, &mut order, 0);
        }
        for sid in order {
            let d = &prog.structs[sid];
            let mut s = format!("Public Type {}\n", self.struct_name(sid));
            for (fname, fty) in &d.fields {
                let (t, arr) = self.vtype(fty);
                s.push_str(&format!(
                    "    {}{} As {t}\n",
                    names::member_ident(fname),
                    if arr { "()" } else { "" }
                ));
            }
            s.push_str("End Type");
            self.out[mi].decls.push(s);
        }
        let declares: Vec<FnId> = md
            .fns
            .iter()
            .copied()
            .filter(|&f| matches!(prog.fns[f].kind, FnKind::Declare { .. }))
            .collect();
        if !declares.is_empty() {
            let block = self.declare_block(&declares, false);
            self.out[mi].decls.push(block);
        }
        for &f in &md.fns {
            let d = &prog.fns[f];
            if !matches!(d.kind, FnKind::Free | FnKind::Main | FnKind::ModInit)
                || d.is_generic()
                || d.decl.as_ref().is_none_or(|x| x.body.is_none())
            {
                continue;
            }
            // Non-exported `inline` functions are fully expanded and not emitted (spec §8.1).
            if d.inline == InlineMode::Inline
                && !d.exported
                && self.opt > 0
                && d.inline_blockers.is_empty()
            {
                continue;
            }
            self.emit_fn_instance(f, &[]);
        }
        for &c in &md.classes {
            if prog.classes[c].type_params.is_empty() {
                self.class_name(c, &[]);
            }
        }
        if md.is_form() {
            // Variables of a form are initialised by `UserForm_Initialize` (see `emit_form_events`).
            self.emit_form_events(m);
            return;
        }
        // The events of a sheet's (the workbook's) file are generated by `emit_events`; its
        // variables are initialised with those of the standard modules.
        if let Some(f) = md.init_fn {
            self.init_calls
                .push(format!("{}.{}", md.vba_name, prog.fns[f].name));
        }
        for &f in &md.fns {
            if prog.fns[f].attrs.init {
                self.init_calls
                    .push(format!("{}.{}", md.vba_name, prog.fns[f].proc_ident()));
                self.force_public.insert(f);
            }
        }
    }

    pub(crate) fn enum_member_name(&self, e: EnumId, member: &str) -> String {
        let d = &self.prog.enums[e];
        if d.host {
            member.to_string()
        } else {
            format!("{}_{}", names::member_ident(&d.name), member)
        }
    }

    /// `Declare` statements for both VBA7 and VBA6 (spec §11.3).
    fn declare_block(&mut self, fns: &[FnId], force_public: bool) -> String {
        let prog = self.prog;
        let mut vba7 = Vec::new();
        let mut vba6 = Vec::new();
        let mut win64 = Vec::new();
        for &f in fns {
            let d = &prog.fns[f];
            let FnKind::Declare { lib, alias } = &d.kind else {
                continue;
            };
            let vis = if d.exported || force_public {
                "Public"
            } else {
                "Private"
            };
            let uses_long = d.params.iter().any(|p| p.ty == Type::Long) || d.ret_ty() == Type::Long;
            let line = |ptr: &str, ptrsafe: bool| {
                let params = d
                    .params
                    .iter()
                    .map(|p| {
                        let (ty, arr) = match &p.ty {
                            Type::Ptr | Type::WStr => (ptr.to_string(), false),
                            Type::Array(e) => (decl_scalar(e, ptr), true),
                            other => (decl_scalar(other, ptr), false),
                        };
                        let byref =
                            p.mode == PassMode::ByRef || arr || matches!(p.ty, Type::Struct(_));
                        let ty = match &p.ty {
                            Type::Struct(s) => self.struct_name(*s),
                            _ => ty,
                        };
                        format!(
                            "{} {}{} As {ty}",
                            if byref { "ByRef" } else { "ByVal" },
                            names::member_ident(&p.name),
                            if arr { "()" } else { "" }
                        )
                    })
                    .collect::<Vec<_>>()
                    .join(", ");
                let ret = d.ret_ty();
                let kw = if ret == Type::Void { "Sub" } else { "Function" };
                let alias = alias
                    .as_ref()
                    .map(|a| format!(" Alias \"{a}\""))
                    .unwrap_or_default();
                let ret_s = if ret == Type::Void {
                    String::new()
                } else {
                    format!(" As {}", decl_scalar(&ret, ptr))
                };
                format!(
                    "{vis} Declare {}{kw} {} Lib \"{lib}\"{alias} ({params}){ret_s}",
                    if ptrsafe { "PtrSafe " } else { "" },
                    names::member_ident(&d.name)
                )
            };
            if uses_long {
                win64.push(line("LongPtr", true));
            } else {
                vba7.push(line("LongPtr", true));
                vba6.push(line("Long", false));
            }
        }
        let mut s = String::new();
        if !vba7.is_empty() {
            if self.cfg.target.vba6 {
                s.push_str("#If VBA7 Then\n");
                s.push_str(&vba7.join("\n"));
                s.push_str("\n#Else\n");
                s.push_str(&vba6.join("\n"));
                s.push_str("\n#End If");
            } else {
                s.push_str(&vba7.join("\n"));
            }
        }
        if !win64.is_empty() {
            if !s.is_empty() {
                s.push('\n');
            }
            // `long` (LongLong) only exists in 64-bit Office.
            s.push_str("#If Win64 Then\n");
            s.push_str(&win64.join("\n"));
            s.push_str("\n#End If");
        }
        s
    }

    /// `RAS_Len_<Struct>(arr)`: length of an array of structs (0 when unallocated).
    fn emit_struct_helpers(&mut self) {
        for s in self.struct_len_used.clone() {
            let d = &self.prog.structs[s];
            let tname = self.struct_name(s);
            let fname = format!("RAS_Len_{tname}");
            let mi = self.std_vmodule(d.module);
            let arr = || vname("arr");
            self.out[mi].procs.push(VProc {
                kind: ProcKind::Function,
                public: true,
                friend: false,
                name: fname.clone(),
                params: vec![VParam {
                    name: "arr".into(),
                    byref: true,
                    ty: tname,
                    is_array: true,
                    optional_default: None,
                }],
                ret: Some("Long".into()),
                locals: vec![],
                body: vec![
                    VStmt::Raw("On Error Resume Next".into()),
                    assign(
                        vname(fname),
                        bin(
                            "+",
                            bin(
                                "-",
                                call("UBound", vec![arr()]),
                                call("LBound", vec![arr()]),
                            ),
                            long(1),
                        ),
                    ),
                ],
                leading_comments: vec![],
                attributes: vec![],
                source: None,
            });
        }
    }

    fn emit_std_declares(&mut self) {
        let prog = self.prog;
        let mut by_module: HashMap<ModuleId, Vec<FnId>> = HashMap::new();
        let mut used: Vec<FnId> = self.used_declares.iter().copied().collect();
        used.sort_unstable();
        for f in used {
            by_module.entry(prog.fns[f].module).or_default().push(f);
        }
        let mut mods: Vec<ModuleId> = by_module.keys().copied().collect();
        mods.sort_unstable();
        for m in mods {
            let mi = self.std_vmodule(m);
            let block = self.declare_block(&by_module[&m], true);
            self.out[mi].decls.insert(0, block);
        }
    }

    pub(crate) fn is_demoted(&self, v: VarId) -> bool {
        let d = &self.prog.vars[v];
        self.prog.modules[d.module].main_fn.is_some()
            && !d.used_outside_main
            && !d.exported
            && !d.declared
    }

    /// VBA `Const` text for a constant variable, when it can be one.
    pub(crate) fn const_literal(&self, v: VarId) -> Option<String> {
        let d = &self.prog.vars[v];
        if !d.is_const || d.declared {
            return None;
        }
        let val = d
            .const_value
            .clone()
            .or_else(|| d.init.as_ref().and_then(|i| const_value(i)))?;
        match (&val, &d.ty) {
            (ConstVal::Int(i), Type::Int | Type::Enum(_)) => i32::try_from(*i)
                .ok()
                .filter(|x| *x != i32::MIN)
                .map(|x| format!("{x}&")),
            (ConstVal::Int(i), Type::Short | Type::Byte) => Some(format!("{i}")),
            (ConstVal::Int(i), Type::Double) => Some(format!("{i}#")),
            (ConstVal::Int(i), Type::Float) => Some(format!("{i}!")),
            (ConstVal::Float(f), Type::Double) => Some(expr::double_lit(*f)),
            (ConstVal::Bool(b), Type::Bool) => {
                Some(if *b { "True".into() } else { "False".into() })
            }
            (ConstVal::Str(s), Type::Str) => {
                let plain =
                    s.chars().all(|c| !c.is_control()) && s.len() < 500 && self.encodable(s);
                plain.then(|| format!("\"{}\"", s.replace('"', "\"\"")))
            }
            _ => None,
        }
    }

    pub(crate) fn encodable(&self, s: &str) -> bool {
        crate::emit::encodable(s, self.cfg.target.codepage)
    }

    fn emit_module_vars(&mut self) {
        let prog = self.prog;
        for m in 0..prog.modules.len() {
            if prog.modules[m].kind == ModuleKind::TypeDef {
                continue;
            }
            let mut lines = Vec::new();
            for &v in &prog.modules[m].vars {
                let d = &prog.vars[v];
                if d.declared || self.is_demoted(v) {
                    continue;
                }
                let public = d.exported || self.force_public_vars.contains(&v);
                let vis = if public { "Public" } else { "Private" };
                let name = names::module_ident(&d.name);
                let (ty, arr) = self.vtype(&d.ty);
                if self.with_events.contains(&EventSource::Var(v)) {
                    // Its object's events are handled (spec §12.2.2).
                    lines.push(format!("Private WithEvents {name} As {ty}"));
                } else if let Some(c) = self.const_literal(v) {
                    lines.push(format!("{vis} Const {name} As {ty} = {c}"));
                } else {
                    lines.push(ptr_conditional(
                        &format!("{vis} {name}{} As {ty}", if arr { "()" } else { "" }),
                        self.cfg.target.vba6,
                    ));
                }
            }
            if !lines.is_empty() {
                if prog.modules[m].kind == ModuleKind::Std
                    && !self.index.contains_key(&prog.modules[m].vba_name)
                {
                    continue;
                }
                let mi = self.std_vmodule(m);
                self.out[mi].decls.push(lines.join("\n"));
            }
        }
    }

    // ------------------------------------------------------------------ types and names

    /// VBA type for a declaration: (type name, is array).
    pub(crate) fn vtype(&mut self, t: &Type) -> (String, bool) {
        match t {
            Type::Array(_) => {
                let elem = t.elem().unwrap().clone();
                if elem.is_array() && !self.reported_3d {
                    self.reported_3d = true;
                    let m = self.src_module();
                    self.diag(
                        Severity::Error,
                        m,
                        Span::default(),
                        "E0603",
                        "3 次元以上の配列には対応していません（`T[][]` までです）",
                    );
                }
                // Arrays of `any` live in a plain Variant so that Range data is never copied (spec §9.2.1).
                if matches!(elem, Type::Any | Type::Error | Type::Null | Type::Decimal) {
                    ("Variant".into(), false)
                } else {
                    (self.vscalar(&elem), true)
                }
            }
            other => (self.vscalar(other), false),
        }
    }

    fn vscalar(&mut self, t: &Type) -> String {
        match t {
            Type::Byte => "Byte".into(),
            Type::Short => "Integer".into(),
            Type::Int => "Long".into(),
            Type::Long => "LongLong".into(),
            Type::Float => "Single".into(),
            Type::Double => "Double".into(),
            Type::Currency => "Currency".into(),
            Type::Ptr | Type::WStr => "LongPtr".into(),
            Type::Date => "Date".into(),
            Type::Str | Type::AStr => "String".into(),
            Type::Bool => "Boolean".into(),
            Type::Object | Type::Map(..) | Type::Func(_) => "Object".into(),
            // A union of object types lives in an `Object`, any other union in a `Variant` (spec §2.6.1).
            Type::Union(_) if t.is_object() => "Object".into(),
            Type::Class(c, args) => self.class_name(*c, args),
            Type::Enum(e) => {
                let d = &self.prog.enums[*e];
                if d.host && !d.late_bound {
                    d.name.clone()
                } else {
                    // Private enums of other modules are not visible as a type; `Long` is equivalent.
                    "Long".into()
                }
            }
            Type::Struct(s) => self.struct_name(*s),
            _ => "Variant".into(),
        }
    }

    /// Name of the VBA `Type` of a struct. A `Public Type` is visible in the whole project, so
    /// structs of different modules that share a name get the module name appended (all but the first).
    pub(crate) fn struct_name(&self, s: StructId) -> String {
        let d = &self.prog.structs[s];
        let base = names::member_ident(&d.name);
        let first = self
            .prog
            .structs
            .iter()
            .find(|o| o.name.eq_ignore_ascii_case(&d.name))
            .map(|o| o.id);
        if first == Some(s) {
            base
        } else {
            format!("{base}_{}", self.prog.modules[d.module].vba_name)
        }
    }

    /// Name fragment used in mangled generic names (spec §2.4.1).
    fn mangle_one(&mut self, t: &Type) -> String {
        match t {
            // Structural, so that `T = int[]` and `T = int[][]` get different implementations.
            Type::Array(inner) => format!("{}Arr", self.mangle_one(inner)),
            other => self.vscalar(other),
        }
    }

    pub(crate) fn mangle(&mut self, targs: &[Type]) -> String {
        targs
            .iter()
            .map(|t| self.mangle_one(t))
            .collect::<Vec<_>>()
            .join("_")
    }

    pub(crate) fn shorten(name: String, max: usize) -> String {
        if name.chars().count() <= max {
            return name;
        }
        let hash = crc32fast::hash(name.as_bytes());
        let keep: String = name.chars().take(max - 9).collect();
        format!("{keep}_{hash:08X}")
    }

    /// VBA class name of a class instantiation; schedules its emission.
    pub(crate) fn class_name(&mut self, c: ClassId, args: &[Type]) -> String {
        let def = &self.prog.classes[c];
        if def.kind == ClassKind::Host {
            return match def.namespace.as_deref() {
                // What `on` returns is an instance of one of the listener classes (spec §12.2.3).
                None if def.name == "RAS_Listener" => "Object".into(),
                // (`model::referenced_namespace` names the same libraries.)
                Some("Excel") | Some("VBA") | Some("Office") | None => def.name.clone(),
                // `StdFont` is stdole's; the rest are qualified (`Label`, `Image` … would be ambiguous).
                Some("MSForms") if def.name == "StdFont" => def.name.clone(),
                Some("MSForms") => format!("MSForms.{}", def.name),
                // `Scripting`, `VBScript` and libraries from `ras typegen` are late-bound (spec §15.3).
                Some(_) => "Object".into(),
            };
        }
        let mangled = self.mangle(args);
        if let Some(n) = self.class_insts.get(&(c, mangled.clone())) {
            return n.clone();
        }
        let base = names::member_ident(&def.name);
        let name = if args.is_empty() {
            base
        } else {
            format!("{base}__{mangled}")
        };
        // Class module names are limited to 31 characters.
        let name = Self::shorten(name, 31);
        self.class_insts.insert((c, mangled), name.clone());
        self.class_queue.push((c, args.to_vec()));
        name
    }

    /// Unqualified VBA procedure name for a function instantiation; schedules emission when needed.
    pub(crate) fn fn_name(&mut self, f: FnId, targs: &[Type]) -> String {
        let def = &self.prog.fns[f];
        let base = match &def.kind {
            FnKind::Ctor { .. } => "RAS_Init".to_string(),
            FnKind::Method {
                class,
                is_static: true,
            }
            | FnKind::Getter {
                class,
                is_static: true,
            }
            | FnKind::Setter {
                class,
                is_static: true,
            } => {
                format!(
                    "{}_{}",
                    names::member_ident(&self.prog.classes[*class].name),
                    def.name
                )
            }
            FnKind::Host { .. } => def.vba_name().to_string(),
            _ => def.proc_ident(),
        };
        if let FnKind::Declare { .. } = def.kind {
            if self.prog.modules[def.module].kind == ModuleKind::Std {
                self.used_declares.insert(f);
            }
            return base;
        }
        let needs_queue = def.is_generic() || self.prog.modules[def.module].kind == ModuleKind::Std;
        if !needs_queue {
            return base;
        }
        let mangled = self.mangle(targs);
        if let Some(n) = self.fn_insts.get(&(f, mangled.clone())) {
            return n.clone();
        }
        let name = if targs.is_empty() {
            base
        } else {
            Self::shorten(format!("{base}__{mangled}"), 255)
        };
        self.fn_insts.insert((f, mangled), name.clone());
        if matches!(def.kind, FnKind::Free) {
            let count = {
                let c = self.instance_counts.entry(f).or_default();
                if def.is_generic() {
                    *c += 1;
                }
                *c
            };
            // A recursion whose type arguments keep growing (`f<T>` calling `f<T[]>`) never ends (spec §2.4.1).
            if count == 257 {
                self.diag(Severity::Error, def.module, def.name_span, "E0602", format!("ジェネリック関数 `{}` の展開が終わりません（呼び出しのたびに型引数が大きくなる再帰は使えません）", def.name));
            }
            if count <= 256 {
                self.fn_queue.push((f, targs.to_vec()));
            }
        }
        name
    }

    /// Expression naming a free function from the current procedure (module-qualified when needed).
    pub(crate) fn fn_callee(&mut self, f: FnId, targs: &[Type]) -> VExpr {
        let name = self.fn_name(f, targs);
        let def = &self.prog.fns[f];
        let same_module =
            def.module == self.pc.emit_module && !self.pc.in_class && !self.pc.in_closure;
        if same_module {
            return VExpr::Name(name);
        }
        if !def.exported {
            self.force_public.insert(f);
        }
        if let Some(me) = self.closure_form(def.module) {
            return me.member(name);
        }
        VExpr::Name(self.prog.modules[def.module].vba_name.clone()).member(name)
    }

    /// Inside an arrow function of a form's code-behind, the form the closure was created in:
    /// the functions and variables of the code-behind are that instance's, not those of the
    /// form's default instance (which naming the form would load).
    fn closure_form(&self, module: ModuleId) -> Option<VExpr> {
        if self.pc.in_closure
            && self.prog.modules[module].is_form()
            && self.frame().module == module
        {
            self.frame().this.clone()
        } else {
            None
        }
    }

    pub(crate) fn var_ref(&mut self, v: VarId) -> VExpr {
        let d = &self.prog.vars[v];
        if d.declared {
            return VExpr::Name(d.vba_name().to_string());
        }
        let name = names::module_ident(&d.name);
        let same_module =
            d.module == self.pc.emit_module && !self.pc.in_class && !self.pc.in_closure;
        if same_module {
            return VExpr::Name(name);
        }
        // A constant of a form (a sheet) stays private there: its value.
        let object_module = self.prog.modules[d.module].form.is_some();
        if object_module && !d.exported {
            if let Some(c) = self.const_literal(v) {
                return lit(c);
            }
            // VBA has no `Public` arrays and structures in an object module.
            let ty = d.ty.clone();
            if self.vtype(&ty).1 || ty.has_struct() {
                let m = self.src_module();
                let span = self.pc.closure_span.unwrap_or_default();
                self.diag(Severity::Error, m, span, "E0642", format!("アロー関数の中から、フォーム・シートのファイルの配列・構造体の変数 `{}` は参照できません（アロー関数は別のクラスになり、VBA はオブジェクト モジュールの配列・構造体を公開できないためです）。アロー関数の外で値を取り出すか、関数を通して読み書きしてください", d.name));
            }
        }
        if !d.exported {
            self.force_public_vars.insert(v);
        }
        if let Some(me) = self.closure_form(d.module) {
            return me.member(name);
        }
        VExpr::Name(self.prog.modules[d.module].vba_name.clone()).member(name)
    }

    /// Initial value of a type (`None` for arrays, structs and other types without a literal default).
    pub(crate) fn default_value(&mut self, t: &Type) -> Option<VExpr> {
        Some(match t {
            Type::Byte | Type::Short => lit("0"),
            Type::Int | Type::Enum(_) => lit("0&"),
            Type::Long => lit("0^"),
            Type::Float => lit("0!"),
            Type::Double => lit("0#"),
            Type::Currency => lit("0@"),
            Type::Ptr => lit("0"),
            Type::Date => call("CDate", vec![lit("0")]),
            Type::Str => lit("vbNullString"),
            Type::Bool => lit("False"),
            Type::Any | Type::Decimal | Type::Error | Type::Null | Type::Param(_) => lit("Empty"),
            Type::Object | Type::Class(..) | Type::Map(..) | Type::Func(_) => lit("Nothing"),
            Type::Union(_) if t.is_object() => lit("Nothing"),
            Type::Union(_) => lit("Empty"),
            Type::Array(_) if self.vtype(t) == ("Variant".to_string(), false) => lit("Empty"),
            _ => return None,
        })
    }

    // ------------------------------------------------------------------ procedure context helpers

    pub(crate) fn frame(&self) -> &Frame {
        self.pc.frames.last().expect("no frame")
    }

    pub(crate) fn src_module(&self) -> ModuleId {
        self.pc
            .frames
            .last()
            .map(|f| f.module)
            .unwrap_or(self.pc.emit_module)
    }

    pub(crate) fn info(&self) -> &'p TypeInfo {
        &self.prog.modules[self.src_module()].info
    }

    /// Static type of an expression in the current instantiation.
    pub(crate) fn ty(&self, e: &Expr) -> Type {
        let t = self.info().ty(e.id);
        match self.pc.frames.last() {
            Some(f) if !f.env.is_empty() => t.subst(&f.env),
            _ => t,
        }
    }

    pub(crate) fn subst(&self, t: &Type) -> Type {
        match self.pc.frames.last() {
            Some(f) if !f.env.is_empty() => t.subst(&f.env),
            _ => t.clone(),
        }
    }

    pub(crate) fn emit(&mut self, s: VStmt) {
        self.pc.stmt_count += 1;
        self.pc.blocks.last_mut().expect("no block").push(s);
    }

    /// Runs `f`, collecting the statements it emits into a separate block.
    pub(crate) fn capture<R>(&mut self, f: impl FnOnce(&mut Self) -> R) -> (Vec<VStmt>, R) {
        self.pc.blocks.push(Vec::new());
        let r = f(self);
        (self.pc.blocks.pop().unwrap(), r)
    }

    pub(crate) fn out_len(&self) -> usize {
        self.pc.blocks.last().map_or(0, |b| b.len())
    }

    /// Reserves a procedure-unique identifier based on `base` (`x`, `x_2`, …; spec §13.1).
    pub(crate) fn unique(&mut self, base: &str) -> String {
        let mut candidate = base.to_string();
        let mut n = 1;
        loop {
            let lower = candidate.to_lowercase();
            if !self.pc.names.contains(&lower) && !self.global_names.contains(&lower) {
                self.pc.names.insert(lower);
                return candidate;
            }
            n += 1;
            candidate = format!("{base}_{n}");
        }
    }

    /// Declares a hoisted local (`Dim` at the top of the procedure).
    pub(crate) fn declare(&mut self, base: &str, t: &Type) -> String {
        let name = self.unique(base);
        let (ty, arr) = self.vtype(t);
        self.pc.locals.push((name.clone(), ty, arr));
        name
    }

    pub(crate) fn declare_raw(&mut self, base: &str, ty: &str, arr: bool) -> String {
        let name = self.unique(base);
        self.pc.locals.push((name.clone(), ty.to_string(), arr));
        name
    }

    pub(crate) fn temp(&mut self, t: &Type) -> String {
        self.pc.temp_n += 1;
        let base = format!("ras_t{}", self.pc.temp_n);
        self.declare(&base, t)
    }

    pub(crate) fn label(&mut self, prefix: &str) -> String {
        self.pc.label_n += 1;
        let base = format!("{prefix}_{}", self.pc.label_n);
        self.unique(&base)
    }

    /// Name for a user local in the current frame (`ras_inl{N}_{name}` inside inlined bodies).
    /// VBA identifier of a local variable or parameter. Besides the reserved and well-known names,
    /// the name of any bundled host global is avoided: a local `workbooks` would shadow `Workbooks`,
    /// and the VBE would rewrite the casing of every use of it (spec §13.6).
    pub(crate) fn local_ident(&self, name: &str) -> String {
        let base = names::local_ident(name);
        if base == name
            && self
                .prog
                .host_global_names
                .contains_key(&name.to_lowercase())
        {
            format!("{base}_")
        } else {
            base
        }
    }

    pub(crate) fn local_name(&mut self, name: &str, t: &Type) -> String {
        let base = match self.frame().inl {
            Some(n) => format!("ras_inl{n}_{}", names::valid_ident(name)),
            None => self.local_ident(name),
        };
        self.declare(&base, t)
    }

    pub(crate) fn bind(&mut self, local: LocalId, b: Binding) {
        self.pc.frames.last_mut().unwrap().locals.insert(local, b);
    }

    pub(crate) fn local_expr(&self, local: LocalId) -> VExpr {
        match self.frame().locals.get(&local) {
            Some(Binding::Var(n)) => VExpr::Name(n.clone()),
            Some(Binding::Expr(e)) => e.clone(),
            None => VExpr::Name(format!("ras_unbound_{local}")),
        }
    }

    pub(crate) fn local_ty(&self, local: LocalId) -> Type {
        let f = self.frame();
        match f.fn_id {
            Some(id) => self.prog.fns[id]
                .locals
                .get(local)
                .map(|l| l.ty.subst(&f.env))
                .unwrap_or(Type::Error),
            None => Type::Error,
        }
    }

    pub(crate) fn line_of(&self, span: Span) -> u32 {
        self.prog.modules[self.src_module()].line_of(u32::from(span.start()))
    }

    // ------------------------------------------------------------------ functions

    fn yield_plan(&mut self, f: FnId) -> std::rc::Rc<HashSet<u32>> {
        if let Some(p) = self.yield_plans.get(&f) {
            return p.clone();
        }
        let def = &self.prog.fns[f];
        let mut set = HashSet::new();
        let enabled = self.cfg.doevents.enabled
            && !def.attrs.nodoevents
            && (!self.no_yield.contains(&f) || def.attrs.doevents);
        if let Some(body) = def.decl.as_ref().and_then(|d| d.body.as_ref()) {
            stmt::plan_yields(self, f, &body.stmts, enabled, &mut set);
        }
        let rc = std::rc::Rc::new(set);
        self.yield_plans.insert(f, rc.clone());
        rc
    }

    pub(crate) fn current_yield_plan(&mut self) -> std::rc::Rc<HashSet<u32>> {
        match self.frame().fn_id {
            Some(f) => self.yield_plan(f),
            None => std::rc::Rc::new(HashSet::new()),
        }
    }

    fn doc_lines(doc: &Option<String>) -> (Vec<String>, Option<String>) {
        let Some(doc) = doc else {
            return (Vec::new(), None);
        };
        // The summary is the first paragraph (spec §16.1.3).
        let summary: Vec<&str> = doc
            .lines()
            .take_while(|l| !l.trim().is_empty() && !l.trim_start().starts_with('@'))
            .collect();
        if summary.is_empty() {
            return (Vec::new(), None);
        }
        let one_line = summary.join(" ");
        (
            summary.iter().map(|s| s.to_string()).collect(),
            Some(one_line),
        )
    }

    fn emit_fn_instance(&mut self, f: FnId, targs: &[Type]) {
        let prog = self.prog;
        let def = &prog.fns[f];
        let mi = self.std_vmodule(def.module);
        let name = self.fn_name(f, targs);
        let env: HashMap<String, Type> = def
            .type_params
            .iter()
            .zip(targs.iter())
            .map(|(p, a)| (p.name.clone(), a.clone()))
            .collect();
        // In a form, only `export` makes a member public: handlers are called from the form itself.
        // In a sheet's file too, where the handlers are called from the sheet's own module.
        let public = def.exported
            || (!def.attrs.form_member
                && (def.attrs.macro_entry
                    || (!def.attrs.on_events.is_empty() && !def.attrs.doc_member)
                    || matches!(def.kind, FnKind::Main | FnKind::ModInit)));
        if let Some(p) = self.build_proc(f, env, None, name.clone(), public) {
            self.fn_procs.entry(f).or_default().push((mi, name));
            self.out[mi].procs.push(p);
        }
        if def.is_macro_entry() {
            self.macros.push(format!(
                "{}.{}",
                prog.modules[def.module].vba_name,
                def.proc_ident()
            ));
        }
    }

    /// Builds the VBA procedure for function `f` instantiated with `env`.
    pub(crate) fn build_proc(
        &mut self,
        f: FnId,
        env: HashMap<String, Type>,
        class: Option<(ClassId, Vec<Type>)>,
        name: String,
        public: bool,
    ) -> Option<VProc> {
        let prog = self.prog;
        let def = &prog.fns[f];
        let decl = def.decl.clone()?;
        let body = decl.body.as_ref()?;
        let md = &prog.modules[def.module];
        let saved = std::mem::take(&mut self.pc);
        self.pc.emit_module = def.module;
        self.pc.in_class = class.is_some();
        self.pc.blocks.push(Vec::new());
        self.pc.names.insert(name.to_lowercase());
        if let Some((c, _)) = &class {
            for fld in &prog.classes[*c].fields {
                self.pc.names.insert(fld.name.to_lowercase());
            }
            for &m in &prog.classes[*c].methods {
                self.pc.names.insert(prog.fns[m].name.to_lowercase());
            }
        }
        let ret_ty = def.ret_ty().subst(&env);
        let env_copy = env.clone();
        let is_instance = matches!(
            def.kind,
            FnKind::Method {
                is_static: false,
                ..
            } | FnKind::Getter {
                is_static: false,
                ..
            } | FnKind::Setter {
                is_static: false,
                ..
            } | FnKind::Ctor { .. }
        );
        let mut frame = Frame {
            module: def.module,
            fn_id: Some(f),
            env,
            locals: HashMap::new(),
            this: is_instance.then(|| lit("Me")),
            inl: None,
        };
        let mut params = Vec::new();
        for p in &def.params {
            let pty = p.ty.subst(&frame.env);
            let base = if def.locals[p.local].kind == LocalKind::This {
                "this_".to_string()
            } else {
                self.local_ident(&p.name)
            };
            let pname = self.unique(&base);
            let (ty, is_array) = self.vtype(&pty);
            let byref = p.mode == PassMode::ByRef || is_array || matches!(pty, Type::Struct(_));
            // Constant defaults are also declared as `Optional` so that callers outside RAS can omit them.
            let optional_default = if byref || matches!(def.kind, FnKind::Setter { .. }) {
                None
            } else {
                match &p.default {
                    Some(d) => expr::literal_default(d, &pty),
                    None if p.optional => self
                        .default_value(&pty)
                        .map(|v| v.print())
                        .filter(|s| s != "Nothing" && !s.starts_with("CDate")),
                    None => None,
                }
            };
            frame.locals.insert(p.local, Binding::Var(pname.clone()));
            params.push(VParam {
                name: pname,
                byref,
                ty,
                is_array,
                optional_default,
            });
        }
        if matches!(def.kind, FnKind::ModInit) && !def.attrs.form_member {
            // A dummy parameter keeps the initializer out of Excel's macro list.
            params.push(VParam {
                name: "ras_dummy".into(),
                byref: false,
                ty: "Long".into(),
                is_array: false,
                optional_default: Some("0&".into()),
            });
        }
        // `Optional` parameters must be trailing in VBA.
        let mut seen_required = false;
        for p in params.iter_mut().rev() {
            if p.optional_default.is_none() {
                seen_required = true;
            } else if seen_required {
                p.optional_default = None;
            }
        }
        let (kind, exit_kw) = match (&def.kind, &ret_ty) {
            (FnKind::Getter { .. }, _) => (ProcKind::PropertyGet, "Property"),
            (FnKind::Setter { .. }, _) => {
                let obj = def
                    .params
                    .first()
                    .is_some_and(|p| p.ty.subst(&frame.env).is_object());
                (
                    if obj {
                        ProcKind::PropertySet
                    } else {
                        ProcKind::PropertyLet
                    },
                    "Property",
                )
            }
            (_, Type::Void) => (ProcKind::Sub, "Sub"),
            _ => (ProcKind::Function, "Function"),
        };
        let entry = def.is_macro_entry();
        let fast = def.attrs.fast;
        self.pc.exit_kw = exit_kw;
        self.pc.proc_name = name.clone();
        self.pc.ret_ty = ret_ty.clone();
        self.pc.display_name = format!("{}.{}", md.vba_name, def.name);
        self.pc.uses_handler = entry || fast || stmt::needs_handler(&body.stmts);
        if entry || fast {
            let l = self.unique("RAS_Exit");
            self.pc.exit_label = Some(l);
        }
        self.pc.frames.push(frame);
        self.pc.rets.push(RetMode::Proc);
        self.pc.inline_stack.push(f);

        let mut prologue = Vec::new();
        let mut epilogue = Vec::new();
        if entry {
            let mut flags = 0;
            if def.attrs.reentrant {
                flags |= 1;
            }
            if def.attrs.silent {
                flags |= 2;
            }
            if def.attrs.nocancel {
                flags |= 4;
            }
            let mut args = vec![lit(format!("\"{}\"", self.pc.display_name))];
            if flags != 0 {
                args.push(long(flags));
            }
            prologue.push(VStmt::Raw(format!(
                "If Not {} Then Exit {exit_kw}",
                call("RAS_Enter", args).print()
            )));
        }
        if self.pc.uses_handler {
            prologue.push(VStmt::Raw("On Error GoTo RAS_Handler".into()));
            self.pc.names.insert("ras_handler".into());
        }
        if fast {
            prologue.push(call_stmt("RAS_SuspendAll", vec![]));
            epilogue.push(call_stmt("RAS_ResumeAll", vec![]));
        }
        if entry {
            epilogue.push(call_stmt("RAS_Leave", vec![]));
        }
        for s in prologue {
            self.emit(s);
        }
        let has_epilogue = !epilogue.is_empty();
        self.stmts(&body.stmts, true);
        if let Some(l) = self.pc.exit_label.clone() {
            if self.pc.exit_label_used || has_epilogue {
                self.emit(VStmt::Label(l));
            }
        }
        for s in epilogue.clone() {
            self.emit(s);
        }
        if self.pc.uses_handler {
            self.emit(VStmt::Exit(exit_kw));
            self.emit(VStmt::Label("RAS_Handler".into()));
            let file = lit(format!("\"{}\"", md.rel_path.replace('"', "")));
            self.emit(call_stmt(
                "RAS_ErrCapture",
                vec![
                    name_member("Err", "Number"),
                    name_member("Err", "Source"),
                    name_member("Err", "Description"),
                    lit("Erl"),
                    file,
                ],
            ));
            let arms = std::mem::take(&mut self.pc.handler_arms);
            if !arms.is_empty() {
                let cases = arms
                    .into_iter()
                    .map(|(id, body)| (vec![long(id as i64)], body))
                    .collect();
                self.emit(VStmt::Select {
                    expr: vname("ras_tryId"),
                    cases,
                    els: None,
                });
            }
            if entry {
                // Outermost entry: restore application state and report; nested: propagate to the caller.
                let exit_label = self.pc.exit_label.clone().unwrap();
                let mut fail_args = vec![lit(format!("\"{}\"", self.pc.display_name))];
                if let Some(h) = &def.attrs.on_error {
                    fail_args.push(lit(format!("\"{}\"", h.replace('"', ""))));
                }
                self.emit(VStmt::Raw(format!(
                    "If {} Then Resume {exit_label}",
                    call("RAS_Fail", fail_args).print()
                )));
            }
            for s in epilogue {
                self.emit(s);
            }
            self.emit(call_stmt("RAS_Rethrow", vec![]));
        }
        if self.pc.uses_try_id {
            self.pc
                .locals
                .insert(0, ("ras_tryId".into(), "Long".into(), false));
        }
        let body_stmts = self.pc.blocks.pop().unwrap();
        let locals = std::mem::take(&mut self.pc.locals);
        self.pc = saved;

        let (leading_comments, summary) = Self::doc_lines(&def.doc);
        let mut attributes = Vec::new();
        if let (true, Some(s)) = (public, summary) {
            // Characters outside the target code page cannot appear in an Attribute line.
            let text: String = s
                .chars()
                .filter(|c| !c.is_control() && self.encodable(c.encode_utf8(&mut [0; 4])))
                .take(240)
                .collect();
            if !text.trim().is_empty() {
                attributes.push(format!(
                    "Attribute {name}.VB_Description = \"{}\"",
                    text.replace('"', "\"\"")
                ));
            }
        }
        let (rt, rarr) = self.vtype(&ret_ty);
        let ret = match kind {
            ProcKind::Function | ProcKind::PropertyGet => {
                Some(if rarr { format!("{rt}()") } else { rt })
            }
            _ => None,
        };
        // A class module cannot expose a `Type` of a standard module through a public member.
        let friend = class.is_some()
            && public
            && (ret_ty.has_struct()
                || def
                    .params
                    .iter()
                    .any(|p| p.ty.subst(&env_copy).has_struct()));
        Some(VProc {
            kind,
            public,
            friend,
            name,
            params,
            ret,
            locals,
            body: body_stmts,
            leading_comments,
            attributes,
            source: Some(0),
        })
    }

    // ------------------------------------------------------------------ classes

    fn emit_class(&mut self, c: ClassId, targs: &[Type]) {
        let prog = self.prog;
        let def = &prog.classes[c];
        let name = self.class_name(c, targs);
        let env = prog.class_env(c, targs);
        let md = &prog.modules[def.module];
        let ci = self.vmodule(&name, VModuleKind::Class);
        self.out[ci].sources.push(md.rel_path.clone());
        if def.predeclared {
            self.out[ci]
                .class_attrs
                .push(("VB_PredeclaredId".into(), "True".into()));
        }
        let (doc, _) = Self::doc_lines(&def.doc);
        self.out[ci].leading_comments = doc;

        if def.kind == ClassKind::Interface {
            for fld in &def.fields {
                let (ty, arr) = self.vtype(&fld.ty.subst(&env));
                let ty = if arr { "Variant".to_string() } else { ty };
                self.out[ci]
                    .decls
                    .push(format!("Public {} As {ty}", names::member_ident(&fld.name)));
            }
            for &m in &def.methods {
                let p = self.signature_stub(m, &env, names::member_ident(&prog.fns[m].name), true);
                self.out[ci].procs.push(p);
            }
            return;
        }

        for iface in &def.implements {
            if let Type::Class(i, iargs) = iface.subst(&env) {
                let iname = self.class_name(i, &iargs);
                self.out[ci].implements.push(iname);
            }
        }
        let mi = self.std_vmodule(def.module);
        let mut init_stmts: Vec<(usize, bool)> = Vec::new();
        for (idx, fld) in def.fields.iter().enumerate() {
            let fty = fld.ty.subst(&env);
            let (ty, arr) = self.vtype(&fty);
            let fname = names::member_ident(&fld.name);
            if fld.is_static {
                let vis = "Public";
                let sname = format!("{}_{}", names::member_ident(&def.name), fld.name);
                self.out[mi].decls.push(ptr_conditional(
                    &format!("{vis} {sname}{} As {ty}", if arr { "()" } else { "" }),
                    self.cfg.target.vba6,
                ));
            } else if arr && fld.vis == Visibility::Public {
                // VBA does not allow public array fields: private backing field + Variant property.
                self.out[ci]
                    .decls
                    .push(format!("Private f_{fname}() As {ty}"));
                self.out[ci].procs.push(VProc {
                    kind: ProcKind::PropertyGet,
                    public: true,
                    friend: false,
                    name: fname.clone(),
                    params: vec![],
                    ret: Some("Variant".into()),
                    locals: vec![],
                    body: vec![assign(vname(fname.clone()), vname(format!("f_{fname}")))],
                    leading_comments: vec![],
                    attributes: vec![],
                    source: None,
                });
                self.out[ci].procs.push(VProc {
                    kind: ProcKind::PropertyLet,
                    public: true,
                    friend: false,
                    name: fname.clone(),
                    params: vec![VParam {
                        name: "ras_v".into(),
                        byref: false,
                        ty: "Variant".into(),
                        is_array: false,
                        optional_default: None,
                    }],
                    ret: None,
                    locals: vec![],
                    body: vec![assign(vname(format!("f_{fname}")), vname("ras_v"))],
                    leading_comments: vec![],
                    attributes: vec![],
                    source: None,
                });
            } else {
                let vis = if fld.vis == Visibility::Public {
                    "Public"
                } else {
                    "Private"
                };
                let with_events = if self.with_events.contains(&EventSource::Field(c, idx)) {
                    "WithEvents "
                } else {
                    ""
                };
                self.out[ci].decls.push(ptr_conditional(
                    &format!(
                        "{vis} {with_events}{fname}{} As {ty}",
                        if arr { "()" } else { "" }
                    ),
                    self.cfg.target.vba6,
                ));
            }
            if fld.init.is_some() {
                init_stmts.push((idx, fld.is_static));
            }
        }
        // Field initializers run in Class_Initialize (instance) / the init glue (static).
        for is_static in [false, true] {
            let fields: Vec<usize> = init_stmts
                .iter()
                .filter(|x| x.1 == is_static)
                .map(|x| x.0)
                .collect();
            if fields.is_empty() {
                continue;
            }
            let saved = std::mem::take(&mut self.pc);
            self.pc.emit_module = def.module;
            self.pc.in_class = !is_static;
            self.pc.blocks.push(Vec::new());
            self.pc.exit_kw = "Sub";
            self.pc.frames.push(Frame {
                module: def.module,
                fn_id: None,
                env: env.clone(),
                locals: HashMap::new(),
                this: (!is_static).then(|| lit("Me")),
                inl: None,
            });
            self.pc.rets.push(RetMode::Proc);
            for idx in fields {
                let fld = &def.fields[idx];
                let fty = fld.ty.subst(&env);
                let target = self.field_target(c, idx, &fty, None);
                let init = fld.init.clone().unwrap();
                self.expr_into(&init, &target, &fty);
            }
            let body = self.pc.blocks.pop().unwrap();
            let locals = std::mem::take(&mut self.pc.locals);
            self.pc = saved;
            if is_static {
                let pname = format!("RAS_StaticInit_{name}");
                self.out[mi].procs.push(VProc {
                    kind: ProcKind::Sub,
                    public: true,
                    friend: false,
                    name: pname.clone(),
                    params: vec![VParam {
                        name: "ras_dummy".into(),
                        byref: false,
                        ty: "Long".into(),
                        is_array: false,
                        optional_default: Some("0&".into()),
                    }],
                    ret: None,
                    locals,
                    body,
                    leading_comments: vec![],
                    attributes: vec![],
                    source: Some(0),
                });
                self.init_calls.push(format!("{}.{pname}", md.vba_name));
            } else {
                self.out[ci].procs.push(VProc {
                    kind: ProcKind::Sub,
                    public: false,
                    friend: false,
                    name: "Class_Initialize".into(),
                    params: vec![],
                    ret: None,
                    locals,
                    body,
                    leading_comments: vec![],
                    attributes: vec![],
                    source: Some(0),
                });
            }
        }
        // Event procedures of the fields whose objects are listened to (spec §12.2.2).
        let handlers: Vec<FnId> = def
            .methods
            .iter()
            .copied()
            .filter(|&m| !prog.fns[m].attrs.on_events.is_empty())
            .collect();
        self.emit_with_events(ci, &name, &handlers, true);
        for &m in &def.methods {
            let d = &prog.fns[m];
            if d.is_generic() {
                self.diag(Severity::Error, d.module, d.name_span, "E0601", "クラスのジェネリックメソッドにはまだ対応していません。モジュールのジェネリック関数（拡張メソッド）を使ってください");
                continue;
            }
            let is_static = matches!(
                d.kind,
                FnKind::Method {
                    is_static: true,
                    ..
                } | FnKind::Getter {
                    is_static: true,
                    ..
                } | FnKind::Setter {
                    is_static: true,
                    ..
                }
            );
            let pname = self.fn_name(m, &[]);
            if is_static {
                if let Some(p) = self.build_proc(m, env.clone(), None, pname, true) {
                    self.out[mi].procs.push(p);
                }
            } else if let Some(p) = self.build_proc(
                m,
                env.clone(),
                Some((c, targs.to_vec())),
                pname,
                d.vis == Visibility::Public,
            ) {
                self.out[ci].procs.push(p);
            }
        }
        // Constructor + factory function (`New_Foo(args)`).
        let mut factory_params = Vec::new();
        let mut factory_args = Vec::new();
        if let Some(ctor) = def.ctor {
            if let Some(p) = self.build_proc(
                ctor,
                env.clone(),
                Some((c, targs.to_vec())),
                "RAS_Init".into(),
                true,
            ) {
                factory_params = p.params.clone();
                factory_args = p.params.iter().map(|x| vname(x.name.clone())).collect();
                self.out[ci].procs.push(p);
            }
        }
        let fname = format!("New_{name}");
        let mut body = vec![set(vname(fname.clone()), lit(format!("New {name}")))];
        if def.ctor.is_some() {
            body.push(VStmt::Call {
                target: vname(fname.clone()).member("RAS_Init"),
                args: factory_args,
            });
        }
        self.out[mi].procs.push(VProc {
            kind: ProcKind::Function,
            public: true,
            friend: false,
            name: fname,
            params: factory_params,
            ret: Some(name.clone()),
            locals: vec![],
            body,
            leading_comments: vec![],
            attributes: vec![],
            source: None,
        });
        // Interface forwarding stubs: `Private Function IShape_area() ...`.
        for iface in &def.implements {
            let Type::Class(i, iargs) = iface.subst(&env) else {
                continue;
            };
            let iname = self.class_name(i, &iargs);
            let ienv = prog.class_env(i, &iargs);
            let idef = &prog.classes[i];
            for &m in &idef.methods {
                let mname = names::member_ident(&prog.fns[m].name);
                let mut p = self.signature_stub(m, &ienv, format!("{iname}_{mname}"), false);
                let args: Vec<VExpr> = p.params.iter().map(|x| vname(x.name.clone())).collect();
                let target = lit("Me").member(mname);
                let ret = prog.fns[m].ret_ty().subst(&ienv);
                p.body = if ret == Type::Void {
                    vec![VStmt::Call { target, args }]
                } else {
                    let value = if args.is_empty() {
                        target
                    } else {
                        target.index(args)
                    };
                    vec![VStmt::Assign {
                        target: vname(p.name.clone()),
                        value,
                        set: ret.is_object(),
                    }]
                };
                self.out[ci].procs.push(p);
            }
            for fld in &idef.fields {
                let fty = fld.ty.subst(&ienv);
                let (ty, arr) = self.vtype(&fty);
                let ty = if arr { "Variant".to_string() } else { ty };
                let pname = format!("{iname}_{}", names::member_ident(&fld.name));
                let me = lit("Me").member(names::member_ident(&fld.name));
                let obj = fty.is_object();
                self.out[ci].procs.push(VProc {
                    kind: ProcKind::PropertyGet,
                    public: false,
                    friend: false,
                    name: pname.clone(),
                    params: vec![],
                    ret: Some(ty.clone()),
                    locals: vec![],
                    body: vec![VStmt::Assign {
                        target: vname(pname.clone()),
                        value: me.clone(),
                        set: obj,
                    }],
                    leading_comments: vec![],
                    attributes: vec![],
                    source: None,
                });
                if !fld.readonly {
                    self.out[ci].procs.push(VProc {
                        kind: if obj {
                            ProcKind::PropertySet
                        } else {
                            ProcKind::PropertyLet
                        },
                        public: false,
                        friend: false,
                        name: pname,
                        params: vec![VParam {
                            name: "ras_v".into(),
                            byref: false,
                            ty,
                            is_array: false,
                            optional_default: None,
                        }],
                        ret: None,
                        locals: vec![],
                        body: vec![VStmt::Assign {
                            target: me,
                            value: vname("ras_v"),
                            set: obj,
                        }],
                        leading_comments: vec![],
                        attributes: vec![],
                        source: None,
                    });
                }
            }
        }
    }

    /// An empty procedure with the signature of `f` (interface members and forwarding stubs).
    fn signature_stub(
        &mut self,
        f: FnId,
        env: &HashMap<String, Type>,
        name: String,
        public: bool,
    ) -> VProc {
        let def = &self.prog.fns[f];
        let mut params = Vec::new();
        for p in &def.params {
            let pty = p.ty.subst(env);
            let (ty, is_array) = self.vtype(&pty);
            let byref = p.mode == PassMode::ByRef || is_array;
            params.push(VParam {
                name: self.local_ident(&p.name),
                byref,
                ty,
                is_array,
                optional_default: None,
            });
        }
        let ret_ty = def.ret_ty().subst(env);
        let (rt, rarr) = self.vtype(&ret_ty);
        let (kind, ret) = if ret_ty == Type::Void {
            (ProcKind::Sub, None)
        } else {
            (
                ProcKind::Function,
                Some(if rarr { format!("{rt}()") } else { rt }),
            )
        };
        let (leading_comments, _) = Self::doc_lines(&def.doc);
        VProc {
            kind,
            public,
            friend: false,
            name,
            params,
            ret,
            locals: vec![],
            body: vec![],
            leading_comments,
            attributes: vec![],
            source: None,
        }
    }

    /// Assignment target for a field of class `c` (`obj.field`, bare `field` inside the class, or the static variable).
    pub(crate) fn field_target(
        &mut self,
        c: ClassId,
        idx: usize,
        fty: &Type,
        object: Option<VExpr>,
    ) -> VExpr {
        let def = &self.prog.classes[c];
        let fld = &def.fields[idx];
        if def.kind == ClassKind::Host {
            // A property of a host object keeps the type library's spelling, reserved word or not (`ws.Type`).
            let fname = fld.emitted_name().to_string();
            return match object {
                Some(o) => o.member(fname),
                None => VExpr::Name(fname),
            };
        }
        let fname = names::member_ident(&fld.name);
        if fld.is_static {
            let sname = format!("{}_{}", names::member_ident(&def.name), fld.name);
            let md = &self.prog.modules[def.module];
            return VExpr::Name(md.vba_name.clone()).member(sname);
        }
        let public_array = fld.vis == Visibility::Public && self.vtype(fty).1;
        match object {
            // Inside the class the (backing) field is addressed directly.
            None => VExpr::Name(if public_array {
                format!("f_{fname}")
            } else {
                fname
            }),
            Some(o) if o.is_lit("Me") => VExpr::Name(if public_array {
                format!("f_{fname}")
            } else {
                fname
            }),
            Some(o) => o.member(fname),
        }
    }

    // ------------------------------------------------------------------ events and init glue (spec §12)

    /// Event procedures of the sheets and of the workbook (spec §12.2): the handlers of
    /// `objects/<CodeName>.ras` (`@on("Worksheet.Change")`, `@on("CommandButton1.Click")`) and of
    /// `objects/ThisWorkbook.ras` (`@on("Workbook.Open")`) are called from the event procedure
    /// generated in the same module. `ThisWorkbook` also gets the calls that make sure the
    /// project is initialised (spec §12.3), file or not.
    fn emit_events(&mut self) {
        let prog = self.prog;
        // `@on` belongs to the file of what raises the event.
        for f in &prog.fns {
            if prog.modules[f.module].form.is_some()
                || prog.modules[f.module].kind != ModuleKind::User
            {
                continue;
            }
            // (Methods of a class: `emit_with_events`.)
            if !matches!(f.kind, FnKind::Free) {
                continue;
            }
            for ev in &f.attrs.on_events {
                if let Some(EventSource::Var(v)) = prog.event_source(f.id, &ev.target) {
                    self.diag(Severity::Error, f.module, ev.span, "E0610", format!("`@on(\"{}.{}\")`: 標準モジュールの変数 `{}` のイベントは受け取れません（VBA の `WithEvents` は、標準モジュールに置けません）。変数とハンドラを、クラスのフィールドとメソッドにするか、`objects/` のシート・`ThisWorkbook` のファイルに移してください", ev.target, ev.event, prog.vars[v].name));
                    continue;
                }
                let wb = ev.target.eq_ignore_ascii_case("Workbook")
                    || ev.target.eq_ignore_ascii_case("ThisWorkbook");
                let file = if wb {
                    "objects/ThisWorkbook.ras".to_string()
                } else if ev.target.eq_ignore_ascii_case("Worksheet") {
                    "objects/<シートのコード名>.ras".to_string()
                } else {
                    format!("objects/{}.ras", ev.target)
                };
                let on = if wb {
                    format!("Workbook.{}", ev.event)
                } else {
                    format!("Worksheet.{}", ev.event)
                };
                self.diag(Severity::Error, f.module, ev.span, "E0610", format!("`@on(\"{}.{}\")`: イベントハンドラは、そのイベントを起こすシート・ブックのファイルに書きます。`{file}` に `@on(\"{on}\")` として書いてください（`src/` のファイルは標準モジュールで、イベントを受け取れません）", ev.target, ev.event));
            }
        }
        // (module, object, event procedure) → handlers in declaration order
        struct Proc {
            module: ModuleId,
            object: String,
            event: String,
            params: Vec<(String, String, bool)>,
            handlers: Vec<(FnId, bool)>,
        }
        let mut table: Vec<Proc> = Vec::new();
        // (module, object) of the shapes and form controls whose macro is generated: public.
        let mut macros: Vec<(ModuleId, String)> = Vec::new();
        let docs: Vec<ModuleId> = (0..prog.modules.len())
            .filter(|&m| prog.modules[m].document().is_some())
            .collect();
        for &m in &docs {
            let md = &prog.modules[m];
            let Some(doc) = md.document() else { continue };
            let is_wb = doc.kind == ObjectKind::Workbook;
            let own = if is_wb { "Workbook" } else { "Worksheet" };
            for &f in &md.fns {
                let d = &prog.fns[f];
                for ev in &d.attrs.on_events {
                    // A variable whose object is listened to: `emit_with_events` below.
                    if prog.event_source(f, &ev.target).is_some() {
                        continue;
                    }
                    let other = if is_wb { "Worksheet" } else { "Workbook" };
                    if ev.target.eq_ignore_ascii_case(other)
                        || (!is_wb && ev.target.eq_ignore_ascii_case("ThisWorkbook"))
                    {
                        let msg = if is_wb {
                            format!("`@on(\"{}.{}\")`: シートのイベントは、そのシートのファイル（`objects/<シートのコード名>.ras`）に書きます。すべてのシートのイベントをブックで受け取るには `@on(\"Workbook.Sheet{}\")` を使います", ev.target, ev.event, ev.event)
                        } else {
                            format!("`@on(\"{}.{}\")`: ブックのイベントは `objects/ThisWorkbook.ras` に書きます", ev.target, ev.event)
                        };
                        self.diag(Severity::Error, m, ev.span, "E0610", msg);
                        continue;
                    }
                    let own_event = ev.target.eq_ignore_ascii_case(own)
                        || (is_wb && ev.target.eq_ignore_ascii_case("ThisWorkbook"))
                        || (!is_wb && ev.target.to_lowercase() == md.vba_name.to_lowercase());
                    let (object, event, params): (String, String, Vec<(String, String, bool)>) =
                        if own_event {
                            match document_event_sig(is_wb, &ev.event) {
                                Some((name, sig)) => (
                                    own.to_string(),
                                    name.to_string(),
                                    sig.into_iter()
                                        .map(|p| (p.0.to_string(), p.1.to_string(), p.2))
                                        .collect(),
                                ),
                                None => {
                                    self.diag(
                                        Severity::Error,
                                        m,
                                        ev.span,
                                        "E0611",
                                        format!(
                                            "`@on(\"{}.{}\")`: イベント `{}` には対応していません",
                                            ev.target, ev.event, ev.event
                                        ),
                                    );
                                    continue;
                                }
                            }
                        } else if let Some(name) = self.sheet_shape(doc, &ev.target, &ev.event) {
                            // A shape or a form control: the macro its OnAction runs, public
                            // and named as Excel names a new one (`ボタン1_Click`).
                            macros.push((m, name.clone()));
                            (name, "Click".to_string(), Vec::new())
                        } else if let Some(name) = self.untyped_control(doc, &ev.target) {
                            // A control whose events are not known (not an MSForms control, or no
                            // workbook to read the controls from): the event procedure is declared
                            // as the handler is (spec §12.2.1).
                            let params = self
                                .fn_procs
                                .get(&f)
                                .and_then(|procs| procs.first())
                                .and_then(|(mi, proc)| {
                                    self.out[*mi].procs.iter().find(|p| &p.name == proc)
                                })
                                .map(|p| {
                                    p.params
                                        .iter()
                                        .map(|x| (x.name.clone(), x.ty.clone(), x.byref))
                                        .collect()
                                })
                                .unwrap_or_default();
                            (name, ev.event.clone(), params)
                        } else if let Some(control) = doc.decl.as_ref().and_then(|x| {
                            x.controls
                                .iter()
                                .find(|c| c.name.eq_ignore_ascii_case(&ev.target))
                        }) {
                            match sheet_control_event(&control.control_type, &ev.event) {
                                Some((name, params)) => (control.name.clone(), name, params),
                                None => {
                                    let known = sheet_control_events(&control.control_type);
                                    let hint = if known.is_empty() {
                                        format!("{} のイベントは扱えません", control.control_type)
                                    } else {
                                        format!("使えるイベント: {}", known.join(", "))
                                    };
                                    self.diag(Severity::Error, m, ev.span, "E0611", format!("`@on(\"{}.{}\")`: {} にイベント `{}` はありません（{hint}）", ev.target, ev.event, control.name, ev.event));
                                    continue;
                                }
                            }
                        } else {
                            let what = if is_wb {
                                "ブック".to_string()
                            } else {
                                format!("シート `{}` ", md.vba_name)
                            };
                            self.diag(Severity::Error, m, ev.span, "E0613", format!("`@on(\"{}.{}\")`: {what}にコントロール `{}` がありません（{}自身のイベントは `{own}.イベント`）", ev.target, ev.event, ev.target, if is_wb { "ブック" } else { "シート" }));
                            continue;
                        };
                    if d.params.len() > params.len() {
                        self.diag(Severity::Error, m, ev.span, "E0612", format!("イベント `{object}_{event}` の引数は {} 個ですが、ハンドラ `{}` は {} 個受け取ります", params.len(), d.name, d.params.len()));
                        continue;
                    }
                    match table
                        .iter_mut()
                        .find(|t| t.module == m && t.object == object && t.event == event)
                    {
                        Some(t) => t.handlers.push((f, ev.reentrant)),
                        None => table.push(Proc {
                            module: m,
                            object,
                            event,
                            params,
                            handlers: vec![(f, ev.reentrant)],
                        }),
                    }
                }
            }
        }
        for &m in &docs {
            let vi = self.std_vmodule(m);
            let name = prog.modules[m].vba_name.clone();
            let handlers = prog.modules[m].fns.clone();
            self.emit_with_events(vi, &name, &handlers, true);
        }
        let needs_init = !self.init_calls.is_empty()
            || !table.is_empty()
            || docs.iter().any(|&m| {
                prog.modules[m]
                    .fns
                    .iter()
                    .any(|&f| !prog.fns[f].attrs.on_events.is_empty())
            });
        // Initialisation must run from every entry path (spec §12.3).
        let wb_module = docs.iter().copied().find(|&m| {
            prog.modules[m]
                .document()
                .is_some_and(|d| d.kind == ObjectKind::Workbook)
        });
        // `usize::MAX`: `ThisWorkbook` of a project without `objects/ThisWorkbook.ras`.
        let wb_key = wb_module.unwrap_or(usize::MAX);
        if needs_init {
            for ev in ["Open", "Activate", "WindowActivate"] {
                if !table
                    .iter()
                    .any(|t| t.module == wb_key && t.object == "Workbook" && t.event == ev)
                {
                    let params = document_event_sig(true, ev)
                        .map(|(_, sig)| {
                            sig.into_iter()
                                .map(|p| (p.0.to_string(), p.1.to_string(), p.2))
                                .collect()
                        })
                        .unwrap_or_default();
                    table.push(Proc {
                        module: wb_key,
                        object: "Workbook".into(),
                        event: ev.to_string(),
                        params,
                        handlers: Vec::new(),
                    });
                }
            }
        }
        let mut has_open_handlers = false;
        for t in table {
            let is_wb = t.module == wb_key;
            let (di, doc) = if t.module == usize::MAX {
                (
                    self.vmodule("ThisWorkbook", VModuleKind::Document),
                    "ThisWorkbook".to_string(),
                )
            } else {
                (
                    self.std_vmodule(t.module),
                    prog.modules[t.module].vba_name.clone(),
                )
            };
            let proc_name = format!("{}_{}", t.object, t.event);
            if t.module != usize::MAX {
                // A function of the file named like the event procedure would clash with it.
                for &f in &prog.modules[t.module].fns {
                    let d = &prog.fns[f];
                    if d.proc_ident().eq_ignore_ascii_case(&proc_name) {
                        self.diag(Severity::Error, t.module, d.name_span, "E0614", format!("関数 `{}` の名前はイベント プロシージャ `{proc_name}` と同じです。別の名前にしてください（`@on(\"{}.{}\")` で受け取ります）", d.name, t.object, t.event));
                    }
                }
            }
            let open = is_wb && t.object == "Workbook" && t.event == "Open";
            let mut calls = Vec::new();
            let mut locals = Vec::new();
            for (i, (h, reentrant)) in t.handlers.iter().enumerate() {
                let d = &prog.fns[*h];
                let args: Vec<VExpr> = t
                    .params
                    .iter()
                    .take(d.params.len())
                    .map(|p| vname(p.0.clone()))
                    .collect();
                let call = VStmt::Call {
                    target: VExpr::Name(d.proc_ident()),
                    args,
                };
                if *reentrant || open {
                    calls.push(call);
                } else {
                    // Re-entrancy guard: ignore events raised by the handler itself (spec §12.2).
                    let flag = format!("ras_busy{}", i + 1);
                    locals.push(flag.clone());
                    calls.push(if_then(
                        vname(flag.clone()).not(),
                        vec![
                            assign(vname(flag.clone()), lit("True")),
                            call,
                            assign(vname(flag), lit("False")),
                        ],
                    ));
                }
            }
            let event_name = if t.object == "Workbook" || t.object == "Worksheet" {
                format!("{doc}.{}", t.event)
            } else {
                format!("{doc}.{proc_name}")
            };
            let error = call_stmt(
                "RAS_EventError",
                vec![
                    lit(format!("\"{event_name}\"")),
                    name_member("Err", "Number"),
                    name_member("Err", "Source"),
                    name_member("Err", "Description"),
                ],
            );
            let mut body = Vec::new();
            if open {
                // At once, or when the workbook is activated if it is leaving the protected
                // view (spec §12.3 path 1): `RAS_Glue` calls the handlers back.
                body.push(call_stmt("RAS_WorkbookOpen", vec![]));
                if !calls.is_empty() {
                    has_open_handlers = true;
                    let params = vec![VParam {
                        name: "ras_dummy".into(),
                        byref: false,
                        ty: "Long".into(),
                        is_array: false,
                        optional_default: Some("0&".into()),
                    }];
                    self.out[di].procs.push(VProc {
                        kind: ProcKind::Sub,
                        public: true,
                        friend: false,
                        name: "RAS_WorkbookOpened".into(),
                        params,
                        ret: None,
                        locals: vec![],
                        body: calls,
                        leading_comments: vec![],
                        attributes: vec![],
                        source: None,
                    });
                }
            } else {
                for flag in &locals {
                    body.push(VStmt::Raw(format!("Static {flag} As Boolean")));
                }
                body.push(VStmt::Raw("On Error GoTo RAS_OnError".into()));
                body.push(call_stmt("RAS_EnsureInit", vec![]));
                if is_wb
                    && needs_init
                    && t.object == "Workbook"
                    && (t.event == "Activate" || t.event == "WindowActivate")
                {
                    // What `Workbook_Open` left for now (the workbook was in the protected view).
                    body.push(call_stmt("RAS_OpenDeferred", vec![]));
                }
                body.extend(calls);
                body.push(VStmt::Exit("Sub"));
                body.push(VStmt::Label("RAS_OnError".into()));
                for flag in &locals {
                    body.push(assign(vname(flag.clone()), lit("False")));
                }
                body.push(error);
            }
            let params = t
                .params
                .iter()
                .map(|p| VParam {
                    name: p.0.clone(),
                    byref: p.2,
                    ty: p.1.clone(),
                    is_array: false,
                    optional_default: None,
                })
                .collect();
            let public = macros
                .iter()
                .any(|(mm, o)| *mm == t.module && *o == t.object);
            self.out[di].procs.push(VProc {
                kind: ProcKind::Sub,
                public,
                friend: false,
                name: proc_name,
                params,
                ret: None,
                locals: vec![],
                body,
                leading_comments: vec![],
                attributes: vec![],
                source: None,
            });
        }
        let deferred_open: Vec<VStmt> = if has_open_handlers {
            vec![VStmt::Call {
                target: vname("ThisWorkbook").member("RAS_WorkbookOpened"),
                args: vec![],
            }]
        } else {
            Vec::new()
        };
        if needs_init {
            let ri = self.vmodule("RAS_Glue", VModuleKind::Standard);
            self.out[ri].private_module = true;
            let mut body = vec![
                VStmt::Raw("On Error GoTo RAS_OnError".into()),
                call_stmt("RAS_EnsureInit", vec![]),
            ];
            body.extend(deferred_open);
            body.push(VStmt::Exit("Sub"));
            body.push(VStmt::Label("RAS_OnError".into()));
            body.push(call_stmt(
                "RAS_EventError",
                vec![
                    lit("\"Workbook.Open\""),
                    name_member("Err", "Number"),
                    name_member("Err", "Source"),
                    name_member("Err", "Description"),
                ],
            ));
            let simple = |name: &str, body: Vec<VStmt>| VProc {
                kind: ProcKind::Sub,
                public: true,
                friend: false,
                name: name.into(),
                params: vec![],
                ret: None,
                locals: vec![],
                body,
                leading_comments: vec![],
                attributes: vec![],
                source: None,
            };
            self.out[ri].procs.push(simple("RAS_DeferredInit", body));
            // `Workbook_Open` while the workbook leaves the protected view ("Enable Editing") comes
            // before Excel is ready for it: calls of the object model fail with error 1004. As
            // Microsoft says to ("Object Model calls may fail from WorkbookOpen event when exiting
            // Protected View"), what the event does then waits for `Workbook_Activate`. Not
            // `Application.OnTime`: Excel opens a workbook again to make a call it was asked for,
            // when the workbook is closed before it, and the call cannot be taken back while the
            // workbook closes (seen with Excel 2021).
            self.out[ri]
                .decls
                .push("Private ras_open_deferred As Boolean".into());
            let open = vec![
                VStmt::Raw("If RAS_InProtectedView() Then".into()),
                VStmt::Raw("    ras_open_deferred = True".into()),
                VStmt::Raw("Else".into()),
                VStmt::Raw("    RAS_DeferredInit".into()),
                VStmt::Raw("End If".into()),
            ];
            self.out[ri].procs.push(simple("RAS_WorkbookOpen", open));
            let later = vec![
                VStmt::Raw("If Not ras_open_deferred Then Exit Sub".into()),
                VStmt::Raw("ras_open_deferred = False".into()),
                VStmt::Raw("RAS_DeferredInit".into()),
            ];
            self.out[ri].procs.push(simple("RAS_OpenDeferred", later));
            // Late-bound: Excel 2007 has no protected view (and no `ProtectedViewWindows`).
            let in_view = vec![
                VStmt::Raw("Dim app As Object".into()),
                VStmt::Raw("Dim w As Object".into()),
                VStmt::Raw("On Error Resume Next".into()),
                VStmt::Raw("Set app = Application".into()),
                VStmt::Raw("Set w = app.ProtectedViewWindows.Item(ThisWorkbook.Name)".into()),
                VStmt::Raw("RAS_InProtectedView = Not w Is Nothing".into()),
            ];
            let mut f = simple("RAS_InProtectedView", in_view);
            f.kind = ProcKind::Function;
            f.public = false;
            f.ret = Some("Boolean".into());
            self.out[ri].procs.push(f);
            // spec §12.3 path 2
            self.out[ri].procs.push(simple(
                "Auto_Open",
                vec![call_stmt("RAS_EnsureInit", vec![])],
            ));
        }
    }

    /// The class whose events a `WithEvents` variable of type `t` receives: one host class that
    /// VBA can declare by name (a late-bound one is an `Object`, which has no events).
    fn event_class(&mut self, t: &Type) -> Option<ClassId> {
        match t {
            Type::Class(c, args)
                if args.is_empty()
                    && self.prog.classes[*c].kind == ClassKind::Host
                    && self.class_name(*c, &[]) != "Object" =>
            {
                Some(*c)
            }
            _ => None,
        }
    }

    /// Finds the variables and fields named by `@on("<name>.<Event>")` (spec §12.2.2) and
    /// checks that VBA can declare them `WithEvents`.
    fn collect_with_events(&mut self) {
        let prog = self.prog;
        for f in &prog.fns {
            if prog.modules[f.module].kind != ModuleKind::User || prog.modules[f.module].is_test() {
                continue;
            }
            for ev in &f.attrs.on_events {
                let Some(source) = prog.event_source(f.id, &ev.target) else {
                    continue;
                };
                if self.with_events.contains(&source) {
                    continue;
                }
                let (name, ty, problem) = match source {
                    EventSource::Var(v) => {
                        let d = &prog.vars[v];
                        // A standard module: reported by `emit_events`.
                        if prog.modules[f.module].form.is_none() {
                            continue;
                        }
                        let problem = if d.is_const {
                            Some("`const` にはできません（代入したオブジェクトのイベントを受け取るため、`let` で宣言します）")
                        } else if d.exported {
                            Some("`export` できません")
                        } else {
                            None
                        };
                        (d.name.clone(), d.ty.clone(), problem)
                    }
                    EventSource::Field(c, i) => {
                        let fld = &prog.classes[c].fields[i];
                        let problem = if !prog.classes[c].type_params.is_empty() {
                            Some("ジェネリック クラスのフィールドにはできません")
                        } else {
                            None
                        };
                        (fld.name.clone(), fld.ty.clone(), problem)
                    }
                };
                let problem = match problem {
                    Some(p) => Some(p.to_string()),
                    None if self.event_class(&ty).is_none() => Some(format!("型 `{}` のイベントは受け取れません（`Excel.Application`・`Excel.Workbook`・`Excel.Worksheet`・`MSForms.CommandButton` のように、イベントを持つクラスの型を 1 つだけ書きます。`object`・`any`・ユニオン型・配列・RAS のクラスは使えません）", prog.show(&ty))),
                    None => None,
                };
                match problem {
                    Some(p) => self.diag(
                        Severity::Error,
                        f.module,
                        ev.span,
                        "E0615",
                        format!(
                            "`@on(\"{}.{}\")`: `{name}` のイベントは受け取れません: {p}",
                            ev.target, ev.event
                        ),
                    ),
                    None => {
                        self.with_events.insert(source);
                    }
                }
            }
        }
    }

    /// Event procedures for the objects that variables (fields) hold (spec §12.2.2): the
    /// `@on("<variable>.<Event>")` handlers among `fns` are called from
    /// `Private Sub <variable>_<Event>(…)`, generated in the module `vi` (named `owner`) that
    /// declares the variable `WithEvents`.
    fn emit_with_events(&mut self, vi: usize, owner: &str, fns: &[FnId], ensure_init: bool) {
        let prog = self.prog;
        struct Proc {
            object: String,
            event: String,
            params: Vec<(String, String, bool)>,
            handlers: Vec<(FnId, bool)>,
        }
        let mut table: Vec<Proc> = Vec::new();
        for &f in fns {
            let d = &prog.fns[f];
            for ev in &d.attrs.on_events {
                let Some(source) = prog.event_source(f, &ev.target) else {
                    continue;
                };
                if !self.with_events.contains(&source) {
                    continue;
                }
                let (object, ty) = match source {
                    EventSource::Var(v) => (
                        names::module_ident(&prog.vars[v].name),
                        prog.vars[v].ty.clone(),
                    ),
                    EventSource::Field(c, i) => (
                        names::member_ident(&prog.classes[c].fields[i].name),
                        prog.classes[c].fields[i].ty.clone(),
                    ),
                };
                let Some(class) = self.event_class(&ty) else {
                    continue;
                };
                let (event, params) = match host_events(prog, class) {
                    Some(events) => match events
                        .into_iter()
                        .find(|e| e.0.eq_ignore_ascii_case(&ev.event))
                    {
                        Some(e) => e,
                        None => {
                            self.diag(Severity::Error, d.module, ev.span, "E0611", format!("`@on(\"{}.{}\")`: `{}` にイベント `{}` はありません（または対応していません）", ev.target, ev.event, prog.class_display(class), ev.event));
                            continue;
                        }
                    },
                    // A class whose events are not known: declared as the handler is.
                    None => {
                        let mut params = Vec::new();
                        for p in &d.params {
                            let (ty, is_array) = self.vtype(&p.ty);
                            params.push((
                                names::member_ident(&p.name),
                                ty,
                                p.mode == PassMode::ByRef || is_array,
                            ));
                        }
                        (ev.event.clone(), params)
                    }
                };
                if d.params.len() > params.len() {
                    self.diag(Severity::Error, d.module, ev.span, "E0612", format!("イベント `{object}_{event}` の引数は {} 個ですが、ハンドラ `{}` は {} 個受け取ります", params.len(), d.name, d.params.len()));
                    continue;
                }
                match table
                    .iter_mut()
                    .find(|t| t.object.eq_ignore_ascii_case(&object) && t.event == event)
                {
                    Some(t) => t.handlers.push((f, ev.reentrant)),
                    None => table.push(Proc {
                        object,
                        event,
                        params,
                        handlers: vec![(f, ev.reentrant)],
                    }),
                }
            }
        }
        for t in table {
            let proc_name = format!("{}_{}", t.object, t.event);
            for &f in fns {
                let d = &prog.fns[f];
                if d.proc_ident().eq_ignore_ascii_case(&proc_name) {
                    self.diag(Severity::Error, d.module, d.name_span, "E0614", format!("関数 `{}` の名前はイベント プロシージャ `{proc_name}` と同じです。別の名前にしてください（`@on(\"{}.{}\")` で受け取ります）", d.name, t.object, t.event));
                }
            }
            let mut calls = Vec::new();
            let mut locals = Vec::new();
            for (i, (h, reentrant)) in t.handlers.iter().enumerate() {
                let d = &prog.fns[*h];
                let args: Vec<VExpr> = t
                    .params
                    .iter()
                    .take(d.params.len())
                    .map(|p| vname(p.0.clone()))
                    .collect();
                let call = VStmt::Call {
                    target: VExpr::Name(d.proc_ident()),
                    args,
                };
                if *reentrant {
                    calls.push(call);
                } else {
                    let flag = format!("ras_busy{}", i + 1);
                    locals.push(flag.clone());
                    calls.push(if_then(
                        vname(flag.clone()).not(),
                        vec![
                            assign(vname(flag.clone()), lit("True")),
                            call,
                            assign(vname(flag), lit("False")),
                        ],
                    ));
                }
            }
            let mut body = Vec::new();
            for flag in &locals {
                body.push(VStmt::Raw(format!("Static {flag} As Boolean")));
            }
            body.push(VStmt::Raw("On Error GoTo RAS_OnError".into()));
            if ensure_init {
                body.push(call_stmt("RAS_EnsureInit", vec![]));
            }
            body.extend(calls);
            body.push(VStmt::Exit("Sub"));
            body.push(VStmt::Label("RAS_OnError".into()));
            for flag in &locals {
                body.push(assign(vname(flag.clone()), lit("False")));
            }
            body.push(call_stmt(
                "RAS_EventError",
                vec![
                    lit(format!("\"{owner}.{proc_name}\"")),
                    name_member("Err", "Number"),
                    name_member("Err", "Source"),
                    name_member("Err", "Description"),
                ],
            ));
            let params = t
                .params
                .iter()
                .map(|p| VParam {
                    name: p.0.clone(),
                    byref: p.2,
                    ty: p.1.clone(),
                    is_array: false,
                    optional_default: None,
                })
                .collect();
            self.out[vi].procs.push(VProc {
                kind: ProcKind::Sub,
                public: false,
                friend: false,
                name: proc_name,
                params,
                ret: None,
                locals: vec![],
                body,
                leading_comments: vec![],
                attributes: vec![],
                source: None,
            });
        }
    }

    /// The shape or form control `target` of a sheet whose macro (`OnAction`) an
    /// `@on("<name>.Click")` handler is called from (spec §12.2.1): a name that is not an ActiveX
    /// control of the sheet (one that is no VBA name, `ボタン 1`, whether the controls are known
    /// or not). The object of its macro, `ボタン1` (the macro is `ボタン1_Click`).
    fn sheet_shape(&self, doc: &FormModule, target: &str, event: &str) -> Option<String> {
        if doc.kind != ObjectKind::Sheet || !event.eq_ignore_ascii_case("Click") {
            return None;
        }
        let control = doc.decl.as_ref().is_some_and(|x| {
            x.controls
                .iter()
                .any(|c| c.name.eq_ignore_ascii_case(target))
        });
        let vba_name = !target.is_empty()
            && target.chars().next().is_some_and(|c| c.is_alphabetic())
            && target.chars().all(|c| c.is_alphanumeric() || c == '_');
        if control || (vba_name && !self.cfg.sheet_controls_known) {
            return None;
        }
        Some(shape_macro_object(target))
    }

    /// The name of the control `target` of a sheet when its events are not known: a control
    /// that is not an MSForms one, or any name when the controls of the workbook were not read.
    fn untyped_control(&self, doc: &FormModule, target: &str) -> Option<String> {
        if doc.kind != ObjectKind::Sheet {
            return None;
        }
        match doc.decl.as_ref().and_then(|x| {
            x.controls
                .iter()
                .find(|c| c.name.eq_ignore_ascii_case(target))
        }) {
            Some(c) if crate::userform::events_of(&c.control_type).is_empty() => {
                Some(c.name.clone())
            }
            Some(_) => None,
            None if !self.cfg.sheet_controls_known => Some(target.to_string()),
            None => None,
        }
    }

    /// Event procedures of a form (spec §14.11.2): `@on("btnOK.Click")` handlers are called from
    /// `Private Sub btnOK_Click()`, generated in the form's code module with the same re-entrancy
    /// guard and error dialog as document-module events (spec §12.2). `UserForm_Initialize`
    /// also runs the initialisers of the form's variables.
    fn emit_form_events(&mut self, m: ModuleId) {
        let prog = self.prog;
        let md = &prog.modules[m];
        let Some(form) = &md.form else { return };
        let Some(decl) = &form.decl else { return };
        let fi = self.std_vmodule(m);
        // (object, event) → handlers in declaration order.
        let mut table: Vec<(
            (String, &'static crate::userform::EventDef),
            Vec<(FnId, bool)>,
        )> = Vec::new();
        for &f in &md.fns {
            let d = &prog.fns[f];
            for ev in &d.attrs.on_events {
                // A variable whose object is listened to (spec §12.2.2): `emit_with_events`.
                if prog.event_source(f, &ev.target).is_some() {
                    continue;
                }
                let host_type = if ev.target.eq_ignore_ascii_case("UserForm")
                    || ev.target.eq_ignore_ascii_case(&decl.name)
                {
                    Some(("UserForm".to_string(), "UserForm".to_string()))
                } else {
                    decl.controls
                        .iter()
                        .find(|c| c.name.eq_ignore_ascii_case(&ev.target))
                        .map(|c| (c.control_type.clone(), c.name.clone()))
                };
                let Some((host_type, object)) = host_type else {
                    self.diag(Severity::Error, m, ev.span, "E0613", format!("`@on(\"{}.{}\")`: フォーム `{}` にコントロール `{}` がありません（フォーム自身のイベントは `UserForm.イベント`）", ev.target, ev.event, decl.name, ev.target));
                    continue;
                };
                let Some(def) = crate::userform::find_event(&host_type, &ev.event) else {
                    let known: Vec<&str> = crate::userform::events_of(&host_type)
                        .iter()
                        .map(|e| e.name.as_str())
                        .collect();
                    let hint = if known.is_empty() {
                        if host_type.eq_ignore_ascii_case("Page") {
                            "ページのイベントは親の MultiPage で受け取ります".to_string()
                        } else {
                            format!("{host_type} のイベントは扱えません")
                        }
                    } else {
                        format!("使えるイベント: {}", known.join(", "))
                    };
                    self.diag(
                        Severity::Error,
                        m,
                        ev.span,
                        "E0611",
                        format!(
                            "`@on(\"{}.{}\")`: {} にイベント `{}` はありません（{hint}）",
                            ev.target, ev.event, object, ev.event
                        ),
                    );
                    continue;
                };
                if d.params.len() > def.params.len() {
                    self.diag(Severity::Error, m, ev.span, "E0612", format!("イベント `{object}_{}` の引数は {} 個（{}）ですが、ハンドラ `{}` は {} 個受け取ります", def.name, def.params.len(), def.ras_params(), d.name, d.params.len()));
                    continue;
                }
                let key = (object, def);
                match table
                    .iter_mut()
                    .find(|t| t.0 .0 == key.0 && t.0 .1.name == key.1.name)
                {
                    Some(t) => t.1.push((f, ev.reentrant)),
                    None => table.push((key, vec![(f, ev.reentrant)])),
                }
            }
        }
        let handlers = md.fns.clone();
        self.emit_with_events(fi, &decl.name, &handlers, false);
        let init_fn = md.init_fn;
        if init_fn.is_some()
            && !table
                .iter()
                .any(|t| t.0 .0 == "UserForm" && t.0 .1.name == "Initialize")
        {
            if let Some(def) = crate::userform::find_event("UserForm", "Initialize") {
                table.insert(0, (("UserForm".to_string(), def), Vec::new()));
            }
        }
        // The handlers the form's code attached with `on` go with the form (spec §12.2.3).
        let listens = self.form_listens.contains(&m);
        if listens
            && !table
                .iter()
                .any(|t| t.0 .0 == "UserForm" && t.0 .1.name == "Terminate")
        {
            if let Some(def) = crate::userform::find_event("UserForm", "Terminate") {
                table.push((("UserForm".to_string(), def), Vec::new()));
            }
        }
        for ((object, def), handlers) in table {
            let proc_name = format!("{object}_{}", def.name);
            // A function of the code-behind named like the event procedure would clash with it.
            for &f in &md.fns {
                let d = &prog.fns[f];
                if d.proc_ident().eq_ignore_ascii_case(&proc_name) {
                    self.diag(Severity::Error, m, d.name_span, "E0614", format!("関数 `{}` の名前はイベント プロシージャ `{proc_name}` と同じです。別の名前にしてください（`@on(\"{object}.{}\")` で受け取ります）", d.name, def.name));
                }
            }
            let mut body = Vec::new();
            let mut calls = Vec::new();
            let mut locals = Vec::new();
            if object == "UserForm" && def.name == "Initialize" {
                if let Some(f) = init_fn {
                    calls.push(call_stmt(prog.fns[f].proc_ident(), vec![]));
                }
            }
            for (i, (h, reentrant)) in handlers.iter().enumerate() {
                let d = &prog.fns[*h];
                let args: Vec<VExpr> = def
                    .params
                    .iter()
                    .take(d.params.len())
                    .map(|p| vname(p.name.clone()))
                    .collect();
                let call = VStmt::Call {
                    target: VExpr::Name(d.proc_ident()),
                    args,
                };
                if *reentrant {
                    calls.push(call);
                } else {
                    // Re-entrancy guard: ignore events raised by the handler itself (spec §12.2).
                    let flag = format!("ras_busy{}", i + 1);
                    locals.push(flag.clone());
                    calls.push(if_then(
                        vname(flag.clone()).not(),
                        vec![
                            assign(vname(flag.clone()), lit("True")),
                            call,
                            assign(vname(flag), lit("False")),
                        ],
                    ));
                }
            }
            for flag in &locals {
                body.push(VStmt::Raw(format!("Static {flag} As Boolean")));
            }
            let unlisten = listens && object == "UserForm" && def.name == "Terminate";
            body.push(VStmt::Raw("On Error GoTo RAS_OnError".into()));
            body.extend(calls);
            if unlisten {
                body.push(call_stmt("RAS_UnlistenOwner", vec![lit("Me")]));
            }
            body.push(VStmt::Exit("Sub"));
            body.push(VStmt::Label("RAS_OnError".into()));
            for flag in &locals {
                body.push(assign(vname(flag.clone()), lit("False")));
            }
            if unlisten {
                body.push(call_stmt("RAS_UnlistenOwner", vec![lit("Me")]));
            }
            body.push(call_stmt(
                "RAS_EventError",
                vec![
                    lit(format!("\"{}.{proc_name}\"", decl.name)),
                    name_member("Err", "Number"),
                    name_member("Err", "Source"),
                    name_member("Err", "Description"),
                ],
            ));
            let params = def
                .params
                .iter()
                .map(|p| VParam {
                    name: p.name.clone(),
                    byref: p.byref,
                    ty: p.vba_type.clone(),
                    is_array: false,
                    optional_default: None,
                })
                .collect();
            self.out[fi].procs.push(VProc {
                kind: ProcKind::Sub,
                public: false,
                friend: false,
                name: proc_name,
                params,
                ret: None,
                locals: vec![],
                body,
                leading_comments: vec![],
                attributes: vec![],
                source: None,
            });
        }
    }

    fn emit_init_glue(&mut self) {
        // `RAS_InitAll` is called once by the runtime's `RAS_EnsureInit`.
        let ri = self.vmodule("RAS_Glue", VModuleKind::Standard);
        self.out[ri].private_module = true;
        let mut body: Vec<VStmt> = Vec::new();
        for c in std::mem::take(&mut self.init_calls) {
            let (m, p) = c.split_once('.').unwrap();
            body.push(VStmt::Call {
                target: vname(m).member(p),
                args: vec![],
            });
        }
        if body.is_empty() && self.out[ri].procs.is_empty() {
            // Nothing to initialise: the runtime gets an empty stub only if something refers to it.
            self.index.remove("RAS_Glue");
            self.out.remove(ri);
            for v in self.index.values_mut() {
                if *v > ri {
                    *v -= 1;
                }
            }
            return;
        }
        self.out[ri].procs.insert(
            0,
            VProc {
                kind: ProcKind::Sub,
                public: true,
                friend: false,
                name: "RAS_InitAll".into(),
                params: vec![VParam {
                    name: "ras_dummy".into(),
                    byref: false,
                    ty: "Long".into(),
                    is_array: false,
                    optional_default: Some("0&".into()),
                }],
                ret: None,
                locals: vec![],
                body,
                leading_comments: vec![],
                attributes: vec![],
                source: None,
            },
        );
    }
}

pub(crate) fn name_member(obj: &str, member: &str) -> VExpr {
    VExpr::Name(obj.to_string()).member(member)
}

/// VBA type used in a `Declare` statement.
fn decl_scalar(t: &Type, ptr: &str) -> String {
    match t {
        Type::Byte => "Byte".into(),
        Type::Short => "Integer".into(),
        Type::Int | Type::Enum(_) => "Long".into(),
        Type::Long => "LongLong".into(),
        Type::Float => "Single".into(),
        Type::Double => "Double".into(),
        Type::Currency => "Currency".into(),
        Type::Ptr | Type::WStr => ptr.into(),
        Type::Date => "Date".into(),
        Type::Str | Type::AStr => "String".into(),
        Type::Bool => "Boolean".into(),
        Type::Object | Type::Class(..) => "Object".into(),
        _ => "Any".into(),
    }
}

/// Wraps a declaration that mentions `LongPtr` in `#If VBA7` when VBA6 output is enabled (spec §11.3).
pub(crate) fn ptr_conditional(decl: &str, vba6: bool) -> String {
    if vba6 && decl.contains("LongPtr") {
        format!(
            "#If VBA7 Then\n{decl}\n#Else\n{}\n#End If",
            decl.replace("LongPtr", "Long")
        )
    } else if decl.contains("LongLong") {
        format!("#If Win64 Then\n{decl}\n#End If")
    } else {
        decl.to_string()
    }
}

#[allow(dead_code)]
fn _keep(_: &ast::Block, _: ExprId) {}

/// The object part of the macro generated for a shape or form control named `name` (spec
/// §12.2.1), as Excel names a new macro: its letters, digits and `_` (`ボタン 1` → `ボタン1`),
/// `Shape` before one that does not start with a letter.
pub fn shape_macro_object(name: &str) -> String {
    let kept: String = name
        .chars()
        .filter(|c| c.is_alphanumeric() || *c == '_')
        .collect();
    if kept.chars().next().is_some_and(|c| c.is_alphabetic()) {
        kept
    } else {
        format!("Shape{kept}")
    }
}
