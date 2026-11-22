//! Event handlers attached at run time: `target.on("<Event>", handler)` (spec §12.2.3).
//!
//! An object with events is listened to by an instance of a generated class that holds it in
//! a `WithEvents` variable (`RAS_On_<Class>_<Event>`) and calls the handler, a closure object,
//! from the event procedure. The run time keeps the listeners alive. An object with an
//! `OnAction` property gets the name of a macro.

use std::collections::BTreeSet;

use ras_syntax::ast::{Expr, ExprKind};

use super::Lowerer;
use crate::model::*;
use crate::vba::*;

/// A listener class to generate: one per host class and event.
pub(crate) struct ListenerClass {
    pub name: String,
    /// VBA name of the host class (`MSForms.CommandButton`).
    pub vba_type: String,
    /// `<Class>.<Event>`, for the error dialog.
    pub display: String,
    pub event: String,
    /// Parameters of the event: name, VBA type, by reference.
    pub params: Vec<(String, String, bool)>,
    /// What the handlers take: the target first (handlers of the members of an array or a
    /// collection) or not, and how many of the event's arguments.
    pub argcs: BTreeSet<(bool, usize)>,
}

/// Added to the number of arguments a listener is bound with when its handler takes the
/// target before the arguments of the event.
const SENDER: usize = 256;

/// The procedures that listen to an event of a target whose class is known at run time only
/// (a `MSForms.Control`) and of the members of an array or a collection.
pub(crate) struct Dispatcher {
    pub event: String,
    /// `TypeName` of a target → its listener class.
    pub classes: Vec<(String, String)>,
    /// `RAS_OnEach_<Event>` is needed too.
    pub each: bool,
}

impl<'p> Lowerer<'p> {
    /// `target.on("<Event>", handler)`.
    pub(crate) fn event_on(
        &mut self,
        e: &Expr,
        recv: &Expr,
        args: &[Expr],
        want: bool,
    ) -> Option<VExpr> {
        let Some(bind) = self.info().event_binds.get(&e.id).cloned() else {
            // An error of the checker's.
            return want.then(|| lit("Nothing"));
        };
        let target = self.expr(recv);
        let handler_arg = args.get(1)?;
        match bind.kind {
            EventBindKind::OnAction { each } => {
                let f = bind.handler_fn?;
                let proc = format!(
                    "{}.{}",
                    self.prog.modules[self.prog.fns[f].module].vba_name,
                    self.fn_name(f, &[])
                );
                self.force_public.insert(f);
                let macro_name = call("RAS_MacroName", vec![lit(format!("\"{proc}\""))]);
                if each {
                    self.emit(call_stmt("RAS_OnActionEach", vec![target, macro_name]));
                } else {
                    self.emit(assign(target.member("OnAction"), macro_name));
                }
                None
            }
            EventBindKind::Listener {
                classes,
                dynamic,
                each,
                event,
                params,
                argc,
            } => {
                // One class known now: its listener is created here. Otherwise a generated
                // procedure finds the listener class of each target.
                // (`argc` as the listener gets it: `SENDER` is added when the target comes first.)
                let sender = each && argc >= 1;
                let event_argc = argc - usize::from(sender);
                let argc_code = (event_argc + if sender { SENDER } else { 0 }) as i64;
                let class_name = if dynamic || each {
                    self.dispatcher(&classes, &event, event_argc, sender, each)
                } else {
                    self.listener_class(classes[0], &event, &params, event_argc, sender)
                };
                // The form whose code attaches the handler: its listeners go when it is unloaded.
                let lexical = self.frame().module;
                let owner = if self.prog.modules[lexical].is_form() {
                    let me = self
                        .frame()
                        .this
                        .clone()
                        .or_else(|| (!self.pc.in_closure).then(|| lit("Me")));
                    if me.is_some() {
                        self.form_listens.insert(lexical);
                    }
                    me
                } else {
                    None
                };
                let handler = match (bind.handler_fn, &handler_arg.peel().kind) {
                    (Some(f), _) => self.fn_closure(f, argc),
                    (None, ExprKind::Arrow(arrow)) => {
                        let quiet = std::mem::replace(&mut self.quiet_closures, true);
                        let v = self.closure_value(handler_arg.peel(), arrow);
                        self.quiet_closures = quiet;
                        v
                    }
                    (None, _) => self.expr(handler_arg),
                };
                let owner = owner.unwrap_or_else(|| lit("Nothing"));
                if each {
                    self.emit(call_stmt(
                        format!("RAS_OnEach_{event}"),
                        vec![target, handler, long(argc_code), owner],
                    ));
                    return None;
                }
                if dynamic {
                    let argv = vec![target, handler, long(argc_code), owner];
                    if want {
                        return Some(call(class_name, argv));
                    }
                    self.emit(call_stmt(class_name, argv));
                    return None;
                }
                let argv = vec![
                    lit(format!("New {class_name}")),
                    target,
                    handler,
                    long(argc_code),
                    owner,
                ];
                if want {
                    Some(call("RAS_Listen", argv))
                } else {
                    self.emit(call_stmt("RAS_Listen", argv));
                    None
                }
            }
        }
    }

    /// The class that listens to `event` of the host class `class`; it is generated once the
    /// whole program is lowered (`emit_listeners`), for the numbers of arguments in use.
    fn listener_class(
        &mut self,
        class: ClassId,
        event: &str,
        params: &[(String, String, bool)],
        argc: usize,
        sender: bool,
    ) -> String {
        let def = &self.prog.classes[class];
        let name = Self::shorten(format!("RAS_On_{}_{event}", def.name), 31);
        if let Some(l) = self.listeners.iter_mut().find(|l| l.name == name) {
            l.argcs.insert((sender, argc));
            return name;
        }
        let display = format!("{}.{event}", def.name);
        let vba_type = self.class_name(class, &[]);
        self.listeners.push(ListenerClass {
            name: name.clone(),
            vba_type,
            display,
            event: event.to_string(),
            params: params.to_vec(),
            argcs: BTreeSet::from([(sender, argc)]),
        });
        name
    }

    /// `RAS_OnAny_<Event>`, the procedure that listens to `event` of a target of one of
    /// `classes`, whichever it is at run time (`RAS_OnEach_<Event>` does so for the members of
    /// an array or a collection). Returns its name.
    fn dispatcher(
        &mut self,
        classes: &[ClassId],
        event: &str,
        argc: usize,
        sender: bool,
        each: bool,
    ) -> String {
        let prog = self.prog;
        let mut entries = Vec::new();
        for &c in classes {
            // The listener class declares the event as this class does.
            let Some(params) = super::host_events(prog, c)
                .and_then(|events| events.into_iter().find(|e| e.0.eq_ignore_ascii_case(event)))
                .map(|e| e.1)
            else {
                continue;
            };
            let listener = self.listener_class(c, event, &params, argc, sender);
            entries.push((prog.classes[c].name.clone(), listener));
        }
        match self.dispatchers.iter_mut().find(|d| d.event == event) {
            Some(d) => {
                d.each |= each;
                for e in entries {
                    if !d.classes.contains(&e) {
                        d.classes.push(e);
                    }
                }
            }
            None => self.dispatchers.push(Dispatcher {
                event: event.to_string(),
                classes: entries,
                each,
            }),
        }
        format!("RAS_OnAny_{event}")
    }

    /// The procedures of `dispatcher`, in the standard module `RAS_Events`.
    pub(crate) fn emit_dispatchers(&mut self) {
        let dispatchers = std::mem::take(&mut self.dispatchers);
        if dispatchers.is_empty() {
            return;
        }
        let mi = self.vmodule("RAS_Events", VModuleKind::Standard);
        self.out[mi].private_module = true;
        let raw = |s: String| VStmt::Raw(s);
        let by_val = |name: &str, ty: &str| VParam {
            name: name.into(),
            byref: false,
            ty: ty.into(),
            is_array: false,
            optional_default: None,
        };
        for d in dispatchers {
            let any = format!("RAS_OnAny_{}", d.event);
            // `TypeName`, not `TypeOf … Is`: a check box, an option button and a toggle button
            // share their interface.
            let mut body = vec![raw("Select Case TypeName(target)".into())];
            for (type_name, listener) in &d.classes {
                body.push(raw(format!("Case \"{type_name}\"")));
                body.push(raw(format!(
                    "    Set {any} = RAS_Listen(New {listener}, target, handler, argc, owner)"
                )));
            }
            body.push(raw("End Select".into()));
            let params = vec![
                by_val("target", "Object"),
                by_val("handler", "Object"),
                by_val("argc", "Long"),
                by_val("owner", "Object"),
            ];
            self.out[mi].procs.push(VProc {
                kind: ProcKind::Function,
                public: true,
                friend: false,
                name: any.clone(),
                params,
                ret: Some("Object".into()),
                locals: vec![],
                body,
                leading_comments: vec![],
                attributes: vec![],
                source: None,
            });
            if d.each {
                let body = vec![
                    raw("Dim t As Variant".into()),
                    raw("If IsArray(targets) Then".into()),
                    raw("    If RAS_Len(targets) = 0& Then Exit Sub".into()),
                    raw("End If".into()),
                    raw("For Each t In targets".into()),
                    raw(format!("    {any} t, handler, argc, owner")),
                    raw("Next t".into()),
                ];
                let params = vec![
                    VParam {
                        name: "targets".into(),
                        byref: true,
                        ty: "Variant".into(),
                        is_array: false,
                        optional_default: None,
                    },
                    by_val("handler", "Object"),
                    by_val("argc", "Long"),
                    by_val("owner", "Object"),
                ];
                self.out[mi].procs.push(VProc {
                    kind: ProcKind::Sub,
                    public: true,
                    friend: false,
                    name: format!("RAS_OnEach_{}", d.event),
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
    }

    /// A closure object that calls the function `f` with the first `argc` arguments of an
    /// event (`btn.on("Click", ok)`): the same class with an `Apply` method an arrow function
    /// becomes (spec §4.3). A function of a form's code-behind is called on the form the
    /// closure was created in.
    fn fn_closure(&mut self, f: FnId, argc: usize) -> VExpr {
        let def = &self.prog.fns[f];
        let home = &self.prog.modules[def.module];
        self.lambda_n += 1;
        let cname = format!("RAS_Lambda_{:03}", self.lambda_n);
        let pname = self.fn_name(f, &[]);
        if !def.exported {
            self.force_public.insert(f);
        }
        // A function of the form this code is in: called on this instance.
        let me = if home.is_form() && self.frame().module == def.module {
            self.frame()
                .this
                .clone()
                .or_else(|| (!self.pc.in_closure).then(|| lit("Me")))
        } else {
            None
        };
        let callee = match &me {
            Some(_) => lit("Me").member("Captured_Me").member(pname),
            None => name(home.vba_name.clone()).member(pname),
        };
        let params: Vec<VParam> = (0..argc)
            .map(|i| VParam {
                name: format!("a{}", i + 1),
                byref: false,
                ty: "Variant".into(),
                is_array: false,
                optional_default: None,
            })
            .collect();
        let argv: Vec<VExpr> = params.iter().map(|p| name(p.name.clone())).collect();
        let ret = def.ret_ty();
        let body = if ret == Type::Void {
            vec![VStmt::Call {
                target: callee,
                args: argv,
            }]
        } else {
            let value = if argv.is_empty() {
                callee
            } else {
                callee.index(argv)
            };
            vec![VStmt::Assign {
                target: name("Apply"),
                value,
                set: ret.is_object(),
            }]
        };
        let ci = self.vmodule(&cname, VModuleKind::Class);
        let src_path = self.prog.modules[self.src_module()].rel_path.clone();
        self.out[ci].sources.push(src_path);
        if me.is_some() {
            self.out[ci]
                .decls
                .push("Public Captured_Me As Object".into());
        }
        self.out[ci].procs.push(VProc {
            kind: ProcKind::Function,
            public: true,
            friend: false,
            name: "Apply".into(),
            params,
            ret: Some("Variant".into()),
            locals: vec![],
            body,
            leading_comments: vec![],
            attributes: vec![],
            source: None,
        });
        let t = VExpr::Name(self.temp(&Type::Object));
        self.emit(set(t.clone(), lit(format!("New {cname}"))));
        if let Some(me) = me {
            self.emit(set(t.clone().member("Captured_Me"), me));
        }
        t
    }

    /// The listener classes of the `on` calls (spec §12.2.3).
    pub(crate) fn emit_listeners(&mut self) {
        for l in std::mem::take(&mut self.listeners) {
            let ci = self.vmodule(&l.name, VModuleKind::Class);
            let raw = |s: &str| VStmt::Raw(s.to_string());
            self.out[ci].decls.push(
                [
                    format!("Private WithEvents ras_target As {}", l.vba_type),
                    "Private ras_handler As Object".into(),
                    "Private ras_argc As Long".into(),
                    "Private ras_busy As Boolean".into(),
                    "Public RAS_Owner As Object".into(),
                ]
                .join("\n"),
            );
            let by_val = |name: &str, ty: &str| VParam {
                name: name.into(),
                byref: false,
                ty: ty.into(),
                is_array: false,
                optional_default: None,
            };
            let proc = |name: &str, public: bool, params: Vec<VParam>, body: Vec<VStmt>| VProc {
                kind: ProcKind::Sub,
                public,
                friend: false,
                name: name.into(),
                params,
                ret: None,
                locals: vec![],
                body,
                leading_comments: vec![],
                attributes: vec![],
                source: None,
            };
            let bind = vec![
                raw("Set ras_target = target"),
                raw("Set ras_handler = handler"),
                raw("ras_argc = argc"),
                raw("Set RAS_Owner = owner"),
            ];
            self.out[ci].procs.push(proc(
                "RAS_Bind",
                true,
                vec![
                    by_val("target", "Object"),
                    by_val("handler", "Object"),
                    by_val("argc", "Long"),
                    by_val("owner", "Object"),
                ],
                bind,
            ));
            self.out[ci].procs.push(proc(
                "Off",
                true,
                vec![],
                vec![
                    raw("Set ras_target = Nothing"),
                    raw("Set ras_handler = Nothing"),
                    raw("Set RAS_Owner = Nothing"),
                    raw("RAS_Unlisten Me"),
                ],
            ));
            // A handler that returns `false` cancels an event with a `Cancel` argument by reference.
            let cancel = l
                .params
                .iter()
                .find(|p| p.2 && p.1 == "Boolean")
                .map(|p| p.0.clone());
            let apply = |(sender, argc): (bool, usize)| -> String {
                let mut argv: Vec<&str> =
                    l.params.iter().take(argc).map(|p| p.0.as_str()).collect();
                if sender {
                    argv.insert(0, "ras_target");
                }
                match &cancel {
                    Some(c) => format!(
                        "If RAS_Cancels(ras_handler.Apply({})) Then {c} = True",
                        argv.join(", ")
                    ),
                    None if argv.is_empty() => "ras_handler.Apply".to_string(),
                    None => format!("ras_handler.Apply {}", argv.join(", ")),
                }
            };
            let mut body = vec![
                raw("If ras_busy Or ras_handler Is Nothing Then Exit Sub"),
                raw("On Error GoTo RAS_OnError"),
                raw("ras_busy = True"),
            ];
            if l.argcs.len() == 1 {
                body.push(VStmt::Raw(apply(*l.argcs.iter().next().unwrap())));
            } else {
                body.push(raw("Select Case ras_argc"));
                for &key in &l.argcs {
                    body.push(VStmt::Raw(format!(
                        "Case {}&",
                        key.1 + if key.0 { SENDER } else { 0 }
                    )));
                    body.push(VStmt::Raw(format!("    {}", apply(key))));
                }
                body.push(raw("End Select"));
            }
            body.push(raw("ras_busy = False"));
            body.push(VStmt::Exit("Sub"));
            body.push(VStmt::Label("RAS_OnError".into()));
            body.push(raw("ras_busy = False"));
            body.push(call_stmt(
                "RAS_EventError",
                vec![
                    lit(format!("\"{}\"", l.display)),
                    super::name_member("Err", "Number"),
                    super::name_member("Err", "Source"),
                    super::name_member("Err", "Description"),
                ],
            ));
            let params = l
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
            self.out[ci].procs.push(proc(
                &format!("ras_target_{}", l.event),
                false,
                params,
                body,
            ));
        }
    }
}
