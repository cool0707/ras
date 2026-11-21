//! Language server for RAS (`ras lsp --stdio`, spec §16.2).
//!
//! Transport is `lsp-server`; the protocol payloads are handled as JSON
//! values. Every change re-analyses the project (sources on disk overlaid
//! with the unsaved editor buffers).

use std::collections::{HashMap, HashSet};
use std::error::Error;
use std::path::{Path, PathBuf};

use lsp_server::{Connection, ErrorCode, Message, Notification, Request, RequestId, Response, ResponseError};
use ras_compiler::model::ModuleKind;
use ras_compiler::{hints, query, refs, BuildOptions, BuildOutput, Config, Diagnostic, ModuleId, Program, Severity, SourceInput};
use serde_json::{json, Value};

mod sync;

type Res<T> = Result<T, Box<dyn Error + Send + Sync>>;

const COMMANDS: &[&str] = &["ras.showGeneratedVba", "ras.syncNow", "ras.openInExcel", "ras.runMacro", "ras.showForm", "ras.listMacros", "ras.sheet"];

fn percent_decode(s: &str) -> String {
    let bytes = s.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 3 <= bytes.len() {
            if let Some(v) = std::str::from_utf8(&bytes[i + 1..i + 3]).ok().and_then(|h| u8::from_str_radix(h, 16).ok()) {
                out.push(v);
                i += 3;
                continue;
            }
        }
        out.push(bytes[i]);
        i += 1;
    }
    String::from_utf8_lossy(&out).into_owned()
}

/// `file:///c%3A/dir/a.ras` → `c:/dir/a.ras`; `file:///home/a.ras` → `/home/a.ras`.
pub fn uri_to_path(uri: &str) -> Option<PathBuf> {
    let rest = uri.strip_prefix("file://")?;
    let rest = rest.strip_prefix("localhost").unwrap_or(rest);
    let decoded = percent_decode(rest);
    let bytes = decoded.as_bytes();
    // Windows drive letter: `/c:/...`
    if bytes.len() >= 3 && bytes[0] == b'/' && bytes[1].is_ascii_alphabetic() && bytes[2] == b':' {
        return Some(PathBuf::from(&decoded[1..]));
    }
    Some(PathBuf::from(decoded))
}

pub fn path_to_uri(path: &Path) -> String {
    let s = path.to_string_lossy().replace('\\', "/");
    let mut out = String::from("file://");
    if !s.starts_with('/') {
        out.push('/');
    }
    for b in s.bytes() {
        match b {
            b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'/' | b'-' | b'_' | b'.' | b'~' | b':' => out.push(b as char),
            _ => out.push_str(&format!("%{b:02X}")),
        }
    }
    out
}

fn same_file(a: &Path, b: &Path) -> bool {
    let norm = |p: &Path| p.to_string_lossy().replace('\\', "/").to_lowercase();
    a == b || norm(a) == norm(b) || matches!((a.canonicalize(), b.canonicalize()), (Ok(x), Ok(y)) if x == y)
}

/// The project on disk, with the sheet code names of its workbook when the file can be read.
fn load_project(root: &Path) -> Result<(Config, Vec<SourceInput>), String> {
    let (mut config, sources) = ras_compiler::load_project(root)?;
    if let Some(book) = config.project.workbook.as_ref().map(|w| root.join(w)).filter(|b| b.is_file()) {
        let _ = ras_sync::read_workbook(&mut config, root, &book);
    }
    Ok((config, sources))
}

struct Server {
    root: Option<PathBuf>,
    /// Unsaved editor contents: path → (uri as sent by the client, text).
    open: HashMap<PathBuf, (String, String)>,
    prog: Option<Program>,
    build: Option<BuildOutput>,
    /// URIs that currently show diagnostics (to clear them when they are fixed).
    published: HashSet<String>,
}

impl Server {
    fn find_root(&mut self, hint: &Path) {
        if self.root.as_ref().is_some_and(|r| r.join("ras.toml").is_file()) {
            return;
        }
        let mut dir = hint.parent();
        while let Some(d) = dir {
            if d.join("ras.toml").is_file() {
                self.root = Some(d.to_path_buf());
                return;
            }
            dir = d.parent();
        }
    }

    /// Re-analyses the project with the open buffers overlaid.
    fn analyze(&mut self) {
        let (config, mut sources) = match self.root.as_ref().map(|r| load_project(r)) {
            Some(Ok(x)) => x,
            _ => (Config::default(), Vec::new()),
        };
        for (path, (_, text)) in &self.open {
            match sources.iter_mut().find(|s| same_file(&s.path, path)) {
                Some(s) => s.text = text.clone(),
                None => {
                    // A file outside any project (or not saved yet): analysed on its own.
                    let name = path.file_name().map(|n| n.to_string_lossy().to_string()).unwrap_or_default();
                    let is_typedef = name.ends_with(".d.ras");
                    // A new file of `objects/`: the module of a sheet (spec §12.2).
                    let objects = self.root.as_ref().map(|r| r.join(ras_compiler::OBJECTS_DIR));
                    if let (Some(stem), true) = (name.strip_suffix(".ras"), !is_typedef && path.parent().is_some_and(|d| objects.as_ref().is_some_and(|o| same_file(d, o)))) {
                        sources.push(ras_compiler::object_source(path.clone(), stem, text.clone()));
                        continue;
                    }
                    sources.push(SourceInput {
                        path: path.clone(),
                        rel_path: name.clone(),
                        key: name.trim_end_matches(".ras").to_string(),
                        kind: if is_typedef { ModuleKind::TypeDef } else { ModuleKind::User },
                        text: text.clone(),
                    });
                }
            }
        }
        let prog = ras_compiler::analyze(sources, config);
        self.build = Some(ras_compiler::build(&prog, &BuildOptions::default()));
        self.prog = Some(prog);
    }

    fn module_for(&self, path: &Path) -> Option<ModuleId> {
        let prog = self.prog.as_ref()?;
        prog.modules.iter().position(|m| m.kind != ModuleKind::Std && same_file(&m.path, path))
    }

    fn uri_for(&self, prog: &Program, m: ModuleId) -> String {
        let path = &prog.modules[m].path;
        match self.open.iter().find(|(p, _)| same_file(p, path)) {
            Some((_, (uri, _))) => uri.clone(),
            None => path_to_uri(path),
        }
    }

    fn range(prog: &Program, m: ModuleId, span: ras_syntax::TextRange) -> Value {
        let md = &prog.modules[m];
        let (sl, sc) = md.line_col_utf16(span.start().into());
        let (el, ec) = md.line_col_utf16(span.end().into());
        json!({ "start": { "line": sl, "character": sc }, "end": { "line": el, "character": ec } })
    }

    /// "Run in Excel" / "Debug" above each function that can be started on its own. The
    /// commands are the VS Code extension's: `ras.client.runMacro` runs the macro in the
    /// workbook open in Excel (live sync, Windows only), `ras.debug.start` debugs the
    /// function with the JavaScript run time (spec §14.10).
    fn run_lenses(prog: &Program, m: ModuleId, uri: &str) -> Vec<Value> {
        let md = &prog.modules[m];
        let mut lenses = Vec::new();
        // A form's code-behind: one pair at the top that shows the form (spec §14.10.7). Its
        // functions need the form loaded and are not started on their own.
        if let Some(form) = md.form.as_ref().filter(|_| md.is_form()) {
            if form.decl.is_none() {
                return lenses;
            }
            let range = json!({ "start": { "line": 0, "character": 0 }, "end": { "line": 0, "character": 0 } });
            let name = &md.vba_name;
            if cfg!(windows) {
                lenses.push(json!({ "range": range, "command": { "title": "$(play) Excelで実行", "tooltip": format!("{name} を Excel で表示します（ブックへ同期してから表示）"), "command": "ras.client.showForm", "arguments": [name] } }));
            }
            let entry = format!("{name}.show");
            lenses.push(json!({ "range": range, "command": { "title": "$(debug-alt) デバッグ実行", "tooltip": format!("{name} を表示してデバッグ実行します（フォームは VS Code に表示されます）"), "command": "ras.debug.start", "arguments": [{ "entry": entry, "uri": uri }] } }));
            return lenses;
        }
        for &f in &md.fns {
            let d = &prog.fns[f];
            let Some(decl) = d.decl.as_ref().filter(|x| x.body.is_some()) else { continue };
            if !d.is_debug_entry() {
                continue;
            }
            // Above the declaration with its decorators (the first statement for top-level code).
            let range = Self::range(prog, m, ras_syntax::TextRange::empty(decl.span.start()));
            if cfg!(windows) && d.is_macro_entry() {
                let name = format!("{}.{}", md.vba_name, d.proc_ident());
                lenses.push(json!({ "range": range, "command": { "title": "$(play) Excelで実行", "tooltip": format!("{name} を Excel で実行します（ブックへ同期してから実行）"), "command": "ras.client.runMacro", "arguments": [name] } }));
            }
            let entry = format!("{}.{}", md.vba_name, d.name);
            lenses.push(json!({ "range": range, "command": { "title": "$(debug-alt) デバッグ実行", "tooltip": format!("{entry} をデバッグ実行します"), "command": "ras.debug.start", "arguments": [{ "entry": entry, "uri": uri }] } }));
        }
        lenses
    }

    fn publish_diagnostics(&mut self, conn: &Connection) -> Res<()> {
        let Some(prog) = &self.prog else { return Ok(()) };
        let mut by_uri: HashMap<String, Vec<Value>> = HashMap::new();
        let to_json = |d: &Diagnostic, m: ModuleId| {
            let severity = match d.severity {
                Severity::Error => 1,
                Severity::Warning => 2,
                Severity::Info => 3,
                Severity::Hint => 4,
            };
            json!({ "range": Self::range(prog, m, d.span), "severity": severity, "code": d.code, "source": "ras", "message": d.message })
        };
        let build_diags = self.build.iter().flat_map(|b| b.diagnostics.iter());
        for d in prog.diagnostics.iter().chain(build_diags) {
            let Some(m) = d.module else { continue };
            if prog.modules[m].kind == ModuleKind::Std || prog.modules[m].path.starts_with("<stdlib>") {
                continue;
            }
            by_uri.entry(self.uri_for(prog, m)).or_default().push(to_json(d, m));
        }
        // Optimisations that could not be applied are shown as hints (spec §6 rule 13).
        for (m, note) in self.build.iter().flat_map(|b| b.notes.iter()) {
            if !note.applied && prog.modules[*m].kind == ModuleKind::User {
                by_uri.entry(self.uri_for(prog, *m)).or_default().push(json!({
                    "range": Self::range(prog, *m, note.span), "severity": 4, "code": note.kind, "source": "ras", "message": note.message
                }));
            }
        }
        let mut uris: HashSet<String> = by_uri.keys().cloned().collect();
        uris.extend(self.published.iter().cloned());
        for uri in &uris {
            let diagnostics = by_uri.remove(uri).unwrap_or_default();
            conn.sender.send(Message::Notification(Notification::new("textDocument/publishDiagnostics".into(), json!({ "uri": uri, "diagnostics": diagnostics }))))?;
        }
        self.published = uris;
        Ok(())
    }

    /// (module, byte offset) of a `textDocument/position` request.
    fn locate(&self, params: &Value) -> Option<(ModuleId, u32)> {
        let uri = params["textDocument"]["uri"].as_str()?;
        let m = self.module_for(&uri_to_path(uri)?)?;
        let md = &self.prog.as_ref()?.modules[m];
        let line = params["position"]["line"].as_u64()? as u32;
        let ch = params["position"]["character"].as_u64()? as u32;
        Some((m, md.offset_of(line, ch)))
    }

    /// Location (`{ uri, range }`) of a span, or `None` inside the bundled library.
    fn location(&self, prog: &Program, m: ModuleId, span: ras_syntax::TextRange) -> Option<Value> {
        if prog.modules[m].path.starts_with("<stdlib>") {
            return None;
        }
        Some(json!({ "uri": self.uri_for(prog, m), "range": Self::range(prog, m, span) }))
    }

    /// The module of a request's `textDocument`.
    fn doc_module(&self, p: &Value) -> Option<ModuleId> {
        p["textDocument"]["uri"].as_str().and_then(uri_to_path).and_then(|path| self.module_for(&path))
    }

    /// Requests that can fail with a message for the user (rename).
    fn handle_request(&mut self, req: &Request) -> Result<Value, String> {
        let p = &req.params;
        let Some(prog) = &self.prog else { return Ok(Value::Null) };
        match req.method.as_str() {
            "textDocument/prepareRename" => {
                let Some((m, off)) = self.locate(p) else { return Ok(Value::Null) };
                let (span, name) = refs::prepare_rename(prog, m, off)?;
                Ok(json!({ "range": Self::range(prog, m, span), "placeholder": name }))
            }
            "textDocument/rename" => {
                let Some((m, off)) = self.locate(p) else { return Ok(Value::Null) };
                let new = p["newName"].as_str().unwrap_or("").trim();
                let edits = refs::rename(prog, m, off, new)?;
                let mut changes: serde_json::Map<String, Value> = serde_json::Map::new();
                for e in edits {
                    let edit = json!({ "range": Self::range(prog, e.module, e.span), "newText": e.text });
                    changes.entry(self.uri_for(prog, e.module)).or_insert_with(|| json!([])).as_array_mut().unwrap().push(edit);
                }
                Ok(json!({ "changes": changes }))
            }
            _ => Ok(self.handle(req)),
        }
    }

    fn handle(&mut self, req: &Request) -> Value {
        let p = &req.params;
        let Some(prog) = &self.prog else { return Value::Null };
        match req.method.as_str() {
            "textDocument/hover" => {
                let Some((m, off)) = self.locate(p) else { return Value::Null };
                match query::hover(prog, m, off) {
                    Some(h) => json!({ "contents": { "kind": "markdown", "value": h.markdown }, "range": Self::range(prog, m, h.span) }),
                    None => Value::Null,
                }
            }
            "textDocument/completion" => {
                let Some((m, off)) = self.locate(p) else { return json!([]) };
                // Typing `(` or `,` only offers the names of the arguments (spec §3.3); the
                // list is incomplete so that the next keystroke asks for everything again.
                let by_char = p["context"]["triggerKind"] == 2 && matches!(p["context"]["triggerCharacter"].as_str(), Some("(" | ","));
                let list = if by_char { query::named_arg_completions(prog, m, off) } else { query::completions(prog, m, off) };
                let items: Vec<Value> = list
                    .into_iter()
                    .map(|c| {
                        let mut v = json!({ "label": c.label, "kind": c.kind, "detail": c.detail, "sortText": format!("{}{}", c.sort, c.label) });
                        if let Some(doc) = c.doc {
                            v["documentation"] = json!({ "kind": "markdown", "value": doc });
                        }
                        if let Some(t) = c.insert {
                            v["insertText"] = json!(t);
                        }
                        if let Some(t) = c.filter {
                            v["filterText"] = json!(t);
                        }
                        if let Some((span, text)) = c.import {
                            v["additionalTextEdits"] = json!([{ "range": Self::range(prog, m, span), "newText": text }]);
                        }
                        v
                    })
                    .collect();
                json!({ "isIncomplete": by_char, "items": items })
            }
            "textDocument/definition" => {
                let Some((m, off)) = self.locate(p) else { return Value::Null };
                match query::definition(prog, m, off) {
                    Some((tm, span)) if !prog.modules[tm].path.starts_with("<stdlib>") => json!({ "uri": self.uri_for(prog, tm), "range": Self::range(prog, tm, span) }),
                    _ => Value::Null,
                }
            }
            "textDocument/typeDefinition" => {
                let Some((m, off)) = self.locate(p) else { return Value::Null };
                query::type_definition(prog, m, off).and_then(|(tm, span)| self.location(prog, tm, span)).unwrap_or(Value::Null)
            }
            "textDocument/references" => {
                let Some((m, off)) = self.locate(p) else { return json!([]) };
                let with_decl = p["context"]["includeDeclaration"].as_bool().unwrap_or(true);
                let locations: Vec<Value> = refs::references(prog, m, off).iter().filter(|r| with_decl || !r.decl).filter_map(|r| self.location(prog, r.module, r.span)).collect();
                Value::Array(locations)
            }
            "textDocument/documentHighlight" => {
                let Some((m, off)) = self.locate(p) else { return json!([]) };
                let highlights: Vec<Value> = refs::references(prog, m, off)
                    .iter()
                    .filter(|r| r.module == m)
                    // DocumentHighlightKind: 2 = Read, 3 = Write.
                    .map(|r| json!({ "range": Self::range(prog, m, r.span), "kind": if r.decl || r.write { 3 } else { 2 } }))
                    .collect();
                Value::Array(highlights)
            }
            "textDocument/signatureHelp" => {
                let Some((m, off)) = self.locate(p) else { return Value::Null };
                let Some(h) = query::signature_help(prog, m, off) else { return Value::Null };
                let params: Vec<Value> = h
                    .params
                    .iter()
                    .map(|((s, e), doc)| {
                        let mut v = json!({ "label": [s, e] });
                        if let Some(d) = doc {
                            v["documentation"] = json!({ "kind": "markdown", "value": d });
                        }
                        v
                    })
                    .collect();
                // Outside the parameters: none is highlighted.
                let active = h.active.unwrap_or(params.len() as u32);
                let mut sig = json!({ "label": h.label, "parameters": params, "activeParameter": active });
                if let Some(d) = h.doc {
                    sig["documentation"] = json!({ "kind": "markdown", "value": d });
                }
                json!({ "signatures": [sig], "activeSignature": 0, "activeParameter": active })
            }
            "workspace/symbol" => {
                let q = p["query"].as_str().unwrap_or("");
                let list: Vec<Value> = query::workspace_symbols(prog, q)
                    .into_iter()
                    .take(1000)
                    .filter_map(|(m, s, container)| {
                        let mut v = json!({ "name": s.name, "kind": s.kind, "location": self.location(prog, m, s.name_span)? });
                        if let Some(c) = container {
                            v["containerName"] = json!(c);
                        }
                        Some(v)
                    })
                    .collect();
                Value::Array(list)
            }
            "textDocument/inlayHint" => {
                let Some(m) = self.doc_module(p) else { return json!([]) };
                let md = &prog.modules[m];
                let pos = |v: &Value| md.offset_of(v["line"].as_u64().unwrap_or(0) as u32, v["character"].as_u64().unwrap_or(0) as u32);
                let (from, to) = (pos(&p["range"]["start"]), pos(&p["range"]["end"]));
                let list: Vec<Value> = hints::inlay_hints(prog, m, from, to)
                    .into_iter()
                    .map(|h| {
                        let (line, character) = md.line_col_utf16(h.offset);
                        let mut v = json!({ "position": { "line": line, "character": character }, "label": h.label, "kind": h.kind, "paddingLeft": h.pad_left, "paddingRight": h.pad_right });
                        if let Some(t) = h.tooltip {
                            v["tooltip"] = json!(t);
                        }
                        v
                    })
                    .collect();
                Value::Array(list)
            }
            "textDocument/semanticTokens/full" => {
                let Some(m) = self.doc_module(p) else { return Value::Null };
                json!({ "data": hints::semantic_tokens(prog, m) })
            }
            "textDocument/foldingRange" => {
                let Some(m) = self.doc_module(p) else { return json!([]) };
                let list: Vec<Value> = hints::folding_ranges(prog, m)
                    .into_iter()
                    .map(|f| {
                        let mut v = json!({ "startLine": f.start, "endLine": f.end });
                        if let Some(k) = f.kind {
                            v["kind"] = json!(k);
                        }
                        v
                    })
                    .collect();
                Value::Array(list)
            }
            "textDocument/documentSymbol" => {
                let Some(m) = p["textDocument"]["uri"].as_str().and_then(uri_to_path).and_then(|path| self.module_for(&path)) else { return json!([]) };
                fn conv(prog: &Program, m: ModuleId, s: &query::Symbol) -> Value {
                    json!({
                        "name": if s.name.is_empty() { "(無名)" } else { &s.name }, "kind": s.kind, "detail": s.detail,
                        "range": Server::range(prog, m, s.span.cover(s.name_span)), "selectionRange": Server::range(prog, m, s.name_span),
                        "children": s.children.iter().map(|c| conv(prog, m, c)).collect::<Vec<_>>(),
                    })
                }
                Value::Array(query::symbols(prog, m).iter().map(|s| conv(prog, m, s)).collect())
            }
            "textDocument/formatting" => {
                let Some(m) = p["textDocument"]["uri"].as_str().and_then(uri_to_path).and_then(|path| self.module_for(&path)) else { return Value::Null };
                let md = &prog.modules[m];
                // The mechanical fixes (`ras check --fix`: `Cells` → `cells`, W0304) are applied before formatting.
                let mut seen = HashSet::new();
                let fixes: Vec<&ras_compiler::Fix> = prog.fixes.iter().filter(|f| f.module == m && seen.insert(u32::from(f.span.start()))).collect();
                let fixed = ras_compiler::apply_fixes(&md.text, &fixes);
                match ras_syntax::format(&fixed).or(Some(fixed)) {
                    Some(text) if text != *md.text => {
                        let end = md.line_col_utf16(md.text.len() as u32);
                        json!([{ "range": { "start": { "line": 0, "character": 0 }, "end": { "line": end.0, "character": end.1 } }, "newText": text }])
                    }
                    _ => json!([]),
                }
            }
            "textDocument/codeLens" => {
                let Some(m) = p["textDocument"]["uri"].as_str().and_then(uri_to_path).and_then(|path| self.module_for(&path)) else { return json!([]) };
                let mut lenses = Self::run_lenses(prog, m, p["textDocument"]["uri"].as_str().unwrap_or(""));
                lenses.extend(
                    self.build
                        .iter()
                        .flat_map(|b| b.notes.iter())
                        .filter(|(nm, note)| *nm == m && note.applied)
                        .map(|(_, note)| json!({ "range": Self::range(prog, m, note.span), "command": { "title": note.message, "command": "" } })),
                );
                Value::Array(lenses)
            }
            "textDocument/codeAction" => {
                // Quick fixes for the diagnostics that can be corrected mechanically (`cells` for `Cells`).
                let Some(uri) = p["textDocument"]["uri"].as_str() else { return json!([]) };
                let Some(m) = uri_to_path(uri).and_then(|path| self.module_for(&path)) else { return json!([]) };
                let md = &prog.modules[m];
                let pos = |v: &Value| md.offset_of(v["line"].as_u64().unwrap_or(0) as u32, v["character"].as_u64().unwrap_or(0) as u32);
                let (from, to) = (pos(&p["range"]["start"]), pos(&p["range"]["end"]));
                let mut seen = HashSet::new();
                let fixes: Vec<&ras_compiler::Fix> = prog.fixes.iter().filter(|f| f.module == m && seen.insert(u32::from(f.span.start()))).collect();
                let edit = |f: &ras_compiler::Fix| json!({ "range": Self::range(prog, m, f.span), "newText": f.replacement });
                let mut actions: Vec<Value> = fixes
                    .iter()
                    .filter(|f| u32::from(f.span.start()) <= to && from <= u32::from(f.span.end()))
                    .map(|f| json!({ "title": format!("`{}` に修正", f.replacement), "kind": "quickfix", "isPreferred": true, "edit": { "changes": { uri: [edit(f)] } } }))
                    .collect();
                if !actions.is_empty() && fixes.len() > 1 {
                    let all: Vec<Value> = fixes.iter().map(|f| edit(f)).collect();
                    actions.push(json!({ "title": format!("このファイルの同種の警告をすべて修正（{} 箇所）", all.len()), "kind": "quickfix", "edit": { "changes": { uri: all } } }));
                }
                // `source.fixAll` (run on save through `editor.codeActionsOnSave`) applies every mechanical fix of the file.
                let only: Vec<&str> = p["context"]["only"].as_array().map(|a| a.iter().filter_map(|k| k.as_str()).collect()).unwrap_or_default();
                if !only.is_empty() && !only.contains(&"quickfix") {
                    let wants_fix_all = only.iter().any(|k| matches!(*k, "source" | "source.fixAll" | "source.fixAll.ras"));
                    if !wants_fix_all || fixes.is_empty() {
                        return json!([]);
                    }
                    let all: Vec<Value> = fixes.iter().map(|f| edit(f)).collect();
                    return json!([{ "title": format!("機械的に直せる警告をすべて修正（{} 箇所）", all.len()), "kind": "source.fixAll.ras", "edit": { "changes": { uri: all } } }]);
                }
                // Names used without their `import`.
                for fix in query::import_fixes(prog, m, from, to) {
                    let edit = json!({ "range": Self::range(prog, m, fix.span), "newText": fix.text });
                    actions.push(json!({ "title": fix.title, "kind": "quickfix", "edit": { "changes": { uri: [edit] } } }));
                }
                Value::Array(actions)
            }
            "workspace/executeCommand" => self.execute_command(prog, p),
            _ => Value::Null,
        }
    }

    /// The commands that talk to Excel are carried out by the sync thread, which answers the request itself.
    fn sync_job(req: &Request) -> Option<sync::Job> {
        if req.method != "workspace/executeCommand" {
            return None;
        }
        let args = &req.params["arguments"];
        let id = req.id.clone();
        match req.params["command"].as_str()? {
            "ras.syncNow" => Some(sync::Job::Sync { id, overwrite: args[0]["overwrite"].as_bool().unwrap_or(false) }),
            "ras.openInExcel" => Some(sync::Job::Open { id }),
            "ras.runMacro" => Some(sync::Job::Run { id, name: args[0].as_str()?.to_string() }),
            "ras.showForm" => Some(sync::Job::ShowForm { id, name: args[0].as_str()?.to_string() }),
            // `{ action: "add" | "copy" | "remove" | "rename" | "list", … }` (spec §14.16).
            "ras.sheet" => sync::sheet_edit(&args[0]).map(|(edit, force)| sync::Job::Sheet { id, edit, force }),
            _ => None,
        }
    }

    fn execute_command(&self, prog: &Program, p: &Value) -> Value {
        let args = &p["arguments"];
        match p["command"].as_str().unwrap_or("") {
            "ras.showGeneratedVba" => {
                let Some((m, off)) = self.locate(&args[0]) else { return Value::Null };
                let md = &prog.modules[m];
                let Some(build) = &self.build else { return Value::Null };
                let Some(file) = build.files.iter().find(|f| f.name == md.vba_name) else { return Value::Null };
                let line = file.vba_line_for(&md.rel_path, md.line_of(off)).map(|l| l as i64 - 1).unwrap_or(-1);
                json!({ "moduleName": file.name, "text": file.code, "line": line })
            }
            "ras.listMacros" => json!(self.build.as_ref().map(|b| b.macros.clone()).unwrap_or_default()),
            _ => Value::Null,
        }
    }
}

/// Runs the language server on stdin / stdout until the client shuts it down.
pub fn run_stdio() -> Res<()> {
    let (conn, io_threads) = Connection::stdio();
    run(&conn)?;
    drop(conn);
    io_threads.join()?;
    Ok(())
}

fn respond(conn: &Connection, id: RequestId, result: Value) -> Res<()> {
    conn.sender.send(Message::Response(Response { id, result: Some(result), error: None }))?;
    Ok(())
}

pub fn run(conn: &Connection) -> Res<()> {
    let (id, params) = conn.initialize_start()?;
    let root = params["workspaceFolders"][0]["uri"].as_str().or(params["rootUri"].as_str()).and_then(uri_to_path);
    conn.initialize_finish(
        id,
        json!({
            "capabilities": {
                "textDocumentSync": 1,
                "hoverProvider": true,
                "completionProvider": { "triggerCharacters": [".", "(", ","] },
                "definitionProvider": true,
                "typeDefinitionProvider": true,
                "referencesProvider": true,
                "documentHighlightProvider": true,
                "renameProvider": { "prepareProvider": true },
                "signatureHelpProvider": { "triggerCharacters": ["(", ","], "retriggerCharacters": [")"] },
                "documentSymbolProvider": true,
                "workspaceSymbolProvider": true,
                "inlayHintProvider": true,
                "semanticTokensProvider": { "legend": { "tokenTypes": hints::TOKEN_TYPES, "tokenModifiers": hints::TOKEN_MODIFIERS }, "full": true },
                "foldingRangeProvider": true,
                "documentFormattingProvider": true,
                "codeActionProvider": { "codeActionKinds": ["quickfix", "source.fixAll.ras"] },
                "codeLensProvider": { "resolveProvider": false },
                "executeCommandProvider": { "commands": COMMANDS },
            },
            "serverInfo": { "name": "ras", "version": env!("CARGO_PKG_VERSION") },
        }),
    )?;
    let mut server = Server { root, open: HashMap::new(), prog: None, build: None, published: HashSet::new() };
    server.analyze();
    server.publish_diagnostics(conn)?;
    let sender = conn.sender.clone();
    let jobs = sync::spawn(Box::new(move |m| {
        let _ = sender.send(m);
    }));
    if let Some(root) = server.root.clone().filter(|r| r.join("ras.toml").is_file()) {
        let _ = jobs.send(sync::Job::Root(root));
    }

    for msg in &conn.receiver {
        match msg {
            Message::Request(req) => {
                if conn.handle_shutdown(&req)? {
                    return Ok(());
                }
                if let Some(job) = Server::sync_job(&req) {
                    let _ = jobs.send(job);
                    continue;
                }
                match server.handle_request(&req) {
                    Ok(result) => respond(conn, req.id, result)?,
                    Err(message) => {
                        let error = ResponseError { code: ErrorCode::RequestFailed as i32, message, data: None };
                        conn.sender.send(Message::Response(Response { id: req.id, result: None, error: Some(error) }))?;
                    }
                }
            }
            Message::Notification(n) => {
                let p = &n.params;
                let doc_path = p["textDocument"]["uri"].as_str().and_then(uri_to_path);
                let changed = match n.method.as_str() {
                    "textDocument/didOpen" => match (doc_path, p["textDocument"]["uri"].as_str(), p["textDocument"]["text"].as_str()) {
                        (Some(path), Some(uri), Some(text)) => {
                            let before = server.root.clone();
                            server.find_root(&path);
                            if let Some(root) = server.root.clone().filter(|r| Some(r) != before.as_ref()) {
                                let _ = jobs.send(sync::Job::Root(root));
                            }
                            server.open.insert(path, (uri.to_string(), text.to_string()));
                            true
                        }
                        _ => false,
                    },
                    "textDocument/didChange" => match (doc_path, p["textDocument"]["uri"].as_str(), p["contentChanges"].as_array().and_then(|c| c.last()).and_then(|c| c["text"].as_str())) {
                        (Some(path), Some(uri), Some(text)) => {
                            server.open.insert(path, (uri.to_string(), text.to_string()));
                            true
                        }
                        _ => false,
                    },
                    "textDocument/didClose" => {
                        if let Some(path) = doc_path {
                            server.open.remove(&path);
                        }
                        true
                    }
                    "textDocument/didSave" | "workspace/didChangeWatchedFiles" => {
                        let _ = jobs.send(sync::Job::Saved);
                        true
                    }
                    _ => false,
                };
                if changed {
                    server.analyze();
                    server.publish_diagnostics(conn)?;
                }
            }
            Message::Response(_) => {}
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn uri_round_trip() {
        assert_eq!(uri_to_path("file:///home/u/a%20b/%E8%A1%A8.ras").unwrap(), PathBuf::from("/home/u/a b/表.ras"));
        assert_eq!(uri_to_path("file:///c%3A/Users/u/a.ras").unwrap(), PathBuf::from("c:/Users/u/a.ras"));
        assert_eq!(path_to_uri(Path::new("/home/u/a b/表.ras")), "file:///home/u/a%20b/%E8%A1%A8.ras");
        assert_eq!(path_to_uri(Path::new("C:\\Users\\u\\a.ras")), "file:///C:/Users/u/a.ras");
        assert!(uri_to_path("untitled:Untitled-1").is_none());
    }

    /// Drives the server over an in-memory connection, as an editor would.
    #[test]
    fn serves_an_editing_session() {
        let dir = std::env::temp_dir().join(format!("ras-lsp-test-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(dir.join("src")).unwrap();
        std::fs::write(dir.join("ras.toml"), "[project]\nname = \"T\"\n").unwrap();
        let file = dir.join("src").join("main.ras");
        let good = "/** 2 倍にする。 */\nexport function twice(x: int): int {\n    return x * 2;\n}\n\nexport function run() {\n    const words = [\"a\"];\n    msgBox(words.join(\",\") + twice(2));\n}\n\n/** 後で実行する。 */\n@fast\nexport function later() {\n}\n";
        std::fs::write(&file, good).unwrap();
        let uri = path_to_uri(&file);

        let (server_conn, client) = Connection::memory();
        let handle = std::thread::spawn(move || run(&server_conn).map_err(|e| e.to_string()));
        let send = |m: Message| client.sender.send(m).unwrap();
        let next_id = std::cell::Cell::new(0);
        let request = |method: &str, params: Value| -> Value {
            next_id.set(next_id.get() + 1);
            let id = next_id.get();
            send(Message::Request(Request { id: id.into(), method: method.into(), params }));
            loop {
                match client.receiver.recv_timeout(std::time::Duration::from_secs(30)).unwrap() {
                    Message::Response(r) if r.id == id.into() => return r.result.unwrap_or(Value::Null),
                    _ => {}
                }
            }
        };
        let caps = request("initialize", json!({ "rootUri": path_to_uri(&dir), "capabilities": {} }));
        assert_eq!(caps["capabilities"]["hoverProvider"], true);
        send(Message::Notification(Notification::new("initialized".into(), json!({}))));

        // hover on `twice` at its call site shows the signature and the JSDoc
        let hover = request("textDocument/hover", json!({ "textDocument": { "uri": uri }, "position": { "line": 7, "character": 30 } }));
        let text = hover["contents"]["value"].as_str().unwrap();
        assert!(text.contains("function twice(x: int): int") && text.contains("2 倍にする。"), "{text}");

        // go to definition
        let def = request("textDocument/definition", json!({ "textDocument": { "uri": uri }, "position": { "line": 7, "character": 30 } }));
        assert_eq!(def["range"]["start"]["line"], 1);

        // member completion after `words.` while typing (the buffer has a syntax error)
        let broken = good.replace("msgBox(words.join(\",\") + twice(2));", "words.");
        send(Message::Notification(Notification::new("textDocument/didOpen".into(), json!({ "textDocument": { "uri": uri, "languageId": "ras", "version": 1, "text": broken } }))));
        let comp = request("textDocument/completion", json!({ "textDocument": { "uri": uri }, "position": { "line": 7, "character": 10 } }));
        let labels: Vec<&str> = comp["items"].as_array().unwrap().iter().map(|i| i["label"].as_str().unwrap()).collect();
        assert!(labels.contains(&"map") && labels.contains(&"length") && !labels.contains(&"msgBox"), "{labels:?}");

        // host members are completed in camelCase; a collection also offers the array-style methods
        let host = good.replace("msgBox(words.join(\",\") + twice(2));", "worksheets.");
        send(Message::Notification(Notification::new("textDocument/didChange".into(), json!({ "textDocument": { "uri": uri, "version": 2 }, "contentChanges": [{ "text": host }] }))));
        let comp = request("textDocument/completion", json!({ "textDocument": { "uri": uri }, "position": { "line": 7, "character": 15 } }));
        let labels: Vec<&str> = comp["items"].as_array().unwrap().iter().map(|i| i["label"].as_str().unwrap()).collect();
        assert!(labels.contains(&"count") && labels.contains(&"filter") && !labels.contains(&"Count"), "{labels:?}");

        // the type library's spelling is accepted with a warning and a quick fix
        let cased = good.replace("msgBox(", "MsgBox(");
        send(Message::Notification(Notification::new("textDocument/didChange".into(), json!({ "textDocument": { "uri": uri, "version": 2 }, "contentChanges": [{ "text": cased }] }))));
        let actions = request("textDocument/codeAction", json!({ "textDocument": { "uri": uri }, "range": { "start": { "line": 7, "character": 5 }, "end": { "line": 7, "character": 5 } }, "context": { "diagnostics": [] } }));
        assert_eq!(actions[0]["edit"]["changes"][&uri][0]["newText"], "msgBox", "{actions}");
        let hover = request("textDocument/hover", json!({ "textDocument": { "uri": uri }, "position": { "line": 7, "character": 5 } }));
        assert!(hover["contents"]["value"].as_str().unwrap().contains("function msgBox(") && hover["contents"]["value"].as_str().unwrap().contains("// VBA: MsgBox"), "{hover}");

        // diagnostics follow the edits: an error appears and is cleared again
        let with_error = good.replace("return x * 2;", "return \"s\";");
        send(Message::Notification(Notification::new("textDocument/didChange".into(), json!({ "textDocument": { "uri": uri, "version": 2 }, "contentChanges": [{ "text": with_error }] }))));
        let mut saw_error = false;
        send(Message::Notification(Notification::new("textDocument/didChange".into(), json!({ "textDocument": { "uri": uri, "version": 3 }, "contentChanges": [{ "text": good }] }))));
        let symbols = {
            next_id.set(next_id.get() + 1);
            let id = next_id.get();
            send(Message::Request(Request { id: id.into(), method: "textDocument/documentSymbol".into(), params: json!({ "textDocument": { "uri": uri } }) }));
            let mut last_diags: Option<usize> = None;
            loop {
                match client.receiver.recv_timeout(std::time::Duration::from_secs(30)).unwrap() {
                    Message::Notification(n) if n.method == "textDocument/publishDiagnostics" && n.params["uri"] == uri => {
                        let errors = n.params["diagnostics"].as_array().unwrap().iter().filter(|d| d["severity"] == 1).count();
                        saw_error |= errors > 0;
                        last_diags = Some(errors);
                    }
                    Message::Response(r) if r.id == id.into() => {
                        assert_eq!(last_diags, Some(0), "errors must be cleared after the fix");
                        break r.result.unwrap();
                    }
                    _ => {}
                }
            }
        };
        assert!(saw_error, "the type error must have been published");
        let names: Vec<&str> = symbols.as_array().unwrap().iter().map(|s| s["name"].as_str().unwrap()).collect();
        assert_eq!(names, vec!["twice", "run", "later"]);

        // generated VBA for the cursor position, and the macro list
        let vba = request("workspace/executeCommand", json!({ "command": "ras.showGeneratedVba", "arguments": [{ "textDocument": { "uri": uri }, "position": { "line": 2, "character": 4 } }] }));
        assert_eq!(vba["moduleName"], "Main");
        let line = vba["line"].as_i64().unwrap();
        assert!(vba["text"].as_str().unwrap().lines().nth(line as usize).unwrap().contains("twice = x * 2&"), "{vba}");
        let macros = request("workspace/executeCommand", json!({ "command": "ras.listMacros", "arguments": [] }));
        assert_eq!(macros, json!(["Main.run", "Main.later"]));

        // run / debug lenses above the functions that can be started on their own
        let lenses = request("textDocument/codeLens", json!({ "textDocument": { "uri": uri } }));
        let runs: Vec<(u64, String, Value)> = lenses
            .as_array()
            .unwrap()
            .iter()
            .filter(|l| l["command"]["command"] != "")
            .map(|l| (l["range"]["start"]["line"].as_u64().unwrap(), l["command"]["command"].as_str().unwrap().to_string(), l["command"]["arguments"][0].clone()))
            .collect();
        let mut expected = vec![];
        if cfg!(windows) {
            expected.push((5, "ras.client.runMacro".to_string(), json!("Main.run")));
        }
        expected.push((5, "ras.debug.start".to_string(), json!({ "entry": "Main.run", "uri": uri })));
        if cfg!(windows) {
            expected.push((11, "ras.client.runMacro".to_string(), json!("Main.later")));
        }
        expected.push((11, "ras.debug.start".to_string(), json!({ "entry": "Main.later", "uri": uri })));
        assert_eq!(runs, expected, "twice takes an argument, so it gets no lens; later's sits on its decorator: {lenses}");
        let sync = request("workspace/executeCommand", json!({ "command": "ras.syncNow", "arguments": [] }));
        assert_eq!(sync["ok"], false);

        // formatting
        let messy = good.replace("    return x * 2;", "return x * 2;");
        send(Message::Notification(Notification::new("textDocument/didChange".into(), json!({ "textDocument": { "uri": uri, "version": 4 }, "contentChanges": [{ "text": messy }] }))));
        let edits = request("textDocument/formatting", json!({ "textDocument": { "uri": uri }, "options": { "tabSize": 4, "insertSpaces": true } }));
        assert_eq!(edits[0]["newText"].as_str().unwrap(), good);

        // formatting also applies the W0304 fixes, and so does `source.fixAll` on save
        let messy_cased = messy.replace("msgBox(", "MsgBox(");
        send(Message::Notification(Notification::new("textDocument/didChange".into(), json!({ "textDocument": { "uri": uri, "version": 5 }, "contentChanges": [{ "text": messy_cased }] }))));
        let edits = request("textDocument/formatting", json!({ "textDocument": { "uri": uri }, "options": { "tabSize": 4, "insertSpaces": true } }));
        assert_eq!(edits[0]["newText"].as_str().unwrap(), good);
        let fix_all = request("textDocument/codeAction", json!({ "textDocument": { "uri": uri }, "range": { "start": { "line": 0, "character": 0 }, "end": { "line": 0, "character": 0 } }, "context": { "diagnostics": [], "only": ["source.fixAll"] } }));
        assert_eq!(fix_all[0]["kind"], "source.fixAll.ras", "{fix_all}");
        assert_eq!(fix_all[0]["edit"]["changes"][&uri][0]["newText"], "msgBox", "{fix_all}");
        let organize = request("textDocument/codeAction", json!({ "textDocument": { "uri": uri }, "range": { "start": { "line": 0, "character": 0 }, "end": { "line": 0, "character": 0 } }, "context": { "diagnostics": [], "only": ["source.organizeImports"] } }));
        assert_eq!(organize, json!([]));

        request("shutdown", Value::Null);
        send(Message::Notification(Notification::new("exit".into(), Value::Null)));
        handle.join().unwrap().unwrap();
        let _ = std::fs::remove_dir_all(&dir);
    }

    /// Argument names are offered inside an argument list (spec §3.3).
    #[test]
    fn completes_named_arguments() {
        let dir = std::env::temp_dir().join(format!("ras-lsp-named-test-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(dir.join("src")).unwrap();
        std::fs::write(dir.join("ras.toml"), "[project]\nname = \"T\"\n").unwrap();
        let file = dir.join("src").join("main.ras");
        let head = "function label(a: int, b: int = 2, c: string = \"x\"): string {\n    return c;\n}\n\nfunction first<T>(xs: T[], fallback?: T): T {\n    return xs[0];\n}\n\nfunction lastRow(this: Excel.Worksheet, col: int = 1): int {\n    return col;\n}\n\nclass Point {\n    x: double;\n    y: double;\n    constructor(x: double, y: double = 0) {\n        this.x = x;\n        this.y = y;\n    }\n}\n\nexport function run(ws: Excel.Worksheet) {\n    ";
        std::fs::write(&file, format!("{head}\n}}\n")).unwrap();
        let uri = path_to_uri(&file);

        let (server_conn, client) = Connection::memory();
        let handle = std::thread::spawn(move || run(&server_conn).map_err(|e| e.to_string()));
        let send = |m: Message| client.sender.send(m).unwrap();
        let next_id = std::cell::Cell::new(0);
        let request = |method: &str, params: Value| -> Value {
            next_id.set(next_id.get() + 1);
            let id = next_id.get();
            send(Message::Request(Request { id: id.into(), method: method.into(), params }));
            loop {
                match client.receiver.recv_timeout(std::time::Duration::from_secs(30)).unwrap() {
                    Message::Response(r) if r.id == id.into() => return r.result.unwrap_or(Value::Null),
                    _ => {}
                }
            }
        };
        request("initialize", json!({ "rootUri": path_to_uri(&dir), "capabilities": {} }));
        send(Message::Notification(Notification::new("initialized".into(), json!({}))));
        let version = std::cell::Cell::new(0);
        // Completion at the end of `typed` (written in `run`), optionally triggered by a character.
        let complete = |typed: &str, trigger: Option<&str>| -> (Vec<String>, Value) {
            version.set(version.get() + 1);
            let text = format!("{head}{typed}\n}}\n");
            send(Message::Notification(Notification::new("textDocument/didChange".into(), json!({ "textDocument": { "uri": uri, "version": version.get() }, "contentChanges": [{ "text": text }] }))));
            let line = text[..head.len() + typed.len()].matches('\n').count();
            let character = typed.rsplit('\n').next().unwrap().encode_utf16().count() + 4;
            let mut params = json!({ "textDocument": { "uri": uri }, "position": { "line": line, "character": character } });
            if let Some(c) = trigger {
                params["context"] = json!({ "triggerKind": 2, "triggerCharacter": c });
            }
            let comp = request("textDocument/completion", params);
            let labels = comp["items"].as_array().unwrap().iter().map(|i| i["label"].as_str().unwrap().to_string()).collect();
            (labels, comp)
        };
        let named = |labels: &[String]| -> Vec<String> { labels.iter().filter(|l| l.ends_with(':')).cloned().collect() };

        // typing `(` offers only the names, inserted as `name: `
        let (labels, comp) = complete("label(", Some("("));
        assert_eq!(labels, ["a:", "b:", "c:"]);
        assert_eq!(comp["isIncomplete"], true);
        assert_eq!(comp["items"][0]["insertText"], "a: ");
        assert_eq!(comp["items"][0]["filterText"], "a");
        assert!(comp["items"][1]["detail"].as_str().unwrap().contains("省略可"), "{comp}");

        // parameters given positionally are left out; ordinary names are offered too
        let (labels, _) = complete("label(1, ", Some(","));
        assert_eq!(labels, ["b:", "c:"]);
        let (labels, _) = complete("label(1, b", None);
        assert_eq!(named(&labels), ["b:", "c:"]);
        assert!(labels.iter().any(|l| l == "ws"), "{labels:?}");

        // after a named argument only the remaining names can follow
        let (labels, _) = complete("label(c: \"a, b\", ", None);
        assert_eq!(labels, ["a:", "b:"]);
        let (labels, _) = complete("const s = label(a: label(1), ", Some(","));
        assert_eq!(labels, ["b:", "c:"]);

        // the value of a named argument, nested parentheses and other brackets are not argument names
        assert!(named(&complete("label(a: ", None).0).is_empty());
        assert!(complete("label((", Some("(")).0.is_empty());
        assert!(complete("const xs = [1, ", Some(",")).0.is_empty());
        assert!(complete("if (", Some("(")).0.is_empty());

        // extension methods (without `this`), constructors and host functions
        assert_eq!(complete("ws.lastRow(", Some("(")).0, ["col:"]);
        assert_eq!(complete("lastRow(", Some("(")).0, ["col:"]);
        assert_eq!(complete("const n = first<int>([1], ", Some(",")).0, ["fallback:"]);
        assert_eq!(complete("const p = new Point(1, ", Some(",")).0, ["y:"]);
        let (labels, _) = complete("msgBox(\"x\", ", Some(","));
        assert!(labels.iter().any(|l| l == "title:") && !labels.iter().any(|l| l == "prompt:"), "{labels:?}");
        let (labels, _) = complete("ws.range(\"A:A\").find(\"x\", ", Some(","));
        assert!(labels.iter().any(|l| l == "lookAt:") && !labels.iter().any(|l| l == "what:"), "{labels:?}");

        request("shutdown", Value::Null);
        send(Message::Notification(Notification::new("exit".into(), Value::Null)));
        handle.join().unwrap().unwrap();
        let _ = std::fs::remove_dir_all(&dir);
    }

    /// References, highlights, rename, parameter information and the other navigation requests (spec §16.2).
    #[test]
    fn finds_references_and_renames() {
        let dir = std::env::temp_dir().join(format!("ras-lsp-refs-test-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(dir.join("src")).unwrap();
        std::fs::write(dir.join("ras.toml"), "[project]\nname = \"T\"\n").unwrap();
        let util_file = dir.join("src").join("util.ras");
        let main_file = dir.join("src").join("main.ras");
        let util = "/**\n * 2 倍にする。\n * @param x 元の値\n */\nexport function twice(x: int, byref out: int): int {\n    out = x * 2;\n    return out;\n}\n\nexport class Point {\n    x: double = 0;\n}\n";
        let main = "import { twice, Point } from \"./util\";\n\nexport function run() {\n    let r = 0;\n    const a = twice(1, r);\n    const p: Point = new Point();\n    msgBox(twice(a, r) + p.x);\n}\n";
        std::fs::write(&util_file, util).unwrap();
        std::fs::write(&main_file, main).unwrap();
        let (util_uri, main_uri) = (path_to_uri(&util_file), path_to_uri(&main_file));

        let (server_conn, client) = Connection::memory();
        let handle = std::thread::spawn(move || run(&server_conn).map_err(|e| e.to_string()));
        let send = |m: Message| client.sender.send(m).unwrap();
        let next_id = std::cell::Cell::new(0);
        let call = |method: &str, params: Value| -> Response {
            next_id.set(next_id.get() + 1);
            let id = next_id.get();
            send(Message::Request(Request { id: id.into(), method: method.into(), params }));
            loop {
                match client.receiver.recv_timeout(std::time::Duration::from_secs(30)).unwrap() {
                    Message::Response(r) if r.id == id.into() => return r,
                    _ => {}
                }
            }
        };
        let request = |method: &str, params: Value| -> Value { call(method, params).result.unwrap_or(Value::Null) };
        let at = |uri: &str, line: u32, character: u32| json!({ "textDocument": { "uri": uri }, "position": { "line": line, "character": character } });
        let caps = request("initialize", json!({ "rootUri": path_to_uri(&dir), "capabilities": {} }));
        for cap in ["referencesProvider", "documentHighlightProvider", "renameProvider", "signatureHelpProvider", "typeDefinitionProvider", "workspaceSymbolProvider", "inlayHintProvider", "semanticTokensProvider", "foldingRangeProvider"] {
            assert!(!caps["capabilities"][cap].is_null(), "{cap}: {caps}");
        }
        send(Message::Notification(Notification::new("initialized".into(), json!({}))));
        send(Message::Notification(Notification::new("textDocument/didOpen".into(), json!({ "textDocument": { "uri": main_uri, "languageId": "ras", "version": 1, "text": main } }))));

        // references of `twice` from a call in main.ras: the declaration, the import and two calls
        let refs = request("textDocument/references", json!({ "textDocument": { "uri": main_uri }, "position": { "line": 4, "character": 15 }, "context": { "includeDeclaration": true } }));
        let found: Vec<(String, u64)> = refs.as_array().unwrap().iter().map(|r| (r["uri"].as_str().unwrap().to_string(), r["range"]["start"]["line"].as_u64().unwrap())).collect();
        assert_eq!(found, vec![(main_uri.clone(), 0), (main_uri.clone(), 4), (main_uri.clone(), 6), (util_uri.clone(), 4)], "{refs}");
        let refs = request("textDocument/references", json!({ "textDocument": { "uri": main_uri }, "position": { "line": 4, "character": 15 }, "context": { "includeDeclaration": false } }));
        assert_eq!(refs.as_array().unwrap().len(), 3);

        // highlights of `r` in run: the declaration and the byref uses
        let hl = request("textDocument/documentHighlight", at(&main_uri, 3, 8));
        let kinds: Vec<u64> = hl.as_array().unwrap().iter().map(|h| h["kind"].as_u64().unwrap()).collect();
        assert_eq!(kinds, vec![3, 2, 2], "{hl}");

        // rename `twice` across both files
        let prep = request("textDocument/prepareRename", at(&main_uri, 6, 12));
        assert_eq!(prep["placeholder"], "twice", "{prep}");
        let edit = request("textDocument/rename", json!({ "textDocument": { "uri": main_uri }, "position": { "line": 6, "character": 12 }, "newName": "doubled" }));
        assert_eq!(edit["changes"][&main_uri].as_array().unwrap().len(), 3, "{edit}");
        assert_eq!(edit["changes"][&util_uri].as_array().unwrap().len(), 1, "{edit}");
        assert_eq!(edit["changes"][&util_uri][0]["newText"], "doubled");
        // a rename the server refuses comes back as an error with the reason
        let refused = call("textDocument/rename", json!({ "textDocument": { "uri": main_uri }, "position": { "line": 6, "character": 12 }, "newName": "Dim" }));
        let error = refused.error.expect("an error");
        assert!(error.message.contains("予約語"), "{}", error.message);
        let refused = call("textDocument/prepareRename", at(&main_uri, 6, 6));
        assert!(refused.error.is_some(), "msgBox is declared in the type definitions");

        // parameter information inside `twice(a, |`
        let help = request("textDocument/signatureHelp", at(&main_uri, 6, 20));
        assert_eq!(help["signatures"][0]["label"], "twice(x: int, byref out: int): int", "{help}");
        assert_eq!(help["activeParameter"], 1, "{help}");
        assert!(help["signatures"][0]["documentation"]["value"].as_str().unwrap().contains("2 倍にする。"), "{help}");
        assert_eq!(help["signatures"][0]["parameters"][0]["label"], json!([6, 12]));
        assert_eq!(help["signatures"][0]["parameters"][0]["documentation"]["value"], "元の値");

        // type definition of `p` goes to class Point
        let def = request("textDocument/typeDefinition", at(&main_uri, 6, 25));
        assert_eq!(def["uri"], util_uri, "{def}");
        assert_eq!(def["range"]["start"]["line"], 9);

        // workspace symbols
        let syms = request("workspace/symbol", json!({ "query": "twi" }));
        assert_eq!(syms.as_array().unwrap().len(), 1, "{syms}");
        assert_eq!(syms[0]["name"], "twice");

        // inlay hints: the inferred type of `a` and `byref` at both calls
        let hints = request("textDocument/inlayHint", json!({ "textDocument": { "uri": main_uri }, "range": { "start": { "line": 0, "character": 0 }, "end": { "line": 8, "character": 0 } } }));
        let labels: Vec<&str> = hints.as_array().unwrap().iter().map(|h| h["label"].as_str().unwrap()).collect();
        assert_eq!(labels, vec![": int", ": int", "byref", "byref"], "{hints}");

        // a name of another module that is not imported: a quick fix and a completion add the import
        let unimported = "export function run() {\n    const q = new Point();\n    tw\n}\n";
        send(Message::Notification(Notification::new("textDocument/didChange".into(), json!({ "textDocument": { "uri": main_uri, "version": 2 }, "contentChanges": [{ "text": unimported }] }))));
        let actions = request("textDocument/codeAction", json!({ "textDocument": { "uri": main_uri }, "range": { "start": { "line": 1, "character": 20 }, "end": { "line": 1, "character": 20 } }, "context": { "diagnostics": [] } }));
        assert_eq!(actions[0]["title"], "\"./util\" から `Point` をインポート", "{actions}");
        assert_eq!(actions[0]["edit"]["changes"][&main_uri][0]["newText"], "import { Point } from \"./util\";\n");
        let comp = request("textDocument/completion", at(&main_uri, 2, 6));
        let item = comp["items"].as_array().unwrap().iter().find(|i| i["label"] == "twice").expect("twice is offered");
        assert_eq!(item["additionalTextEdits"][0]["newText"], "import { twice } from \"./util\";\n", "{item}");
        send(Message::Notification(Notification::new("textDocument/didChange".into(), json!({ "textDocument": { "uri": main_uri, "version": 3 }, "contentChanges": [{ "text": main }] }))));

        // semantic tokens and folding ranges
        let toks = request("textDocument/semanticTokens/full", json!({ "textDocument": { "uri": main_uri } }));
        assert!(toks["data"].as_array().is_some_and(|d| !d.is_empty() && d.len() % 5 == 0), "{toks}");
        let folds = request("textDocument/foldingRange", json!({ "textDocument": { "uri": main_uri } }));
        assert!(folds.as_array().unwrap().iter().any(|f| f["startLine"] == 2 && f["endLine"] == 6), "{folds}");

        request("shutdown", Value::Null);
        send(Message::Notification(Notification::new("exit".into(), Value::Null)));
        handle.join().unwrap().unwrap();
        let _ = std::fs::remove_dir_all(&dir);
    }

    /// Without the workbook open in Excel the sync thread reports "not connected" and the commands say why.
    #[test]
    fn reports_sync_status_when_the_workbook_is_not_open() {
        let dir = std::env::temp_dir().join(format!("ras-lsp-sync-test-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(dir.join("src")).unwrap();
        std::fs::write(dir.join("ras.toml"), "[project]\nname = \"T\"\nworkbook = \"Book.xlsm\"\n").unwrap();
        std::fs::write(dir.join("src").join("main.ras"), "export function run() {\n}\n").unwrap();

        let (server_conn, client) = Connection::memory();
        let handle = std::thread::spawn(move || run(&server_conn).map_err(|e| e.to_string()));
        let send = |m: Message| client.sender.send(m).unwrap();
        let mut status = None;
        let mut wait_for = |id: i32| loop {
            match client.receiver.recv_timeout(std::time::Duration::from_secs(30)).unwrap() {
                Message::Response(r) if r.id == id.into() => return r.result.unwrap_or(Value::Null),
                Message::Notification(n) if n.method == "ras/syncStatus" => status = Some(n.params),
                _ => {}
            }
        };
        send(Message::Request(Request { id: 1.into(), method: "initialize".into(), params: json!({ "rootUri": path_to_uri(&dir), "capabilities": {} }) }));
        wait_for(1);
        send(Message::Notification(Notification::new("initialized".into(), json!({}))));
        send(Message::Notification(Notification::new("textDocument/didSave".into(), json!({ "textDocument": { "uri": path_to_uri(&dir.join("src").join("main.ras")) } }))));
        send(Message::Request(Request { id: 2.into(), method: "workspace/executeCommand".into(), params: json!({ "command": "ras.syncNow", "arguments": [] }) }));
        let sync = wait_for(2);
        assert_eq!(sync["ok"], false, "{sync}");
        assert!(sync["message"].as_str().is_some_and(|m| !m.is_empty()), "{sync}");
        assert_eq!(sync["offline"], true, "no Excel has the workbook: the editor writes to the file instead: {sync}");
        send(Message::Request(Request { id: 3.into(), method: "workspace/executeCommand".into(), params: json!({ "command": "ras.runMacro", "arguments": ["Main.run"] }) }));
        assert_eq!(wait_for(3)["ok"], false);
        send(Message::Request(Request { id: 4.into(), method: "shutdown".into(), params: Value::Null }));
        wait_for(4);
        let status = status.expect("a ras/syncStatus notification");
        assert_eq!(status["connected"], false, "{status}");
        assert_eq!(status["pending"], 0, "{status}");
        assert!(status["workbook"].as_str().is_some_and(|w| w.ends_with("Book.xlsm")), "{status}");
        send(Message::Notification(Notification::new("exit".into(), Value::Null)));
        handle.join().unwrap().unwrap();
        let _ = std::fs::remove_dir_all(&dir);
    }

    /// A form's code-behind gets one run / debug pair at its top, which shows the form (spec §14.10.7).
    #[test]
    fn form_code_behind_has_a_run_lens_for_the_form() {
        let dir = std::env::temp_dir().join(format!("ras-lsp-form-lens-test-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(dir.join("src")).unwrap();
        std::fs::create_dir_all(dir.join("forms")).unwrap();
        std::fs::write(dir.join("ras.toml"), "[project]\nname = \"T\"\n").unwrap();
        std::fs::write(dir.join("src").join("main.ras"), "export function run() {\n    frmInput.show();\n}\n").unwrap();
        std::fs::write(
            dir.join("forms").join("frmInput.form.json"),
            r#"{ "schemaVersion": 1, "kind": "ras.userform", "form": { "name": "frmInput" }, "controls": [{ "id": "ok", "name": "btnOK", "type": "CommandButton" }] }"#,
        )
        .unwrap();
        let file = dir.join("forms").join("frmInput.ras");
        let text = "@on(\"btnOK.Click\")\nfunction ok() {\n    unload(this);\n}\n\nexport function value(): string {\n    return \"\";\n}\n";
        std::fs::write(&file, text).unwrap();
        let uri = path_to_uri(&file);

        let (server_conn, client) = Connection::memory();
        let handle = std::thread::spawn(move || run(&server_conn).map_err(|e| e.to_string()));
        let send = |m: Message| client.sender.send(m).unwrap();
        let next_id = std::cell::Cell::new(0);
        let request = |method: &str, params: Value| -> Value {
            next_id.set(next_id.get() + 1);
            let id = next_id.get();
            send(Message::Request(Request { id: id.into(), method: method.into(), params }));
            loop {
                match client.receiver.recv_timeout(std::time::Duration::from_secs(30)).unwrap() {
                    Message::Response(r) if r.id == id.into() => return r.result.unwrap_or(Value::Null),
                    _ => {}
                }
            }
        };
        request("initialize", json!({ "rootUri": path_to_uri(&dir), "capabilities": {} }));
        send(Message::Notification(Notification::new("initialized".into(), json!({}))));
        send(Message::Notification(Notification::new("textDocument/didOpen".into(), json!({ "textDocument": { "uri": uri, "languageId": "ras", "version": 1, "text": text } }))));
        let lenses = request("textDocument/codeLens", json!({ "textDocument": { "uri": uri } }));
        let runs: Vec<(u64, String, Value)> = lenses
            .as_array()
            .unwrap()
            .iter()
            .filter(|l| l["command"]["command"] != "")
            .map(|l| (l["range"]["start"]["line"].as_u64().unwrap(), l["command"]["command"].as_str().unwrap().to_string(), l["command"]["arguments"][0].clone()))
            .collect();
        let mut expected = vec![];
        if cfg!(windows) {
            expected.push((0, "ras.client.showForm".to_string(), json!("frmInput")));
        }
        expected.push((0, "ras.debug.start".to_string(), json!({ "entry": "frmInput.show", "uri": uri })));
        assert_eq!(runs, expected, "the exported function of the form gets no lens of its own: {lenses}");
        // The server advertises the command that shows a form in Excel.
        assert!(COMMANDS.contains(&"ras.showForm"));

        request("shutdown", Value::Null);
        send(Message::Notification(Notification::new("exit".into(), Value::Null)));
        handle.join().unwrap().unwrap();
        let _ = std::fs::remove_dir_all(&dir);
    }
}
