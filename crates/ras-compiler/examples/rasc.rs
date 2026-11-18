//! Compiles the given `.ras` files as one project and prints the generated VBA
//! (`cargo run -p ras-compiler --example rasc -- a.ras b.ras`).
use ras_compiler::{compile_str, render_diagnostics, BuildOptions, Config};

fn main() {
    let mut cfg = Config::default();
    let mut opts = BuildOptions::default();
    let mut files: Vec<(String, String)> = Vec::new();
    for arg in std::env::args().skip(1) {
        match arg.as_str() {
            "-O0" => cfg.opt_override = Some(0),
            "-O1" => cfg.opt_override = Some(1),
            "-O2" => cfg.opt_override = Some(2),
            "--no-erl" => opts.erl = Some(false),
            "--no-doevents" => cfg.doevents.enabled = false,
            s if s.starts_with("--sheet=") => cfg.sheet_code_names.push(s["--sheet=".len()..].to_string()),
            path => {
                let text = std::fs::read_to_string(path).unwrap_or_else(|e| panic!("{path}: {e}"));
                let key = std::path::Path::new(path).file_name().unwrap().to_string_lossy().trim_end_matches(".ras").to_string();
                files.push((key, text));
            }
        }
    }
    let refs: Vec<(&str, &str)> = files.iter().map(|(k, t)| (k.as_str(), t.as_str())).collect();
    let (prog, out) = compile_str(&refs, cfg, &opts);
    let user: Vec<_> = prog.diagnostics.iter().filter(|d| d.module.is_none_or(|m| prog.modules[m].kind == ras_compiler::ModuleKind::User)).cloned().collect();
    eprint!("{}", render_diagnostics(&prog, &user));
    let std_errors = prog.diagnostics.len() - user.len();
    if std_errors > 0 {
        eprintln!("({std_errors} diagnostics in bundled files)");
        if std::env::var_os("RAS_SHOW_STD").is_some() {
            let all: Vec<_> = prog.diagnostics.iter().filter(|d| !user.iter().any(|u| std::ptr::eq(u, *d))).cloned().collect();
            eprint!("{}", render_diagnostics(&prog, &all));
        }
    }
    eprint!("{}", render_diagnostics(&prog, &out.diagnostics));
    for f in &out.files {
        println!("===== {} =====", f.file_name);
        print!("{}", f.code.replace("\r\n", "\n"));
    }
}
