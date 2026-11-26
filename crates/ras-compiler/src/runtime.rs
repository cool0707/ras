//! The VBA runtime library (`RAS_Runtime.bas`, `RAS_Error.cls`).
//!
//! Only the procedures that the generated code actually refers to are emitted
//! (tree shaking, spec §7.5). Sections of `runtime.vba` start with
//! `'@@ <kind> <names...>`.

use std::collections::HashSet;

use crate::config::Config;

const SOURCE: &str = include_str!("runtime.vba");

struct Section {
    kind: &'static str,
    names: Vec<&'static str>,
    body: String,
}

fn sections() -> Vec<Section> {
    let mut out: Vec<Section> = Vec::new();
    for line in SOURCE.lines() {
        if let Some(head) = line.strip_prefix("'@@ ") {
            let mut parts = head.split_whitespace();
            let kind = parts.next().unwrap_or("proc");
            out.push(Section {
                kind,
                names: parts.collect(),
                body: String::new(),
            });
        } else if let Some(s) = out.last_mut() {
            s.body.push_str(line);
            s.body.push('\n');
        }
    }
    out
}

/// Identifiers (`[A-Za-z_][A-Za-z0-9_]*`) outside string literals and comments.
pub fn identifiers(code: &str) -> HashSet<&str> {
    let mut set = HashSet::new();
    for line in code.lines() {
        let bytes = line.as_bytes();
        let mut i = 0;
        let mut in_str = false;
        while i < bytes.len() {
            let b = bytes[i];
            if in_str {
                if b == b'"' {
                    in_str = false;
                }
                i += 1;
            } else if b == b'"' {
                in_str = true;
                i += 1;
            } else if b == b'\'' {
                break;
            } else if b.is_ascii_alphabetic() || b == b'_' {
                let start = i;
                while i < bytes.len() && (bytes[i].is_ascii_alphanumeric() || bytes[i] == b'_') {
                    i += 1;
                }
                set.insert(&line[start..i]);
            } else {
                i += 1;
            }
        }
    }
    set
}

/// Drops the VBA6 half of `#If VBA7 Then ... #Else ... #End If` blocks (`target.vba6 = false`).
fn strip_vba6(body: &str) -> String {
    let mut out = String::new();
    let mut state = 0; // 0 = outside, 1 = VBA7 branch, 2 = VBA6 branch
    for line in body.lines() {
        let t = line.trim();
        match (state, t) {
            (0, "#If VBA7 Then") => state = 1,
            (1, "#Else") => state = 2,
            (1 | 2, "#End If") => state = 0,
            (2, _) => {}
            _ => {
                out.push_str(line);
                out.push('\n');
            }
        }
    }
    out
}

pub struct Runtime {
    /// Code of `RAS_Runtime.bas` (without header), empty when nothing is needed.
    pub module: String,
    /// Code of `RAS_Error.cls` when the error class is used.
    pub error_class: Option<String>,
}

/// Builds the runtime for code that references the identifiers in `used`.
pub fn build(used: &HashSet<String>, cfg: &Config, has_init_glue: bool) -> Runtime {
    let japanese = cfg.target.codepage == 932;
    let msg = |ja: &'static str, en: &'static str| if japanese { ja } else { en };
    let secs = sections();
    let mut needed: Vec<bool> = secs
        .iter()
        .map(|s| s.names.iter().any(|n| used.contains(*n)))
        .collect();
    // Transitive closure over references between sections.
    loop {
        let mut changed = false;
        for i in 0..secs.len() {
            if !needed[i] {
                continue;
            }
            let ids = identifiers(&secs[i].body);
            for (j, other) in secs.iter().enumerate() {
                if !needed[j] && other.names.iter().any(|n| ids.contains(n)) {
                    needed[j] = true;
                    changed = true;
                }
            }
        }
        if !changed {
            break;
        }
    }
    let render = |body: &str| -> String {
        let fix_events = if cfg.init.fix_enable_events {
            format!(
                "    If Not Application.EnableEvents Then\n        Debug.Print \"[RAS] {}\"\n        Application.EnableEvents = True\n    End If",
                msg("EnableEvents が False のまま残っていたため True に戻しました", "EnableEvents was left False; restored to True")
            )
        } else {
            String::new()
        };
        let init_call = if has_init_glue {
            "    RAS_Glue.RAS_InitAll"
        } else {
            ""
        };
        let text = body
            .replace("{{YIELD_MS}}", &cfg.doevents.interval_ms.to_string())
            .replace("{{FIX_EVENTS}}", &fix_events)
            .replace("{{INIT_CALL}}", init_call)
            .replace("{{MSG_BUSY}}", msg("処理中です", "Busy"))
            .replace(
                "{{MSG_CANCELLED}}",
                msg("ユーザーにより中断されました", "Cancelled by the user"),
            )
            .replace("{{MSG_WHERE}}", msg("場所", "Location"))
            .replace("{{MSG_NUMBER}}", msg("エラー番号", "Error number"));
        let text = if cfg.target.vba6 {
            text
        } else {
            strip_vba6(&text)
        };
        // Blank placeholder lines are dropped.
        text.lines()
            .filter(|l| !l.trim().is_empty() || !body.contains("{{"))
            .map(|l| format!("{l}\n"))
            .collect()
    };
    let mut decls = String::new();
    let mut procs = String::new();
    let mut error_class = None;
    for (s, n) in secs.iter().zip(needed.iter()) {
        if !*n {
            continue;
        }
        match s.kind {
            "decl" => decls.push_str(&render(&s.body)),
            "class" => error_class = Some(render(&s.body)),
            _ => {
                procs.push('\n');
                procs.push_str(&render(&s.body));
            }
        }
    }
    let module = if decls.is_empty() && procs.is_empty() {
        String::new()
    } else {
        format!("{decls}{procs}")
    };
    Runtime {
        module,
        error_class,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn used(names: &[&str]) -> HashSet<String> {
        names.iter().map(|s| s.to_string()).collect()
    }

    #[test]
    fn tree_shaking_follows_dependencies() {
        let cfg = Config::default();
        let rt = build(&used(&["RAS_Sort"]), &cfg, false);
        assert!(rt.module.contains("Sub RAS_Sort("));
        assert!(rt.module.contains("Sub RAS_QSort("));
        assert!(rt.module.contains("Sub RAS_Swap("));
        assert!(rt.module.contains("Function RAS_Len("));
        assert!(!rt.module.contains("RAS_Yield"));
        assert!(rt.error_class.is_none());
    }

    #[test]
    fn yield_pulls_in_declares_and_state() {
        let cfg = Config::default();
        let rt = build(&used(&["RAS_Yield", "ras_yieldMask"]), &cfg, false);
        assert!(rt.module.contains("Declare PtrSafe Function GetTickCount"));
        assert!(rt.module.contains("Declare Function GetInputState"));
        assert!(rt.module.contains("Public ras_yieldMask As Long"));
        assert!(rt.module.contains("RAS_YIELD_MS As Double = 100#"));
    }

    #[test]
    fn sizable_forms_pull_in_the_window_api() {
        let mut cfg = Config::default();
        cfg.target.vba6 = false;
        let rt = build(&used(&["RAS_FormSizable"]), &cfg, false);
        assert!(rt
            .module
            .contains("Public Sub RAS_FormSizable(ByVal form As Object)"));
        assert!(rt.module.contains(
            "Private Declare PtrSafe Function RAS_WindowOfApi Lib \"shlwapi\" Alias \"#172\""
        ));
        assert!(rt
            .module
            .contains("Declare PtrSafe Function RAS_SetWindowPosApi"));
        assert!(rt.module.contains("    Dim hwnd As LongPtr\n"));
        assert!(!rt.module.contains("#Else"));
    }

    #[test]
    fn nothing_used_means_no_runtime() {
        let rt = build(&used(&["foo"]), &Config::default(), false);
        assert!(rt.module.is_empty());
    }

    #[test]
    fn error_class_and_vba6_stripping() {
        let mut cfg = Config::default();
        cfg.target.vba6 = false;
        let rt = build(&used(&["RAS_CaughtError", "RAS_Sleep"]), &cfg, true);
        assert!(rt.error_class.unwrap().contains("Public code As Long"));
        assert!(!rt.module.contains("#Else"));
        assert!(rt.module.contains("Declare PtrSafe Sub RAS_SleepApi"));
    }

    #[test]
    fn placeholders_are_all_replaced() {
        let all: HashSet<String> = sections()
            .iter()
            .flat_map(|s| s.names.iter().map(|n| n.to_string()))
            .collect();
        let rt = build(&all, &Config::default(), true);
        assert!(!rt.module.contains("{{"), "unreplaced placeholder");
        assert!(rt.module.contains("RAS_Glue.RAS_InitAll"));
    }
}
