//! `ras.toml` (spec §15.2).

use serde::Deserialize;

#[derive(Debug, Clone, Deserialize, Default)]
#[serde(default)]
pub struct Config {
    pub project: ProjectCfg,
    pub target: TargetCfg,
    pub build: BuildCfg,
    pub pack: PackCfg,
    pub doevents: DoEventsCfg,
    pub sync: SyncCfg,
    pub lint: LintCfg,
    pub init: InitCfg,
    pub test: TestCfg,
    /// The test files of the project (`rel_path`s matching `[test] include`), as opposed to
    /// the helper modules next to them (filled in by the driver).
    #[serde(skip)]
    pub test_files: Vec<String>,
    /// Sheet code names of the target workbook (filled in by the driver, not read from the file).
    #[serde(skip)]
    pub sheet_code_names: Vec<String>,
    /// `-O` override from the command line.
    #[serde(skip)]
    pub opt_override: Option<u8>,
    /// User forms of the project (`forms/<Name>.form.json`), whose code-behind
    /// `forms/<Name>.ras` is compiled as the form's code module (spec §14.11).
    #[serde(skip)]
    pub forms: Vec<FormDecl>,
    /// The ActiveX controls on the sheets of the target workbook (spec §12.2.1): one entry per
    /// sheet that has any, named by the sheet's code name (filled in by the driver).
    #[serde(skip)]
    pub sheet_controls: Vec<FormDecl>,
    /// `sheet_controls` was read from the workbook: a control it does not list is not on the
    /// sheet. Otherwise (no workbook, or one that could not be read) the controls are unknown.
    #[serde(skip)]
    pub sheet_controls_known: bool,
}

/// What the compiler needs to know about a user form: its name and controls.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FormDecl {
    /// `form.name` (the VBA name of the form).
    pub name: String,
    pub controls: Vec<FormControl>,
}

/// A control of a user form.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FormControl {
    /// VBA name (`txtName`).
    pub name: String,
    /// FormSpec control type (`TextBox`, `CommandButton`, `Page`, `Control` …).
    pub control_type: String,
}

impl FormDecl {
    /// Reads the name and controls from the text of a `.form.json` file (FormSpec);
    /// `None` when the text is not JSON with a `form.name`.
    pub fn from_form_json(text: &str) -> Option<FormDecl> {
        let value: serde_json::Value = serde_json::from_str(text.strip_prefix('\u{feff}').unwrap_or(text)).ok()?;
        let name = value.get("form")?.get("name")?.as_str()?.to_string();
        let mut controls = Vec::new();
        fn walk(list: Option<&serde_json::Value>, out: &mut Vec<FormControl>) {
            for c in list.and_then(|v| v.as_array()).into_iter().flatten() {
                let name = c.get("name").and_then(|v| v.as_str()).unwrap_or_default();
                let ty = c.get("type").and_then(|v| v.as_str()).unwrap_or_default();
                if !name.is_empty() {
                    out.push(FormControl { name: name.to_string(), control_type: ty.trim().to_string() });
                }
                // Legacy nested specs keep children inside their container.
                walk(c.get("controls"), out);
            }
        }
        walk(value.get("controls"), &mut controls);
        Some(FormDecl { name, controls })
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
pub struct ProjectCfg {
    pub name: String,
    pub workbook: Option<String>,
}

impl Default for ProjectCfg {
    fn default() -> Self {
        ProjectCfg { name: "RASProject".into(), workbook: None }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
pub struct TargetCfg {
    pub host: String,
    /// Also emit VBA6 (32-bit Office 2007 and earlier) `Declare`s.
    pub vba6: bool,
    pub codepage: u16,
}

impl Default for TargetCfg {
    fn default() -> Self {
        TargetCfg { host: "excel".into(), vba6: true, codepage: 932 }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(default, rename_all = "kebab-case")]
pub struct BuildCfg {
    /// 0: no expansion / 1: auto-inline in loops / 2: aggressive.
    pub opt_level: u8,
    pub inline_threshold: Option<u32>,
    /// Emit line numbers for `Erl`.
    pub erl: bool,
    pub keep_comments: bool,
    /// Warn when one generic definition expands to more than this many instances (spec §2.4.1).
    pub generic_instance_limit: usize,
}

impl Default for BuildCfg {
    fn default() -> Self {
        BuildCfg { opt_level: 1, inline_threshold: None, erl: true, keep_comments: true, generic_instance_limit: 16 }
    }
}

impl Config {
    pub fn opt_level(&self) -> u8 {
        self.opt_override.unwrap_or(self.build.opt_level).min(2)
    }

    /// Cost threshold for automatic inlining inside loops (spec §8.2).
    pub fn inline_threshold(&self) -> u32 {
        match self.build.inline_threshold {
            Some(t) => t,
            None if self.opt_level() >= 2 => 120,
            None => 40,
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(default, rename_all = "kebab-case")]
pub struct PackCfg {
    pub backup_generations: usize,
    pub sign: Option<String>,
}

impl Default for PackCfg {
    fn default() -> Self {
        PackCfg { backup_generations: 5, sign: None }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(default, rename_all = "kebab-case")]
pub struct DoEventsCfg {
    pub enabled: bool,
    /// Minimum interval between two `DoEvents` calls.
    pub interval_ms: u32,
    /// Smallest constant trip count that gets a yield point.
    pub loop_threshold: u32,
}

impl Default for DoEventsCfg {
    fn default() -> Self {
        DoEventsCfg { enabled: true, interval_ms: 100, loop_threshold: 10_000 }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(default, rename_all = "kebab-case")]
pub struct SyncCfg {
    pub enabled: bool,
    #[serde(rename = "compileCheck", alias = "compile-check")]
    pub compile_check: bool,
    pub shared_memory_mb: u32,
    pub pull_unmanaged: bool,
    /// Live sync creates `objects/<CodeName>.ras` for each sheet (and `ThisWorkbook`) that has
    /// no file yet (spec §14.7.2).
    pub sheet_files: bool,
}

impl Default for SyncCfg {
    fn default() -> Self {
        SyncCfg { enabled: true, compile_check: true, shared_memory_mb: 64, pull_unmanaged: true, sheet_files: true }
    }
}

#[derive(Debug, Clone, Deserialize, Default)]
#[serde(default, rename_all = "kebab-case")]
pub struct LintCfg {
    pub require_docs: bool,
    /// Warn about what the stand-in for Excel does not have (`ras check --target stand-in`).
    pub stand_in: bool,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(default, rename_all = "kebab-case")]
pub struct InitCfg {
    /// Re-enable `Application.EnableEvents` during initialisation when it was left off (spec §12.3).
    pub fix_enable_events: bool,
}

impl Default for InitCfg {
    fn default() -> Self {
        InitCfg { fix_enable_events: true }
    }
}

/// `[test]` (spec §14.12.9).
#[derive(Debug, Clone, Deserialize)]
#[serde(default, rename_all = "kebab-case")]
pub struct TestCfg {
    /// Patterns of the test files, relative to the project root.
    pub include: Vec<String>,
    /// Time a test may take.
    pub timeout_ms: u64,
    /// `standIn` (the workbook's stand-in) or `com` (the real Excel).
    pub excel: String,
    /// `failed`, `always` or `never`.
    pub save_workbooks: String,
    /// `fail` or `ignore`: what an error an event handler does not catch does to the test.
    pub event_errors: String,
}

impl Default for TestCfg {
    fn default() -> Self {
        TestCfg { include: vec!["test/**/*.test.ras".into()], timeout_ms: 10_000, excel: "standIn".into(), save_workbooks: "failed".into(), event_errors: "fail".into() }
    }
}

impl TestCfg {
    /// Whether a path relative to the project root (`/`-separated) is a test file.
    pub fn is_test_file(&self, rel_path: &str) -> bool {
        self.include.iter().any(|p| glob_match(p, rel_path))
    }
}

/// Matches a path against a pattern with `*` (within a name), `?` and `**` (any folders).
pub fn glob_match(pattern: &str, path: &str) -> bool {
    fn segs(p: &[&str], s: &[&str]) -> bool {
        match p.split_first() {
            None => s.is_empty(),
            Some((&"**", rest)) => (0..=s.len()).any(|i| segs(rest, &s[i..])),
            Some((first, rest)) => s.split_first().is_some_and(|(name, more)| seg(first.as_bytes(), name.as_bytes()) && segs(rest, more)),
        }
    }
    fn seg(p: &[u8], s: &[u8]) -> bool {
        match p.split_first() {
            None => s.is_empty(),
            Some((b'*', rest)) => (0..=s.len()).any(|i| seg(rest, &s[i..])),
            Some((b'?', rest)) => !s.is_empty() && seg(rest, &s[1..]),
            Some((c, rest)) => s.first() == Some(c) && seg(rest, &s[1..]),
        }
    }
    let p: Vec<&str> = pattern.trim_start_matches("./").split('/').collect();
    let s: Vec<&str> = path.split('/').collect();
    segs(&p, &s)
}

impl Config {
    pub fn from_toml(text: &str) -> Result<Config, String> {
        let cfg: Config = toml::from_str(text).map_err(|e| e.to_string())?;
        for (key, value, allowed) in [
            ("excel", &cfg.test.excel, &["standIn", "com"][..]),
            ("save-workbooks", &cfg.test.save_workbooks, &["failed", "always", "never"][..]),
            ("event-errors", &cfg.test.event_errors, &["fail", "ignore"][..]),
        ] {
            if !allowed.contains(&value.as_str()) {
                return Err(format!("test.{key} = \"{value}\" は指定できません（{}）", allowed.join(" / ")));
            }
        }
        if cfg.target.host != "excel" {
            return Err(format!("target.host = \"{}\" には対応していません（現状 \"excel\" のみ）", cfg.target.host));
        }
        Ok(cfg)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn spec_example_parses() {
        let cfg = Config::from_toml(
            r#"
[project]
name = "SalesReport"
workbook = "C:/work/SalesReport.xlsm"

[target]
host = "excel"
vba6 = true
codepage = 932

[build]
opt-level = 1
inline-threshold = 40
erl = true

[pack]
backup-generations = 5

[doevents]
enabled = true
interval-ms = 100
loop-threshold = 10000

[sync]
enabled = true
compileCheck = true
shared-memory-mb = 64
pull-unmanaged = true
"#,
        )
        .unwrap();
        assert_eq!(cfg.project.name, "SalesReport");
        assert_eq!(cfg.build.inline_threshold, Some(40));
        assert!(cfg.sync.compile_check);
        assert_eq!(cfg.doevents.loop_threshold, 10000);
    }

    #[test]
    fn defaults() {
        let cfg = Config::from_toml("").unwrap();
        assert_eq!(cfg.target.codepage, 932);
        assert_eq!(cfg.inline_threshold(), 40);
        assert!(Config::from_toml("[target]\nhost = \"word\"").is_err());
    }
}
