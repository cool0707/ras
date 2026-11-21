//! What was added to a generated module in the VBE (RAS language specification §14.7.3).
//!
//! A module is split into the statements of its declaration section and its procedures.
//! [`diff`] compares the module as it is now with the code RAS wrote: procedures and
//! declarations that the written code does not have were added in the VBE and can be taken
//! over into the module's RAS file; a written procedure that changed or is gone is drift.

/// A procedure (`Sub`, `Function`, `Property`) with the comment lines right above it.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Procedure {
    /// Its name; for a `Property` with its accessor (`Get Name`), as one name has up to three.
    pub name: String,
    /// Its lines, from the comments above it to its `End` line (LF line ends).
    pub text: String,
}

/// A module split into the statements of its declaration section and its procedures.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Parts {
    /// Statements before the first procedure: a declaration with its continuation lines, a
    /// `Type` / `Enum` block, a conditional block (`#If` … `#End If`). `Option` statements, the
    /// hash line of RAS and comments are left out.
    pub declarations: Vec<String>,
    /// The procedures, in order.
    pub procedures: Vec<Procedure>,
}

/// Modifiers that may precede `Sub` / `Function` / `Property`.
const MODIFIERS: [&str; 4] = ["public", "private", "friend", "static"];

/// The kind and name of the procedure a line declares.
fn procedure_start(line: &str) -> Option<(String, String)> {
    let mut words = strip_line_number(line).split_whitespace().peekable();
    while words.peek().is_some_and(|w| MODIFIERS.iter().any(|m| w.eq_ignore_ascii_case(m))) {
        words.next();
    }
    let keyword = words.next()?.to_ascii_lowercase();
    let (end, prefix) = match keyword.as_str() {
        "sub" | "function" => (keyword.clone(), String::new()),
        "property" => {
            let accessor = words.next()?.to_ascii_lowercase();
            if !["get", "let", "set"].contains(&accessor.as_str()) {
                return None;
            }
            ("property".to_string(), format!("{}{} ", accessor[..1].to_uppercase(), &accessor[1..]))
        }
        _ => return None,
    };
    let name = words.next()?.split('(').next()?;
    if name.is_empty() {
        return None;
    }
    Some((end, format!("{prefix}{name}")))
}

/// A line without the line number RAS puts in front of statements (for `Erl`).
fn strip_line_number(line: &str) -> &str {
    let trimmed = line.trim_start();
    let digits = trimmed.len() - trimmed.trim_start_matches(|c: char| c.is_ascii_digit()).len();
    if digits > 0 && trimmed[digits..].starts_with([' ', '\t']) {
        trimmed[digits..].trim_start()
    } else {
        trimmed
    }
}

fn is_comment(line: &str) -> bool {
    let t = line.trim_start();
    t.starts_with('\'') || t.to_ascii_lowercase().starts_with("rem ")
}

fn first_words(line: &str, count: usize) -> String {
    let mut words = line.split_whitespace().map(str::to_ascii_lowercase).filter(|w| !MODIFIERS.contains(&w.as_str()));
    (0..count).filter_map(|_| words.next()).collect::<Vec<_>>().join(" ")
}

/// Splits the code of a module (with or without its `Attribute` header).
pub fn split(code: &str) -> Parts {
    let text = code.replace("\r\n", "\n").replace('\r', "\n");
    let lines: Vec<&str> = text.lines().collect();
    let mut parts = Parts::default();
    let mut i = 0;
    // Comment lines seen since the last statement: they belong to the procedure that follows.
    let mut comments: Vec<&str> = Vec::new();
    while i < lines.len() {
        let line = lines[i];
        let trimmed = line.trim();
        if trimmed.is_empty() {
            comments.clear();
            i += 1;
            continue;
        }
        if is_comment(line) {
            comments.push(line);
            i += 1;
            continue;
        }
        if let Some((end, name)) = procedure_start(line) {
            let mut text: Vec<&str> = std::mem::take(&mut comments);
            let end_line = format!("end {end}");
            let mut j = i;
            loop {
                text.push(lines[j]);
                let done = strip_line_number(lines[j]).to_ascii_lowercase().split('\'').next().unwrap_or_default().trim() == end_line;
                j += 1;
                if done || j >= lines.len() {
                    break;
                }
            }
            parts.procedures.push(Procedure { name, text: text.iter().map(|l| format!("{l}\n")).collect() });
            i = j;
            continue;
        }
        comments.clear();
        let lower = trimmed.to_ascii_lowercase();
        if trimmed.starts_with("Attribute ") || trimmed.starts_with("VERSION ") || lower.starts_with("option ") {
            i += 1;
            continue;
        }
        // A block, or a statement with its continuation lines.
        let block_end = match first_words(trimmed, 1).as_str() {
            "type" => Some("end type"),
            "enum" => Some("end enum"),
            w if w.starts_with("#if") => Some("#end if"),
            _ => None,
        };
        let mut j = i;
        let mut statement: Vec<&str> = Vec::new();
        loop {
            statement.push(lines[j]);
            let this = lines[j].trim().to_ascii_lowercase();
            let done = match block_end {
                Some(end) => this.split('\'').next().unwrap_or_default().trim() == end,
                None => !lines[j].trim_end().ends_with(" _"),
            };
            j += 1;
            if done || j >= lines.len() {
                break;
            }
        }
        // Declarations after a procedure do not exist in VBA; whatever is there stays out.
        if parts.procedures.is_empty() {
            parts.declarations.push(statement.iter().map(|l| format!("{l}\n")).collect());
        }
        i = j;
    }
    parts
}

/// A text as the comparison sees it: the VBE rewrites the case of names and the spacing, and
/// RAS numbers the lines; none of that is a change (spec §14.7.3).
fn normalized(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for line in text.lines() {
        let line = strip_line_number(line);
        if line.is_empty() || line.starts_with("Attribute ") {
            continue;
        }
        let mut in_string = false;
        let mut pending_space = false;
        for c in line.chars() {
            if c == '"' {
                in_string = !in_string;
            }
            if !in_string && c.is_whitespace() {
                pending_space = true;
                continue;
            }
            if pending_space && !out.ends_with('\n') && !out.is_empty() {
                out.push(' ');
            }
            pending_space = false;
            if in_string {
                out.push(c);
            } else {
                out.extend(c.to_lowercase());
            }
        }
        out.push('\n');
    }
    out
}

/// How a module differs from the code RAS wrote into it.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Diff {
    /// Declarations the written code does not have.
    pub added_declarations: Vec<String>,
    /// Procedures of a name the written code does not have.
    pub added_procedures: Vec<Procedure>,
    /// Names of written procedures whose text changed.
    pub changed: Vec<String>,
    /// Names of written procedures, and written declarations, that are gone.
    pub removed: Vec<String>,
}

impl Diff {
    /// Something was added in the VBE.
    pub fn has_additions(&self) -> bool {
        !self.added_declarations.is_empty() || !self.added_procedures.is_empty()
    }

    /// What RAS wrote was changed or removed: drift that cannot be taken over (spec §14.7).
    pub fn has_drift(&self) -> bool {
        !self.changed.is_empty() || !self.removed.is_empty()
    }

    /// The added code as one piece of VBA: the declarations, then the procedures.
    pub fn added_code(&self) -> String {
        let mut out: String = self.added_declarations.concat();
        for p in &self.added_procedures {
            if !out.is_empty() {
                out.push('\n');
            }
            out.push_str(&p.text);
        }
        out
    }

    /// The names of what was added, for messages.
    pub fn added_names(&self) -> Vec<String> {
        let mut names: Vec<String> = self.added_procedures.iter().map(|p| p.name.clone()).collect();
        if !self.added_declarations.is_empty() {
            names.push(format!("宣言 {} 件", self.added_declarations.len()));
        }
        names
    }

    /// A short hash of the added code, to tell whether it was taken over before.
    pub fn added_hash(&self) -> String {
        let mut hash: u32 = 0x811c_9dc5;
        for b in normalized(&self.added_code()).bytes() {
            hash = (hash ^ b as u32).wrapping_mul(0x0100_0193);
        }
        format!("{hash:08x}")
    }
}

/// Compares the module `current` with `written`, the code RAS wrote into it.
pub fn diff(written: &str, current: &str) -> Diff {
    let (base, now) = (split(written), split(current));
    let mut out = Diff::default();
    let base_declarations: Vec<String> = base.declarations.iter().map(|d| normalized(d)).collect();
    let now_declarations: Vec<String> = now.declarations.iter().map(|d| normalized(d)).collect();
    for (text, key) in now.declarations.iter().zip(&now_declarations) {
        if !base_declarations.contains(key) {
            out.added_declarations.push(text.clone());
        }
    }
    for (text, key) in base.declarations.iter().zip(&base_declarations) {
        if !now_declarations.contains(key) {
            out.removed.push(text.lines().next().unwrap_or_default().trim().to_string());
        }
    }
    let same = |a: &str, b: &str| a.to_lowercase() == b.to_lowercase();
    for p in &now.procedures {
        match base.procedures.iter().find(|b| same(&b.name, &p.name)) {
            None => out.added_procedures.push(p.clone()),
            Some(b) if normalized(&b.text) != normalized(&p.text) => out.changed.push(p.name.clone()),
            Some(_) => {}
        }
    }
    for b in &base.procedures {
        if !now.procedures.iter().any(|p| same(&p.name, &b.name)) {
            out.removed.push(b.name.clone());
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    const WRITTEN: &str = "' RAS-GENERATED hash=0000ABCD\r\nOption Explicit\r\nOption Compare Binary\r\n\r\nPrivate lastEdited As String\r\nPrivate Const LIMIT As Long = 10&\r\n\r\n' 入力欄を消す。\r\nPublic Sub clearInput()\r\n    If Not RAS_Enter(\"Sheet1.clearInput\") Then Exit Sub\r\n5   Me.Range(\"A2:C100\").ClearContents\r\nEnd Sub\r\n\r\nPrivate Sub Worksheet_Change(ByVal Target As Range)\r\n    onChange Target\r\nEnd Sub\r\n";

    #[test]
    fn a_module_in_parts() {
        let parts = split(WRITTEN);
        assert_eq!(parts.declarations, ["Private lastEdited As String\n", "Private Const LIMIT As Long = 10&\n"]);
        let names: Vec<&str> = parts.procedures.iter().map(|p| p.name.as_str()).collect();
        assert_eq!(names, ["clearInput", "Worksheet_Change"]);
        // The comment above a procedure belongs to it.
        assert!(parts.procedures[0].text.starts_with("' 入力欄を消す。\nPublic Sub clearInput()\n"), "{}", parts.procedures[0].text);
        // Blocks and continuation lines are one statement; properties are told apart by accessor.
        let parts = split("Private Type Point\n    x As Long\nEnd Type\nPrivate Declare PtrSafe Function GetTickCount _\n    Lib \"kernel32\" () As Long\n#If VBA7 Then\nPrivate p As LongPtr\n#End If\nProperty Get Name() As String\nEnd Property\nProperty Let Name(v As String)\nEnd Property\n");
        assert_eq!(parts.declarations.len(), 3, "{:?}", parts.declarations);
        let names: Vec<&str> = parts.procedures.iter().map(|p| p.name.as_str()).collect();
        assert_eq!(names, ["Get Name", "Let Name"]);
    }

    #[test]
    fn what_the_vbe_rewrites_is_no_change() {
        // Other case and spacing, no line numbers, another line end.
        let current = WRITTEN.replace("\r\n", "\n").replace("5   Me.Range", "    me.range").replace("lastEdited As String", "LastEdited  As String").replace("onChange Target", "OnChange  target");
        let d = diff(WRITTEN, &current);
        assert_eq!(d, Diff::default());
        // A string is compared as it is.
        let d = diff(WRITTEN, &WRITTEN.replace("A2:C100", "a2:c100"));
        assert_eq!(d.changed, ["clearInput"]);
    }

    #[test]
    fn added_changed_and_removed() {
        let current = format!("{}\r\nPrivate Sub Worksheet_Activate()\r\n    MsgBox \"hi\"\r\nEnd Sub\r\n", WRITTEN.replace("Private Const LIMIT As Long = 10&\r\n", "Private Const LIMIT As Long = 10&\r\nPrivate counter As Long\r\n"));
        let d = diff(WRITTEN, &current);
        assert!(d.has_additions() && !d.has_drift(), "{d:?}");
        assert_eq!(d.added_names(), ["Worksheet_Activate", "宣言 1 件"]);
        assert_eq!(d.added_code(), "Private counter As Long\n\nPrivate Sub Worksheet_Activate()\n    MsgBox \"hi\"\nEnd Sub\n");
        assert_eq!(d.added_hash(), diff(WRITTEN, &current.replace("MsgBox \"hi\"", "msgbox  \"hi\"")).added_hash());

        let current = WRITTEN.replace("    onChange Target\r\n", "    onChange Target\r\n    Beep\r\n").replace("Private lastEdited As String\r\n", "");
        let d = diff(WRITTEN, &current);
        assert!(!d.has_additions() && d.has_drift(), "{d:?}");
        assert_eq!(d.changed, ["Worksheet_Change"]);
        assert_eq!(d.removed, ["Private lastEdited As String"]);
        let d = diff(WRITTEN, "Option Explicit\r\n");
        assert_eq!(d.removed.len(), 4, "{d:?}");
    }
}
