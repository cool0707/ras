//! Dialogs of Excel and the VBE, as the watcher sees them (spec §14.10.6).
//!
//! A COM call into Excel does not return while Excel shows a modal dialog: a run-time error of
//! hand-written VBA (an event handler fired by a cell write), a compile error, a `MsgBox`, a
//! confirmation of Excel itself. Excel is usually behind the editor at that moment, so the
//! program looks hung. The watcher (`watch.rs`) finds such windows; this module decides what
//! they are, how to describe them and — when asked to — which button answers them safely.
//!
//! The rules follow what xlflow's dialog watcher established on real Excel installations:
//! buttons are recognised by control id and access key rather than by their (localised)
//! caption, a run-time error is answered with End rather than Debug (Debug leaves the VBE in
//! break mode), and a dialog without an unambiguous cancelling answer is left alone.

/// Standard dialog control ids (`IDOK` …).
pub const ID_OK: i32 = 1;
pub const ID_CANCEL: i32 = 2;
pub const ID_CLOSE: i32 = 8;
pub const ID_HELP: i32 = 9;

/// Start of every line the watcher prints; the VS Code extension turns such lines into notifications.
pub const LOG_PREFIX: &str = "[ras] ";

/// What to do with a dialog that blocks a call.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Mode {
    /// Ignore dialogs.
    Off,
    /// Report the dialog and bring it to the front; the user answers it.
    Notify,
    /// Report the dialog and answer it with its cancelling button (unattended runs).
    Dismiss,
}

impl Mode {
    pub fn parse(s: &str) -> Option<Mode> {
        match s.trim().to_ascii_lowercase().as_str() {
            "off" | "none" | "false" | "0" => Some(Mode::Off),
            "" | "notify" | "report" => Some(Mode::Notify),
            "dismiss" | "cancel" | "close" => Some(Mode::Dismiss),
            _ => None,
        }
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Button {
    /// Window handle; 0 when the button is known by name only (UI Automation) and cannot be clicked.
    pub handle: isize,
    /// Caption as shown, including the access-key marker (`&End`, `終了(&E)`).
    pub text: String,
    pub id: i32,
    pub enabled: bool,
}

impl Button {
    /// The access key: the character after `&`, or the one in a trailing `(X)`.
    pub fn access_key(&self) -> Option<char> {
        access_key(&self.text)
    }

    /// Caption without the access-key marker.
    pub fn label(&self) -> String {
        plain_label(&self.text)
    }

    fn clickable(&self) -> bool {
        self.handle != 0 && self.enabled
    }

    fn is_help(&self) -> bool {
        self.id == ID_HELP || self.access_key() == Some('H')
    }
}

#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Dialog {
    pub handle: isize,
    /// Window class: `#32770` (standard dialog), `NUIDialog` / `bosa_sdm_…` (Office's own).
    pub class: String,
    pub title: String,
    pub visible: bool,
    /// Whether the window has an owner (modal dialogs do).
    pub owned: bool,
    /// Static texts, top to bottom.
    pub texts: Vec<String>,
    pub buttons: Vec<Button>,
    /// Has a text box (`InputBox`).
    pub has_edit: bool,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Kind {
    /// "Run-time error …" of VBA: Continue / End / Debug / Help.
    RuntimeError,
    /// Another message of the VBE, in practice a compile error.
    VbeMessage,
    /// A dialog that asks for text (`InputBox`).
    Input,
    /// A message with buttons (`MsgBox`, alerts built from standard controls).
    Message,
    /// A dialog whose controls are not windows (most dialogs of Excel itself).
    Other,
}

/// What `Mode::Dismiss` does with a dialog.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Action {
    /// Press this button (index into `Dialog::buttons`).
    Click(usize),
    /// Ask the window to close, which Office dialogs treat as Cancel.
    Close,
    /// No answer is safe to give on the user's behalf.
    Leave,
}

fn access_key(text: &str) -> Option<char> {
    let chars: Vec<char> = text.chars().collect();
    for (i, c) in chars.iter().enumerate() {
        if *c == '&' {
            match chars.get(i + 1) {
                Some('&') | None => {}
                Some(k) if k.is_ascii_alphanumeric() => return Some(k.to_ascii_uppercase()),
                Some(_) => {}
            }
        }
    }
    // Localised captions carry the key in parentheses: `終了(E)`, `デバッグ（D）`.
    for (i, c) in chars.iter().enumerate() {
        if (*c == '(' || *c == '（') && matches!(chars.get(i + 2), Some(')') | Some('）')) {
            if let Some(k) = chars.get(i + 1).filter(|k| k.is_ascii_alphanumeric()) {
                return Some(k.to_ascii_uppercase());
            }
        }
    }
    None
}

fn plain_label(text: &str) -> String {
    let mut s = text.trim().replace('&', "");
    // `終了(E)` → `終了`
    if let Some(open) = s.rfind(['(', '（']) {
        let tail: Vec<char> = s[open..].chars().collect();
        if open > 0 && tail.len() == 3 && (tail[1].is_ascii_uppercase() || tail[1].is_ascii_digit()) && (tail[2] == ')' || tail[2] == '）') {
            s.truncate(open);
        }
    }
    s.trim().to_string()
}

/// Whether a window of this class is a dialog at all.
pub fn is_dialog_class(class: &str) -> bool {
    class == "#32770" || class.eq_ignore_ascii_case("NUIDialog") || class.to_ascii_lowercase().starts_with("bosa_sdm_")
}

impl Dialog {
    /// Whether the window is a dialog that is waiting for an answer.
    ///
    /// Visibility alone is not the test: Excel creates the run-time error dialog but may not
    /// paint it until its window gets the focus, so an owned standard dialog with buttons
    /// counts even while hidden. Hidden windows of the other classes are leftovers Excel keeps.
    pub fn is_waiting(&self) -> bool {
        if !is_dialog_class(&self.class) {
            return false;
        }
        self.visible || (self.class == "#32770" && self.owned && self.buttons.iter().any(|b| b.handle != 0))
    }

    pub fn kind(&self) -> Kind {
        let standard = self.class == "#32770";
        let keys: Vec<char> = self.buttons.iter().filter_map(Button::access_key).collect();
        if standard && (3..=4).contains(&self.buttons.len()) && ['C', 'E', 'D'].iter().all(|k| keys.contains(k)) {
            return Kind::RuntimeError;
        }
        if standard && self.title.contains("Microsoft Visual Basic") && !self.has_edit {
            return Kind::VbeMessage;
        }
        if self.has_edit {
            return Kind::Input;
        }
        if self.buttons.iter().any(|b| b.handle != 0) {
            return Kind::Message;
        }
        Kind::Other
    }

    /// The answer `Mode::Dismiss` gives.
    pub fn dismissal(&self) -> Action {
        let by_key = |key: char| self.buttons.iter().position(|b| b.clickable() && b.access_key() == Some(key));
        let by_id = |id: i32| self.buttons.iter().position(|b| b.clickable() && b.id == id);
        if self.kind() == Kind::RuntimeError {
            // End stops the macro; Debug would leave the VBE in break mode and Excel unusable for COM.
            return by_key('E').map(Action::Click).unwrap_or(Action::Leave);
        }
        if let Some(i) = by_id(ID_CANCEL).or_else(|| by_id(ID_CLOSE)) {
            return Action::Click(i);
        }
        if !self.buttons.iter().any(|b| b.handle != 0) {
            return Action::Close;
        }
        // A single answer (apart from Help) is no choice: an information box, a compile error.
        let mut answers = self.buttons.iter().enumerate().filter(|(_, b)| b.clickable() && !b.is_help());
        match (answers.next(), answers.next()) {
            (Some((i, _)), None) => Action::Click(i),
            _ => Action::Leave,
        }
    }

    /// `「title」 text [button] [button]` on one line.
    pub fn summary(&self) -> String {
        let one_line = |s: &str| s.split_whitespace().collect::<Vec<_>>().join(" ");
        let mut out = format!("「{}」", one_line(&self.title));
        let texts: Vec<String> = self.texts.iter().map(|t| one_line(t)).filter(|t| !t.is_empty()).collect();
        if !texts.is_empty() {
            out.push(' ');
            out.push_str(&texts.join(" / "));
        }
        for b in self.buttons.iter().filter(|b| !b.label().is_empty()) {
            out.push_str(&format!(" [{}]", b.label()));
        }
        out
    }

    fn what(&self) -> &'static str {
        match self.kind() {
            Kind::RuntimeError => "VBA の実行時エラーのダイアログ",
            Kind::VbeMessage => "VBE のダイアログ（コンパイル エラーなど）",
            Kind::Input => "入力ダイアログ",
            Kind::Message | Kind::Other => "ダイアログ",
        }
    }

    /// The line printed when the dialog is found while a call is waiting.
    pub fn report(&self) -> String {
        format!("{LOG_PREFIX}Excel が{}を表示して応答を待っています: {}  — Excel 側で応答してください", self.what(), self.summary())
    }

    /// The line printed when `Mode::Dismiss` has answered (or could not answer) the dialog.
    pub fn dismissed_report(&self, action: &Action) -> String {
        match action {
            Action::Click(i) => format!("{LOG_PREFIX}Excel の{}を「{}」で閉じました: {}", self.what(), self.buttons[*i].label(), self.summary()),
            Action::Close => format!("{LOG_PREFIX}Excel の{}を閉じました: {}", self.what(), self.summary()),
            Action::Leave => format!("{LOG_PREFIX}Excel が{}を表示して応答を待っています（自動では閉じられません）: {}  — Excel 側で応答してください", self.what(), self.summary()),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn button(handle: isize, text: &str, id: i32) -> Button {
        Button { handle, text: text.into(), id, enabled: true }
    }

    fn dialog(class: &str, title: &str, buttons: Vec<Button>) -> Dialog {
        Dialog { handle: 1, class: class.into(), title: title.into(), visible: true, owned: true, texts: vec![], buttons, has_edit: false }
    }

    #[test]
    fn access_keys_and_labels() {
        assert_eq!(access_key("&End"), Some('E'));
        assert_eq!(access_key("終了(&E)"), Some('E'));
        assert_eq!(access_key("デバッグ(D)"), Some('D'));
        assert_eq!(access_key("継続（C）"), Some('C'));
        assert_eq!(access_key("Save && Close"), None);
        assert_eq!(access_key("OK"), None);
        assert_eq!(plain_label("終了(&E)"), "終了");
        assert_eq!(plain_label("&Debug"), "Debug");
        assert_eq!(plain_label("ヘルプ(H)"), "ヘルプ");
        assert_eq!(plain_label("f(x)"), "f(x)", "only a trailing access key is removed");
        assert_eq!(plain_label("OK"), "OK");
    }

    #[test]
    fn modes() {
        assert_eq!(Mode::parse("Dismiss"), Some(Mode::Dismiss));
        assert_eq!(Mode::parse(""), Some(Mode::Notify));
        assert_eq!(Mode::parse("off"), Some(Mode::Off));
        assert_eq!(Mode::parse("whatever"), None);
    }

    #[test]
    fn runtime_error_is_ended_not_debugged() {
        let en = dialog("#32770", "Microsoft Visual Basic", vec![button(10, "&Continue", 4800), button(11, "&End", 4801), button(12, "&Debug", 4802), button(13, "&Help", 4803)]);
        assert_eq!(en.kind(), Kind::RuntimeError);
        assert_eq!(en.dismissal(), Action::Click(1));
        let mut ja = dialog("#32770", "Microsoft Visual Basic", vec![button(10, "継続(&C)", 0), button(11, "終了(&E)", 0), button(12, "デバッグ(&D)", 0), button(13, "ヘルプ(&H)", 0)]);
        ja.texts = vec!["実行時エラー '9':".into(), "インデックスが有効範囲にありません。".into()];
        assert_eq!(ja.kind(), Kind::RuntimeError);
        assert_eq!(ja.dismissal(), Action::Click(1));
        assert_eq!(ja.summary(), "「Microsoft Visual Basic」 実行時エラー '9': / インデックスが有効範囲にありません。 [継続] [終了] [デバッグ] [ヘルプ]");
        assert!(ja.report().starts_with("[ras] Excel がVBA の実行時エラーのダイアログを表示して応答を待っています: "));
        assert!(ja.dismissed_report(&Action::Click(1)).contains("「終了」で閉じました"));
        // With End disabled nothing else is pressed.
        ja.buttons[1].enabled = false;
        assert_eq!(ja.dismissal(), Action::Leave);
    }

    #[test]
    fn compile_error_and_information_boxes_have_one_answer() {
        let compile = dialog("#32770", "Microsoft Visual Basic for Applications", vec![button(10, "OK", ID_OK), button(11, "ヘルプ", ID_HELP)]);
        assert_eq!(compile.kind(), Kind::VbeMessage);
        assert_eq!(compile.dismissal(), Action::Click(0));
        let info = dialog("#32770", "Book1", vec![button(10, "OK", ID_OK)]);
        assert_eq!(info.kind(), Kind::Message);
        assert_eq!(info.dismissal(), Action::Click(0));
    }

    #[test]
    fn cancel_is_preferred_and_a_real_choice_is_left_alone() {
        let ok_cancel = dialog("#32770", "Microsoft Excel", vec![button(10, "OK", ID_OK), button(11, "キャンセル", ID_CANCEL)]);
        assert_eq!(ok_cancel.dismissal(), Action::Click(1));
        let yes_no = dialog("#32770", "Microsoft Excel", vec![button(10, "はい(&Y)", 6), button(11, "いいえ(&N)", 7)]);
        assert_eq!(yes_no.dismissal(), Action::Leave);
        assert!(yes_no.dismissed_report(&Action::Leave).contains("自動では閉じられません"));
        let mut input = dialog("#32770", "入力", vec![button(10, "OK", ID_OK), button(11, "キャンセル", ID_CANCEL)]);
        input.has_edit = true;
        assert_eq!(input.kind(), Kind::Input);
        assert_eq!(input.dismissal(), Action::Click(1));
    }

    #[test]
    fn office_dialogs_without_child_windows_are_closed() {
        // Buttons known through UI Automation only have no handle.
        let alert = dialog("NUIDialog", "Microsoft Excel", vec![button(0, "保存", 0), button(0, "保存しない", 0), button(0, "キャンセル", 0)]);
        assert_eq!(alert.kind(), Kind::Other);
        assert_eq!(alert.dismissal(), Action::Close);
        assert_eq!(alert.summary(), "「Microsoft Excel」 [保存] [保存しない] [キャンセル]");
    }

    #[test]
    fn which_windows_count() {
        let mut d = dialog("#32770", "Microsoft Visual Basic", vec![button(10, "&End", 0)]);
        assert!(d.is_waiting());
        // Not painted yet (Excel in the background): still a dialog that blocks.
        d.visible = false;
        assert!(d.is_waiting());
        d.owned = false;
        assert!(!d.is_waiting());
        let mut sdm = dialog("bosa_sdm_XL9", "セルの書式設定", vec![]);
        assert!(sdm.is_waiting());
        sdm.visible = false;
        assert!(!sdm.is_waiting(), "hidden leftovers are not dialogs");
        assert!(!dialog("XLMAIN", "Book1 - Excel", vec![]).is_waiting());
    }
}
