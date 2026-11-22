//! Watches for the dialogs Excel shows while a COM call is waiting for it (spec §14.10.6).
//!
//! The JavaScript thread is stuck inside `IDispatch::Invoke` for as long as Excel shows a
//! modal dialog, so the watching happens on a thread of its own. It looks at the top-level
//! windows of the watched processes only while a call has been outstanding for a moment,
//! reports each dialog once on stderr and, depending on the mode, brings it to the front or
//! answers it. What counts as a dialog and how it is answered is decided in `dialog.rs`.

use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::sync::{Mutex, Once, OnceLock};
use std::time::{Duration, Instant};

use napi::Result;
use napi_derive::napi;
use windows::core::BOOL;
use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CLSCTX_INPROC_SERVER, COINIT_MULTITHREADED,
};
use windows::Win32::UI::Accessibility::{
    CUIAutomation, IUIAutomation, TreeScope_Descendants, UIA_ButtonControlTypeId,
    UIA_TextControlTypeId,
};
use windows::Win32::UI::Input::KeyboardAndMouse::IsWindowEnabled;
use windows::Win32::UI::WindowsAndMessaging::{
    EnumChildWindows, EnumWindows, GetClassNameW, GetDlgCtrlID, GetWindow, GetWindowTextW,
    GetWindowThreadProcessId, IsWindow, IsWindowVisible, PostMessageW, SendMessageTimeoutW,
    SetForegroundWindow, BM_CLICK, GW_OWNER, SMTO_ABORTIFHUNG, WM_CLOSE, WM_COMMAND, WM_GETTEXT,
};

use crate::dialog::{is_dialog_class, Action, Button, Dialog, Mode, LOG_PREFIX};

/// A call must have been outstanding this long before windows are looked at.
const GRACE_MS: u64 = 300;
const POLL: Duration = Duration::from_millis(120);
/// How long Excel may reject calls without a dialog in sight before the user is told.
const BUSY_HINT_MS: u64 = 3_000;
/// Pause between attempts to answer a dialog that is still there.
const RETRY_MS: u64 = 700;
const MAX_ATTEMPTS: u32 = 3;

struct Config {
    pids: Vec<u32>,
    mode: Mode,
}

static CONFIG: Mutex<Config> = Mutex::new(Config {
    pids: Vec::new(),
    mode: Mode::Notify,
});
/// Start of the outstanding call in milliseconds since `epoch()`, plus one; 0 when no call is outstanding.
static CALL_SINCE: AtomicU64 = AtomicU64::new(0);
static CALL_SEQ: AtomicU64 = AtomicU64::new(0);
/// The outstanding call is being rejected by Excel (`RPC_E_CALL_REJECTED`) and retried.
static REJECTED: AtomicBool = AtomicBool::new(false);
static STARTED: Once = Once::new();

fn epoch() -> Instant {
    static EPOCH: OnceLock<Instant> = OnceLock::new();
    *EPOCH.get_or_init(Instant::now)
}

fn now_ms() -> u64 {
    epoch().elapsed().as_millis() as u64
}

/// Marks a COM call as outstanding for as long as it lives.
pub(crate) struct CallGuard(());

impl CallGuard {
    pub(crate) fn new() -> CallGuard {
        CALL_SEQ.fetch_add(1, Ordering::SeqCst);
        REJECTED.store(false, Ordering::SeqCst);
        CALL_SINCE.store(now_ms() + 1, Ordering::SeqCst);
        CallGuard(())
    }

    /// Excel has just refused the call because it is busy.
    pub(crate) fn rejected(&self) {
        REJECTED.store(true, Ordering::SeqCst);
    }
}

impl Drop for CallGuard {
    fn drop(&mut self) {
        CALL_SINCE.store(0, Ordering::SeqCst);
        REJECTED.store(false, Ordering::SeqCst);
    }
}

fn config() -> (Vec<u32>, Mode) {
    let c = CONFIG.lock().unwrap_or_else(|e| e.into_inner());
    (c.pids.clone(), c.mode)
}

// ---------------------------------------------------------------------------- windows

fn hwnd(handle: isize) -> HWND {
    HWND(handle as *mut core::ffi::c_void)
}

unsafe extern "system" fn collect(window: HWND, list: LPARAM) -> BOOL {
    let list = &mut *(list.0 as *mut Vec<isize>);
    list.push(window.0 as isize);
    BOOL(1)
}

fn class_name(window: HWND) -> String {
    let mut buf = [0u16; 128];
    let n = unsafe { GetClassNameW(window, &mut buf) };
    String::from_utf16_lossy(&buf[..n.max(0) as usize])
}

/// Text of a window. A control of another process only answers `WM_GETTEXT`; the timeout
/// keeps a process that does not pump messages from stalling the watcher.
fn window_text(window: HWND, control: bool) -> String {
    let mut buf = [0u16; 1024];
    let n = unsafe { GetWindowTextW(window, &mut buf) };
    if n > 0 || !control {
        return String::from_utf16_lossy(&buf[..n.max(0) as usize]);
    }
    let mut copied = 0usize;
    let sent = unsafe {
        SendMessageTimeoutW(
            window,
            WM_GETTEXT,
            WPARAM(buf.len()),
            LPARAM(buf.as_mut_ptr() as isize),
            SMTO_ABORTIFHUNG,
            200,
            Some(&mut copied),
        )
    };
    if sent.0 == 0 {
        return String::new();
    }
    String::from_utf16_lossy(&buf[..copied.min(buf.len())])
}

fn process_of(window: HWND) -> u32 {
    let mut pid = 0u32;
    unsafe { GetWindowThreadProcessId(window, Some(&mut pid)) };
    pid
}

fn describe(window: HWND, uia: Option<&IUIAutomation>) -> Dialog {
    let mut d = Dialog {
        handle: window.0 as isize,
        class: class_name(window),
        title: window_text(window, false),
        visible: unsafe { IsWindowVisible(window) }.as_bool(),
        owned: unsafe { GetWindow(window, GW_OWNER) }.is_ok_and(|o| !o.0.is_null()),
        ..Dialog::default()
    };
    let mut children: Vec<isize> = Vec::new();
    unsafe {
        let _ = EnumChildWindows(
            Some(window),
            Some(collect),
            LPARAM(&mut children as *mut Vec<isize> as isize),
        );
    }
    for child in children.into_iter().take(64).map(hwnd) {
        let class = class_name(child);
        if class.eq_ignore_ascii_case("Button") {
            let text = window_text(child, true);
            // Group boxes and unlabeled owner-drawn buttons are not answers.
            if !text.trim().is_empty() {
                d.buttons.push(Button {
                    handle: child.0 as isize,
                    text,
                    id: unsafe { GetDlgCtrlID(child) },
                    enabled: unsafe { IsWindowEnabled(child) }.as_bool(),
                });
            }
        } else if class.eq_ignore_ascii_case("Static") {
            let text = window_text(child, true);
            if !text.trim().is_empty() {
                d.texts.push(text);
            }
        } else if class.eq_ignore_ascii_case("Edit")
            || class.to_ascii_lowercase().starts_with("richedit")
        {
            d.has_edit = true;
        }
    }
    // Office draws its own dialogs without child windows; UI Automation still knows their content.
    if d.texts.is_empty() && d.buttons.is_empty() && d.visible {
        if let Some(uia) = uia {
            let _ = describe_by_automation(uia, window, &mut d);
        }
    }
    d
}

fn describe_by_automation(
    uia: &IUIAutomation,
    window: HWND,
    d: &mut Dialog,
) -> windows::core::Result<()> {
    unsafe {
        let element = uia.ElementFromHandle(window)?;
        let all = element.FindAll(TreeScope_Descendants, &uia.CreateTrueCondition()?)?;
        for i in 0..all.Length()?.min(200) {
            let Ok(item) = all.GetElement(i) else {
                continue;
            };
            let (Ok(kind), Ok(name)) = (item.CurrentControlType(), item.CurrentName()) else {
                continue;
            };
            let name = name.to_string();
            if name.trim().is_empty() {
                continue;
            }
            if kind == UIA_TextControlTypeId && d.texts.len() < 8 {
                d.texts.push(name);
            } else if kind == UIA_ButtonControlTypeId && d.buttons.len() < 8 {
                d.buttons.push(Button {
                    handle: 0,
                    text: name,
                    id: 0,
                    enabled: true,
                });
            }
        }
    }
    Ok(())
}

/// The dialogs the given processes are showing.
fn scan(pids: &[u32], uia: Option<&IUIAutomation>) -> Vec<Dialog> {
    let mut windows: Vec<isize> = Vec::new();
    unsafe {
        let _ = EnumWindows(
            Some(collect),
            LPARAM(&mut windows as *mut Vec<isize> as isize),
        );
    }
    windows
        .into_iter()
        .map(hwnd)
        .filter(|w| {
            pids.contains(&process_of(*w))
                && is_dialog_class(&class_name(*w))
                && unsafe { IsWindowEnabled(*w) }.as_bool()
        })
        .map(|w| describe(w, uia))
        .filter(Dialog::is_waiting)
        .collect()
}

fn click(dialog: &Dialog, button: &Button, attempt: u32) {
    let target = hwnd(button.handle);
    unsafe {
        if attempt % 2 == 1 {
            let _ = SendMessageTimeoutW(
                target,
                BM_CLICK,
                WPARAM(0),
                LPARAM(0),
                SMTO_ABORTIFHUNG,
                1000,
                None,
            );
        } else {
            // BM_CLICK is ignored by a dialog that is not active; the notification it would send is not.
            let _ = PostMessageW(
                Some(hwnd(dialog.handle)),
                WM_COMMAND,
                WPARAM(button.id as u16 as usize),
                LPARAM(button.handle),
            );
        }
    }
}

fn act(dialog: &Dialog, action: &Action, attempt: u32) {
    match action {
        Action::Click(i) => click(dialog, &dialog.buttons[*i], attempt),
        Action::Close => unsafe {
            let _ = PostMessageW(Some(hwnd(dialog.handle)), WM_CLOSE, WPARAM(0), LPARAM(0));
        },
        Action::Leave => {}
    }
}

// ---------------------------------------------------------------------------- the watcher

struct Seen {
    handle: isize,
    attempts: u32,
    last_ms: u64,
    gave_up: bool,
}

fn run() {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_MULTITHREADED);
    }
    let uia: Option<IUIAutomation> =
        unsafe { CoCreateInstance(&CUIAutomation, None, CLSCTX_INPROC_SERVER) }.ok();
    let mut seq = 0u64;
    let mut seen: Vec<Seen> = Vec::new();
    let mut hinted = false;
    loop {
        std::thread::sleep(POLL);
        let since = CALL_SINCE.load(Ordering::SeqCst);
        if since == 0 {
            continue;
        }
        let current = CALL_SEQ.load(Ordering::SeqCst);
        if current != seq {
            seq = current;
            seen.clear();
            hinted = false;
        }
        let now = now_ms();
        let waited = (now + 1).saturating_sub(since);
        if waited < GRACE_MS {
            continue;
        }
        let (pids, mode) = config();
        if mode == Mode::Off || pids.is_empty() {
            continue;
        }
        let dialogs = scan(&pids, uia.as_ref());
        // The call may have returned while the windows were examined: then the dialog is not ours to answer.
        if CALL_SINCE.load(Ordering::SeqCst) == 0 || CALL_SEQ.load(Ordering::SeqCst) != seq {
            continue;
        }
        if dialogs.is_empty() {
            if REJECTED.load(Ordering::SeqCst) && waited >= BUSY_HINT_MS && !hinted {
                hinted = true;
                eprintln!("{LOG_PREFIX}Excel が応答しません（セルの編集中か、メニューやダイアログが開いています）。Excel 側の操作を終えると続行します");
            }
            continue;
        }
        for d in &dialogs {
            let action = if mode == Mode::Dismiss {
                d.dismissal()
            } else {
                Action::Leave
            };
            match seen.iter_mut().find(|s| s.handle == d.handle) {
                None => {
                    seen.push(Seen {
                        handle: d.handle,
                        attempts: 1,
                        last_ms: now,
                        gave_up: false,
                    });
                    if mode == Mode::Dismiss {
                        eprintln!("{}", d.dismissed_report(&action));
                        act(d, &action, 1);
                    } else {
                        eprintln!("{}", d.report());
                        // Usually refused for a background process; the taskbar button flashes instead.
                        unsafe {
                            let _ = SetForegroundWindow(hwnd(d.handle));
                        }
                    }
                }
                Some(s) => {
                    if action == Action::Leave
                        || s.gave_up
                        || now.saturating_sub(s.last_ms) < RETRY_MS
                    {
                        continue;
                    }
                    if s.attempts >= MAX_ATTEMPTS {
                        s.gave_up = true;
                        eprintln!("{LOG_PREFIX}Excel のダイアログを閉じられませんでした: {}  — Excel 側で応答してください", d.summary());
                        continue;
                    }
                    s.attempts += 1;
                    s.last_ms = now;
                    act(d, &action, s.attempts);
                }
            }
        }
    }
}

/// One line about the dialogs the watched processes show right now (for error messages).
pub(crate) fn open_dialog_summary() -> Option<String> {
    let (pids, mode) = config();
    if mode == Mode::Off || pids.is_empty() {
        return None;
    }
    let dialogs = scan(&pids, None);
    if dialogs.is_empty() {
        return None;
    }
    Some(
        dialogs
            .iter()
            .map(Dialog::summary)
            .collect::<Vec<_>>()
            .join("、"),
    )
}

// ---------------------------------------------------------------------------- JavaScript API

/// Starts watching the dialogs of a process. `mode`: `"notify"` (report and bring to the
/// front), `"dismiss"` (report and answer with the cancelling button) or `"off"`.
/// The mode applies to every watched process.
#[napi]
pub fn watch_dialogs(pid: u32, mode: String) -> Result<()> {
    let mode = Mode::parse(&mode).ok_or_else(|| {
        napi::Error::from_reason(format!(
            "unknown dialog mode `{mode}` (notify, dismiss or off)"
        ))
    })?;
    {
        let mut c = CONFIG.lock().unwrap_or_else(|e| e.into_inner());
        c.mode = mode;
        if pid != 0 && !c.pids.contains(&pid) {
            c.pids.push(pid);
        }
    }
    if mode != Mode::Off {
        STARTED.call_once(|| {
            epoch();
            let _ = std::thread::Builder::new()
                .name("ras-dialog-watcher".into())
                .spawn(run);
        });
    }
    Ok(())
}

/// The process that owns a window (`Application.Hwnd`); 0 when there is no such window.
#[napi]
pub fn process_of_window(window: f64) -> u32 {
    // Excel reports the handle as a 32-bit Long; handles are sign-extended to pointer size.
    let handle = hwnd(window as i64 as i32 as isize);
    if unsafe { IsWindow(Some(handle)) }.as_bool() {
        process_of(handle)
    } else {
        0
    }
}

/// Descriptions of the dialogs a process is showing right now.
#[napi]
pub fn open_dialogs(pid: u32) -> Vec<String> {
    scan(&[pid], None).iter().map(Dialog::summary).collect()
}
