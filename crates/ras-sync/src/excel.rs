//! The workbook that is open in Excel, and its VBA project (spec §14.3, §14.6).

use std::ffi::c_void;
use std::path::Path;

use windows::core::{Interface, PCWSTR};
use windows::Win32::Foundation::HWND;
use windows::Win32::Globalization::{MultiByteToWideChar, WideCharToMultiByte, CP_ACP, MULTI_BYTE_TO_WIDE_CHAR_FLAGS};
use windows::Win32::System::Com::{CLSIDFromProgID, CoTaskMemFree, CreateBindCtx, GetRunningObjectTable, IDispatch, IMoniker};
use windows::Win32::System::Ole::GetActiveObject;
use windows::Win32::UI::Shell::ShellExecuteW;
use windows::Win32::UI::WindowsAndMessaging::{SetForegroundWindow, SW_SHOWNORMAL};

use crate::disp::{init_com, wide, Arg, ComError, Obj};
use crate::plan::{Component, Kind, VbProject};
use crate::SyncError;

/// `vbext_ComponentType`.
const CT_STANDARD: i32 = 1;
const CT_CLASS: i32 = 2;
const CT_FORM: i32 = 3;
const CT_DOCUMENT: i32 = 100;
/// `vbext_VBAMode`: the project is neither running nor in break mode.
const VM_DESIGN: i32 = 2;
/// `vbext_ProjectProtection`: locked for viewing.
const PP_LOCKED: i32 = 1;
/// How long a macro start waits for Excel to become free.
const RUN_WAIT_MS: u128 = 30_000;
/// How long `open` waits for Excel to show the workbook.
const OPEN_WAIT_MS: u128 = 60_000;

fn same_path(a: &str, b: &str) -> bool {
    let norm = |s: &str| s.replace('/', "\\").to_lowercase();
    norm(a) == norm(b)
}

fn failed(what: &str, e: windows::core::Error) -> SyncError {
    SyncError::Failed(format!("{what}: {} (0x{:08X})", e.message(), e.code().0 as u32))
}

pub struct Book {
    app: Obj,
    workbook: Obj,
}

/// The workbook registered under `path` in the Running Object Table, in any Excel instance.
fn find_in_rot(path: &str) -> Result<Option<Obj>, SyncError> {
    unsafe {
        let rot = GetRunningObjectTable(0).map_err(|e| failed("Running Object Table を開けません", e))?;
        let monikers = rot.EnumRunning().map_err(|e| failed("Running Object Table を列挙できません", e))?;
        let ctx = CreateBindCtx(0).map_err(|e| failed("CreateBindCtx", e))?;
        loop {
            let mut item: [Option<IMoniker>; 1] = [None];
            let mut fetched = 0u32;
            if monikers.Next(&mut item, Some(&mut fetched)).is_err() || fetched == 0 {
                return Ok(None);
            }
            let Some(moniker) = item[0].take() else { continue };
            let Ok(display) = moniker.GetDisplayName(&ctx, None) else { continue };
            let name = display.to_string().unwrap_or_default();
            CoTaskMemFree(Some(display.0 as *const c_void));
            if !same_path(&name, path) {
                continue;
            }
            if let Ok(unknown) = rot.GetObject(&moniker) {
                if let Ok(ptr) = unknown.cast::<IDispatch>() {
                    return Ok(Some(Obj(ptr)));
                }
            }
        }
    }
}

/// Workbooks in a OneDrive / SharePoint folder are not registered under their local path:
/// the active Excel is asked for a workbook of that file name (spec §14.10).
fn find_by_name(path: &Path) -> Option<Obj> {
    let file_name = path.file_name()?.to_str()?;
    let prog_id = wide("Excel.Application");
    let app = unsafe {
        let clsid = CLSIDFromProgID(PCWSTR(prog_id.as_ptr())).ok()?;
        let mut unknown = None;
        GetActiveObject(&clsid, None, &mut unknown).ok()?;
        Obj(unknown?.cast::<IDispatch>().ok()?)
    };
    let workbook = app.object("Workbooks", &[]).ok()?.object("Item", &[Arg::Str(file_name)]).ok()?;
    let full_name = workbook.get("FullName", &[]).ok()?.string();
    let is_url = full_name.to_lowercase().starts_with("http");
    (is_url || same_path(&full_name, &path.to_string_lossy())).then_some(workbook)
}

/// Finds the open workbook; `None` when no Excel has it open.
pub fn find(path: &Path) -> Result<Option<Book>, SyncError> {
    init_com();
    let workbook = match find_in_rot(&path.to_string_lossy())? {
        Some(w) => Some(w),
        None => find_by_name(path),
    };
    let Some(workbook) = workbook else { return Ok(None) };
    let app = workbook.object("Application", &[])?;
    Ok(Some(Book { app, workbook }))
}

/// Opens the workbook the way a double click does, and waits until Excel shows it.
pub fn open(path: &Path) -> Result<Book, SyncError> {
    init_com();
    if !path.is_file() {
        return Err(SyncError::Failed(format!("{} が見つかりません", path.display())));
    }
    let file = wide(&path.to_string_lossy());
    let verb = wide("open");
    let instance = unsafe { ShellExecuteW(None, PCWSTR(verb.as_ptr()), PCWSTR(file.as_ptr()), PCWSTR::null(), PCWSTR::null(), SW_SHOWNORMAL) };
    if instance.0 as usize <= 32 {
        return Err(SyncError::Failed(format!("{} を開けませんでした（Excel がインストールされているか確認してください）", path.display())));
    }
    let started = std::time::Instant::now();
    while started.elapsed().as_millis() < OPEN_WAIT_MS {
        std::thread::sleep(std::time::Duration::from_millis(500));
        if let Some(book) = find(path)? {
            return Ok(book);
        }
    }
    Err(SyncError::Failed(format!("{} が Excel で開かれたことを確認できませんでした（保護ビューで開いた場合は「編集を有効にする」を押してください）", path.display())))
}

impl Book {
    /// False once the workbook was closed or Excel is gone. A busy Excel still has it.
    pub fn is_alive(&self) -> bool {
        match self.workbook.get("Name", &[]) {
            Ok(_) => true,
            Err(e) => e.is_busy(),
        }
    }

    /// The VBA project, when it can be written to right now.
    pub fn project(&self) -> Result<Project, SyncError> {
        let project = self.workbook.object("VBProject", &[]).map_err(|e| {
            if e.is_busy() {
                return SyncError::from(e);
            }
            SyncError::Failed(format!(
                "VBA プロジェクトにアクセスできません。Excel の［ファイル］→［オプション］→［トラスト センター］→［トラスト センターの設定］→［マクロの設定］で「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を有効にしてください。有効になっている場合は、プログラムから起動された Excel（以前のバージョンの RAS のデバッグ実行が起動したものなど）の可能性があります。Excel を終了し、ブックを開き直してください（{}）",
                e.description
            ))
        })?;
        if project.get("Protection", &[])?.int() == Some(PP_LOCKED) {
            return Err(SyncError::Failed("VBA プロジェクトがパスワードで保護されているため、同期できません".into()));
        }
        // Rewriting code under a running or halted macro resets its state or crashes it (spec §14.6).
        if project.get("Mode", &[])?.int() != Some(VM_DESIGN) {
            return Err(SyncError::Busy("マクロの実行中か、中断モードです".into()));
        }
        let book = self.workbook.get("Name", &[])?.string();
        Ok(Project { components: project.object("VBComponents", &[])?, app: self.app.clone(), book })
    }

    /// The code name of the workbook (`ThisWorkbook`).
    pub fn code_name(&self) -> Result<String, SyncError> {
        Ok(self.workbook.get("CodeName", &[])?.string())
    }

    /// `Application.Run` with the workbook-qualified name (spec §14.5); returns when the macro ends.
    pub fn run(&self, name: &str) -> Result<(), SyncError> {
        let book = self.workbook.get("Name", &[])?.string();
        let qualified = format!("'{}'!{name}", book.replace('\'', "''"));
        self.app.call_patiently("Run", &[Arg::Str(&qualified)], RUN_WAIT_MS).map_err(|e| {
            if e.is_busy() {
                SyncError::from(e)
            } else {
                SyncError::Failed(format!("マクロ `{name}` を実行できませんでした: {}", e.description))
            }
        })?;
        Ok(())
    }

    /// Shows the user form `name` (`RAS_Runtime.RAS_ShowForm`) with Excel in front; returns
    /// when a modal form is closed, or at once for a modeless one.
    pub fn show_form(&self, name: &str) -> Result<(), SyncError> {
        self.activate();
        let book = self.workbook.get("Name", &[])?.string();
        let qualified = format!("'{}'!RAS_Runtime.RAS_ShowForm", book.replace('\'', "''"));
        // A modal form keeps the call waiting for as long as it is open.
        self.app.call_patiently("Run", &[Arg::Str(&qualified), Arg::Str(name)], RUN_WAIT_MS).map_err(|e| {
            if e.is_busy() {
                SyncError::from(e)
            } else {
                SyncError::Failed(format!("ユーザーフォーム `{name}` を表示できませんでした: {}", e.description))
            }
        })?;
        Ok(())
    }

    /// Shows Excel with the workbook in front. Windows may refuse to change the foreground window.
    pub fn activate(&self) {
        let _ = self.app.put("Visible", Arg::Bool(true));
        let _ = self.workbook.call("Activate", &[]);
        if let Some(handle) = self.app.get("Hwnd", &[]).ok().and_then(|v| v.int()) {
            unsafe {
                let _ = SetForegroundWindow(HWND(handle as isize as *mut c_void));
            }
        }
    }
}

pub struct Project {
    components: Obj,
    /// `Application`, to run the helper that unloads forms.
    app: Obj,
    /// `Workbook.Name`, to qualify that helper.
    book: String,
}

/// Name of the temporary module that unloads forms before they are replaced.
const UNLOAD_MODULE: &str = "RAS_SyncUnload";

/// VBA that unloads every loaded instance of the forms `names` and returns how many are still
/// loaded (a `QueryClose` handler can refuse to close).
fn unload_code(names: &[String]) -> String {
    let test = names.iter().map(|n| format!("n = \"{}\"", n.to_lowercase().replace('"', "\"\""))).collect::<Vec<_>>().join(" Or ");
    format!(
        "Public Function RAS_UnloadForms() As Long\r\n\
         \x20   Dim i As Long, n As String\r\n\
         \x20   On Error Resume Next\r\n\
         \x20   For i = VBA.UserForms.Count - 1 To 0 Step -1\r\n\
         \x20       n = LCase$(TypeName(VBA.UserForms(i)))\r\n\
         \x20       If {test} Then Unload VBA.UserForms(i)\r\n\
         \x20   Next\r\n\
         \x20   For i = 0 To VBA.UserForms.Count - 1\r\n\
         \x20       n = LCase$(TypeName(VBA.UserForms(i)))\r\n\
         \x20       If {test} Then RAS_UnloadForms = RAS_UnloadForms + 1\r\n\
         \x20   Next\r\n\
         End Function\r\n"
    )
}

impl Project {
    fn component(&self, name: &str) -> Result<Obj, ComError> {
        self.components.object("Item", &[Arg::Str(name)])
    }
}

/// Text in the ANSI code page of this Windows, which is what `VBComponents.Import` reads (spec §14.6).
fn to_ansi(text: &str) -> Vec<u8> {
    let units: Vec<u16> = text.encode_utf16().collect();
    if units.is_empty() {
        return Vec::new();
    }
    unsafe {
        let len = WideCharToMultiByte(CP_ACP, 0, &units, None, windows::core::PCSTR::null(), None);
        let mut out = vec![0u8; len.max(0) as usize];
        let written = WideCharToMultiByte(CP_ACP, 0, &units, Some(&mut out), windows::core::PCSTR::null(), None);
        out.truncate(written.max(0) as usize);
        out
    }
}

/// Text in the ANSI code page of this Windows (what `VBComponent.Export` writes), decoded.
fn from_ansi(bytes: &[u8]) -> String {
    if bytes.is_empty() {
        return String::new();
    }
    unsafe {
        let len = MultiByteToWideChar(CP_ACP, MULTI_BYTE_TO_WIDE_CHAR_FLAGS(0), bytes, None);
        let mut out = vec![0u16; len.max(0) as usize];
        let written = MultiByteToWideChar(CP_ACP, MULTI_BYTE_TO_WIDE_CHAR_FLAGS(0), bytes, Some(&mut out));
        out.truncate(written.max(0) as usize);
        String::from_utf16_lossy(&out)
    }
}

impl Project {
    /// The user form `name` as `VBComponent.Export` writes it (`.frm` / `.frx`), read back with
    /// its designer storage. `codepage` is the code page its designer properties are stored in.
    pub fn export_form(&mut self, name: &str, codepage: u16) -> Result<ras_ovba::UserForm, SyncError> {
        let dir = std::env::temp_dir().join(format!("ras-sync-export-{}", std::process::id()));
        let _ = std::fs::remove_dir_all(&dir);
        std::fs::create_dir_all(&dir).map_err(|e| SyncError::Failed(format!("{}: {e}", dir.display())))?;
        let frm = dir.join(format!("{name}.frm"));
        let result = (|| {
            self.component(name)?.call("Export", &[Arg::Str(&frm.to_string_lossy())]).map_err(|e| {
                if e.is_busy() {
                    SyncError::from(e)
                } else {
                    SyncError::Failed(format!("ユーザーフォーム `{name}` を書き出せませんでした: {}", e.description))
                }
            })?;
            let text = std::fs::read(&frm).map_err(|e| SyncError::Failed(format!("{}: {e}", frm.display())))?;
            let frx_path = frm.with_extension("frx");
            let frx = std::fs::read(&frx_path).map_err(|e| SyncError::Failed(format!("{}: {e}", frx_path.display())))?;
            ras_ovba::import_user_form(&from_ansi(&text), &frx, codepage).map_err(|e| SyncError::Failed(format!("ユーザーフォーム `{name}` を読み取れませんでした: {e}")))
        })();
        let _ = std::fs::remove_dir_all(&dir);
        result
    }
}

impl VbProject for Project {
    fn components(&mut self) -> Result<Vec<Component>, SyncError> {
        let count = self.components.get("Count", &[])?.int().unwrap_or(0);
        let mut out = Vec::with_capacity(count as usize);
        for i in 1..=count {
            let c = self.components.object("Item", &[Arg::Int(i)])?;
            let kind = match c.get("Type", &[])?.int() {
                Some(CT_STANDARD) => Kind::Standard,
                Some(CT_CLASS) => Kind::Class,
                Some(CT_FORM) => Kind::Form,
                Some(CT_DOCUMENT) => Kind::Document,
                _ => Kind::Other,
            };
            let module = c.object("CodeModule", &[])?;
            let lines = module.get("CountOfLines", &[])?.int().unwrap_or(0);
            let code = if lines > 0 { module.get("Lines", &[Arg::Int(1), Arg::Int(lines)])?.string() } else { String::new() };
            out.push(Component { name: c.get("Name", &[])?.string(), kind, code });
        }
        Ok(out)
    }

    fn set_code(&mut self, name: &str, code: &str) -> Result<(), SyncError> {
        let module = self.component(name)?.object("CodeModule", &[])?;
        // Also for a new module: it starts with `Option Explicit` when the VBE is set to require declarations.
        let lines = module.get("CountOfLines", &[])?.int().unwrap_or(0);
        if lines > 0 {
            module.call("DeleteLines", &[Arg::Int(1), Arg::Int(lines)])?;
        }
        let text = code.trim_end_matches(['\r', '\n']);
        if !text.is_empty() {
            module.call("AddFromString", &[Arg::Str(text)])?;
        }
        Ok(())
    }

    fn add(&mut self, name: &str, kind: Kind) -> Result<(), SyncError> {
        let component = self.components.object("Add", &[Arg::Int(if kind == Kind::Class { CT_CLASS } else { CT_STANDARD })])?;
        if let Err(e) = component.put("Name", Arg::Str(name)) {
            let _ = self.components.call("Remove", &[Arg::Obj(&component)]);
            return Err(SyncError::Failed(format!("モジュール `{name}` を作成できませんでした: {}", e.description)));
        }
        Ok(())
    }

    fn import(&mut self, name: &str, file_text: &str) -> Result<(), SyncError> {
        let dir = std::env::temp_dir().join(format!("ras-sync-{}", std::process::id()));
        std::fs::create_dir_all(&dir).map_err(|e| SyncError::Failed(format!("{}: {e}", dir.display())))?;
        let file = dir.join(format!("{name}.cls"));
        // A `.cls` with LF only is imported as a standard module.
        let text = file_text.replace("\r\n", "\n").replace('\n', "\r\n");
        std::fs::write(&file, to_ansi(&text)).map_err(|e| SyncError::Failed(format!("{}: {e}", file.display())))?;
        let imported = self.components.object("Import", &[Arg::Str(&file.to_string_lossy())]);
        let _ = std::fs::remove_dir_all(&dir);
        let component = imported?;
        let actual = component.get("Name", &[])?.string();
        if actual != name {
            // A component of that name was still there, so the VBE numbered the new one.
            let _ = self.components.call("Remove", &[Arg::Obj(&component)]);
            return Err(SyncError::Failed(format!("クラス `{name}` を取り込めませんでした（`{actual}` という名前になりました）")));
        }
        Ok(())
    }

    fn import_form(&mut self, name: &str, files: &ras_ovba::FormFiles) -> Result<(), SyncError> {
        let dir = std::env::temp_dir().join(format!("ras-sync-{}", std::process::id()));
        std::fs::create_dir_all(&dir).map_err(|e| SyncError::Failed(format!("{}: {e}", dir.display())))?;
        // The `.frm` names its `.frx` by file name: both lie in the same directory.
        let frm = dir.join(format!("{name}.frm"));
        let written = std::fs::write(&frm, to_ansi(&files.frm)).and_then(|()| std::fs::write(dir.join(format!("{name}.frx")), &files.frx));
        if let Err(e) = written {
            let _ = std::fs::remove_dir_all(&dir);
            return Err(SyncError::Failed(format!("{}: {e}", frm.display())));
        }
        let imported = self.components.object("Import", &[Arg::Str(&frm.to_string_lossy())]);
        let _ = std::fs::remove_dir_all(&dir);
        let component = imported.map_err(|e| if e.is_busy() { SyncError::from(e) } else { SyncError::Failed(format!("ユーザーフォーム `{name}` を取り込めませんでした: {}", e.description)) })?;
        let actual = component.get("Name", &[])?.string();
        let is_form = component.get("Type", &[])?.int() == Some(CT_FORM);
        if actual != name || !is_form {
            // A component of that name was still there (the VBE numbered the new one), or the files were not taken as a form.
            let _ = self.components.call("Remove", &[Arg::Obj(&component)]);
            return Err(SyncError::Failed(format!("ユーザーフォーム `{name}` を取り込めませんでした（`{actual}` という{}になりました）", if is_form { "名前" } else { "フォームではないモジュール" })));
        }
        Ok(())
    }

    fn unload_forms(&mut self, names: &[String]) -> Result<(), SyncError> {
        // A designer open in the VBE.
        for name in names {
            let Ok(component) = self.component(name) else { continue };
            if component.get("HasOpenDesigner", &[]).ok().and_then(|v| v.int()).is_some_and(|v| v != 0) {
                if let Ok(window) = component.object("DesignerWindow", &[]) {
                    let _ = window.call("Close", &[]);
                }
            }
        }
        // Loaded instances (a form closed with `Hide`, a modeless form) live in `VBA.UserForms`,
        // which only code of the project sees: a temporary module unloads them.
        if let Ok(stale) = self.component(UNLOAD_MODULE) {
            let _ = self.components.call("Remove", &[Arg::Obj(&stale)]);
        }
        let module = self.components.object("Add", &[Arg::Int(CT_STANDARD)])?;
        let result = (|| -> Result<i32, SyncError> {
            module.put("Name", Arg::Str(UNLOAD_MODULE))?;
            let code = module.object("CodeModule", &[])?;
            code.call("AddFromString", &[Arg::Str(&unload_code(names))])?;
            let qualified = format!("'{}'!{UNLOAD_MODULE}.RAS_UnloadForms", self.book.replace('\'', "''"));
            let left = self.app.call_patiently("Run", &[Arg::Str(&qualified)], RUN_WAIT_MS).map_err(|e| {
                if e.is_busy() {
                    SyncError::from(e)
                } else {
                    SyncError::Failed(format!("読み込まれているユーザーフォームを閉じられませんでした: {}", e.description))
                }
            })?;
            Ok(left.int().unwrap_or(0))
        })();
        let _ = self.components.call("Remove", &[Arg::Obj(&module)]);
        match result? {
            0 => Ok(()),
            _ => Err(SyncError::Failed(format!(
                "ユーザーフォーム {} が読み込まれたままで閉じられないため、更新できません（QueryClose で閉じるのを取り消していないか確認してください。フォームを閉じると同期します）",
                names.join(", ")
            ))),
        }
    }

    fn rename(&mut self, name: &str, new_name: &str) -> Result<(), SyncError> {
        self.component(name)?.put("Name", Arg::Str(new_name)).map_err(|e| if e.is_busy() { SyncError::from(e) } else { SyncError::Failed(format!("`{name}` の名前を変更できませんでした: {}", e.description)) })
    }

    fn remove(&mut self, name: &str) -> Result<(), SyncError> {
        let component = self.component(name)?;
        self.components.call("Remove", &[Arg::Obj(&component)])?;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn unload_code_names_the_forms() {
        let code = unload_code(&["frmInput".to_string(), "UserForm1".to_string()]);
        assert!(code.contains("If n = \"frminput\" Or n = \"userform1\" Then Unload VBA.UserForms(i)\r\n"), "{code}");
        assert!(code.starts_with("Public Function RAS_UnloadForms() As Long\r\n    Dim i As Long"), "{code}");
    }

    #[test]
    fn paths_and_ansi_text() {
        assert!(same_path("C:/Work/Book.xlsm", "c:\\work\\book.xlsm"));
        assert!(!same_path("C:\\Work\\Book.xlsm", "C:\\Work\\Other.xlsm"));
        assert_eq!(to_ansi("Sub A()\r\n"), b"Sub A()\r\n");
        assert!(to_ansi("").is_empty());
        assert_eq!(from_ansi(b"Sub A()\r\n"), "Sub A()\r\n");
        assert_eq!(from_ansi(&to_ansi("Caption = \"x\"")), "Caption = \"x\"");
        assert!(from_ansi(b"").is_empty());
    }

    /// No Excel on the build machines: looking for a workbook must simply find none.
    #[test]
    fn a_workbook_that_is_not_open_is_not_found() {
        let missing = std::env::temp_dir().join("ras-sync-no-such-book.xlsm");
        assert!(find(&missing).unwrap().is_none());
        assert!(matches!(open(&missing), Err(SyncError::Failed(_))));
    }
}
