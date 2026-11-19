//! Late-binding `IDispatch` calls, as far as the sync needs them: strings, integers, booleans and objects.

use std::mem::ManuallyDrop;

use windows::core::{BSTR, GUID, PCWSTR};
use windows::Win32::Foundation::VARIANT_BOOL;
use windows::Win32::System::Com::{CoInitializeEx, IDispatch, COINIT_APARTMENTTHREADED, DISPATCH_FLAGS, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO};
use windows::Win32::System::Variant::{VariantClear, VARENUM, VARIANT, VT_BOOL, VT_BSTR, VT_DISPATCH, VT_I4};

use crate::SyncError;

/// VBA always calls with English (United States); other locales make Excel reinterpret names and values.
const LCID: u32 = 0x0409;
const DISPID_PROPERTYPUT: i32 = -3;
const DISP_E_EXCEPTION: i32 = 0x8002_0009_u32 as i32;
/// Excel refuses calls while a cell is edited or a dialog is open: `RPC_E_CALL_REJECTED`,
/// `RPC_E_SERVERCALL_RETRYLATER`, and `VBA_E_IGNORE` from the object model itself.
const BUSY: [i32; 3] = [0x8001_0001_u32 as i32, 0x8001_010A_u32 as i32, 0x800A_C472_u32 as i32];
/// How long a call is retried while Excel is busy before the sync is put off.
const BUSY_TIMEOUT_MS: u128 = 3_000;

/// A failed call. `number` is the VBA error number when the callee raised one.
#[derive(Debug, Clone)]
pub struct ComError {
    pub hr: i32,
    pub number: i32,
    pub description: String,
}

impl ComError {
    pub fn is_busy(&self) -> bool {
        BUSY.contains(&self.hr)
    }
}

impl From<ComError> for SyncError {
    fn from(e: ComError) -> SyncError {
        if e.is_busy() {
            SyncError::Busy("セルの編集中か、ダイアログが開いています".into())
        } else {
            SyncError::Failed(format!("Excel の呼び出しに失敗しました: {}", e.description))
        }
    }
}

impl From<windows::core::Error> for ComError {
    fn from(e: windows::core::Error) -> ComError {
        ComError { hr: e.code().0, number: 0, description: format!("{} (0x{:08X})", e.message(), e.code().0 as u32) }
    }
}

/// Initialises COM on this thread (once). The thread keeps its apartment until it ends.
pub fn init_com() {
    thread_local! {
        static INIT: () = unsafe {
            let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        };
    }
    INIT.with(|_| ());
}

pub fn wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}

/// An owned `VARIANT`.
#[repr(transparent)]
pub struct Value(VARIANT);

impl Drop for Value {
    fn drop(&mut self) {
        unsafe {
            let _ = VariantClear(&mut self.0);
        }
    }
}

impl Value {
    fn with(vt: VARENUM, set: impl FnOnce(&mut VARIANT)) -> Value {
        let mut v = VARIANT::default();
        unsafe { (*v.Anonymous.Anonymous).vt = vt };
        set(&mut v);
        Value(v)
    }

    fn vt(&self) -> u16 {
        unsafe { self.0.Anonymous.Anonymous.vt.0 }
    }

    /// The string value; empty for anything else.
    pub fn string(&self) -> String {
        if self.vt() != VT_BSTR.0 {
            return String::new();
        }
        unsafe { self.0.Anonymous.Anonymous.Anonymous.bstrVal.to_string() }
    }

    pub fn int(&self) -> Option<i32> {
        unsafe {
            let data = &self.0.Anonymous.Anonymous.Anonymous;
            match self.vt() {
                2 => Some(data.iVal as i32),
                3 | 22 => Some(data.lVal),
                5 => Some(data.dblVal as i32),
                11 => Some(data.boolVal.0 as i32),
                17 => Some(data.bVal as i32),
                18 => Some(data.uiVal as i32),
                19 | 23 => Some(data.ulVal as i32),
                20 => Some(data.llVal as i32),
                _ => None,
            }
        }
    }

    pub fn object(&self) -> Option<Obj> {
        if self.vt() != VT_DISPATCH.0 {
            return None;
        }
        unsafe { (*self.0.Anonymous.Anonymous.Anonymous.pdispVal).clone().map(Obj) }
    }
}

pub enum Arg<'a> {
    Str(&'a str),
    Int(i32),
    Bool(bool),
    Obj(&'a Obj),
}

impl Arg<'_> {
    fn to_value(&self) -> Value {
        match self {
            Arg::Str(s) => {
                let b = BSTR::from(*s);
                Value::with(VT_BSTR, |v| unsafe { (*v.Anonymous.Anonymous).Anonymous.bstrVal = ManuallyDrop::new(b) })
            }
            Arg::Int(i) => Value::with(VT_I4, |v| unsafe { (*v.Anonymous.Anonymous).Anonymous.lVal = *i }),
            Arg::Bool(b) => Value::with(VT_BOOL, |v| unsafe { (*v.Anonymous.Anonymous).Anonymous.boolVal = VARIANT_BOOL(if *b { -1 } else { 0 }) }),
            Arg::Obj(o) => {
                let d = o.0.clone();
                Value::with(VT_DISPATCH, |v| unsafe { (*v.Anonymous.Anonymous).Anonymous.pdispVal = ManuallyDrop::new(Some(d)) })
            }
        }
    }
}

/// A COM object that is driven by name.
#[derive(Clone)]
pub struct Obj(pub IDispatch);

impl Obj {
    pub fn get(&self, name: &str, args: &[Arg]) -> Result<Value, ComError> {
        self.invoke(name, DISPATCH_PROPERTYGET | DISPATCH_METHOD, args, false, BUSY_TIMEOUT_MS)
    }

    /// The property as an object; an error when it is `Nothing` or not an object.
    pub fn object(&self, name: &str, args: &[Arg]) -> Result<Obj, ComError> {
        self.get(name, args)?.object().ok_or_else(|| ComError { hr: 0, number: 91, description: format!("`{name}` はオブジェクトを返しませんでした") })
    }

    pub fn call(&self, name: &str, args: &[Arg]) -> Result<Value, ComError> {
        self.invoke(name, DISPATCH_METHOD | DISPATCH_PROPERTYGET, args, false, BUSY_TIMEOUT_MS)
    }

    /// A call that is retried for as long as `wait_ms` while Excel is busy.
    pub fn call_patiently(&self, name: &str, args: &[Arg], wait_ms: u128) -> Result<Value, ComError> {
        self.invoke(name, DISPATCH_METHOD | DISPATCH_PROPERTYGET, args, false, wait_ms)
    }

    pub fn put(&self, name: &str, value: Arg) -> Result<(), ComError> {
        self.invoke(name, DISPATCH_PROPERTYPUT, &[value], true, BUSY_TIMEOUT_MS).map(|_| ())
    }

    fn invoke(&self, name: &str, flags: DISPATCH_FLAGS, args: &[Arg], put: bool, wait_ms: u128) -> Result<Value, ComError> {
        let w = wide(name);
        let names = [PCWSTR(w.as_ptr())];
        let mut id = 0i32;
        unsafe { self.0.GetIDsOfNames(&GUID::zeroed(), names.as_ptr(), 1, LCID, &mut id) }.map_err(|e| ComError { description: format!("{name}: {}", e.message()), ..ComError::from(e) })?;
        // Arguments are passed last to first; the vector owns them for the duration of the call.
        let mut rg: Vec<Value> = args.iter().rev().map(Arg::to_value).collect();
        let mut named = DISPID_PROPERTYPUT;
        let params = DISPPARAMS {
            rgvarg: if rg.is_empty() { std::ptr::null_mut() } else { rg.as_mut_ptr() as *mut VARIANT },
            rgdispidNamedArgs: if put { &mut named } else { std::ptr::null_mut() },
            cArgs: rg.len() as u32,
            cNamedArgs: if put { 1 } else { 0 },
        };
        let started = std::time::Instant::now();
        let mut delay = 20u64;
        loop {
            let mut result = Value(VARIANT::default());
            let mut ex = EXCEPINFO::default();
            let mut arg_err = 0u32;
            let r = unsafe { self.0.Invoke(id, &GUID::zeroed(), LCID, flags, &params, Some(&mut result.0), Some(&mut ex), Some(&mut arg_err)) };
            let e = match r {
                Ok(()) => return Ok(result),
                Err(e) => e,
            };
            let hr = e.code().0;
            if hr == DISP_E_EXCEPTION {
                let (source, description, scode) = unsafe {
                    if let Some(fill) = ex.pfnDeferredFillIn {
                        let _ = fill(&mut ex);
                    }
                    let source = ManuallyDrop::take(&mut ex.bstrSource).to_string();
                    let description = ManuallyDrop::take(&mut ex.bstrDescription).to_string();
                    drop(ManuallyDrop::take(&mut ex.bstrHelpFile));
                    (source, description, ex.scode)
                };
                let error = ComError {
                    hr: scode,
                    number: if ex.wCode != 0 { ex.wCode as i32 } else { scode & 0xFFFF },
                    description: if description.is_empty() { format!("{name}: {source} (0x{:08X})", scode as u32) } else { description },
                };
                if !error.is_busy() {
                    return Err(error);
                }
            } else if !BUSY.contains(&hr) {
                return Err(ComError { description: format!("{name}: {}", e.message()), ..ComError::from(e) });
            }
            if started.elapsed().as_millis() >= wait_ms {
                return Err(ComError { hr: BUSY[0], number: 0, description: "Excel が応答しません（セルの編集中か、ダイアログが開いています）".into() });
            }
            std::thread::sleep(std::time::Duration::from_millis(delay));
            delay = (delay * 2).min(400);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use windows::Win32::System::Com::{CLSIDFromProgID, CoCreateInstance, CLSCTX_INPROC_SERVER};

    /// Drives a COM object that exists on every Windows machine.
    #[test]
    fn calls_a_real_object() {
        init_com();
        let prog_id = wide("Scripting.Dictionary");
        let dict = unsafe {
            let clsid = CLSIDFromProgID(PCWSTR(prog_id.as_ptr())).unwrap();
            Obj(CoCreateInstance::<_, IDispatch>(&clsid, None, CLSCTX_INPROC_SERVER).unwrap())
        };
        dict.call("Add", &[Arg::Str("名前"), Arg::Str("値 ①")]).unwrap();
        dict.call("Add", &[Arg::Str("n"), Arg::Int(42)]).unwrap();
        assert_eq!(dict.get("Count", &[]).unwrap().int(), Some(2));
        assert_eq!(dict.get("Item", &[Arg::Str("名前")]).unwrap().string(), "値 ①");
        assert_eq!(dict.get("Item", &[Arg::Str("n")]).unwrap().int(), Some(42));
        assert_eq!(dict.call("Exists", &[Arg::Str("n")]).unwrap().int(), Some(-1));
        assert!(dict.put("CompareMode", Arg::Int(0)).is_err(), "not allowed once the dictionary has items");
        let missing = dict.call("NoSuchMember", &[]).err().unwrap();
        assert!(!missing.is_busy() && missing.description.contains("NoSuchMember"), "{missing:?}");
        assert!(dict.call("Remove", &[Arg::Str("missing key")]).is_err());
        assert!(dict.object("Count", &[]).is_err());
    }
}
