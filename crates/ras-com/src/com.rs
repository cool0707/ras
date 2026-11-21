//! The late-binding COM bridge of the addon.
//!
//! It is a thin bridge: property get / put, method call, default member and
//! `For Each` on an `IDispatch`, plus the lookups needed to find the running Excel.
//! Everything else (type conversions with RAS semantics, error objects, application state)
//! lives in the JavaScript run time.
//!
//! Values cross the boundary as follows:
//!
//! | JavaScript                      | VARIANT                                  |
//! |---------------------------------|------------------------------------------|
//! | `undefined`                     | missing optional argument                |
//! | `null`                          | `Empty` (and `Null`, `Nothing` on read)  |
//! | boolean / string                | `Boolean` / `String`                     |
//! | number                          | `Long` when integral and in range, else `Double` |
//! | `ComObject`                     | object                                   |
//! | array, array of arrays          | 1-D / 2-D array of `Variant` (0-based)   |
//! | `{ $t: "date", v }`             | `Date`                                   |
//! | `{ $t: "error", v }`            | `Error` (`CVErr`)                        |
//! | `{ $t: "nothing" }`             | `Nothing`                                |
//!
//! Errors are thrown with a message of the form `RASCOM{"number":…,"source":…,"description":…}`.

use std::ffi::c_void;
use std::mem::ManuallyDrop;

use napi::bindgen_prelude::{ClassInstance, FromNapiValue};
use napi::{Env, Error, JsBoolean, JsNumber, JsObject, JsString, JsUnknown, NapiRaw, NapiValue, Result, Status, ValueType};
use napi_derive::napi;
use windows::core::{Interface, BSTR, GUID, HRESULT, PCWSTR};
use windows::Win32::Foundation::VARIANT_BOOL;
use windows::Win32::System::Com::{
    CLSIDFromProgID, CoCreateInstance, CoGetObject, CoInitializeEx, CoTaskMemFree, CreateBindCtx, GetRunningObjectTable, IDispatch, IMoniker, CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER,
    COINIT_APARTMENTTHREADED, DISPATCH_FLAGS, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPATCH_PROPERTYPUTREF, DISPPARAMS, EXCEPINFO, FUNCDESC, FUNCFLAG_FHIDDEN,
    FUNCFLAG_FRESTRICTED, IMPLTYPEFLAG_FDEFAULT, IMPLTYPEFLAG_FSOURCE, INVOKE_PROPERTYGET, ITypeInfo, SAFEARRAY, SAFEARRAYBOUND, TKIND_COCLASS, TKIND_ENUM, TKIND_INTERFACE, TYPEATTR, TYPEDESC, VARFLAG_FHIDDEN, VARFLAG_FRESTRICTED, VAR_DISPATCH,
};
use windows::Win32::System::Ole::{
    GetActiveObject, IEnumVARIANT, IProvideClassInfo, SafeArrayCreate, SafeArrayGetDim, SafeArrayGetElement, SafeArrayGetLBound, SafeArrayGetUBound, SafeArrayGetVartype, SafeArrayPutElement, VarR8FromDec,
    PARAMFLAG_FLCID, PARAMFLAG_FOPT, PARAMFLAG_FRETVAL, TYPEFLAG_FDUAL,
};
use windows::Win32::System::Variant::{
    VariantClear, VARENUM, VARIANT, VT_ARRAY, VT_BOOL, VT_BSTR, VT_BYREF, VT_CY, VT_DATE, VT_DECIMAL, VT_DISPATCH, VT_ERROR, VT_I1, VT_I2, VT_I4, VT_I8, VT_INT, VT_R4, VT_R8, VT_UI1, VT_UI2, VT_UI4,
    VT_UI8, VT_UINT, VT_USERDEFINED, VT_VARIANT,
};

/// Excel interprets names and values by the locale passed to `Invoke`. VBA always uses
/// English (United States), which also avoids the "old format or invalid type library" error.
const LCID: u32 = 0x0409;
const DISPID_VALUE: i32 = 0;
const DISPID_PROPERTYPUT: i32 = -3;
const DISPID_NEWENUM: i32 = -4;

const DISP_E_MEMBERNOTFOUND: i32 = 0x8002_0003_u32 as i32;
const DISP_E_PARAMNOTFOUND: i32 = 0x8002_0004_u32 as i32;
const DISP_E_TYPEMISMATCH: i32 = 0x8002_0005_u32 as i32;
const DISP_E_UNKNOWNNAME: i32 = 0x8002_0006_u32 as i32;
const DISP_E_EXCEPTION: i32 = 0x8002_0009_u32 as i32;
const DISP_E_OVERFLOW: i32 = 0x8002_000A_u32 as i32;
const DISP_E_BADINDEX: i32 = 0x8002_000B_u32 as i32;
const DISP_E_BADPARAMCOUNT: i32 = 0x8002_000E_u32 as i32;
const DISP_E_PARAMNOTOPTIONAL: i32 = 0x8002_000F_u32 as i32;
const DISP_E_DIVBYZERO: i32 = 0x8002_0012_u32 as i32;
// What VBA makes of the general HRESULTs (`Border.ThemeColor` of a border without one: 5).
const E_INVALIDARG: i32 = 0x8007_0057_u32 as i32;
const E_OUTOFMEMORY: i32 = 0x8007_000E_u32 as i32;
const E_ACCESSDENIED: i32 = 0x8007_0005_u32 as i32;
const E_NOTIMPL: i32 = 0x8000_4001_u32 as i32;
const E_FILENOTFOUND: i32 = 0x8007_0002_u32 as i32;
const E_PATHNOTFOUND: i32 = 0x8007_0003_u32 as i32;
const RPC_E_CALL_REJECTED: i32 = 0x8001_0001_u32 as i32;
const RPC_E_SERVERCALL_RETRYLATER: i32 = 0x8001_010A_u32 as i32;
const RPC_E_CHANGED_MODE: i32 = 0x8001_0106_u32 as i32;
/// The server process is gone: `RPC_E_DISCONNECTED`, `RPC_S_SERVER_UNAVAILABLE`, `RPC_S_CALL_FAILED`, `CO_E_OBJNOTCONNECTED`.
const SERVER_GONE: [i32; 4] = [0x8001_0108_u32 as i32, 0x8007_06BA_u32 as i32, 0x8007_06BE_u32 as i32, 0x8004_01FD_u32 as i32];
/// How long a call waits while Excel is busy (cell being edited, modal dialog) before it fails.
const BUSY_TIMEOUT_MS: u64 = 60_000;

// ---------------------------------------------------------------------------- errors

fn json_string(s: &str) -> String {
    let mut out = String::with_capacity(s.len() + 2);
    out.push('"');
    for c in s.chars() {
        match c {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if (c as u32) < 0x20 => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out.push('"');
    out
}

/// An error in the form the JavaScript run time turns into a RAS error object.
fn com_error(number: i32, source: &str, description: &str) -> Error {
    Error::new(Status::GenericFailure, format!("RASCOM{{\"number\":{number},\"source\":{},\"description\":{}}}", json_string(source), json_string(description)))
}

/// The VBA error number VBA reports for a standard dispatch HRESULT; other values unchanged.
fn vba_number(hr: i32) -> i32 {
    match hr {
        DISP_E_UNKNOWNNAME | DISP_E_MEMBERNOTFOUND => 438,
        DISP_E_TYPEMISMATCH => 13,
        DISP_E_OVERFLOW => 6,
        DISP_E_BADINDEX => 9,
        DISP_E_BADPARAMCOUNT => 450,
        DISP_E_PARAMNOTFOUND => 448,
        DISP_E_PARAMNOTOPTIONAL => 449,
        DISP_E_DIVBYZERO => 11,
        E_INVALIDARG => 5,
        E_OUTOFMEMORY => 7,
        E_ACCESSDENIED => 70,
        E_NOTIMPL => 445,
        E_FILENOTFOUND => 53,
        E_PATHNOTFOUND => 76,
        other => other,
    }
}

/// VBA error number of a failed call that did not raise an exception of its own.
fn hresult_error(hr: HRESULT, what: &str) -> Error {
    if hr.0 == RPC_E_CALL_REJECTED || hr.0 == RPC_E_SERVERCALL_RETRYLATER {
        let mut message = String::from("Excel が応答しません（セルの編集中か、ダイアログが開いています）。Excel 側の操作を終えてからもう一度実行してください");
        if let Some(dialogs) = crate::watch::open_dialog_summary() {
            message.push_str(&format!("。開いているダイアログ: {dialogs}"));
        }
        return com_error(hr.0, "RAS", &message);
    }
    if SERVER_GONE.contains(&hr.0) {
        return com_error(hr.0, "RAS", "Excel との接続が切れました（Excel が終了したか、応答しなくなって強制終了されました）。デバッグをやり直してください");
    }
    let number = vba_number(hr.0);
    let detail = hr.message();
    com_error(number, "", &if what.is_empty() { detail } else { format!("{what}: {detail}") })
}

fn win(e: windows::core::Error) -> Error {
    hresult_error(e.code(), "")
}

fn init_com() {
    thread_local! {
        static INIT: () = unsafe {
            // The host process may already have initialised COM in another mode: that is fine.
            let hr = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            debug_assert!(hr.is_ok() || hr.0 == RPC_E_CHANGED_MODE);
        };
    }
    INIT.with(|_| ());
}

// ---------------------------------------------------------------------------- variants

/// An owned `VARIANT`.
#[repr(transparent)]
struct Var(VARIANT);

impl Var {
    fn empty() -> Self {
        Var(VARIANT::default())
    }

    fn with(vt: VARENUM, set: impl FnOnce(&mut VARIANT)) -> Self {
        let mut v = VARIANT::default();
        unsafe { (*v.Anonymous.Anonymous).vt = vt };
        set(&mut v);
        Var(v)
    }

    fn vt(&self) -> u16 {
        unsafe { self.0.Anonymous.Anonymous.vt.0 }
    }
}

impl Drop for Var {
    fn drop(&mut self) {
        unsafe {
            let _ = VariantClear(&mut self.0);
        }
    }
}

fn wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}

fn disp_of(env: Env, obj: &JsObject) -> Result<Option<IDispatch>> {
    if !ComObject::instance_of(env, obj)? {
        return Ok(None);
    }
    let inst: ClassInstance<ComObject> = unsafe { FromNapiValue::from_napi_value(env.raw(), obj.raw())? };
    Ok(Some(inst.ptr.clone()))
}

fn to_variant(env: Env, value: JsUnknown) -> Result<Var> {
    Ok(match value.get_type()? {
        ValueType::Undefined => Var::with(VT_ERROR, |v| unsafe { (*v.Anonymous.Anonymous).Anonymous.scode = DISP_E_PARAMNOTFOUND }),
        ValueType::Null => Var::empty(),
        ValueType::Boolean => {
            let b = unsafe { value.cast::<JsBoolean>() }.get_value()?;
            Var::with(VT_BOOL, |v| unsafe { (*v.Anonymous.Anonymous).Anonymous.boolVal = VARIANT_BOOL(if b { -1 } else { 0 }) })
        }
        ValueType::Number => {
            let f = unsafe { value.cast::<JsNumber>() }.get_double()?;
            if f.fract() == 0.0 && f >= i32::MIN as f64 && f <= i32::MAX as f64 {
                Var::with(VT_I4, |v| unsafe { (*v.Anonymous.Anonymous).Anonymous.lVal = f as i32 })
            } else {
                Var::with(VT_R8, |v| unsafe { (*v.Anonymous.Anonymous).Anonymous.dblVal = f })
            }
        }
        ValueType::String => {
            let s = unsafe { value.cast::<JsString>() }.into_utf16()?;
            // The UTF-16 buffer napi hands out ends with a terminator that is not part of the string.
            let units = s.as_slice();
            let units = units.strip_suffix(&[0]).unwrap_or(units);
            let b = BSTR::from_wide(units);
            Var::with(VT_BSTR, |v| unsafe { (*v.Anonymous.Anonymous).Anonymous.bstrVal = ManuallyDrop::new(b) })
        }
        ValueType::Object => {
            let obj = unsafe { value.cast::<JsObject>() };
            if let Some(d) = disp_of(env, &obj)? {
                Var::with(VT_DISPATCH, |v| unsafe { (*v.Anonymous.Anonymous).Anonymous.pdispVal = ManuallyDrop::new(Some(d)) })
            } else if obj.is_array()? {
                array_to_variant(env, &obj)?
            } else if obj.has_named_property("$t")? {
                let tag = obj.get_named_property::<JsString>("$t")?.into_utf8()?.into_owned()?;
                let number = |name: &str| -> Result<f64> { obj.get_named_property::<JsNumber>(name)?.get_double() };
                match tag.as_str() {
                    "date" => {
                        let d = number("v")?;
                        Var::with(VT_DATE, |v| unsafe { (*v.Anonymous.Anonymous).Anonymous.date = d })
                    }
                    "error" => {
                        let code = number("v")? as i32;
                        Var::with(VT_ERROR, |v| unsafe { (*v.Anonymous.Anonymous).Anonymous.scode = (0x800A_0000_u32 | (code as u32 & 0xFFFF)) as i32 })
                    }
                    _ => Var::with(VT_DISPATCH, |v| unsafe { (*v.Anonymous.Anonymous).Anonymous.pdispVal = ManuallyDrop::new(None) }),
                }
            } else {
                return Err(com_error(13, "RAS", "この値は Excel に渡せません（RAS のクラスや Map は COM の値に変換できません）"));
            }
        }
        _ => return Err(com_error(13, "RAS", "この値は Excel に渡せません")),
    })
}

/// A JS array (of arrays) as a 0-based `SAFEARRAY` of `Variant`.
fn array_to_variant(env: Env, arr: &JsObject) -> Result<Var> {
    let rows = arr.get_array_length()?;
    let first_row = if rows > 0 {
        let first = arr.get_element::<JsUnknown>(0)?;
        if first.get_type()? == ValueType::Object {
            let o = unsafe { first.cast::<JsObject>() };
            if o.is_array()? { Some(o.get_array_length()?) } else { None }
        } else {
            None
        }
    } else {
        None
    };
    let mut bounds = vec![SAFEARRAYBOUND { cElements: rows, lLbound: 0 }];
    if let Some(cols) = first_row {
        bounds.push(SAFEARRAYBOUND { cElements: cols, lLbound: 0 });
    }
    let psa = unsafe { SafeArrayCreate(VT_VARIANT, bounds.len() as u32, bounds.as_ptr()) };
    if psa.is_null() {
        return Err(com_error(7, "RAS", "配列を確保できません"));
    }
    // Owned by the VARIANT from here on (VariantClear destroys it).
    let out = Var::with(VARENUM(VT_ARRAY.0 | VT_VARIANT.0), |v| unsafe { (*v.Anonymous.Anonymous).Anonymous.parray = psa });
    let put = |index: &[i32], item: JsUnknown| -> Result<()> {
        let v = to_variant(env, item)?;
        unsafe { SafeArrayPutElement(psa, index.as_ptr(), &v.0 as *const VARIANT as *const c_void) }.map_err(win)
    };
    for r in 0..rows {
        let item = arr.get_element::<JsUnknown>(r)?;
        match first_row {
            None => put(&[r as i32], item)?,
            Some(cols) => {
                if item.get_type()? != ValueType::Object {
                    return Err(com_error(13, "RAS", "2 次元配列の各行は配列でなければなりません"));
                }
                let row = unsafe { item.cast::<JsObject>() };
                let n = row.get_array_length()?.min(cols);
                for c in 0..n {
                    put(&[r as i32, c as i32], row.get_element::<JsUnknown>(c)?)?;
                }
            }
        }
    }
    Ok(out)
}

fn tagged(env: Env, tag: &str, value: f64) -> Result<JsUnknown> {
    let mut o = env.create_object()?;
    o.set_named_property("$t", env.create_string(tag)?)?;
    o.set_named_property("v", env.create_double(value)?)?;
    Ok(o.into_unknown())
}

fn object_to_js(env: Env, d: Option<IDispatch>) -> Result<JsUnknown> {
    match d {
        Some(ptr) => Ok(ComObject { ptr }.into_instance(env)?.as_object(env).into_unknown()),
        None => Ok(env.get_null()?.into_unknown()),
    }
}

fn from_variant(env: Env, v: &VARIANT) -> Result<JsUnknown> {
    unsafe {
        let inner = &v.Anonymous.Anonymous;
        let vt = inner.vt.0;
        let data = &inner.Anonymous;
        if vt & VT_BYREF.0 != 0 {
            return if vt == VT_BYREF.0 | VT_VARIANT.0 && !data.pvarVal.is_null() { from_variant(env, &*data.pvarVal) } else { Ok(env.get_null()?.into_unknown()) };
        }
        if vt & VT_ARRAY.0 != 0 {
            return array_to_js(env, data.parray);
        }
        let num = |f: f64| -> Result<JsUnknown> { Ok(env.create_double(f)?.into_unknown()) };
        match vt {
            2 => num(data.iVal as f64),
            3 | 22 => num(data.lVal as f64),
            4 => num(data.fltVal as f64),
            5 => num(data.dblVal),
            6 => num(data.cyVal.int64 as f64 / 10000.0),
            7 => tagged(env, "date", data.date),
            8 => Ok(env.create_string_utf16(&data.bstrVal)?.into_unknown()),
            9 => object_to_js(env, (*data.pdispVal).clone()),
            10 => tagged(env, "error", (data.scode & 0xFFFF) as f64),
            11 => Ok(env.get_boolean(data.boolVal.0 != 0)?.into_unknown()),
            13 => object_to_js(env, (*data.punkVal).as_ref().and_then(|u| u.cast::<IDispatch>().ok())),
            14 => num(VarR8FromDec(&v.Anonymous.decVal).map_err(win)?),
            16 => num(data.cVal as f64),
            17 => num(data.bVal as f64),
            18 => num(data.uiVal as f64),
            19 | 23 => num(data.ulVal as f64),
            20 => num(data.llVal as f64),
            21 => num(data.ullVal as f64),
            // Empty, Null and anything that has no JavaScript counterpart.
            _ => Ok(env.get_null()?.into_unknown()),
        }
    }
}

/// A 1-D or 2-D `SAFEARRAY` as a 0-based JS array (of row arrays), whatever its lower bounds.
fn array_to_js(env: Env, psa: *mut SAFEARRAY) -> Result<JsUnknown> {
    if psa.is_null() {
        return Ok(env.create_array_with_length(0)?.into_unknown());
    }
    unsafe {
        let dims = SafeArrayGetDim(psa);
        if dims == 0 || dims > 2 {
            return Err(com_error(13, "RAS", "3 次元以上の配列は扱えません"));
        }
        let elem_vt = SafeArrayGetVartype(psa).map_err(win)?;
        let element = |index: &[i32]| -> Result<JsUnknown> {
            let mut v = Var::empty();
            if elem_vt == VT_VARIANT {
                SafeArrayGetElement(psa, index.as_ptr(), &mut v.0 as *mut VARIANT as *mut c_void).map_err(win)?;
            } else {
                // Typed array: the element is copied into the data part of a VARIANT of that type.
                let slot = &mut (*v.0.Anonymous.Anonymous).Anonymous as *mut _ as *mut c_void;
                SafeArrayGetElement(psa, index.as_ptr(), slot).map_err(win)?;
                (*v.0.Anonymous.Anonymous).vt = elem_vt;
            }
            from_variant(env, &v.0)
        };
        let (lb1, ub1) = (SafeArrayGetLBound(psa, 1).map_err(win)?, SafeArrayGetUBound(psa, 1).map_err(win)?);
        let rows = (ub1 - lb1 + 1).max(0);
        let mut out = env.create_array_with_length(rows as usize)?;
        if dims == 1 {
            for i in 0..rows {
                out.set_element(i as u32, element(&[lb1 + i])?)?;
            }
        } else {
            let (lb2, ub2) = (SafeArrayGetLBound(psa, 2).map_err(win)?, SafeArrayGetUBound(psa, 2).map_err(win)?);
            let cols = (ub2 - lb2 + 1).max(0);
            for r in 0..rows {
                let mut row = env.create_array_with_length(cols as usize)?;
                for c in 0..cols {
                    row.set_element(c as u32, element(&[lb1 + r, lb2 + c])?)?;
                }
                out.set_element(r as u32, row)?;
            }
        }
        Ok(out.into_unknown())
    }
}

// ---------------------------------------------------------------------------- IDispatch

fn dispid(disp: &IDispatch, name: &str) -> Result<i32> {
    if name.is_empty() {
        return Ok(DISPID_VALUE);
    }
    let w = wide(name);
    let names = [PCWSTR(w.as_ptr())];
    let mut id = 0i32;
    unsafe { disp.GetIDsOfNames(&GUID::zeroed(), names.as_ptr(), 1, LCID, &mut id) }.map_err(|e| hresult_error(e.code(), name))?;
    Ok(id)
}

/// `Invoke` with the arguments in call order. Retries while Excel rejects the call because it is busy.
fn invoke(disp: &IDispatch, id: i32, name: &str, flags: DISPATCH_FLAGS, mut args: Vec<Var>, put: bool) -> Result<Var> {
    // Arguments are passed last to first. The callee may convert them in place, so the vector
    // handed to it is the one and only owner of the values (`Var` is a transparent VARIANT).
    args.reverse();
    let mut rg = args;
    let mut named = DISPID_PROPERTYPUT;
    let params = DISPPARAMS {
        rgvarg: if rg.is_empty() { std::ptr::null_mut() } else { rg.as_mut_ptr() as *mut VARIANT },
        rgdispidNamedArgs: if put { &mut named } else { std::ptr::null_mut() },
        cArgs: rg.len() as u32,
        cNamedArgs: if put { 1 } else { 0 },
    };
    let started = std::time::Instant::now();
    let mut delay = 20u64;
    // From here until the call returns, the dialog watcher looks for what Excel is waiting for.
    let outstanding = crate::watch::CallGuard::new();
    let outcome = loop {
        let mut result = Var::empty();
        let mut ex = EXCEPINFO::default();
        let mut arg_err = 0u32;
        let r = unsafe { disp.Invoke(id, &GUID::zeroed(), LCID, flags, &params, Some(&mut result.0), Some(&mut ex), Some(&mut arg_err)) };
        match r {
            Ok(()) => break Ok(result),
            Err(e) => {
                let hr = e.code();
                let busy = hr.0 == RPC_E_CALL_REJECTED || hr.0 == RPC_E_SERVERCALL_RETRYLATER;
                if busy && (started.elapsed().as_millis() as u64) < BUSY_TIMEOUT_MS {
                    outstanding.rejected();
                    std::thread::sleep(std::time::Duration::from_millis(delay));
                    delay = (delay * 2).min(500);
                    continue;
                }
                if hr.0 == DISP_E_EXCEPTION {
                    unsafe {
                        if let Some(fill) = ex.pfnDeferredFillIn {
                            let _ = fill(&mut ex);
                        }
                        let source = ManuallyDrop::take(&mut ex.bstrSource).to_string();
                        let description = ManuallyDrop::take(&mut ex.bstrDescription).to_string();
                        drop(ManuallyDrop::take(&mut ex.bstrHelpFile));
                        // Automation errors of FACILITY_CONTROL carry the VBA error number in the low word;
                        // a dispatch HRESULT (e.g. `Worksheets("missing")` -> DISP_E_BADINDEX) maps as in VBA.
                        let number = if ex.wCode != 0 {
                            ex.wCode as i32
                        } else if (ex.scode as u32) & 0xFFFF_0000 == 0x800A_0000 {
                            ex.scode & 0xFFFF
                        } else {
                            vba_number(ex.scode)
                        };
                        break Err(com_error(number, &source, &description));
                    }
                }
                break Err(hresult_error(hr, name));
            }
        }
    };
    drop(outstanding);
    drop(rg);
    outcome
}

fn convert_args(env: Env, args: Vec<JsUnknown>) -> Result<Vec<Var>> {
    args.into_iter().map(|a| to_variant(env, a)).collect()
}

/// A COM object (anything that implements `IDispatch`).
#[napi]
pub struct ComObject {
    ptr: IDispatch,
}

#[napi]
impl ComObject {
    /// Reads a property (with index arguments, if any). An empty name is the default member.
    #[napi]
    pub fn get(&self, env: Env, name: String, args: Vec<JsUnknown>) -> Result<JsUnknown> {
        let id = dispid(&self.ptr, &name)?;
        let flags = if args.is_empty() { DISPATCH_PROPERTYGET } else { DISPATCH_PROPERTYGET | DISPATCH_METHOD };
        let r = invoke(&self.ptr, id, &name, flags, convert_args(env, args)?, false)?;
        from_variant(env, &r.0)
    }

    /// Calls a method. VBA does not tell methods from properties with arguments, so both are tried at once.
    #[napi]
    pub fn call(&self, env: Env, name: String, args: Vec<JsUnknown>) -> Result<JsUnknown> {
        let id = dispid(&self.ptr, &name)?;
        let r = invoke(&self.ptr, id, &name, DISPATCH_METHOD | DISPATCH_PROPERTYGET, convert_args(env, args)?, false)?;
        from_variant(env, &r.0)
    }

    /// Writes a property (`obj.name(args...) = value`). An empty name is the default member.
    #[napi]
    pub fn put(&self, env: Env, name: String, args: Vec<JsUnknown>, value: JsUnknown) -> Result<()> {
        let id = dispid(&self.ptr, &name)?;
        let build = |env: Env| -> Result<Vec<Var>> {
            let mut all = Vec::with_capacity(args.len() + 1);
            for a in &args {
                all.push(to_variant(env, unsafe { JsUnknown::from_raw_unchecked(env.raw(), a.raw()) })?);
            }
            all.push(to_variant(env, unsafe { JsUnknown::from_raw_unchecked(env.raw(), value.raw()) })?);
            Ok(all)
        };
        let all = build(env)?;
        let is_object = all.last().is_some_and(|v| v.vt() == VT_DISPATCH.0);
        match invoke(&self.ptr, id, &name, DISPATCH_PROPERTYPUT, all, true) {
            Ok(_) => Ok(()),
            // Properties that hold an object reference are set with `Set` (PROPERTYPUTREF).
            Err(_) if is_object => invoke(&self.ptr, id, &name, DISPATCH_PROPERTYPUTREF, build(env)?, true).map(|_| ()),
            Err(e) => Err(e),
        }
    }

    /// The members of a collection (`For Each`).
    #[napi]
    pub fn enumerate(&self, env: Env) -> Result<Vec<JsUnknown>> {
        let r = invoke(&self.ptr, DISPID_NEWENUM, "_NewEnum", DISPATCH_METHOD | DISPATCH_PROPERTYGET, Vec::new(), false)?;
        let en: IEnumVARIANT = unsafe {
            let data = &r.0.Anonymous.Anonymous.Anonymous;
            match r.vt() {
                9 => (*data.pdispVal).as_ref().and_then(|d| d.cast::<IEnumVARIANT>().ok()),
                13 => (*data.punkVal).as_ref().and_then(|u| u.cast::<IEnumVARIANT>().ok()),
                _ => None,
            }
        }
        .ok_or_else(|| com_error(438, "RAS", "このオブジェクトは列挙できません"))?;
        let mut out = Vec::new();
        loop {
            let mut item = [VARIANT::default()];
            let mut fetched = 0u32;
            let hr = unsafe { en.Next(&mut item, &mut fetched) };
            if hr.is_err() || fetched == 0 {
                break;
            }
            let [raw] = item;
            let v = Var(raw);
            out.push(from_variant(env, &v.0)?);
        }
        Ok(out)
    }

    /// The class name, as VBA's `TypeName` reports it (`Worksheet`, `Range`).
    #[napi]
    pub fn type_name(&self) -> Result<String> {
        unsafe {
            // The class the object says it is, as VBA asks first.
            if let Some(name) = self.ptr.cast::<IProvideClassInfo>().ok().and_then(|p| p.GetClassInfo().ok()).and_then(|class| type_info_name(&class)) {
                return Ok(name);
            }
            let info = self.ptr.GetTypeInfo(0, LCID).map_err(win)?;
            // Else the class of its library whose default interface this is: `File` for
            // `IFile`, `Dictionary` for `IDictionary`, `FileSystemObject` for `IFileSystem3`.
            if let Some(name) = class_of_interface(&info) {
                return Ok(name);
            }
            let name = type_info_name(&info).ok_or_else(|| com_error(438, "RAS", "型の名前を取得できません"))?;
            Ok(name.strip_prefix('_').unwrap_or(&name).to_string())
        }
    }

    /// The properties that can be read without arguments, from the type information, for the
    /// variables view of the debugger. Hidden and restricted members are left out.
    #[napi]
    pub fn members(&self) -> Result<Vec<ComMember>> {
        unsafe {
            let info = dispatch_info(self.ptr.GetTypeInfo(0, LCID).map_err(win)?);
            let attr = TypeAttr::of(&info)?;
            let (funcs, vars) = ((*attr.0).cFuncs, (*attr.0).cVars);
            let mut out: Vec<ComMember> = Vec::new();
            let mut push = |memid: i32, tdesc: &TYPEDESC| {
                let mut name = BSTR::default();
                let mut context = 0u32;
                if info.GetDocumentation(memid, Some(&mut name), None, &mut context, None).is_err() {
                    return;
                }
                let name = name.to_string();
                if name.is_empty() || name.starts_with('_') || out.iter().any(|m| m.name.eq_ignore_ascii_case(&name)) {
                    return;
                }
                out.push(ComMember { name, scalar: is_scalar(&info, tdesc) });
            };
            for i in 0..u32::from(funcs) {
                let Ok(desc) = info.GetFuncDesc(i) else { continue };
                let f = &*desc;
                let hidden = f.wFuncFlags.0 & (FUNCFLAG_FHIDDEN.0 | FUNCFLAG_FRESTRICTED.0) != 0;
                if f.invkind == INVOKE_PROPERTYGET && !hidden && required_params(f) == 0 {
                    push(f.memid, &f.elemdescFunc.tdesc);
                }
                info.ReleaseFuncDesc(desc);
            }
            for i in 0..u32::from(vars) {
                let Ok(desc) = info.GetVarDesc(i) else { continue };
                let v = &*desc;
                let hidden = v.wVarFlags.0 & (VARFLAG_FHIDDEN.0 | VARFLAG_FRESTRICTED.0) != 0;
                if v.varkind == VAR_DISPATCH && !hidden {
                    push(v.memid, &v.elemdescVar.tdesc);
                }
                info.ReleaseVarDesc(desc);
            }
            Ok(out)
        }
    }

    /// True when both wrap the same COM object (`Is` in VBA).
    #[napi]
    pub fn is_same(&self, other: &ComObject) -> bool {
        match (self.ptr.cast::<windows::core::IUnknown>(), other.ptr.cast::<windows::core::IUnknown>()) {
            (Ok(a), Ok(b)) => a.as_raw() == b.as_raw(),
            _ => false,
        }
    }
}

// ---------------------------------------------------------------------------- type information

/// A readable property of a COM object (see [`ComObject::members`]).
#[napi(object)]
pub struct ComMember {
    /// Name as the type library spells it (`UsedRange`).
    pub name: String,
    /// The property is a number, string, date, boolean or enumeration: reading it is cheap and
    /// shows its value. Objects and `Variant`s (`Range.Value` of a whole sheet) are read on demand.
    pub scalar: bool,
}

unsafe fn type_info_name(info: &ITypeInfo) -> Option<String> {
    let mut name = BSTR::default();
    let mut context = 0u32;
    unsafe { info.GetDocumentation(-1, Some(&mut name), None, &mut context, None) }.ok()?;
    Some(name.to_string()).filter(|n| !n.is_empty())
}

/// The class (coclass) of the type library of `info` that has it as its default interface,
/// remembered by the interface's GUID (a library has hundreds of types).
unsafe fn class_of_interface(info: &ITypeInfo) -> Option<String> {
    thread_local! {
        static KNOWN: std::cell::RefCell<std::collections::HashMap<u128, Option<String>>> = std::cell::RefCell::new(std::collections::HashMap::new());
    }
    unsafe {
        let guid = (*TypeAttr::of(info).ok()?.0).guid;
        if let Some(known) = KNOWN.with(|k| k.borrow().get(&guid.to_u128()).cloned()) {
            return known;
        }
        let find = || -> Option<String> {
            let mut lib = None;
            let mut index = 0u32;
            info.GetContainingTypeLib(&mut lib, &mut index).ok()?;
            let lib = lib?;
            let mut other = None;
            for i in 0..lib.GetTypeInfoCount() {
                if lib.GetTypeInfoType(i).ok()? != TKIND_COCLASS {
                    continue;
                }
                let Ok(class) = lib.GetTypeInfo(i) else { continue };
                let Ok(attr) = TypeAttr::of(&class) else { continue };
                for j in 0..u32::from((*attr.0).cImplTypes) {
                    let Ok(flags) = class.GetImplTypeFlags(j) else { continue };
                    if flags.0 & IMPLTYPEFLAG_FSOURCE.0 != 0 {
                        continue;
                    }
                    let Ok(implemented) = class.GetRefTypeOfImplType(j).and_then(|r| class.GetRefTypeInfo(r)) else { continue };
                    let Ok(its) = TypeAttr::of(&implemented) else { continue };
                    if (*its.0).guid != guid {
                        continue;
                    }
                    if flags.0 & IMPLTYPEFLAG_FDEFAULT.0 != 0 {
                        return type_info_name(&class);
                    }
                    other = other.or_else(|| type_info_name(&class));
                }
            }
            other
        };
        let found = find();
        KNOWN.with(|k| k.borrow_mut().insert(guid.to_u128(), found.clone()));
        found
    }
}

struct TypeAttr<'a>(*mut TYPEATTR, &'a ITypeInfo);

impl<'a> TypeAttr<'a> {
    unsafe fn of(info: &'a ITypeInfo) -> Result<Self> {
        Ok(Self(unsafe { info.GetTypeAttr() }.map_err(win)?, info))
    }
}

impl Drop for TypeAttr<'_> {
    fn drop(&mut self) {
        unsafe { self.1.ReleaseTypeAttr(self.0) };
    }
}

/// The dispatch view of a dual interface, whose descriptions have the real return types
/// (the vtable view returns `HRESULT` with the value as a `retval` parameter).
unsafe fn dispatch_info(info: ITypeInfo) -> ITypeInfo {
    unsafe {
        let Ok(attr) = TypeAttr::of(&info) else { return info };
        let dual = (*attr.0).typekind == TKIND_INTERFACE && i32::from((*attr.0).wTypeFlags) & TYPEFLAG_FDUAL.0 != 0;
        drop(attr);
        if dual {
            if let Ok(dispatch) = info.GetRefTypeOfImplType(u32::MAX).and_then(|r| info.GetRefTypeInfo(r)) {
                return dispatch;
            }
        }
        info
    }
}

/// Parameters a caller must pass (the locale and return-value parameters are filled in by COM).
unsafe fn required_params(f: &FUNCDESC) -> usize {
    let n = usize::try_from(f.cParams).unwrap_or(0);
    let params = if n == 0 || f.lprgelemdescParam.is_null() { &[][..] } else { unsafe { std::slice::from_raw_parts(f.lprgelemdescParam, n) } };
    params.iter().filter(|p| unsafe { p.Anonymous.paramdesc.wParamFlags.0 } & (PARAMFLAG_FLCID.0 | PARAMFLAG_FRETVAL.0 | PARAMFLAG_FOPT.0) == 0).count()
}

unsafe fn is_scalar(info: &ITypeInfo, tdesc: &TYPEDESC) -> bool {
    match tdesc.vt {
        VT_I1 | VT_I2 | VT_I4 | VT_I8 | VT_UI1 | VT_UI2 | VT_UI4 | VT_UI8 | VT_INT | VT_UINT | VT_R4 | VT_R8 | VT_CY | VT_DATE | VT_BSTR | VT_BOOL | VT_DECIMAL => true,
        // An enumeration (`XlSheetVisibility`); other user-defined types are interfaces.
        VT_USERDEFINED => unsafe {
            info.GetRefTypeInfo(tdesc.Anonymous.hreftype).ok().and_then(|t| TypeAttr::of(&t).ok().map(|a| (*a.0).typekind == TKIND_ENUM)).unwrap_or(false)
        },
        _ => false,
    }
}

// ---------------------------------------------------------------------------- lookups

fn clsid(prog_id: &str) -> Result<GUID> {
    let w = wide(prog_id);
    unsafe { CLSIDFromProgID(PCWSTR(w.as_ptr())) }.map_err(|_| com_error(429, "RAS", &format!("`{prog_id}` というクラスは登録されていません")))
}

/// The running instance of an application (`GetObject(, "Excel.Application")`).
#[napi]
pub fn get_active_object(prog_id: String) -> Result<ComObject> {
    init_com();
    let id = clsid(&prog_id)?;
    let mut unknown = None;
    unsafe { GetActiveObject(&id, None, &mut unknown) }.map_err(|_| com_error(429, "RAS", &format!("{prog_id} は起動していません")))?;
    let ptr = unknown.and_then(|u| u.cast::<IDispatch>().ok()).ok_or_else(|| com_error(429, "RAS", &format!("{prog_id} に接続できません")))?;
    Ok(ComObject { ptr })
}

/// `CreateObject(progId)`.
#[napi]
pub fn create_object(prog_id: String) -> Result<ComObject> {
    init_com();
    let id = clsid(&prog_id)?;
    let ptr: IDispatch = unsafe { CoCreateInstance(&id, None, CLSCTX_LOCAL_SERVER | CLSCTX_INPROC_SERVER) }.map_err(|e| hresult_error(e.code(), &prog_id))?;
    Ok(ComObject { ptr })
}

/// `GetObject(path)`: binds to a file, opening it in its application when it is not open yet.
#[napi]
pub fn get_object(path: String) -> Result<ComObject> {
    init_com();
    let w = wide(&path);
    // Opening the file can make its application ask questions (links, read-only, macros).
    let _outstanding = crate::watch::CallGuard::new();
    let ptr: IDispatch = unsafe { CoGetObject(PCWSTR(w.as_ptr()), None) }.map_err(|e| hresult_error(e.code(), &path))?;
    Ok(ComObject { ptr })
}

/// Opens a file the way a double click does (`ShellExecute` with the verb `open`). Excel then
/// starts as the user's own instance, not as an automation server: an Excel that
/// `CreateObject` started differs from one the user started (add-ins, macro security, access
/// to the VBA project), and the instance stays after the debug run for sync and "Run in Excel".
#[napi]
pub fn shell_open(path: String) -> Result<()> {
    use windows::Win32::UI::Shell::ShellExecuteW;
    use windows::Win32::UI::WindowsAndMessaging::SW_SHOWNORMAL;
    init_com();
    let file = wide(&path);
    let verb = wide("open");
    let instance = unsafe { ShellExecuteW(None, PCWSTR(verb.as_ptr()), PCWSTR(file.as_ptr()), PCWSTR::null(), PCWSTR::null(), SW_SHOWNORMAL) };
    if instance.0 as usize <= 32 {
        return Err(com_error(429, "RAS", &format!("{path} を開けませんでした（Excel がインストールされているか確認してください）")));
    }
    Ok(())
}

fn same_path(a: &str, b: &str) -> bool {
    let norm = |s: &str| s.replace('/', "\\").to_lowercase();
    norm(a) == norm(b)
}

/// Finds an open workbook by its full path in the Running Object Table, across every
/// running Excel instance (spec §14.3). Returns null when no instance has it open.
#[napi]
pub fn find_workbook(path: String) -> Result<Option<ComObject>> {
    init_com();
    unsafe {
        let rot = GetRunningObjectTable(0).map_err(win)?;
        let monikers = rot.EnumRunning().map_err(win)?;
        let ctx = CreateBindCtx(0).map_err(win)?;
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
            if !same_path(&name, &path) {
                continue;
            }
            if let Ok(unknown) = rot.GetObject(&moniker) {
                if let Ok(ptr) = unknown.cast::<IDispatch>() {
                    return Ok(Some(ComObject { ptr }));
                }
            }
        }
    }
}
