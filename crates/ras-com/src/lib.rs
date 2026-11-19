//! `ras_com`: the Node addon through which the RAS debug run time (spec §14.10) drives Excel.
//!
//! * [`com`] — the late-binding `IDispatch` bridge (Windows only).
//! * [`watch`] — finds the dialogs Excel shows while a call is waiting for it (Windows only).
//! * [`dialog`] — what a dialog is and what to do about it; plain data, so it is tested everywhere.
//!
//! The addon only exists on Windows; on other platforms the crate holds the dialog model alone.

pub mod dialog;

#[cfg(windows)]
mod com;
#[cfg(windows)]
mod watch;
