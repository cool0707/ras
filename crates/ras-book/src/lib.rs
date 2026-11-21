//! `ras_book`: the Node addon through which the RAS debug run time stands in for Excel where
//! there is none (spec §14.10.8): on macOS and Linux, and on Windows without Excel.
//!
//! The workbook is `ras-workbook`'s; this crate gives it to Node as the class `Book`. The Excel
//! object model itself (`Application`, `Worksheet`, `Range`, ...) is JavaScript, in the run time
//! (`ras-excel.js`).

mod node;
