//! The workbook of the RAS debug run time's stand-in for Excel (spec §14.10.8), which the
//! `ras_book` addon (`crates/ras-book`) gives to Node.
//!
//! [`book::Book`] holds a workbook read with umya-spreadsheet; its formulas are calculated by
//! GridCraft's engine, as xljs uses it (the engine is taken from xljs's `vendor/gridcraft`).

pub mod book;
pub mod diff;
pub mod dump;

pub use gridcraft_core::{CellRef, RangeRef, Value};
