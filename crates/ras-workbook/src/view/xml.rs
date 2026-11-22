//! A small XML tree (local names only) for the parts of a workbook the grid reads whole:
//! drawings, VML, the controls of a sheet. The tree is umya-spreadsheet's (`xml_tree`, in the
//! fork; xljs reads workbooks with it too): elements nested deeper than 256 are left out.
//!
//! `parse` reads as much as can be read, with end tags not matched against their start (VML
//! from old writers is not always well-formed: `<br>` and the like without an end).

pub(crate) use umya_spreadsheet::xml_tree::{Node, parse_lenient as parse};
