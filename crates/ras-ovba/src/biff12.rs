//! Just enough of the BIFF12 record stream of `.xlsb` workbooks
//! (\[MS-XLSB\]) to find sheet names, relationship ids and VBA code names.
//! Binary parts are only ever read, never rewritten.

/// `BrtWsProp`: worksheet properties; ends with the code name.
const BRT_WS_PROP: u32 = 147;
/// `BrtWbProp`: workbook properties; ends with the code name.
const BRT_WB_PROP: u32 = 153;
/// `BrtBundleSh`: one sheet of the workbook.
const BRT_BUNDLE_SH: u32 = 156;
/// `BrtBeginSheetData`: start of the cell table (nothing of interest
/// follows for our purposes).
const BRT_BEGIN_SHEET_DATA: u32 = 145;

/// Iterates over `(record id, record data)`; stops silently at the first
/// malformed record.
fn records(data: &[u8]) -> impl Iterator<Item = (u32, &[u8])> {
    let mut pos = 0usize;
    std::iter::from_fn(move || {
        let first = *data.get(pos)?;
        pos += 1;
        let mut id = first as u32;
        if first & 0x80 != 0 {
            let second = *data.get(pos)?;
            pos += 1;
            id = (first as u32 & 0x7F) | ((second as u32 & 0x7F) << 7);
        }
        let mut len = 0usize;
        for shift in 0..4 {
            let byte = *data.get(pos)?;
            pos += 1;
            len |= ((byte & 0x7F) as usize) << (7 * shift);
            if byte & 0x80 == 0 {
                break;
            }
        }
        let body = data.get(pos..pos.checked_add(len)?)?;
        pos += len;
        Some((id, body))
    })
}

/// Reads an `XLWideString` (32-bit character count + UTF-16LE) at `at`.
/// Returns the string and the offset after it. A count of `0xFFFFFFFF`
/// (null `XLNullableWideString`) yields an empty string.
fn wide_string(data: &[u8], at: usize) -> Option<(String, usize)> {
    let raw = data.get(at..at.checked_add(4)?)?;
    let count = u32::from_le_bytes([raw[0], raw[1], raw[2], raw[3]]);
    if count == u32::MAX {
        return Some((String::new(), at + 4));
    }
    let byte_len = (count as usize).checked_mul(2)?;
    let start = at + 4;
    let bytes = data.get(start..start.checked_add(byte_len)?)?;
    let units: Vec<u16> = bytes.chunks_exact(2).map(|c| u16::from_le_bytes([c[0], c[1]])).collect();
    Some((String::from_utf16_lossy(&units), start + byte_len))
}

/// A sheet listed in `workbook.bin`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct BinarySheet {
    pub name: String,
    pub rel_id: String,
}

/// Workbook code name (if recorded) and the sheets of a `workbook.bin`.
pub(crate) fn workbook_info(data: &[u8]) -> (Option<String>, Vec<BinarySheet>) {
    let mut code_name = None;
    let mut sheets = Vec::new();
    for (id, body) in records(data) {
        match id {
            BRT_WB_PROP => {
                // flags (4 bytes), dwThemeVersion (4 bytes), strName
                if let Some((name, _)) = wide_string(body, 8) {
                    if !name.is_empty() {
                        code_name = Some(name);
                    }
                }
            }
            BRT_BUNDLE_SH => {
                // hsState (4 bytes), iTabID (4 bytes), strRelID, strName
                if let Some((rel_id, next)) = wide_string(body, 8) {
                    if let Some((name, _)) = wide_string(body, next) {
                        sheets.push(BinarySheet { name, rel_id });
                    }
                }
            }
            _ => {}
        }
    }
    (code_name, sheets)
}

/// Code name of a worksheet part (`BrtWsProp.strName`), if recorded.
pub(crate) fn worksheet_code_name(data: &[u8]) -> Option<String> {
    for (id, body) in records(data) {
        match id {
            BRT_WS_PROP => {
                // flags (3 bytes), brtcolorTab (8 bytes), rwSync (4 bytes),
                // colSync (4 bytes), strName
                let (name, _) = wide_string(body, 19)?;
                return if name.is_empty() { None } else { Some(name) };
            }
            BRT_BEGIN_SHEET_DATA => return None,
            _ => {}
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(id: u32, body: &[u8]) -> Vec<u8> {
        let mut out = Vec::new();
        if id < 0x80 {
            out.push(id as u8);
        } else {
            out.push((id & 0x7F) as u8 | 0x80);
            out.push((id >> 7) as u8);
        }
        let mut len = body.len();
        loop {
            let byte = (len & 0x7F) as u8;
            len >>= 7;
            if len == 0 {
                out.push(byte);
                break;
            }
            out.push(byte | 0x80);
        }
        out.extend_from_slice(body);
        out
    }

    fn wide(s: &str) -> Vec<u8> {
        let units: Vec<u16> = s.encode_utf16().collect();
        let mut out = (units.len() as u32).to_le_bytes().to_vec();
        for u in units {
            out.extend_from_slice(&u.to_le_bytes());
        }
        out
    }

    #[test]
    fn reads_workbook_records() {
        let mut data = record(131, &[]); // BrtBeginBook
        let mut wb_prop = vec![0x20, 0, 1, 0, 0x42, 0xE5, 1, 0];
        wb_prop.extend(wide("ThisWorkbook"));
        data.extend(record(BRT_WB_PROP, &wb_prop));
        let mut sheet = vec![0, 0, 0, 0, 1, 0, 0, 0];
        sheet.extend(wide("rId1"));
        sheet.extend(wide("データ"));
        data.extend(record(BRT_BUNDLE_SH, &sheet));
        let mut long = vec![0, 0, 0, 0, 2, 0, 0, 0];
        long.extend(wide("rId2"));
        long.extend(wide(&"x".repeat(100)));
        data.extend(record(BRT_BUNDLE_SH, &long)); // > 127 bytes: two length bytes
        let (code_name, sheets) = workbook_info(&data);
        assert_eq!(code_name.as_deref(), Some("ThisWorkbook"));
        assert_eq!(sheets.len(), 2);
        assert_eq!(sheets[0], BinarySheet { name: "データ".into(), rel_id: "rId1".into() });
        assert_eq!(sheets[1].name.len(), 100);
    }

    #[test]
    fn reads_worksheet_code_name() {
        // Record bytes as written by Excel (code name "Sheet1").
        let body = [
            0xC9, 0x04, 0x02, 0x00, 0x40, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF,
            0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0x06, 0x00, 0x00, 0x00, 0x53, 0x00, 0x68, 0x00, 0x65,
            0x00, 0x65, 0x00, 0x74, 0x00, 0x31, 0x00,
        ];
        let mut data = record(129, &[]);
        data.extend(record(BRT_WS_PROP, &body));
        assert_eq!(worksheet_code_name(&data).as_deref(), Some("Sheet1"));
        // Empty code name.
        let mut empty = body[..19].to_vec();
        empty.extend([0, 0, 0, 0]);
        assert_eq!(worksheet_code_name(&record(BRT_WS_PROP, &empty)), None);
        // Garbage does not panic.
        assert_eq!(worksheet_code_name(&[0x93, 0x01, 0xFF, 0xFF, 0xFF, 0xFF, 0x00]), None);
        assert_eq!(workbook_info(&[0x9C, 0x01, 0x04, 1, 2]).1.len(), 0);
    }
}
