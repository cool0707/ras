//! Formats added to the file's styles: a cell that takes a number format when something is
//! typed into it (a date, a time, a percentage) gets the format it had with that number format.

use super::Result;
use super::scan::{self, Edits};

/// The number formats Excel has without their being written in the file.
const BUILT_IN: &[(u32, &str)] = &[
    (0, "General"),
    (1, "0"),
    (2, "0.00"),
    (3, "#,##0"),
    (4, "#,##0.00"),
    (9, "0%"),
    (10, "0.00%"),
    (11, "0.00E+00"),
    (12, "# ?/?"),
    (13, "# ??/??"),
    (14, "m/d/yyyy"),
    (15, "d-mmm-yy"),
    (16, "d-mmm"),
    (17, "mmm-yy"),
    (18, "h:mm AM/PM"),
    (19, "h:mm:ss AM/PM"),
    (20, "h:mm"),
    (21, "h:mm:ss"),
    (22, "m/d/yyyy h:mm"),
    (37, "#,##0_);(#,##0)"),
    (38, "#,##0_);[Red](#,##0)"),
    (39, "#,##0.00_);(#,##0.00)"),
    (40, "#,##0.00_);[Red](#,##0.00)"),
    (45, "mm:ss"),
    (46, "[h]:mm:ss"),
    (47, "mm:ss.0"),
    (48, "##0.0E+0"),
    (49, "@"),
];

/// How many cell formats (`cellXfs`) the styles have.
pub(super) fn xf_count(doc: &str) -> u32 {
    scan::tree(doc, &[])
        .and_then(|root| {
            root.child("cellXfs")
                .map(|x| x.children.iter().filter(|c| c.local() == "xf").count() as u32)
        })
        .unwrap_or(0)
}

/// The styles with the formats of `new` added after the ones there: each is the format with
/// the index `.0`, with the number format `.1`.
pub(super) fn with_formats(doc: &str, new: &[(u32, String)]) -> Result<String> {
    let root = scan::tree(doc, &[]).ok_or("xl/styles.xml を読み取れません")?;
    let p = scan::prefix(root.name);
    let mut edits = Edits::default();
    // The number formats written in the file, and the ids to give new ones.
    let formats = root.child("numFmts");
    let mut written: Vec<(u32, String)> = formats
        .iter()
        .flat_map(|f| &f.children)
        .filter_map(|f| {
            Some((
                f.attr(doc, "numFmtId")?.parse().ok()?,
                f.attr(doc, "formatCode")?,
            ))
        })
        .collect();
    let had = written.len();
    let mut id_of = |code: &str| -> u32 {
        if let Some((id, _)) = written.iter().find(|(_, c)| c == code) {
            return *id;
        }
        if let Some((id, _)) = BUILT_IN.iter().find(|(_, c)| *c == code) {
            return *id;
        }
        let id = written
            .iter()
            .map(|(id, _)| *id)
            .max()
            .unwrap_or(0)
            .max(163)
            + 1;
        written.push((id, code.to_string()));
        id
    };
    let cell_xfs = root
        .child("cellXfs")
        .ok_or("xl/styles.xml にセルの書式（cellXfs）がありません")?;
    let xfs: Vec<&scan::El> = cell_xfs
        .children
        .iter()
        .filter(|c| c.local() == "xf")
        .collect();
    let mut added = String::new();
    for (base, code) in new {
        let id = id_of(code).to_string();
        match xfs.get(*base as usize) {
            Some(xf) => {
                // The format it is made from, as it is written, with the number format.
                let text = &doc[xf.span()];
                let open = 0..xf.open.len();
                let mut local = Edits::default();
                scan::set_attr(text, open.clone(), "numFmtId", &id, &mut local);
                scan::set_attr(text, open, "applyNumberFormat", "1", &mut local);
                added.push_str(&local.apply(text));
            }
            None => added.push_str(&format!(
                "<{p}xf numFmtId=\"{id}\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>"
            )),
        }
    }
    if cell_xfs.end == cell_xfs.open.end {
        return Err("xl/styles.xml のセルの書式（cellXfs）が空です".into());
    }
    edits.insert(cell_xfs.inner.end, added);
    if cell_xfs.attr(doc, "count").is_some() {
        scan::set_attr(
            doc,
            cell_xfs.open.clone(),
            "count",
            &(xfs.len() + new.len()).to_string(),
            &mut edits,
        );
    }
    let fresh: String = written[had..]
        .iter()
        .map(|(id, code)| {
            format!(
                "<{p}numFmt numFmtId=\"{id}\" formatCode=\"{}\"/>",
                scan::escape(code)
            )
        })
        .collect();
    if !fresh.is_empty() {
        match formats {
            Some(f) if f.end != f.open.end => {
                edits.insert(f.inner.end, fresh);
                if f.attr(doc, "count").is_some() {
                    scan::set_attr(
                        doc,
                        f.open.clone(),
                        "count",
                        &written.len().to_string(),
                        &mut edits,
                    );
                }
            }
            Some(f) => edits.replace(
                f.span(),
                format!(
                    "<{p}numFmts count=\"{}\">{fresh}</{p}numFmts>",
                    written.len()
                ),
            ),
            // The first element of the styles.
            None => edits.insert(
                root.open.end,
                format!(
                    "<{p}numFmts count=\"{}\">{fresh}</{p}numFmts>",
                    written.len()
                ),
            ),
        }
    }
    let out = edits.apply(doc);
    scan::tree(&out, &[]).ok_or("xl/styles.xml を書き換えられませんでした")?;
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_format_is_added_with_its_number_format() {
        let doc = "<styleSheet><fonts count=\"1\"><font/></fonts><cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" applyFont=\"1\"><alignment wrapText=\"1\"/></xf></cellXfs></styleSheet>";
        assert_eq!(xf_count(doc), 2);
        let out = with_formats(doc, &[(1, "m/d/yyyy".into()), (0, "0.0\"件\"".into())]).unwrap();
        assert!(out.contains("<cellXfs count=\"4\">"), "{out}");
        // The built-in date format has its id; the other one is written.
        assert!(
            out.contains("<xf numFmtId=\"14\" fontId=\"1\" applyFont=\"1\" applyNumberFormat=\"1\"><alignment wrapText=\"1\"/></xf>"),
            "{out}"
        );
        assert!(
            out.starts_with("<styleSheet><numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"0.0&quot;件&quot;\"/></numFmts>"),
            "{out}"
        );
        assert!(
            out.contains("<xf numFmtId=\"164\" fontId=\"0\" applyNumberFormat=\"1\"/>"),
            "{out}"
        );
    }
}
