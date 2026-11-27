//! Formats added to the file's styles: a cell that takes a number format when something is
//! typed into it (a date, a time, a percentage) gets the format it had with that number format.

use super::scan::{self, Edits};
use super::{NewXf, Result};
use crate::book::snapshot::Styles;

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

/// The id of a built-in number format of the code `code`.
pub(super) fn built_in_id(code: &str) -> Option<u32> {
    BUILT_IN.iter().find(|(_, c)| *c == code).map(|(id, _)| *id)
}

/// The styles with the cell style ハイパーリンク (built-in 8) Excel gives the cells of a new
/// link, added as Excel adds it where they have none (the standard font underlined, in the
/// theme's colour of links), and the number of its format among the styles' (`xfId`).
pub(super) fn with_hyperlink_style(doc: &str) -> Result<(String, u32)> {
    let root = scan::tree(doc, &[]).ok_or("xl/styles.xml を読み取れません")?;
    let named = root.child("cellStyles");
    if let Some(id) = named
        .iter()
        .flat_map(|s| &s.children)
        .find(|s| s.attr(doc, "builtinId").as_deref() == Some("8"))
        .and_then(|s| s.attr(doc, "xfId")?.parse().ok())
    {
        return Ok((doc.to_string(), id));
    }
    if !scan::prefix(root.name).is_empty() {
        return Err("このブックのスタイル（名前空間の接頭辞つき）には、ハイパーリンクのスタイルを追加できません".into());
    }
    let missing = |what: &str| format!("xl/styles.xml に {what} がありません");
    let fonts = root.child("fonts").filter(|f| f.end != f.open.end);
    let fonts = fonts.ok_or_else(|| missing("fonts"))?;
    let style_xfs = root.child("cellStyleXfs").filter(|f| f.end != f.open.end);
    let style_xfs = style_xfs.ok_or_else(|| missing("cellStyleXfs"))?;
    let named = named.filter(|f| f.end != f.open.end);
    let named = named.ok_or_else(|| missing("cellStyles"))?;
    let count =
        |el: &scan::El, name: &str| el.children.iter().filter(|c| c.local() == name).count();
    let (font_id, xf_id) = (count(fonts, "font"), count(style_xfs, "xf"));

    // The standard font (the first), its size, name and family as they are.
    let standard = fonts.children.iter().find(|c| c.local() == "font");
    let of = |n: &str| {
        standard
            .and_then(|f| f.child(n))
            .map_or("", |c| &doc[c.span()])
    };
    let font = format!(
        "<font><u/>{}<color theme=\"10\"/>{}{}{}{}</font>",
        of("sz"),
        of("name"),
        of("family"),
        of("charset"),
        of("scheme")
    );
    // The standard style's alignment (vertical centring, in a Japanese workbook).
    let align = style_xfs
        .children
        .iter()
        .find(|c| c.local() == "xf")
        .and_then(|x| x.child("alignment"))
        .map(|a| &doc[a.span()]);
    let xf = match align {
        Some(a) => format!(
            "<xf numFmtId=\"0\" fontId=\"{font_id}\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\">{a}</xf>"
        ),
        None => format!(
            "<xf numFmtId=\"0\" fontId=\"{font_id}\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"0\" applyFill=\"0\" applyBorder=\"0\" applyAlignment=\"0\" applyProtection=\"0\"/>"
        ),
    };
    let mut edits = Edits::default();
    for (el, add, n) in [
        (fonts, font, font_id + 1),
        (style_xfs, xf, xf_id + 1),
        (
            named,
            format!("<cellStyle name=\"ハイパーリンク\" xfId=\"{xf_id}\" builtinId=\"8\"/>"),
            count(named, "cellStyle") + 1,
        ),
    ] {
        edits.insert(el.inner.end, add);
        if el.attr(doc, "count").is_some() {
            scan::set_attr(doc, el.open.clone(), "count", &n.to_string(), &mut edits);
        }
    }
    let out = edits.apply(doc);
    scan::tree(&out, &[]).ok_or("xl/styles.xml を書き換えられませんでした")?;
    Ok((out, xf_id as u32))
}

/// How many cell formats (`cellXfs`) the styles have.
pub(super) fn xf_count(doc: &str) -> u32 {
    scan::tree(doc, &[])
        .and_then(|root| {
            root.child("cellXfs")
                .map(|x| x.children.iter().filter(|c| c.local() == "xf").count() as u32)
        })
        .unwrap_or(0)
}

/// A format written out by umya-spreadsheet, as a workbook's styles hold it: its `xf`, and the
/// font, fill, border and number format it points to.
struct Written {
    xf: String,
    font: Option<String>,
    fill: Option<String>,
    border: Option<String>,
    code: Option<String>,
}

/// The formats `styles` as umya-spreadsheet writes them: a workbook with one cell for each is
/// written, and its styles are read back as text.
fn written(styles: &[&umya_spreadsheet::Style]) -> Result<Vec<Written>> {
    use std::io::Read as _;
    let failed = |what: &str| format!("書式を書き出せませんでした（{what}）");
    let mut book = umya_spreadsheet::new_file();
    let sheet = book.sheet_mut(0).map_err(|e| failed(&e.to_string()))?;
    for (i, style) in styles.iter().enumerate() {
        let cell = sheet.cell_mut((1, i as u32 + 1));
        cell.set_value_number(0);
        cell.set_style((*style).clone());
    }
    let mut bytes = Vec::new();
    umya_spreadsheet::writer::xlsx::write_writer(&book, std::io::Cursor::new(&mut bytes))
        .map_err(|e| failed(&e.to_string()))?;
    let mut zip =
        zip::ZipArchive::new(std::io::Cursor::new(&bytes)).map_err(|e| failed(&e.to_string()))?;
    let mut part = |name: &str| -> Result<String> {
        let mut text = String::new();
        zip.by_name(name)
            .map_err(|e| failed(&e.to_string()))?
            .read_to_string(&mut text)
            .map_err(|e| failed(&e.to_string()))?;
        Ok(text)
    };
    let sheet = part("xl/worksheets/sheet1.xml")?;
    let doc = part("xl/styles.xml")?;
    let root = scan::tree(&doc, &[]).ok_or_else(|| failed("styles"))?;
    let list = |name: &str| -> Vec<&scan::El> {
        root.child(name)
            .map(|l| l.children.iter().collect())
            .unwrap_or_default()
    };
    let (fonts, fills, borders, xfs) = (
        list("fonts"),
        list("fills"),
        list("borders"),
        list("cellXfs"),
    );
    let codes: Vec<(String, String)> = list("numFmts")
        .iter()
        .filter_map(|f| Some((f.attr(&doc, "numFmtId")?, f.attr(&doc, "formatCode")?)))
        .collect();
    // The `xf` of each cell, in the order of the cells.
    let mut index: Vec<usize> = vec![0; styles.len()];
    for tag in scan::tags(&sheet, 0).filter(|t| !t.is_close() && t.local() == "c") {
        let span = tag.start..tag.end;
        let row = scan::attr(&sheet, span.clone(), "r")
            .and_then(|r| gridcraft_core::CellRef::parse(&r))
            .map(|at| at.row as usize);
        let s = scan::attr(&sheet, span, "s").and_then(|s| s.parse::<usize>().ok());
        if let (Some(row), Some(s)) = (row, s)
            && row < index.len()
        {
            index[row] = s;
        }
    }
    let mut out = Vec::new();
    for s in index {
        let xf = xfs.get(s).ok_or_else(|| failed("xf"))?;
        let pick = |name: &str, list: &[&scan::El]| -> Option<String> {
            let id = xf.attr(&doc, name)?.parse::<usize>().ok()?;
            list.get(id).map(|el| doc[el.span()].to_string())
        };
        let id = xf.attr(&doc, "numFmtId").unwrap_or_default();
        out.push(Written {
            xf: doc[xf.span()].to_string(),
            font: pick("fontId", &fonts),
            fill: pick("fillId", &fills),
            border: pick("borderId", &borders),
            code: codes
                .iter()
                .find(|(i, _)| *i == id)
                .map(|(_, c)| c.clone())
                .or_else(|| {
                    let id = id.parse::<u32>().ok()?;
                    BUILT_IN
                        .iter()
                        .find(|(i, _)| *i == id)
                        .map(|(_, c)| c.to_string())
                }),
        });
    }
    Ok(out)
}

/// A list of the styles (`fonts`, `fills`, `borders`) to which elements are added: each is
/// found among those there, or put after them.
struct Pool<'a> {
    el: Option<&'a scan::El<'a>>,
    have: Vec<&'a str>,
    added: Vec<String>,
}

impl<'a> Pool<'a> {
    fn new(doc: &'a str, root: &'a scan::El<'a>, name: &str) -> Pool<'a> {
        let el = root.child(name);
        Pool {
            el,
            have: el
                .iter()
                .flat_map(|l| &l.children)
                .map(|c| &doc[c.span()])
                .collect(),
            added: Vec::new(),
        }
    }

    fn id(&mut self, xml: &str) -> usize {
        if let Some(i) = self.have.iter().position(|h| *h == xml) {
            return i;
        }
        match self.added.iter().position(|a| a == xml) {
            Some(i) => self.have.len() + i,
            None => {
                self.added.push(xml.to_string());
                self.have.len() + self.added.len() - 1
            }
        }
    }

    fn write(&self, doc: &str, name: &str, edits: &mut Edits) -> Result<()> {
        if self.added.is_empty() {
            return Ok(());
        }
        let el = self
            .el
            .filter(|el| el.end != el.open.end)
            .ok_or_else(|| format!("xl/styles.xml に {name} がありません"))?;
        edits.insert(el.inner.end, self.added.concat());
        if el.attr(doc, "count").is_some() {
            scan::set_attr(
                doc,
                el.open.clone(),
                "count",
                &(self.have.len() + self.added.len()).to_string(),
                edits,
            );
        }
        Ok(())
    }
}

/// The styles with the formats of `new` added after the ones there.
///
/// Also: the formats of `new` written out whole that the file has already (as text), with the
/// index of the `xf` that is the same.
pub(super) fn with_formats(
    doc: &str,
    new: &[NewXf],
    styles: &Styles,
) -> Result<(String, Vec<(usize, u32)>)> {
    let root = scan::tree(doc, &[]).ok_or("xl/styles.xml を読み取れません")?;
    let p = scan::prefix(root.name);
    let mut edits = Edits::default();
    // The number formats written in the file, and the ids to give new ones.
    let formats = root.child("numFmts");
    let mut codes: Vec<(u32, String)> = formats
        .iter()
        .flat_map(|f| &f.children)
        .filter_map(|f| {
            Some((
                f.attr(doc, "numFmtId")?.parse().ok()?,
                f.attr(doc, "formatCode")?,
            ))
        })
        .collect();
    let had = codes.len();
    let mut id_of = |code: &str| -> u32 {
        if let Some((id, _)) = codes.iter().find(|(_, c)| c == code) {
            return *id;
        }
        if let Some((id, _)) = BUILT_IN.iter().find(|(_, c)| *c == code) {
            return *id;
        }
        let id = codes.iter().map(|(id, _)| *id).max().unwrap_or(0).max(163) + 1;
        codes.push((id, code.to_string()));
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
    let fresh: Vec<&umya_spreadsheet::Style> = new
        .iter()
        .filter_map(|n| match n {
            NewXf::Fresh(style) => styles.get(*style),
            NewXf::Derived(..) => None,
        })
        .collect();
    if !fresh.is_empty() && !p.is_empty() {
        return Err(
            "このブックのスタイル（名前空間の接頭辞つき）には、書式を追加できません".into(),
        );
    }
    let mut fresh = if fresh.is_empty() {
        Vec::new()
    } else {
        written(&fresh)?
    }
    .into_iter();
    let style_xf_count = root.child("cellStyleXfs").map_or(0, |x| {
        x.children.iter().filter(|c| c.local() == "xf").count()
    });
    let mut fonts = Pool::new(doc, &root, "fonts");
    let mut fills = Pool::new(doc, &root, "fills");
    let mut borders = Pool::new(doc, &root, "borders");
    let with = |text: &str, attrs: &[(&str, String)]| -> String {
        let end = text.find('>').map_or(text.len(), |i| i + 1);
        let mut local = Edits::default();
        for (name, value) in attrs {
            scan::set_attr(text, 0..end, name, value, &mut local);
        }
        local.apply(text)
    };
    let mut added = String::new();
    let mut found = Vec::new();
    for n in new {
        match n {
            NewXf::Derived(base, code) => {
                let id = id_of(code).to_string();
                match xfs.get(*base as usize) {
                    // The format it is made from, as it is written, with the number format.
                    Some(xf) => added.push_str(&with(
                        &doc[xf.span()],
                        &[("numFmtId", id), ("applyNumberFormat", "1".into())],
                    )),
                    None => added.push_str(&format!(
                        "<{p}xf numFmtId=\"{id}\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>"
                    )),
                }
            }
            NewXf::Fresh(style) => {
                let w = fresh.next().ok_or("書式を書き出せませんでした")?;
                // Its cell style (ハイパーリンク, ...), when the styles have it.
                let xf_id = styles
                    .get(*style)
                    .map(|s| s.format_id())
                    .filter(|id| (*id as usize) < style_xf_count)
                    .unwrap_or(0);
                let mut attrs: Vec<(&str, String)> = vec![("xfId", xf_id.to_string())];
                let id = match &w.code {
                    Some(code) => id_of(code),
                    None => 0,
                };
                attrs.push(("numFmtId", id.to_string()));
                if id != 0 {
                    attrs.push(("applyNumberFormat", "1".into()));
                }
                for (name, apply, pool, xml) in [
                    ("fontId", "applyFont", &mut fonts, &w.font),
                    ("fillId", "applyFill", &mut fills, &w.fill),
                    ("borderId", "applyBorder", &mut borders, &w.border),
                ] {
                    let id = xml.as_ref().map_or(0, |xml| pool.id(xml));
                    attrs.push((name, id.to_string()));
                    if id != 0 {
                        attrs.push((apply, "1".into()));
                    }
                }
                let xml = with(&w.xf, &attrs);
                if let Some(i) = xfs.iter().position(|xf| doc[xf.span()] == xml) {
                    found.push((*style, i as u32));
                }
                added.push_str(&xml);
            }
        }
    }
    if cell_xfs.end == cell_xfs.open.end {
        return Err("xl/styles.xml のセルの書式（cellXfs）が空です".into());
    }
    fonts.write(doc, "fonts", &mut edits)?;
    fills.write(doc, "fills", &mut edits)?;
    borders.write(doc, "borders", &mut edits)?;
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
    let fresh_codes: String = codes[had..]
        .iter()
        .map(|(id, code)| {
            format!(
                "<{p}numFmt numFmtId=\"{id}\" formatCode=\"{}\"/>",
                scan::escape(code)
            )
        })
        .collect();
    if !fresh_codes.is_empty() {
        let whole = format!(
            "<{p}numFmts count=\"{}\">{fresh_codes}</{p}numFmts>",
            codes.len()
        );
        match formats {
            Some(f) if f.end != f.open.end => {
                edits.insert(f.inner.end, fresh_codes);
                if f.attr(doc, "count").is_some() {
                    scan::set_attr(
                        doc,
                        f.open.clone(),
                        "count",
                        &codes.len().to_string(),
                        &mut edits,
                    );
                }
            }
            Some(f) => edits.replace(f.span(), whole),
            // The first element of the styles.
            None => edits.insert(root.open.end, whole),
        }
    }
    let out = edits.apply(doc);
    scan::tree(&out, &[]).ok_or("xl/styles.xml を書き換えられませんでした")?;
    Ok((out, found))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_style_hyperlink_is_added_as_excel_adds_it() {
        // A Japanese workbook's styles as Excel 2021 writes them, and as it writes them after a
        // link is put in.
        let doc = r#"<styleSheet><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="游ゴシック"/><family val="2"/><charset val="128"/><scheme val="minor"/></font></fonts><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles></styleSheet>"#;
        let (out, id) = with_hyperlink_style(doc).unwrap();
        assert_eq!(id, 1);
        assert!(out.contains(r#"<fonts count="2"><font><sz val="11"/><color theme="1"/><name val="游ゴシック"/><family val="2"/><charset val="128"/><scheme val="minor"/></font><font><u/><sz val="11"/><color theme="10"/><name val="游ゴシック"/><family val="2"/><charset val="128"/><scheme val="minor"/></font></fonts>"#), "{out}");
        assert!(out.contains(r#"<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf><xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"><alignment vertical="center"/></xf></cellStyleXfs>"#), "{out}");
        assert!(out.contains(r#"<cellStyles count="2"><cellStyle name="標準" xfId="0" builtinId="0"/><cellStyle name="ハイパーリンク" xfId="1" builtinId="8"/></cellStyles>"#), "{out}");
        // There once: the same.
        assert_eq!(with_hyperlink_style(&out).unwrap(), (out.clone(), 1));
    }

    #[test]
    fn a_format_is_added_with_its_number_format() {
        let doc = "<styleSheet><fonts count=\"1\"><font/></fonts><cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" applyFont=\"1\"><alignment wrapText=\"1\"/></xf></cellXfs></styleSheet>";
        assert_eq!(xf_count(doc), 2);
        let out = with_formats(
            doc,
            &[
                NewXf::Derived(1, "m/d/yyyy".into()),
                NewXf::Derived(0, "0.0\"件\"".into()),
            ],
            &Styles::default(),
        )
        .unwrap()
        .0;
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
