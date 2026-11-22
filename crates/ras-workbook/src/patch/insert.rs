//! What Excel's 挿入 tab puts on a sheet, written where Excel keeps it: hyperlinks
//! (`<hyperlinks>` of the sheet, an external one with its relationship), notes (the sheet's
//! comments part and the boxes of its VML drawing) and tables (a table part, `<tableParts>`).
//! New parts get their relationships and content types; nothing else of the file changes.

use std::collections::{BTreeMap, HashSet};

use gridcraft_core::{CellRef, RangeRef};

use super::scan::{self, Edits, El};
use super::{Link, Op, Pack, Result, WORKBOOK, around, sheet_index, sheet_parts};
use crate::book::Book;

const REL_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const CONTENT_TYPES: &str = "[Content_Types].xml";

/// The children of a worksheet, in the order the schema has them (`CT_Worksheet`).
const ORDER: [&str; 38] = [
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "sheetCalcPr",
    "sheetProtection",
    "protectedRanges",
    "scenarios",
    "autoFilter",
    "sortState",
    "dataConsolidate",
    "customSheetViews",
    "mergeCells",
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "drawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
];

// ---------------------------------------------------------------------------- the parts

/// A workbook file opened to put something on one of its sheets.
struct Sheet<'a> {
    pack: Pack<'a>,
    /// The sheet's part, and its XML.
    part: String,
    doc: String,
    /// The parts written: new contents, and the new parts.
    out: BTreeMap<String, Vec<u8>>,
}

impl<'a> Sheet<'a> {
    fn open(bytes: &'a [u8], sheet: &str) -> Result<Sheet<'a>> {
        let mut pack = Pack::open(bytes)?;
        let workbook = pack.text(WORKBOOK)?.ok_or("xl/workbook.xml がありません")?;
        let parts = sheet_parts(&mut pack, &workbook)?;
        let part = parts
            .get(sheet)
            .or_else(|| {
                parts
                    .iter()
                    .find(|(n, _)| n.to_lowercase() == sheet.to_lowercase())
                    .map(|(_, p)| p)
            })
            .and_then(|p| pack.find(p))
            .ok_or_else(|| format!("シート `{sheet}` がありません"))?;
        let doc = pack
            .text(&part)?
            .ok_or_else(|| format!("{part} を読めません"))?;
        Ok(Sheet {
            pack,
            part,
            doc,
            out: BTreeMap::new(),
        })
    }

    /// A part as it is now: as written, else as the file has it.
    fn text(&mut self, part: &str) -> Result<Option<String>> {
        if let Some(data) = self.out.get(part) {
            return Ok(Some(String::from_utf8_lossy(data).into_owned()));
        }
        match self.pack.find(part) {
            Some(name) => self.pack.text(&name),
            None => Ok(None),
        }
    }

    fn put(&mut self, part: &str, text: String) {
        let name = self.pack.find(part).unwrap_or_else(|| part.to_string());
        self.out.insert(name, text.into_bytes());
    }

    /// Whether the file (or what is written) has a part.
    fn has(&self, part: &str) -> bool {
        self.pack.find(part).is_some() || self.out.keys().any(|k| k.eq_ignore_ascii_case(part))
    }

    /// The first free name `<head><n><tail>` (n from 1).
    fn free(&self, head: &str, tail: &str) -> String {
        (1..)
            .map(|n| format!("{head}{n}{tail}"))
            .find(|name| !self.has(name))
            .unwrap_or_default()
    }

    /// The relationships part of the sheet.
    fn rels_part(&self) -> String {
        let (dir, file) = self.part.rsplit_once('/').unwrap_or(("", &self.part));
        format!("{dir}/_rels/{file}.rels")
    }

    /// A relationship of the sheet added: its id.
    fn add_rel(&mut self, kind: &str, target: &str, external: bool) -> Result<String> {
        let part = self.rels_part();
        let doc = self.text(&part)?.unwrap_or_else(|| {
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>".to_string()
        });
        let used: HashSet<String> = scan::tags(&doc, 0)
            .filter(|t| !t.is_close() && t.local() == "Relationship")
            .filter_map(|t| scan::attr(&doc, t.start..t.end, "Id"))
            .collect();
        let id = (1..)
            .map(|n| format!("rId{n}"))
            .find(|id| !used.contains(id))
            .unwrap_or_default();
        let root = scan::tree(&doc, &[]).ok_or_else(|| format!("{part} を読み取れません"))?;
        let p = scan::prefix(root.name);
        let mode = if external {
            " TargetMode=\"External\""
        } else {
            ""
        };
        let rel = format!(
            "<{p}Relationship Id=\"{id}\" Type=\"{REL}{kind}\" Target=\"{}\"{mode}/>",
            scan::escape(target)
        );
        let mut edits = Edits::default();
        if root.open.end == root.end {
            let head = doc[root.open.clone()].trim_end_matches("/>");
            edits.replace(root.open.clone(), format!("{head}>{rel}</{}>", root.name));
        } else {
            edits.insert(root.inner.end, rel);
        }
        let text = edits.apply(&doc);
        self.put(&part, text);
        Ok(id)
    }

    /// The relationships of the sheet with these ids taken away.
    fn remove_rels(&mut self, ids: &HashSet<String>) -> Result<()> {
        if ids.is_empty() {
            return Ok(());
        }
        let part = self.rels_part();
        let Some(doc) = self.text(&part)? else {
            return Ok(());
        };
        let mut edits = Edits::default();
        for tag in scan::tags(&doc, 0).filter(|t| !t.is_close() && t.local() == "Relationship") {
            if scan::attr(&doc, tag.start..tag.end, "Id").is_some_and(|id| ids.contains(&id)) {
                let end = if tag.is_empty() {
                    tag.end
                } else {
                    let close = format!("</{}>", tag.name);
                    doc[tag.end..]
                        .find(&close)
                        .map_or(tag.end, |i| tag.end + i + close.len())
                };
                edits.replace(tag.start..end, "");
            }
        }
        if !edits.is_empty() {
            let text = edits.apply(&doc);
            self.put(&part, text);
        }
        Ok(())
    }

    /// The relationships of the sheet: (id, type, target, external).
    fn rels(&mut self) -> Result<Vec<(String, String, String, bool)>> {
        let part = self.rels_part();
        let Some(doc) = self.text(&part)? else {
            return Ok(Vec::new());
        };
        let dir = self
            .part
            .rsplit_once('/')
            .map_or("", |(d, _)| d)
            .to_string();
        Ok(scan::tags(&doc, 0)
            .filter(|t| !t.is_close() && t.local() == "Relationship")
            .filter_map(|t| {
                let span = t.start..t.end;
                let id = scan::attr(&doc, span.clone(), "Id")?;
                let target = scan::attr(&doc, span.clone(), "Target")?;
                let external =
                    scan::attr(&doc, span.clone(), "TargetMode").as_deref() == Some("External");
                let kind = scan::attr(&doc, span, "Type").unwrap_or_default();
                let target = if external {
                    target
                } else {
                    super::resolve(&dir, &target)
                };
                Some((id, kind, target, external))
            })
            .collect())
    }

    /// The content type of a new part (`Override`), or of an extension (`Default`).
    fn content_type(&mut self, part: Option<&str>, ext: Option<&str>, kind: &str) -> Result<()> {
        let doc = self
            .text(CONTENT_TYPES)?
            .ok_or("[Content_Types].xml がありません")?;
        let root = scan::tree(&doc, &[]).ok_or("[Content_Types].xml を読み取れません")?;
        let p = scan::prefix(root.name);
        let entry = match (part, ext) {
            (Some(part), _) => {
                let name = format!("/{part}");
                if root.children.iter().any(|c| {
                    c.local() == "Override"
                        && c.attr(&doc, "PartName")
                            .is_some_and(|n| n.eq_ignore_ascii_case(&name))
                }) {
                    return Ok(());
                }
                format!(
                    "<{p}Override PartName=\"{}\" ContentType=\"{kind}\"/>",
                    scan::escape(&name)
                )
            }
            (None, Some(ext)) => {
                if root.children.iter().any(|c| {
                    c.local() == "Default"
                        && c.attr(&doc, "Extension")
                            .is_some_and(|e| e.eq_ignore_ascii_case(ext))
                }) {
                    return Ok(());
                }
                // The defaults come before the overrides.
                let at = root
                    .children
                    .iter()
                    .rfind(|c| c.local() == "Default")
                    .map_or(root.inner.start, |c| c.end);
                let mut edits = Edits::default();
                edits.insert(
                    at,
                    format!("<{p}Default Extension=\"{ext}\" ContentType=\"{kind}\"/>"),
                );
                let text = edits.apply(&doc);
                self.put(CONTENT_TYPES, text);
                return Ok(());
            }
            (None, None) => return Ok(()),
        };
        let mut edits = Edits::default();
        edits.insert(root.inner.end, entry);
        let text = edits.apply(&doc);
        self.put(CONTENT_TYPES, text);
        Ok(())
    }

    /// The file with the sheet's new XML and the parts written.
    fn finish(mut self, sheet_doc: Option<String>) -> Result<Vec<u8>> {
        if let Some(doc) = sheet_doc {
            let part = self.part.clone();
            self.out.insert(part, doc.into_bytes());
        }
        let out = std::mem::take(&mut self.out);
        self.pack.rewrite(&out, &HashSet::new())
    }
}

/// A part's target from a sheet in `xl/worksheets/` (`../comments1.xml`).
fn relative(from_part: &str, to: &str) -> String {
    let from: Vec<&str> = from_part.split('/').collect();
    let from = &from[..from.len() - 1];
    let to: Vec<&str> = to.split('/').collect();
    let common = from
        .iter()
        .zip(&to)
        .take_while(|(a, b)| a.eq_ignore_ascii_case(b))
        .count();
    let mut out: Vec<&str> = vec![".."; from.len() - common];
    out.extend(&to[common..]);
    out.join("/")
}

/// The sheet's tree (its cells not read).
fn tree<'d>(doc: &'d str, part: &str) -> Result<El<'d>> {
    scan::tree(doc, &["sheetData", "extLst"]).ok_or_else(|| format!("{part} を読み取れません"))
}

/// `r:id="rId1"` for an element of the sheet `doc` (declaring the relationships' namespace
/// when its root does not).
fn rid_attr(doc: &str, root: &El, id: &str) -> String {
    let prefix = scan::attrs(doc, root.open.clone())
        .into_iter()
        .find(|(name, span)| {
            name.starts_with("xmlns:") && scan::unescape(&doc[span.clone()]) == REL_NS
        })
        .map(|(name, _)| name["xmlns:".len()..].to_string());
    match prefix {
        Some(r) => format!("{r}:id=\"{id}\""),
        None => format!("xmlns:r=\"{REL_NS}\" r:id=\"{id}\""),
    }
}

/// Where a new child `name` of the worksheet goes: after the children the schema puts before
/// it.
fn place(root: &El, name: &str) -> usize {
    let rank = |n: &str| ORDER.iter().position(|o| *o == n).unwrap_or(ORDER.len());
    let mine = rank(name);
    root.children
        .iter()
        .take_while(|c| rank(c.local()) < mine)
        .last()
        .map_or(root.inner.start, |c| c.end)
}

/// The value of an attribute named `local` with a namespace prefix (`r:id`).
fn prefixed_attr(doc: &str, el: &El, local: &str) -> Option<String> {
    scan::attrs(doc, el.open.clone())
        .into_iter()
        .find(|(name, _)| {
            name.contains(':') && scan::local(name) == local && !name.starts_with("xmlns")
        })
        .map(|(_, span)| scan::unescape(&doc[span]))
}

// ---------------------------------------------------------------------------- hyperlinks

/// The file with the cells of `range` of the sheet `sheet` linked to `link` (their hyperlinks
/// replaced), or, `link` None, without hyperlinks.
pub(super) fn link(
    bytes: &[u8],
    sheet: &str,
    range: RangeRef,
    link: Option<&Link>,
) -> Result<Vec<u8>> {
    if range.count() > 1_000_000 {
        return Err("リンクを付ける範囲が大きすぎます".into());
    }
    let mut s = Sheet::open(bytes, sheet)?;
    let sdoc = s.doc.clone();
    let root = tree(&sdoc, &s.part)?;
    let p = scan::prefix(root.name).to_string();
    let mut edits = Edits::default();
    // The hyperlinks there go (those of other cells stay, and their relationships with them).
    let (mut gone, mut kept) = (HashSet::new(), HashSet::new());
    let list = root.child("hyperlinks");
    let mut left = 0;
    for h in list
        .iter()
        .flat_map(|l| &l.children)
        .filter(|c| c.local() == "hyperlink")
    {
        let rid = prefixed_attr(&sdoc, h, "id");
        let here = h
            .attr(&s.doc, "ref")
            .and_then(|r| RangeRef::parse(&r.replace('$', "")))
            .is_some_and(|r| r.intersects(&range));
        if here {
            edits.replace(h.span(), "");
            gone.extend(rid);
        } else {
            left += 1;
            kept.extend(rid);
        }
    }
    let unused: HashSet<String> = gone.difference(&kept).cloned().collect();
    if link.is_none() && gone.is_empty() {
        return Ok(bytes.to_vec());
    }
    let new = match link {
        None => String::new(),
        Some(l) => {
            let mut el = format!("<{p}hyperlink ref=\"{}\"", range.a1());
            if l.location {
                el.push_str(&format!(" location=\"{}\"", scan::escape(&l.target)));
            }
            if !l.display.is_empty() {
                el.push_str(&format!(" display=\"{}\"", scan::escape(&l.display)));
            }
            if !l.tip.is_empty() {
                el.push_str(&format!(" tooltip=\"{}\"", scan::escape(&l.tip)));
            }
            if !l.location {
                let id = s.add_rel("hyperlink", &l.target, true)?;
                el.push(' ');
                el.push_str(&rid_attr(&sdoc, &root, &id));
            }
            el.push_str("/>");
            el
        }
    };
    match list {
        Some(list) if new.is_empty() && left == 0 => edits.replace(list.span(), ""),
        Some(list) if list.open.end == list.end => {
            edits.replace(list.span(), format!("<{p}hyperlinks>{new}</{p}hyperlinks>"))
        }
        Some(list) => edits.insert(list.inner.end, new),
        None => edits.insert(
            place(&root, "hyperlinks"),
            format!("<{p}hyperlinks>{new}</{p}hyperlinks>"),
        ),
    }
    let doc = edits.apply(&sdoc);
    s.remove_rels(&unused)?;
    s.finish(Some(doc))
}

// ---------------------------------------------------------------------------- notes

/// The file with the note of the cell `at` of the sheet `sheet` saying `text` (in place of the
/// one it has; by `author` when it is new), or, `text` None, without a note.
pub(super) fn note(
    bytes: &[u8],
    sheet: &str,
    at: CellRef,
    text: Option<&str>,
    author: &str,
) -> Result<Vec<u8>> {
    let mut s = Sheet::open(bytes, sheet)?;
    let rels = s.rels()?;
    let a1 = at.a1();
    // A comment of the threads ("コメント") is Excel's to change: its note is a copy of it.
    for (_, kind, part, _) in rels.iter().filter(|r| r.1.ends_with("/threadedComment")) {
        if let Some(doc) = s.text(part)? {
            if doc.contains(&format!("ref=\"{a1}\"")) {
                return Err(format!(
                    "{a1} にはスレッド形式のコメントがあります。Excel で編集してください（{kind}）"
                ));
            }
        }
    }
    let sdoc = s.doc.clone();
    let root = tree(&sdoc, &s.part)?;
    let p = scan::prefix(root.name).to_string();
    let mut sheet_edits = Edits::default();

    // The comments part: the text of the note.
    let comments = rels
        .iter()
        .find(|r| r.1 == format!("{REL}comments"))
        .map(|r| r.2.clone());
    let had = match &comments {
        Some(part) => {
            let doc = s
                .text(part)?
                .ok_or_else(|| format!("{part} がありません"))?;
            let (doc, had) = comment_edit(&doc, &a1, text, author)
                .ok_or_else(|| format!("{part} を読み取れません"))?;
            s.put(part, doc);
            had
        }
        None => {
            let Some(text) = text else {
                return Ok(bytes.to_vec());
            };
            let part = s.free("xl/comments", ".xml");
            let doc = format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors><author>{}</author></authors><commentList>{}</commentList></comments>",
                scan::escape(author),
                comment_xml("", &a1, 0, text)
            );
            s.put(&part, doc);
            s.add_rel("comments", &relative(&s.part.clone(), &part), false)?;
            s.content_type(
                Some(&part),
                None,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
            )?;
            false
        }
    };
    if text.is_none() && !had {
        return Ok(bytes.to_vec());
    }

    // The VML drawing: the note's box (hidden until the cell is pointed at).
    let legacy = root.child("legacyDrawing");
    let vml = legacy
        .and_then(|l| prefixed_attr(&sdoc, l, "id"))
        .and_then(|id| rels.iter().find(|r| r.0 == id))
        .map(|r| r.2.clone());
    match vml {
        Some(part) => {
            let doc = s
                .text(&part)?
                .ok_or_else(|| format!("{part} がありません"))?;
            if let Some(doc) = vml_edit(&doc, at, text.is_some()) {
                s.put(&part, doc);
            }
        }
        None if text.is_some() => {
            // A drawing of its own, its shapes numbered in a block of 1024 no other has.
            let vmls: Vec<String> = s
                .pack
                .names()
                .iter()
                .filter(|n| n.to_ascii_lowercase().ends_with(".vml"))
                .cloned()
                .collect();
            let mut block = 0;
            for n in vmls {
                if let Some(doc) = s.pack.text(&n)? {
                    block = idmap(&doc).into_iter().fold(block, u32::max);
                }
            }
            let block = block + 1;
            let part = s.free("xl/drawings/vmlDrawing", ".vml");
            let doc = format!(
                "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\n <o:shapelayout v:ext=\"edit\">\n  <o:idmap v:ext=\"edit\" data=\"{block}\"/>\n </o:shapelayout>{}{}</xml>",
                NOTE_SHAPETYPE,
                note_shape(block * 1024 + 1, at)
            );
            s.put(&part, doc);
            let id = s.add_rel("vmlDrawing", &relative(&s.part.clone(), &part), false)?;
            s.content_type(
                None,
                Some("vml"),
                "application/vnd.openxmlformats-officedocument.vmlDrawing",
            )?;
            sheet_edits.insert(
                place(&root, "legacyDrawing"),
                format!("<{p}legacyDrawing {}/>", rid_attr(&sdoc, &root, &id)),
            );
        }
        None => {}
    }
    let doc = if sheet_edits.is_empty() {
        None
    } else {
        Some(sheet_edits.apply(&sdoc))
    };
    s.finish(doc)
}

/// A note, as Excel writes it.
fn comment_xml(p: &str, a1: &str, author: usize, text: &str) -> String {
    format!(
        "<{p}comment ref=\"{a1}\" authorId=\"{author}\">{}</{p}comment>",
        comment_text(p, text)
    )
}

fn comment_text(p: &str, text: &str) -> String {
    format!(
        "<{p}text><{p}r><{p}rPr><{p}sz val=\"9\"/><{p}color indexed=\"81\"/><{p}rFont val=\"MS P ゴシック\"/><{p}family val=\"3\"/><{p}charset val=\"128\"/></{p}rPr><{p}t xml:space=\"preserve\">{}</{p}t></{p}r></{p}text>",
        scan::escape(&text.replace("\r\n", "\n"))
    )
}

/// The comments part with the note of `a1` written (or taken away): the part, and whether the
/// cell had a note. None when the part cannot be read.
fn comment_edit(doc: &str, a1: &str, text: Option<&str>, author: &str) -> Option<(String, bool)> {
    let root = scan::tree(doc, &["text"])?;
    let p = scan::prefix(root.name);
    let mut edits = Edits::default();
    let list = root.child("commentList");
    let found = list.and_then(|l| {
        l.children.iter().find(|c| {
            c.local() == "comment"
                && c.attr(doc, "ref")
                    .is_some_and(|r| r.replace('$', "").eq_ignore_ascii_case(a1))
        })
    });
    match (found, text) {
        (Some(c), None) => edits.replace(c.span(), ""),
        (Some(c), Some(text)) => match c.child("text") {
            Some(t) => edits.replace(t.span(), comment_text(p, text)),
            None => edits.insert(c.inner.end, comment_text(p, text)),
        },
        (None, None) => return Some((doc.to_string(), false)),
        (None, Some(text)) => {
            // The author's number (added to the authors when new).
            let authors = root.child("authors");
            let names: Vec<String> = authors
                .iter()
                .flat_map(|a| &a.children)
                .filter(|a| a.local() == "author")
                .map(|a| a.text(doc))
                .collect();
            let index = match names.iter().position(|n| n == author) {
                Some(i) => i,
                None => {
                    let el = format!("<{p}author>{}</{p}author>", scan::escape(author));
                    match authors {
                        Some(a) if a.open.end == a.end => {
                            edits.replace(a.span(), format!("<{p}authors>{el}</{p}authors>"))
                        }
                        Some(a) => edits.insert(a.inner.end, el),
                        None => {
                            edits.insert(root.inner.start, format!("<{p}authors>{el}</{p}authors>"))
                        }
                    }
                    names.len()
                }
            };
            let el = comment_xml(p, a1, index, text);
            match list {
                Some(l) if l.open.end == l.end => {
                    edits.replace(l.span(), format!("<{p}commentList>{el}</{p}commentList>"))
                }
                Some(l) => edits.insert(l.inner.end, el),
                None => {
                    let at = authors.map_or(root.inner.start, |a| a.end);
                    edits.insert(at, format!("<{p}commentList>{el}</{p}commentList>"));
                }
            }
        }
    }
    Some((edits.apply(doc), found.is_some()))
}

/// The blocks of shape ids a VML drawing has (`<o:idmap data="1,2">`).
fn idmap(doc: &str) -> Vec<u32> {
    scan::tags(doc, 0)
        .filter(|t| !t.is_close() && t.local() == "idmap")
        .filter_map(|t| scan::attr(doc, t.start..t.end, "data"))
        .flat_map(|d| {
            d.split(|c: char| c == ',' || c.is_whitespace())
                .filter_map(|n| n.trim().parse().ok())
                .collect::<Vec<u32>>()
        })
        .collect()
}

const NOTE_SHAPETYPE: &str = "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"\n  path=\"m,l,21600r21600,l21600,xe\">\n  <v:stroke joinstyle=\"miter\"/>\n  <v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/>\n </v:shapetype>";

/// The box of a note, as Excel draws a new one: right of the cell, hidden.
fn note_shape(id: u32, at: CellRef) -> String {
    let (r, c) = (at.row, at.col);
    let top = r.saturating_sub(1);
    // The cell's column and row from 0; the box from the next column, a row above.
    format!(
        "<v:shape id=\"_x0000_s{id}\" type=\"#_x0000_t202\" style='position:absolute;\n  margin-left:{left}pt;margin-top:{mtop}pt;width:108pt;height:59.25pt;z-index:1;\n  visibility:hidden' fillcolor=\"infoBackground [80]\" strokecolor=\"none [81]\" o:insetmode=\"auto\">\n  <v:fill color2=\"infoBackground [80]\"/>\n  <v:shadow color=\"none [81]\" obscured=\"t\"/>\n  <v:path o:connecttype=\"none\"/>\n  <v:textbox style='mso-direction-alt:auto'>\n   <div style='text-align:left'></div>\n  </v:textbox>\n  <x:ClientData ObjectType=\"Note\">\n   <x:MoveWithCells/>\n   <x:SizeWithCells/>\n   <x:Anchor>\n    {c1}, 15, {top}, {toff}, {c2}, 15, {bottom}, 4</x:Anchor>\n   <x:AutoFill>False</x:AutoFill>\n   <x:Row>{r}</x:Row>\n   <x:Column>{c}</x:Column>\n  </x:ClientData>\n </v:shape>",
        left = f64::from(c + 1) * 54.0 + 11.25,
        mtop = f64::from(top) * 13.5 + if r == 0 { 1.5 } else { 7.5 },
        c1 = c + 1,
        toff = if r == 0 { 2 } else { 10 },
        c2 = c + 3,
        bottom = top + 4,
    )
}

/// The VML drawing with a box for the note of `at` (`on`), or without the box it has. None:
/// nothing to change.
fn vml_edit(doc: &str, at: CellRef, on: bool) -> Option<String> {
    let root = scan::tree(doc, &[])?;
    let is_note = |el: &El| {
        el.local() == "shape"
            && el.children.iter().any(|d| {
                d.local() == "ClientData"
                    && d.attr(doc, "ObjectType").as_deref() == Some("Note")
                    && d.child("Row").map(|r| r.text(doc).trim().to_string())
                        == Some(at.row.to_string())
                    && d.child("Column").map(|r| r.text(doc).trim().to_string())
                        == Some(at.col.to_string())
            })
    };
    let found = root.children.iter().find(|c| is_note(c));
    let mut edits = Edits::default();
    match (found, on) {
        (Some(_), true) | (None, false) => return None,
        (Some(shape), false) => edits.replace(shape.span(), ""),
        (None, true) => {
            let blocks = idmap(doc);
            let used: Vec<u32> = scan::tags(doc, 0)
                .filter(|t| !t.is_close() && t.local() == "shape")
                .filter_map(|t| scan::attr(doc, t.start..t.end, "id"))
                .filter_map(|id| id.strip_prefix("_x0000_s")?.parse().ok())
                .collect();
            let base = blocks.first().copied().unwrap_or(1) * 1024;
            let id = used
                .iter()
                .copied()
                .max()
                .map_or(base + 1, |m| m.max(base) + 1);
            let mut add = String::new();
            let typed = root.children.iter().any(|c| {
                c.local() == "shapetype" && c.attr(doc, "id").as_deref() == Some("_x0000_t202")
            });
            if !typed {
                add.push_str(NOTE_SHAPETYPE);
            }
            add.push_str(&note_shape(id, at));
            edits.insert(root.inner.end, add);
        }
    }
    Some(edits.apply(doc))
}

// ---------------------------------------------------------------------------- tables

/// The headings of a table: what its first row shows, each a text of its own (empty: 列1,
/// ...; one shown before: a number after it), as Excel names the columns.
fn headings(shown: &[String]) -> Vec<String> {
    let mut out: Vec<String> = Vec::new();
    let taken =
        |out: &Vec<String>, n: &str| out.iter().any(|o| o.to_lowercase() == n.to_lowercase());
    for (i, text) in shown.iter().enumerate() {
        let text = text.replace(['\r', '\n'], " ");
        let mut name = if text.trim().is_empty() {
            format!("列{}", i + 1)
        } else {
            text.chars().take(255).collect()
        };
        if taken(&out, &name) {
            let base = name.clone();
            name = (2..)
                .map(|n| format!("{base}{n}"))
                .find(|n| !taken(&out, n))
                .unwrap_or(base);
        }
        out.push(name);
    }
    out
}

/// The edits of the cells a table made of `range` needs first — the sheet's filter taken
/// away where it is in the way, a row of headings inserted (`header` false), the headings made
/// texts of their own — and the table's range after them.
pub(super) fn table_cells(
    bytes: &[u8],
    sheet: &str,
    range: RangeRef,
    header: bool,
) -> Result<(Vec<Op>, RangeRef)> {
    let mut book = Book::open_bytes(bytes)?;
    let si = sheet_index(&book, sheet)?;
    let info = book.sheet_summary(si)?;
    let mut range = around(&mut book, si, range);
    if range.is_full_rows() || range.is_full_cols() {
        // Whole rows or columns: as much of them as is used.
        range = info
            .used_range
            .and_then(|u| u.intersection(&range))
            .ok_or("テーブルにするデータがありません")?;
    }
    if range.width() > 16_384 || range.count() > 10_000_000 {
        return Err("テーブルにする範囲が大きすぎます".into());
    }
    for t in &info.tables {
        if RangeRef::parse(&t.range).is_some_and(|r| r.intersects(&range)) {
            return Err(format!(
                "テーブルは重ねられません（{} の範囲 {} と重なります）",
                t.name, t.range
            ));
        }
    }
    for m in &info.merged {
        if RangeRef::parse(&m.replace('$', "")).is_some_and(|r| r.intersects(&range)) {
            return Err(format!(
                "結合したセル（{m}）を含む範囲はテーブルにできません。結合を解除してください"
            ));
        }
    }
    let name = book.sheet_name(si).unwrap_or(sheet).to_string();
    let mut ops = Vec::new();
    if info
        .filter
        .as_deref()
        .and_then(|f| RangeRef::parse(&f.replace('$', "")))
        .is_some_and(|f| f.intersects(&range))
    {
        // The table's own filter takes the place of the sheet's.
        ops.push(Op::Filter {
            sheet: name.clone(),
            range: None,
        });
    }
    let (c1, c2) = (range.start.col, range.end.col);
    let row = |r: u32| RangeRef::new(CellRef::new(r, c1), CellRef::new(r, c2));
    let shown: Vec<(CellRef, String, bool)> = if header {
        book.cells(si, row(range.start.row))?
            .into_iter()
            .map(|c| {
                let plain = c.formula.is_none() && matches!(c.value, crate::book::Val::Text(_));
                (c.at, c.text, plain)
            })
            .collect()
    } else {
        if range.end.row + 1 >= super::MAX_ROWS {
            return Err("見出しの行を挿入できません（シートの最後の行にデータがあります）".into());
        }
        ops.push(Op::Insert {
            sheet: name.clone(),
            range: row(range.start.row),
            right: false,
        });
        range = RangeRef::new(range.start, CellRef::new(range.end.row + 1, range.end.col));
        row(range.start.row)
            .iter()
            .map(|at| (at, String::new(), false))
            .collect()
    };
    let names = headings(&shown.iter().map(|s| s.1.clone()).collect::<Vec<_>>());
    for ((at, text, plain), heading) in shown.iter().zip(&names) {
        if !(*plain && text == heading) {
            ops.push(Op::Set {
                sheet: name.clone(),
                range: RangeRef::cell(*at),
                text: format!("'{heading}"),
            });
        }
    }
    if range.height() == 1 {
        // Only headings: a row for the data under them.
        if range.end.row + 1 >= super::MAX_ROWS {
            return Err("テーブルにはデータの行が必要です".into());
        }
        range = RangeRef::new(range.start, CellRef::new(range.end.row + 1, range.end.col));
    }
    Ok((ops, range))
}

/// Whether `name` can name a table: a letter, `_` or `\` first, then letters, digits, `_` and
/// `.`; not a cell's address (A1, R1C1) nor R or C.
fn table_name_ok(name: &str) -> bool {
    let mut chars = name.chars();
    let Some(first) = chars.next() else {
        return false;
    };
    if !(first.is_alphabetic() || first == '_' || first == '\\') {
        return false;
    }
    if !name
        .chars()
        .all(|c| c.is_alphanumeric() || c == '_' || c == '.' || c == '\\')
    {
        return false;
    }
    if name.chars().count() > 255 || matches!(name.to_ascii_uppercase().as_str(), "R" | "C") {
        return false;
    }
    let upper = name.to_ascii_uppercase();
    let r1c1 = upper
        .strip_prefix('R')
        .map(|rest| {
            let digits = rest.trim_start_matches(|c: char| c.is_ascii_digit());
            digits
                .strip_prefix('C')
                .is_some_and(|d| d.chars().all(|c| c.is_ascii_digit()))
        })
        .unwrap_or(false);
    !(r1c1 || CellRef::parse(name).is_some())
}

/// The file with a table made of `range` of the sheet `sheet`, its headings the texts of its
/// first row (`table_cells` made them so), named `name` (or テーブル1, ...).
pub(super) fn table(
    bytes: &[u8],
    sheet: &str,
    range: RangeRef,
    name: Option<&str>,
) -> Result<Vec<u8>> {
    let mut book = Book::open_bytes(bytes)?;
    let si = sheet_index(&book, sheet)?;
    let head = RangeRef::new(range.start, CellRef::new(range.start.row, range.end.col));
    let shown: Vec<String> = book.cells(si, head)?.into_iter().map(|c| c.text).collect();
    let columns = headings(&shown);

    let mut s = Sheet::open(bytes, sheet)?;
    // The tables of the workbook (ids and names are the workbook's), and its names.
    let mut taken: HashSet<String> = HashSet::new();
    let mut top = 0u32;
    let table_parts: Vec<String> = s
        .pack
        .names()
        .iter()
        .filter(|n| {
            n.to_ascii_lowercase().starts_with("xl/tables/")
                && n.to_ascii_lowercase().ends_with(".xml")
        })
        .cloned()
        .collect();
    for part in table_parts {
        let Some(doc) = s.pack.text(&part)? else {
            continue;
        };
        if let Some(t) = scan::tags(&doc, 0).find(|t| !t.is_close() && t.local() == "table") {
            let span = t.start..t.end;
            if let Some(id) =
                scan::attr(&doc, span.clone(), "id").and_then(|i| i.parse::<u32>().ok())
            {
                top = top.max(id);
            }
            for key in ["name", "displayName"] {
                if let Some(n) = scan::attr(&doc, span.clone(), key) {
                    taken.insert(n.to_lowercase());
                }
            }
        }
    }
    let workbook = s.pack.text(WORKBOOK)?.unwrap_or_default();
    for t in scan::tags(&workbook, 0).filter(|t| !t.is_close() && t.local() == "definedName") {
        if let Some(n) = scan::attr(&workbook, t.start..t.end, "name") {
            taken.insert(n.to_lowercase());
        }
    }
    let name = match name {
        Some(n) => {
            let n = n.trim();
            if !table_name_ok(n) {
                return Err(format!(
                    "テーブル名 `{n}` は使えません（先頭は文字か _、ほかは文字・数字・_・. で、セルの番地と同じものは使えません）"
                ));
            }
            if taken.contains(&n.to_lowercase()) {
                return Err(format!("名前 `{n}` は既に使われています"));
            }
            n.to_string()
        }
        None => (1..)
            .map(|n| format!("テーブル{n}"))
            .find(|n| !taken.contains(&n.to_lowercase()))
            .unwrap_or_default(),
    };
    let id = top + 1;
    let a1 = range.a1();
    let cols: String = columns
        .iter()
        .enumerate()
        .map(|(i, c)| {
            format!(
                "<tableColumn id=\"{}\" name=\"{}\"/>",
                i + 1,
                scan::escape(c)
            )
        })
        .collect();
    let part = s.free("xl/tables/table", ".xml");
    let n = scan::escape(&name);
    s.put(
        &part,
        format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"{id}\" name=\"{n}\" displayName=\"{n}\" ref=\"{a1}\" totalsRowShown=\"0\"><autoFilter ref=\"{a1}\"/><tableColumns count=\"{}\">{cols}</tableColumns><tableStyleInfo name=\"TableStyleMedium2\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/></table>",
            columns.len()
        ),
    );
    let rid = s.add_rel("table", &relative(&s.part.clone(), &part), false)?;
    s.content_type(
        Some(&part),
        None,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
    )?;
    let sdoc = s.doc.clone();
    let root = tree(&sdoc, &s.part)?;
    let p = scan::prefix(root.name).to_string();
    let el = format!("<{p}tablePart {}/>", rid_attr(&sdoc, &root, &rid));
    let mut edits = Edits::default();
    match root.child("tableParts") {
        Some(list) => {
            let count = list
                .children
                .iter()
                .filter(|c| c.local() == "tablePart")
                .count()
                + 1;
            if list.open.end == list.end {
                edits.replace(
                    list.span(),
                    format!("<{p}tableParts count=\"{count}\">{el}</{p}tableParts>"),
                );
            } else {
                scan::set_attr(
                    &sdoc,
                    list.open.clone(),
                    "count",
                    &count.to_string(),
                    &mut edits,
                );
                edits.insert(list.inner.end, el);
            }
        }
        None => edits.insert(
            place(&root, "tableParts"),
            format!("<{p}tableParts count=\"1\">{el}</{p}tableParts>"),
        ),
    }
    let doc = edits.apply(&sdoc);
    s.finish(Some(doc))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn targets_names_and_headings() {
        assert_eq!(
            relative("xl/worksheets/sheet1.xml", "xl/comments2.xml"),
            "../comments2.xml"
        );
        assert_eq!(
            relative("xl/worksheets/sheet1.xml", "xl/tables/table3.xml"),
            "../tables/table3.xml"
        );
        assert_eq!(
            relative("xl/worksheets/sheet1.xml", "xl/worksheets/x.xml"),
            "x.xml"
        );
        assert!(table_name_ok("売上表") && table_name_ok("_t.1") && table_name_ok("Table1"));
        for bad in ["", "1a", "A1", "R1C1", "r2c", "C", "a b", "x-y"] {
            assert!(!table_name_ok(bad), "{bad}");
        }
        let h = |v: &[&str]| headings(&v.iter().map(|s| s.to_string()).collect::<Vec<_>>());
        assert_eq!(
            h(&["a", "", "A", "a2", "1"]),
            ["a", "列2", "A2", "a22", "1"]
        );
    }
}
