//! The objects over the cells of a sheet — shapes, form controls, ActiveX controls — placed and
//! set (`ras book edit`, the grid of the editor): added, moved and resized, given a name, a
//! text, a macro, a linked cell and the like, put before or behind the others, and deleted.
//!
//! The stand-in for Excel does not hold these objects, so they are written straight into the
//! parts of the file that describe them, as Excel saves them:
//!
//! * a shape is an anchor of the sheet's drawing (`xl/drawings/drawingN.xml`, DrawingML);
//! * a form control is a shape of the sheet's VML drawing (`xl/drawings/vmlDrawingN.vml`) with
//!   its `x:ClientData`; from Excel 2010 on also an entry of the sheet's `<controls>` with its
//!   properties (`xl/ctrlProps/ctrlPropN.xml`) and a hidden mirror in the drawing. A control
//!   added here is written the way Excel 2007 saves it (the VML shape alone), which every later
//!   Excel reads; one that has the newer parts is changed in all of them;
//! * an ActiveX control is an entry of the sheet's `<controls>`, its class and persisted data
//!   (`xl/activeX/activeXN.xml` and `.bin`, MS-OFORMS: `crate::oforms`) and a VML shape that
//!   says where it is and which cell it is linked to.
//!
//! Every other byte of the package stays what it was.

use std::collections::{BTreeMap, HashSet};

use gridcraft_core::{CellRef, RangeRef};
use serde_json::Value as Json;

use super::scan::{self, Edits, El};
use super::{Pack, Result};
use crate::book::Book;
use crate::book::view::Sizes;
use crate::view::controls::spid_number;

mod control;
mod shape;
#[cfg(test)]
mod tests;

pub(crate) use control::{decode_name, form_base};

/// A corner of an object on the grid: the column and the pixels into it, the row and the
/// pixels into it (as the grid draws at 100%).
pub type Marker = (u32, f64, u32, f64);

/// Where an object is put.
#[derive(Clone, Debug, PartialEq)]
pub enum Place {
    /// Over the cells of a range, from the top left of its first cell to the bottom right of
    /// its last one.
    Range(RangeRef),
    /// From one corner to the other.
    Markers(Marker, Marker),
    /// At the top left of a cell, with a width and a height in pixels.
    At(CellRef, f64, f64),
}

/// Where an object goes among the others.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum Order {
    Front,
    Back,
    Forward,
    Backward,
}

/// What is set of an object: the keys of `set` and their values.
pub type Settings = Vec<(String, Json)>;

/// An edit of the objects of a sheet.
#[derive(Clone, Debug, PartialEq)]
pub enum What {
    AddShape {
        shape: String,
        place: Place,
        set: Settings,
    },
    AddControl {
        control: String,
        place: Place,
        set: Settings,
    },
    AddActiveX {
        control: String,
        place: Place,
        set: Settings,
    },
    Place {
        object: String,
        place: Place,
    },
    Set {
        object: String,
        set: Settings,
    },
    Delete {
        object: String,
    },
    Order {
        object: String,
        to: Order,
    },
}

fn marker_of(v: &Json, key: &str) -> Result<Marker> {
    let wrong =
        || format!("`{key}` は [列, 列の中のピクセル, 行, 行の中のピクセル] で指定してください");
    let list = v.as_array().filter(|a| a.len() == 4).ok_or_else(wrong)?;
    let n = |i: usize| list[i].as_f64().filter(|x| *x >= 0.0).ok_or_else(wrong);
    let (col, row) = (n(0)?, n(2)?);
    if col.fract() != 0.0 || row.fract() != 0.0 || col >= 16384.0 || row >= 1_048_576.0 {
        return Err(wrong());
    }
    Ok((col as u32, n(1)?, row as u32, n(3)?))
}

fn place_of(o: &Json) -> Result<Place> {
    if let Some(from) = o.get("from") {
        let to = o.get("to").ok_or("`to` がありません")?;
        return Ok(Place::Markers(
            marker_of(from, "from")?,
            marker_of(to, "to")?,
        ));
    }
    if o.get("at").is_some() {
        let at = super::range_of(o, "at")?.start;
        let size = |key: &str| {
            o.get(key)
                .and_then(Json::as_f64)
                .filter(|v| (0.0..=20000.0).contains(v))
                .ok_or_else(|| format!("`{key}`（ピクセル）がありません"))
        };
        return Ok(Place::At(at, size("width")?, size("height")?));
    }
    if o.get("range").is_some() {
        return Ok(Place::Range(super::range_of(o, "range")?));
    }
    Err(
        "位置が指定されていません（`range`、`at` と `width` / `height`、または `from` と `to`）"
            .into(),
    )
}

fn settings_of(o: &Json, required: bool) -> Result<Settings> {
    match o.get("set") {
        Some(Json::Object(map)) => Ok(map.iter().map(|(k, v)| (k.clone(), v.clone())).collect()),
        Some(_) => Err("`set` は { \"キー\": 値 } で指定してください".into()),
        None if required => Err("`set` がありません".into()),
        None => Ok(Vec::new()),
    }
}

/// The edit of the objects that `kind` names, read from its JSON; None for another edit.
pub(super) fn parse(kind: &str, o: &Json) -> Result<Option<What>> {
    let text = |key: &str| {
        o.get(key)
            .and_then(Json::as_str)
            .map(str::to_string)
            .ok_or_else(|| format!("`{key}` がありません"))
    };
    Ok(Some(match kind {
        "addShape" => What::AddShape {
            shape: text("shape")?,
            place: place_of(o)?,
            set: settings_of(o, false)?,
        },
        "addControl" => What::AddControl {
            control: text("control")?,
            place: place_of(o)?,
            set: settings_of(o, false)?,
        },
        "addActiveX" => What::AddActiveX {
            control: text("control")?,
            place: place_of(o)?,
            set: settings_of(o, false)?,
        },
        "placeObject" => What::Place {
            object: text("object")?,
            place: place_of(o)?,
        },
        "setObject" => What::Set {
            object: text("object")?,
            set: settings_of(o, true)?,
        },
        "deleteObject" => What::Delete {
            object: text("object")?,
        },
        "orderObject" => What::Order {
            object: text("object")?,
            to: match text("to")?.as_str() {
                "front" => Order::Front,
                "back" => Order::Back,
                "forward" => Order::Forward,
                "backward" => Order::Backward,
                other => {
                    return Err(format!(
                        "`to` の `{other}` は指定できません（front / back / forward / backward）"
                    ));
                }
            },
        },
        _ => return Ok(None),
    }))
}

// ---------------------------------------------------------------------------- the package

const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const TYPES: &str = "[Content_Types].xml";

/// A relationship of a part: its id, its type and the part it points to.
#[derive(Clone, Debug)]
pub(super) struct Rel {
    pub id: String,
    pub kind: String,
    pub part: String,
}

/// The parts of the file, with those written and removed so far.
pub(super) struct Pkg<'a> {
    pack: Pack<'a>,
    replace: BTreeMap<String, Vec<u8>>,
    remove: HashSet<String>,
}

impl Pkg<'_> {
    /// The name the part has in the file (part names do not differ by case).
    fn entry(&self, part: &str) -> String {
        let part = part.trim_start_matches('/');
        if self.replace.contains_key(part) {
            return part.to_string();
        }
        self.pack.find(part).unwrap_or_else(|| part.to_string())
    }

    pub(super) fn has(&self, part: &str) -> bool {
        let name = self.entry(part);
        !self.remove.contains(&name)
            && (self.replace.contains_key(&name) || self.pack.find(&name).is_some())
    }

    pub(super) fn bytes(&mut self, part: &str) -> Result<Option<Vec<u8>>> {
        let name = self.entry(part);
        if self.remove.contains(&name) {
            return Ok(None);
        }
        if let Some(data) = self.replace.get(&name) {
            return Ok(Some(data.clone()));
        }
        self.pack.bytes(&name)
    }

    pub(super) fn text(&mut self, part: &str) -> Result<Option<String>> {
        match self.bytes(part)? {
            Some(data) => String::from_utf8(data)
                .map(Some)
                .map_err(|_| format!("{part} は UTF-8 ではないため、書き換えられません")),
            None => Ok(None),
        }
    }

    pub(super) fn put(&mut self, part: &str, data: Vec<u8>) {
        let name = self.entry(part);
        self.remove.remove(&name);
        self.replace.insert(name, data);
    }

    pub(super) fn delete(&mut self, part: &str) {
        let name = self.entry(part);
        self.replace.remove(&name);
        if self.pack.find(&name).is_some() {
            self.remove.insert(name);
        }
    }

    fn names(&self) -> Vec<String> {
        let mut names: Vec<String> = self
            .pack
            .names()
            .iter()
            .filter(|n| !self.remove.contains(*n))
            .cloned()
            .collect();
        for name in self.replace.keys() {
            if !names.contains(name) {
                names.push(name.clone());
            }
        }
        names
    }

    /// A name no part has: `xl/drawings/drawing` + N + `.xml`.
    pub(super) fn free(&self, stem: &str, ext: &str) -> String {
        let names: HashSet<String> = self.names().iter().map(|n| n.to_lowercase()).collect();
        (1..)
            .map(|n| format!("{stem}{n}{ext}"))
            .find(|name| !names.contains(&name.to_lowercase()))
            .expect("a number is free")
    }

    /// A part changed where `change` says; an error when it is not there or cannot be read.
    pub(super) fn change(
        &mut self,
        part: &str,
        flat: &[&str],
        change: impl FnOnce(&str, &El, &mut Edits) -> Result<()>,
    ) -> Result<()> {
        let doc = self
            .text(part)?
            .ok_or_else(|| format!("{part} がありません"))?;
        let root = scan::tree(&doc, flat).ok_or_else(|| format!("{part} を読み取れません"))?;
        let mut edits = Edits::default();
        change(&doc, &root, &mut edits)?;
        if edits.is_empty() {
            return Ok(());
        }
        let new = edits.apply(&doc);
        scan::tree(&new, flat).ok_or_else(|| format!("{part} を書き換えられませんでした"))?;
        self.put(part, new.into_bytes());
        Ok(())
    }

    fn rels_part(part: &str) -> String {
        let (dir, file) = part.rsplit_once('/').unwrap_or(("", part));
        format!("{dir}/_rels/{file}.rels")
    }

    pub(super) fn rels(&mut self, part: &str) -> Result<Vec<Rel>> {
        let dir = part.rsplit_once('/').map_or("", |p| p.0);
        let Some(doc) = self.text(&Self::rels_part(part))? else {
            return Ok(Vec::new());
        };
        let mut out = Vec::new();
        for tag in scan::tags(&doc, 0).filter(|t| !t.is_close() && t.local() == "Relationship") {
            let span = tag.start..tag.end;
            if scan::attr(&doc, span.clone(), "TargetMode").as_deref() == Some("External") {
                continue;
            }
            let (Some(id), Some(target)) = (
                scan::attr(&doc, span.clone(), "Id"),
                scan::attr(&doc, span.clone(), "Target"),
            ) else {
                continue;
            };
            out.push(Rel {
                id,
                kind: scan::attr(&doc, span, "Type").unwrap_or_default(),
                part: super::resolve(dir, &target),
            });
        }
        Ok(out)
    }

    /// The part a relationship of `part` points to.
    pub(super) fn target(&mut self, part: &str, id: &str) -> Result<Option<Rel>> {
        Ok(self.rels(part)?.into_iter().find(|r| r.id == id))
    }

    /// A relationship from `part` to `to`; its id.
    pub(super) fn relate(&mut self, part: &str, kind: &str, to: &str) -> Result<String> {
        let rels = Self::rels_part(part);
        // The target as it is written: from the folder of `part`.
        let from: Vec<&str> = part.split('/').collect();
        let dir = &from[..from.len() - 1];
        let target: Vec<&str> = to.split('/').collect();
        let common = dir
            .iter()
            .zip(&target)
            .take_while(|(a, b)| a == b)
            .count()
            .min(target.len() - 1);
        let mut path = vec![".."; dir.len() - common];
        path.extend(&target[common..]);
        let path = path.join("/");
        let doc = match self.text(&rels)? {
            Some(doc) => doc,
            None => "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>".to_string(),
        };
        let used: HashSet<String> = scan::tags(&doc, 0)
            .filter(|t| !t.is_close() && t.local() == "Relationship")
            .filter_map(|t| scan::attr(&doc, t.start..t.end, "Id"))
            .collect();
        let id = (1..)
            .map(|n| format!("rId{n}"))
            .find(|id| !used.contains(id))
            .expect("an id is free");
        let entry = format!(
            "<Relationship Id=\"{id}\" Type=\"{}\" Target=\"{}\"/>",
            scan::escape(kind),
            scan::escape(&path)
        );
        let root = scan::tree(&doc, &[]).ok_or_else(|| format!("{rels} を読み取れません"))?;
        let mut edits = Edits::default();
        if root.end == root.open.end {
            // `<Relationships .../>`
            let open = doc[root.open.clone()].trim_end_matches("/>").to_string();
            edits.replace(root.span(), format!("{open}>{entry}</{}>", root.name));
        } else {
            edits.insert(root.inner.end, entry);
        }
        self.put(&rels, edits.apply(&doc).into_bytes());
        Ok(id)
    }

    pub(super) fn unrelate(&mut self, part: &str, id: &str) -> Result<()> {
        let rels = Self::rels_part(part);
        let Some(doc) = self.text(&rels)? else {
            return Ok(());
        };
        let mut edits = Edits::default();
        for tag in scan::tags(&doc, 0).filter(|t| !t.is_close() && t.local() == "Relationship") {
            if scan::attr(&doc, tag.start..tag.end, "Id").as_deref() == Some(id) {
                edits.replace(tag.start..tag.end, "");
            }
        }
        if !edits.is_empty() {
            self.put(&rels, edits.apply(&doc).into_bytes());
        }
        Ok(())
    }

    /// The content type of a part (`part`), or of the parts with an extension that have none.
    pub(super) fn content_type(&mut self, part: Option<&str>, ext: &str, kind: &str) -> Result<()> {
        self.change(TYPES, &[], |doc, root, edits| {
            let p = scan::prefix(root.name);
            match part {
                Some(part) => {
                    let name = format!("/{part}");
                    let known = root.children.iter().any(|c| {
                        c.local() == "Override"
                            && c.attr(doc, "PartName")
                                .is_some_and(|n| n.eq_ignore_ascii_case(&name))
                    });
                    if !known {
                        edits.insert(
                            root.inner.end,
                            format!("<{p}Override PartName=\"{name}\" ContentType=\"{kind}\"/>"),
                        );
                    }
                }
                None => {
                    let known = root.children.iter().any(|c| {
                        c.local() == "Default"
                            && c.attr(doc, "Extension")
                                .is_some_and(|e| e.eq_ignore_ascii_case(ext))
                    });
                    if !known {
                        edits.insert(
                            root.inner.start,
                            format!("<{p}Default Extension=\"{ext}\" ContentType=\"{kind}\"/>"),
                        );
                    }
                }
            }
            Ok(())
        })
    }

    /// A part and what says its type go.
    pub(super) fn discard(&mut self, part: &str) -> Result<()> {
        self.delete(part);
        self.delete(&Self::rels_part(part));
        let name = format!("/{part}");
        self.change(TYPES, &[], |doc, root, edits| {
            for c in root.children.iter().filter(|c| c.local() == "Override") {
                if c.attr(doc, "PartName")
                    .is_some_and(|n| n.eq_ignore_ascii_case(&name))
                {
                    edits.replace(c.span(), "");
                }
            }
            Ok(())
        })
    }
}

// ---------------------------------------------------------------------------- the sheet

/// The sheet whose objects are edited.
pub(super) struct Sheet {
    pub part: String,
    pub name: String,
    pub sizes: Sizes,
    /// The workbook's default font: its name and size (pt).
    pub font: (String, f64),
}

/// A box on the grid.
#[derive(Clone, Copy, Debug)]
pub(super) struct Rect {
    pub from: Marker,
    pub to: Marker,
    /// In pixels from the top left of A1.
    pub x: f64,
    pub y: f64,
    pub w: f64,
    pub h: f64,
}

impl Sheet {
    pub(super) fn rect(&self, place: &Place) -> Result<Rect> {
        let s = &self.sizes;
        let at = |m: Marker| (s.x(m.0) + m.1, s.y(m.2) + m.3);
        let (from, to) = match *place {
            Place::Range(r) => (
                (r.start.col, 0.0, r.start.row, 0.0),
                s.marker(s.x(r.end.col + 1), s.y(r.end.row + 1)),
            ),
            Place::Markers(from, to) => {
                // (An offset past the end of its column or row is at its end.)
                let (x, y) = at(from);
                let (x2, y2) = at(to);
                (s.marker(x, y), s.marker(x2, y2))
            }
            Place::At(cell, w, h) => {
                let (x, y) = (s.x(cell.col), s.y(cell.row));
                (s.marker(x, y), s.marker(x + w, y + h))
            }
        };
        let (x, y) = at(from);
        let (x2, y2) = at(to);
        if x2 < x || y2 < y || (x2 - x < 1.0 && y2 - y < 1.0) {
            return Err("オブジェクトの大きさがありません（位置の指定を確かめてください）".into());
        }
        Ok(Rect {
            from,
            to,
            x,
            y,
            w: x2 - x,
            h: y2 - y,
        })
    }
}

pub(super) fn emu(px: f64) -> i64 {
    (px * 9525.0).round() as i64
}

/// A number as it is written in a style (`13.5`).
pub(super) fn short(v: f64) -> String {
    let s = format!("{v:.2}");
    s.trim_end_matches('0').trim_end_matches('.').to_string()
}

/// The elements of a worksheet, in the order its schema gives them.
const SHEET_ORDER: &[&str] = &[
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
    "extLst",
];

pub(super) fn find_all<'a, 'd>(el: &'a El<'d>, name: &str, out: &mut Vec<&'a El<'d>>) {
    for c in &el.children {
        if c.local() == name {
            out.push(c);
        }
        find_all(c, name, out);
    }
}

pub(super) fn find<'a, 'd>(el: &'a El<'d>, name: &str) -> Option<&'a El<'d>> {
    for c in &el.children {
        if c.local() == name {
            return Some(c);
        }
        if let Some(found) = find(c, name) {
            return Some(found);
        }
    }
    None
}

/// What a child of a worksheet is, by the schema: what newer Excel wraps (`mc:AlternateContent`)
/// is what it wraps.
fn sheet_child(el: &El) -> usize {
    let name = if el.local() == "AlternateContent" {
        ["controls", "oleObjects"]
            .into_iter()
            .find(|n| find(el, n).is_some())
            .unwrap_or("")
    } else {
        el.local()
    };
    SHEET_ORDER
        .iter()
        .position(|n| *n == name)
        .unwrap_or(usize::MAX)
}

/// Puts a new element `name` into a worksheet, at its place among the others.
pub(super) fn sheet_insert(root: &El, name: &str, xml: String, edits: &mut Edits) {
    let at = SHEET_ORDER
        .iter()
        .position(|n| *n == name)
        .unwrap_or(usize::MAX);
    let before = root
        .children
        .iter()
        .rev()
        .find(|c| sheet_child(c) != usize::MAX && sheet_child(c) < at);
    match before {
        Some(el) => edits.replace(el.end..el.end, xml),
        None => edits.insert(root.inner.start, xml),
    }
}

/// The prefix the document gives the relationships namespace (`r:`), and the declaration to
/// write on an element when the root has none.
pub(super) fn rel_prefix(doc: &str, root: &El) -> (String, String) {
    for (name, value) in scan::attrs(doc, root.open.clone()) {
        if let Some(prefix) = name.strip_prefix("xmlns:")
            && &doc[value] == REL
        {
            return (format!("{prefix}:"), String::new());
        }
    }
    ("r:".into(), format!(" xmlns:r=\"{REL}\""))
}

// ---------------------------------------------------------------------------- finding objects

pub(super) const OBJECTS: [&str; 5] = ["sp", "cxnSp", "grpSp", "pic", "graphicFrame"];

/// An object of the drawing: a child of its root — an anchor, or what wraps anchors
/// (`mc:AlternateContent`: the object for newer Excel and what stands for it in an older one).
pub(super) struct Item<'a, 'd> {
    pub top: &'a El<'d>,
    pub anchors: Vec<&'a El<'d>>,
    pub id: Option<u32>,
    pub name: String,
    /// The control it mirrors.
    pub spid: Option<u32>,
}

fn anchors_of<'a, 'd>(el: &'a El<'d>, out: &mut Vec<&'a El<'d>>) {
    if el.local().ends_with("Anchor") {
        out.push(el);
        return;
    }
    for c in &el.children {
        anchors_of(c, out);
    }
}

/// The object an anchor holds.
pub(super) fn object_of<'a, 'd>(anchor: &'a El<'d>) -> Option<&'a El<'d>> {
    anchor
        .children
        .iter()
        .find(|c| OBJECTS.contains(&c.local()))
        .or_else(|| {
            let wrapped = anchor.child("AlternateContent")?;
            ["Choice", "Fallback"].into_iter().find_map(|side| {
                wrapped
                    .child(side)?
                    .children
                    .iter()
                    .find(|c| OBJECTS.contains(&c.local()))
            })
        })
}

/// The properties every object has (`cNvPr`: its id and name).
pub(super) fn nv_pr<'a, 'd>(obj: &'a El<'d>) -> Option<&'a El<'d>> {
    obj.children
        .iter()
        .find(|c| c.local().starts_with("nv"))
        .and_then(|c| c.child("cNvPr"))
}

pub(super) fn items<'a, 'd>(doc: &'d str, root: &'a El<'d>) -> Vec<Item<'a, 'd>> {
    let mut out = Vec::new();
    for top in &root.children {
        let mut anchors = Vec::new();
        anchors_of(top, &mut anchors);
        let Some(pr) = anchors.iter().find_map(|a| object_of(a).and_then(nv_pr)) else {
            continue;
        };
        out.push(Item {
            top,
            id: pr.attr(doc, "id").and_then(|v| v.parse().ok()),
            name: pr.attr(doc, "name").unwrap_or_default(),
            spid: find(pr, "compatExt")
                .and_then(|c| c.attr(doc, "spid"))
                .and_then(|s| spid_number(&s)),
            anchors,
        });
    }
    out
}

/// The drawing of a sheet (shapes, pictures, charts).
pub(super) fn drawing_part(pkg: &mut Pkg, sheet: &Sheet) -> Result<Option<String>> {
    let mut parts: Vec<String> = pkg
        .rels(&sheet.part)?
        .into_iter()
        .filter(|r| r.kind.ends_with("/drawing"))
        .map(|r| r.part)
        .filter(|p| pkg.has(p))
        .collect();
    parts.sort();
    Ok(parts.into_iter().next())
}

/// The VML drawing of a sheet (controls, the boxes of notes).
pub(super) fn vml_part(pkg: &mut Pkg, sheet: &Sheet) -> Result<Option<String>> {
    let rels = pkg.rels(&sheet.part)?;
    let doc = pkg.text(&sheet.part)?.unwrap_or_default();
    // The one the sheet names (`legacyDrawing`); the header's pictures are another.
    let from = doc.find("</sheetData>").unwrap_or(0);
    let named = scan::tags(&doc, from)
        .find(|t| !t.is_close() && t.local() == "legacyDrawing")
        .and_then(|t| scan::attr(&doc, t.start..t.end, "id"))
        .and_then(|id| rels.iter().find(|r| r.id == id))
        .map(|r| r.part.clone());
    Ok(named.filter(|p| pkg.has(p)))
}

/// A control of a sheet, as its parts name it.
#[derive(Clone, Debug, Default)]
pub(super) struct Ctl {
    pub spid: u32,
    pub name: String,
    /// The object type of its VML shape (`Button`, `Checkbox`, ..., `Pict`).
    pub object: String,
    /// The relationship of its entry in the sheet's `<controls>`: the properties of a form
    /// control, or the ActiveX control.
    pub rel: Option<Rel>,
    /// It has a shape in the VML drawing.
    pub shape: bool,
}

impl Ctl {
    pub(super) fn activex(&self) -> bool {
        self.rel
            .as_ref()
            .is_some_and(|r| r.kind.ends_with("/control"))
    }
}

pub(super) fn controls(pkg: &mut Pkg, sheet: &Sheet) -> Result<Vec<Ctl>> {
    let mut list: Vec<Ctl> = Vec::new();
    let rels = pkg.rels(&sheet.part)?;
    let doc = pkg.text(&sheet.part)?.unwrap_or_default();
    if doc.contains("controls") {
        let from = doc.find("</sheetData>").unwrap_or(0);
        for tag in scan::tags(&doc, from).filter(|t| !t.is_close() && t.local() == "control") {
            let span = tag.start..tag.end;
            let Some(spid) =
                scan::attr(&doc, span.clone(), "shapeId").and_then(|s| spid_number(&s))
            else {
                continue;
            };
            if list.iter().any(|c| c.spid == spid) {
                continue;
            }
            list.push(Ctl {
                spid,
                name: scan::attr(&doc, span.clone(), "name").unwrap_or_default(),
                rel: scan::attr(&doc, span, "id")
                    .and_then(|id| rels.iter().find(|r| r.id == id).cloned()),
                ..Default::default()
            });
        }
    }
    if let Some(part) = vml_part(pkg, sheet)?
        && let Some(vml) = pkg.text(&part)?
        && let Some(root) = scan::tree(&vml, &[])
    {
        let mut shapes = Vec::new();
        find_all(&root, "shape", &mut shapes);
        for sh in shapes {
            let Some(data) = sh.child("ClientData") else {
                continue;
            };
            let object = data.attr(&vml, "ObjectType").unwrap_or_default();
            if object == "Note" {
                continue;
            }
            let Some(spid) = control::shape_id(&vml, sh) else {
                continue;
            };
            let at = match list.iter().position(|c| c.spid == spid) {
                Some(i) => i,
                None => {
                    list.push(Ctl {
                        spid,
                        ..Default::default()
                    });
                    list.len() - 1
                }
            };
            let ctl = &mut list[at];
            ctl.shape = true;
            if ctl.name.is_empty() {
                ctl.name = control::shape_name(&vml, sh)
                    .unwrap_or_else(|| crate::view::controls::default_name(&object, spid));
            }
            ctl.object = object;
        }
    }
    Ok(list)
}

/// What an edit names: an object of the drawing, a control, or a control with its mirror in
/// the drawing.
#[derive(Debug, Default)]
pub(super) struct Target {
    /// The index of the object among those of the drawing ([`items`]).
    pub item: Option<usize>,
    pub ctl: Option<Ctl>,
}

/// The object `object` names: by its name, or by its id on the grid (`s12`: the shape with the
/// id 12, `c1025`: the control with the shape id 1025).
pub(super) fn target(pkg: &mut Pkg, sheet: &Sheet, object: &str) -> Result<Target> {
    let ctls = controls(pkg, sheet)?;
    let drawing = drawing_part(pkg, sheet)?;
    let doc = match &drawing {
        Some(part) => pkg.text(part)?,
        None => None,
    };
    let root = doc.as_deref().and_then(|d| scan::tree(d, &[]));
    let list = match (&doc, &root) {
        (Some(doc), Some(root)) => items(doc, root),
        _ => Vec::new(),
    };
    let wanted = object.trim();
    let same = |name: &str| !name.is_empty() && name.to_lowercase() == wanted.to_lowercase();
    let number = |prefix: char| {
        wanted
            .strip_prefix(prefix)
            .filter(|d| !d.is_empty() && d.bytes().all(|b| b.is_ascii_digit()))
            .and_then(|d| d.parse::<u32>().ok())
    };
    let ctl = ctls
        .iter()
        .find(|c| same(&c.name))
        .or_else(|| number('c').and_then(|id| ctls.iter().find(|c| c.spid == id)));
    let item = match ctl {
        // The mirror of the control.
        Some(c) => list.iter().position(|i| i.spid == Some(c.spid)),
        None => list
            .iter()
            .position(|i| same(&i.name))
            .or_else(|| number('s').and_then(|id| list.iter().position(|i| i.id == Some(id)))),
    };
    // (A mirror named: its control.)
    let ctl = ctl.cloned().or_else(|| {
        let spid = list[item?].spid?;
        ctls.iter().find(|c| c.spid == spid).cloned()
    });
    if item.is_none() && ctl.is_none() {
        let mut names: Vec<String> = list
            .iter()
            .filter(|i| i.spid.is_none())
            .map(|i| i.name.clone())
            .chain(ctls.iter().map(|c| c.name.clone()))
            .filter(|n| !n.is_empty())
            .collect();
        names.truncate(20);
        return Err(format!(
            "シート `{}` にオブジェクト `{wanted}` がありません（{}）",
            sheet.name,
            if names.is_empty() {
                "オブジェクトはありません".to_string()
            } else {
                format!("オブジェクト: {}", names.join(", "))
            }
        ));
    }
    Ok(Target { item, ctl })
}

/// The names of the sheet's objects, to give a new one a name of its own.
pub(super) fn names(pkg: &mut Pkg, sheet: &Sheet) -> Result<HashSet<String>> {
    let mut out: HashSet<String> = controls(pkg, sheet)?
        .into_iter()
        .map(|c| c.name.to_lowercase())
        .collect();
    if let Some(part) = drawing_part(pkg, sheet)?
        && let Some(doc) = pkg.text(&part)?
    {
        for tag in scan::tags(&doc, 0).filter(|t| !t.is_close() && t.local() == "cNvPr") {
            if let Some(name) = scan::attr(&doc, tag.start..tag.end, "name") {
                out.insert(name.to_lowercase());
            }
        }
    }
    Ok(out)
}

/// `base` with the first number from `from` on that makes a name no object has.
pub(super) fn fresh_name(taken: &HashSet<String>, base: &str, from: u32, space: bool) -> String {
    (from.max(1)..)
        .map(|n| format!("{base}{}{n}", if space { " " } else { "" }))
        .find(|name| !taken.contains(&name.to_lowercase()))
        .expect("a number is free")
}

// ---------------------------------------------------------------------------- settings

pub(super) fn wrong(key: &str, value: &Json) -> String {
    format!("`{key}` の値 {value} は指定できません")
}

pub(super) fn text_of(key: &str, value: &Json) -> Result<String> {
    match value {
        Json::String(s) => Ok(s.clone()),
        Json::Null => Ok(String::new()),
        Json::Number(n) => Ok(n.to_string()),
        _ => Err(wrong(key, value)),
    }
}

pub(super) fn flag_of(key: &str, value: &Json) -> Result<bool> {
    value.as_bool().ok_or_else(|| wrong(key, value))
}

pub(super) fn int_of(key: &str, value: &Json) -> Result<i64> {
    value
        .as_f64()
        .filter(|v| v.fract() == 0.0 && v.abs() < 2_147_483_648.0)
        .map(|v| v as i64)
        .ok_or_else(|| wrong(key, value))
}

/// `#RRGGBB` as `RRGGBB`; None for null (no colour, the automatic one).
pub(super) fn colour_of(key: &str, value: &Json) -> Result<Option<String>> {
    match value {
        Json::Null => Ok(None),
        Json::String(s) => {
            let hex = s.trim().trim_start_matches('#');
            if hex.len() == 6 && hex.bytes().all(|b| b.is_ascii_hexdigit()) {
                Ok(Some(hex.to_ascii_uppercase()))
            } else {
                Err(format!(
                    "色 `{s}` を読み取れません（#RRGGBB で指定してください）"
                ))
            }
        }
        _ => Err(wrong(key, value)),
    }
}

/// A reference to cells as the file writes it (`$A$1`, `Sheet2!$A$1:$A$5`); empty for none.
pub(super) fn reference_of(key: &str, value: &Json) -> Result<String> {
    let text = text_of(key, value)?;
    let text = text.trim().trim_start_matches('=');
    if text.is_empty() {
        return Ok(String::new());
    }
    let (sheet, body) = gridcraft_core::split_sheet(text);
    let range = RangeRef::parse(&body.replace('$', "")).ok_or_else(|| {
        format!(
            "`{key}` の参照 `{text}` を読み取れません（A1、Sheet2!A1:A5 のように指定してください）"
        )
    })?;
    let cell = |c: CellRef| {
        format!(
            "${}${}",
            gridcraft_core::addr::col_to_letters(c.col),
            c.row + 1
        )
    };
    let area = if range.is_single() {
        cell(range.start)
    } else {
        format!("{}:{}", cell(range.start), cell(range.end))
    };
    Ok(match sheet {
        Some(s) => format!("{}!{area}", gridcraft_core::addr::quote_sheet(&s)),
        None => area,
    })
}

/// How an object follows the cells: (it moves with them, it is sized with them).
pub(super) fn placement_of(key: &str, value: &Json) -> Result<(bool, bool)> {
    match value.as_str() {
        Some("moveAndSize") => Ok((true, true)),
        Some("move") => Ok((true, false)),
        Some("free") => Ok((false, false)),
        _ => Err(format!(
            "`{key}` の値 {value} は指定できません（moveAndSize / move / free）"
        )),
    }
}

/// The macro of an object as the file names it: of this workbook.
pub(super) fn macro_ref(name: &str) -> String {
    let name = name.trim();
    if name.is_empty() || name.contains('!') {
        name.to_string()
    } else {
        format!("[0]!{name}")
    }
}

// ---------------------------------------------------------------------------- the edits

/// The markers of an anchor (`from`, `to`), written anew.
pub(super) fn set_marker(doc: &str, el: &El, m: Marker, edits: &mut Edits) {
    // (`<xdr:from>` holds `<xdr:col>`; the anchor of a control, `<from>` with `<xdr:col>`.)
    let p = el
        .children
        .first()
        .map_or_else(|| scan::prefix(el.name), |c| scan::prefix(c.name));
    let xml = format!(
        "<{p}col>{}</{p}col><{p}colOff>{}</{p}colOff><{p}row>{}</{p}row><{p}rowOff>{}</{p}rowOff>",
        m.0,
        emu(m.1),
        m.2,
        emu(m.3)
    );
    if el.end == el.open.end {
        let open = doc[el.open.clone()].trim_end_matches("/>").to_string();
        edits.replace(el.span(), format!("{open}>{xml}</{}>", el.name));
    } else {
        edits.replace(el.inner.clone(), xml);
    }
}

/// Moves an object of the drawing to `rect`: its anchor, and the box its shape keeps.
fn place_anchor(doc: &str, anchor: &El, rect: &Rect, edits: &mut Edits) {
    let cell = |name: &str, m: Marker, edits: &mut Edits| {
        if let Some(el) = anchor.child(name) {
            set_marker(doc, el, m, edits);
        }
    };
    let pair = |name: &str, keys: (&str, &str), v: (f64, f64), edits: &mut Edits| {
        if let Some(el) = anchor.child(name) {
            scan::set_attr(doc, el.open.clone(), keys.0, &emu(v.0).to_string(), edits);
            scan::set_attr(doc, el.open.clone(), keys.1, &emu(v.1).to_string(), edits);
        }
    };
    cell("from", rect.from, edits);
    cell("to", rect.to, edits);
    pair("pos", ("x", "y"), (rect.x, rect.y), edits);
    pair("ext", ("cx", "cy"), (rect.w, rect.h), edits);
    let Some(obj) = object_of(anchor) else {
        return;
    };
    let xfrm = obj
        .children
        .iter()
        .find(|c| matches!(c.local(), "spPr" | "grpSpPr"))
        .and_then(|p| p.child("xfrm"))
        .or_else(|| obj.child("xfrm"));
    let Some(xfrm) = xfrm else {
        return;
    };
    // A shape turned on its side keeps the box it has before it is turned.
    let rot = xfrm
        .attr(doc, "rot")
        .and_then(|v| v.parse::<f64>().ok())
        .unwrap_or(0.0)
        / 60000.0;
    let turned = matches!(rot.rem_euclid(180.0), r if (45.0..135.0).contains(&r));
    let (w, h) = if turned {
        (rect.h, rect.w)
    } else {
        (rect.w, rect.h)
    };
    let (x, y) = (rect.x + (rect.w - w) / 2.0, rect.y + (rect.h - h) / 2.0);
    if let Some(off) = xfrm.child("off") {
        scan::set_attr(doc, off.open.clone(), "x", &emu(x).to_string(), edits);
        scan::set_attr(doc, off.open.clone(), "y", &emu(y).to_string(), edits);
    }
    if let Some(ext) = xfrm.child("ext") {
        scan::set_attr(doc, ext.open.clone(), "cx", &emu(w).to_string(), edits);
        scan::set_attr(doc, ext.open.clone(), "cy", &emu(h).to_string(), edits);
    }
}

fn place(pkg: &mut Pkg, sheet: &Sheet, object: &str, place: &Place) -> Result<()> {
    let rect = sheet.rect(place)?;
    let t = target(pkg, sheet, object)?;
    if let (Some(index), Some(part)) = (t.item, drawing_part(pkg, sheet)?) {
        pkg.change(&part, &[], |doc, root, edits| {
            if let Some(item) = items(doc, root).get(index) {
                for a in &item.anchors {
                    place_anchor(doc, a, &rect, edits);
                }
            }
            Ok(())
        })?;
    }
    if let Some(ctl) = &t.ctl {
        control::place(pkg, sheet, ctl, &rect)?;
    }
    Ok(())
}

fn delete(pkg: &mut Pkg, sheet: &Sheet, object: &str) -> Result<()> {
    let t = target(pkg, sheet, object)?;
    if let (Some(index), Some(part)) = (t.item, drawing_part(pkg, sheet)?) {
        pkg.change(&part, &[], |doc, root, edits| {
            if let Some(item) = items(doc, root).get(index) {
                edits.replace(item.top.span(), "");
            }
            Ok(())
        })?;
    }
    if let Some(ctl) = &t.ctl {
        control::delete(pkg, sheet, ctl)?;
    }
    Ok(())
}

/// The order of `count` things after the one at `at` went where `to` says.
fn reorder(count: usize, at: usize, to: Order) -> Vec<usize> {
    let mut order: Vec<usize> = (0..count).collect();
    let moved = order.remove(at);
    let place = match to {
        Order::Front => order.len(),
        Order::Back => 0,
        Order::Forward => (at + 1).min(order.len()),
        Order::Backward => at.saturating_sub(1),
    };
    order.insert(place, moved);
    order
}

fn order(pkg: &mut Pkg, sheet: &Sheet, object: &str, to: Order) -> Result<()> {
    let t = target(pkg, sheet, object)?;
    if let (Some(index), Some(part)) = (t.item, drawing_part(pkg, sheet)?) {
        // The objects of the drawing are drawn in its order: the last one in front.
        return pkg.change(&part, &[], |doc, root, edits| {
            let list = items(doc, root);
            let new = reorder(list.len(), index, to);
            for (slot, from) in new.iter().enumerate() {
                if slot != *from {
                    edits.replace(list[slot].top.span(), &doc[list[*from].top.span()]);
                }
            }
            Ok(())
        });
    }
    match &t.ctl {
        Some(ctl) => control::order(pkg, sheet, ctl, to),
        None => Ok(()),
    }
}

fn set(pkg: &mut Pkg, sheet: &Sheet, object: &str, set: &Settings) -> Result<()> {
    if set.is_empty() {
        return Ok(());
    }
    let t = target(pkg, sheet, object)?;
    match &t.ctl {
        Some(ctl) => control::set(pkg, sheet, ctl, t.item, set),
        None => {
            let part = drawing_part(pkg, sheet)?.ok_or("シートに図形がありません")?;
            shape::set(
                pkg,
                sheet,
                &part,
                t.item.expect("an object of the drawing"),
                set,
            )
        }
    }
}

fn run(pkg: &mut Pkg, sheet: &Sheet, what: &What) -> Result<()> {
    match what {
        What::AddShape {
            shape,
            place,
            set: settings,
        } => {
            let rect = sheet.rect(place)?;
            let name = shape::add(pkg, sheet, shape, &rect)?;
            set(pkg, sheet, &name, settings)
        }
        What::AddControl {
            control,
            place,
            set: settings,
        } => {
            let rect = sheet.rect(place)?;
            let name = control::add_form(pkg, sheet, control, &rect)?;
            set(pkg, sheet, &name, settings)
        }
        What::AddActiveX {
            control,
            place,
            set: settings,
        } => {
            let rect = sheet.rect(place)?;
            let name = control::add_activex(pkg, sheet, control, &rect)?;
            set(pkg, sheet, &name, settings)
        }
        What::Place { object, place: to } => place(pkg, sheet, object, to),
        What::Set {
            object,
            set: settings,
        } => set(pkg, sheet, object, settings),
        What::Delete { object } => delete(pkg, sheet, object),
        What::Order { object, to } => order(pkg, sheet, object, *to),
    }
}

/// The first font of the styles: its name and size.
fn default_font(styles: Option<&str>) -> (String, f64) {
    let mut font = ("游ゴシック".to_string(), 11.0);
    let Some(doc) = styles else {
        return font;
    };
    let Some(start) = doc.find("<font>").or_else(|| doc.find("<font ")) else {
        return font;
    };
    let end = doc[start..]
        .find("</font>")
        .map_or(doc.len(), |i| start + i);
    for tag in scan::tags(&doc[..end], start).filter(|t| !t.is_close()) {
        let value = scan::attr(doc, tag.start..tag.end, "val");
        match (tag.local(), value) {
            ("name", Some(v)) if !v.is_empty() => font.0 = v,
            ("sz", Some(v)) => font.1 = v.parse().ok().filter(|s| *s > 0.0).unwrap_or(font.1),
            _ => {}
        }
    }
    font
}

/// The workbook file `bytes` with the edits of the objects made, in order: (the sheet, the
/// edit).
pub(super) fn apply(bytes: &[u8], ops: &[(&str, &What)]) -> Result<Vec<u8>> {
    let book = Book::open_bytes(bytes)?;
    let mut pkg = Pkg {
        pack: Pack::open(bytes)?,
        replace: BTreeMap::new(),
        remove: HashSet::new(),
    };
    let workbook = pkg
        .text(super::WORKBOOK)?
        .ok_or("ブックに xl/workbook.xml がありません")?;
    let parts = super::sheet_parts(&mut pkg.pack, &workbook)?;
    let font = default_font(pkg.text("xl/styles.xml")?.as_deref());
    for (sheet, what) in ops {
        let si = super::sheet_index(&book, sheet)?;
        let name = book.sheet_name(si).unwrap_or_default().to_string();
        let part = parts
            .get(&name)
            .and_then(|p| pkg.pack.find(p))
            .ok_or_else(|| format!("シート `{name}` のパートが見つかりません"))?;
        if !part.contains("/worksheets/") {
            return Err(format!(
                "シート `{name}` はワークシートではないため、オブジェクトを編集できません"
            ));
        }
        let sheet = Sheet {
            part,
            name,
            sizes: book.sizes(si)?,
            font: font.clone(),
        };
        run(&mut pkg, &sheet, what)?;
    }
    if pkg.replace.is_empty() && pkg.remove.is_empty() {
        return Ok(bytes.to_vec());
    }
    let out = pkg.pack.rewrite(&pkg.replace, &pkg.remove)?;
    // The new file is read as the old one was, or nothing is written.
    Book::open_bytes(&out).map_err(|e| {
        format!("書き換えたブックを読み戻せませんでした（{e}）。ファイルは変更していません")
    })?;
    Ok(out)
}
