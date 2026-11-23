//! Form controls (開発 > 挿入 > フォーム コントロール): put on a sheet, moved, sized and deleted.
//!
//! A form control is kept in up to four places, which must agree: its shape in the sheet's VML
//! drawing (`x:ClientData`: its kind, anchor, macro, linked cell, ...: what Excel 2007 and
//! LibreOffice read), and, as Excel 2010 and later save it, an entry of the sheet's `<controls>`
//! (its name and anchor), its properties (`xl/ctrlProps/ctrlProp<n>.xml`) and a hidden mirror in
//! the sheet's drawing (`a14:compatExt`). A control is changed in all the places it is in. A new
//! one is written as the sheet keeps the others: in VML alone on a sheet without `<controls>`
//! (as Excel 2007 and XlsxWriter write them), in all four on one with them.

use super::drawing::{
    A_NS, XDR_NS, box_of, drawing_of, drop_part, emu, first, marker_point, name_of, ns_prefix,
    sizes_of, units, write_anchor, write_xfrm,
};
use super::insert::{Sheet, idmap, place, prefixed_attr, relative, rid_attr, tree};
use super::scan::{self, Edits, El};
use super::{FormControl, Place, Result};
use crate::book::view::Sizes;

const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const X14_NS: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
const A14_NS: &str = "http://schemas.microsoft.com/office/drawing/2010/main";
const CTRL_PROP_TYPE: &str = "application/vnd.ms-excel.controlproperties+xml";
const VML_TYPE: &str = "application/vnd.openxmlformats-officedocument.vmlDrawing";
/// The VML shape type of form controls.
const CONTROL_SHAPETYPE: &str = "<v:shapetype id=\"_x0000_t201\" coordsize=\"21600,21600\" o:spt=\"201\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/><v:path shadowok=\"f\" o:extrusionok=\"f\" strokeok=\"f\" fillok=\"f\" o:connecttype=\"rect\"/><o:lock v:ext=\"edit\" shapetype=\"t\"/></v:shapetype>";
const FONT: &str = "ＭＳ Ｐゴシック";

/// A kind of form control: its key in the edit, its `ObjectType` in VML and in its properties,
/// the name Excel gives a new one (followed by its number), its size when it is put with a
/// click (px, as Excel 2021 puts it on a sheet of 游ゴシック 11), and whether it shows a text.
pub(super) struct Kind {
    pub key: &'static str,
    vml: &'static str,
    prop: &'static str,
    name: &'static str,
    size: (f64, f64),
    text: bool,
}

pub(super) const KINDS: [Kind; 9] = [
    Kind {
        key: "button",
        vml: "Button",
        prop: "Button",
        name: "ボタン",
        size: (62.0, 30.0),
        text: true,
    },
    Kind {
        key: "checkBox",
        vml: "Checkbox",
        prop: "CheckBox",
        name: "チェック",
        size: (65.0, 25.0),
        text: true,
    },
    Kind {
        key: "optionButton",
        vml: "Radio",
        prop: "Radio",
        name: "オプション",
        size: (77.0, 25.0),
        text: true,
    },
    Kind {
        key: "label",
        vml: "Label",
        prop: "Label",
        name: "ラベル",
        size: (42.0, 19.0),
        text: true,
    },
    Kind {
        key: "groupBox",
        vml: "GBox",
        prop: "GBox",
        name: "グループ",
        size: (98.0, 96.0),
        text: true,
    },
    Kind {
        key: "listBox",
        vml: "List",
        prop: "List",
        name: "リスト",
        size: (96.0, 94.0),
        text: false,
    },
    Kind {
        key: "comboBox",
        vml: "Drop",
        prop: "Drop",
        name: "ドロップ",
        size: (96.0, 96.0),
        text: false,
    },
    Kind {
        key: "spinner",
        vml: "Spin",
        prop: "Spin",
        name: "スピン",
        size: (20.0, 40.0),
        text: false,
    },
    Kind {
        key: "scrollBar",
        vml: "Scroll",
        prop: "Scroll",
        name: "スクロール",
        size: (96.0, 288.0),
        text: false,
    },
];

// ---------------------------------------------------------------------------- where they are

/// The number of a VML shape (1025 of `_x0000_s1025`, by its `o:spid` or its `id`).
fn spid_of(doc: &str, el: &El) -> Option<u32> {
    ["o:spid", "id"]
        .iter()
        .find_map(|n| el.attr(doc, n)?.strip_prefix("_x0000_s")?.parse().ok())
}

/// A control of the sheet's `<controls>`: its element, and what is taken away with it (its
/// `mc:AlternateContent`).
struct Entry<'a, 'd> {
    el: &'a El<'d>,
    span: std::ops::Range<usize>,
}

/// The sheet's `<controls>` (and what is taken away with it, its `mc:AlternateContent`) and
/// its controls.
fn sheet_controls<'a, 'd>(
    root: &'a El<'d>,
) -> Option<(&'a El<'d>, std::ops::Range<usize>, Vec<Entry<'a, 'd>>)> {
    let in_choice = |ac: &'a El<'d>, name: &str| -> Option<&'a El<'d>> {
        ac.children
            .iter()
            .filter(|c| c.local() == "Choice")
            .find_map(|c| c.children.iter().find(|x| x.local() == name))
    };
    let (controls, outer) = root.children.iter().find_map(|c| match c.local() {
        "controls" => Some((c, c.span())),
        "AlternateContent" => in_choice(c, "controls").map(|x| (x, c.span())),
        _ => None,
    })?;
    let list = controls
        .children
        .iter()
        .filter_map(|c| match c.local() {
            "control" => Some(Entry {
                el: c,
                span: c.span(),
            }),
            "AlternateContent" => in_choice(c, "control").map(|x| Entry {
                el: x,
                span: c.span(),
            }),
            _ => None,
        })
        .collect();
    Some((controls, outer, list))
}

fn shape_id(doc: &str, el: &El) -> Option<u32> {
    el.attr(doc, "shapeId")?.parse().ok()
}

/// The VML drawing of the sheet (its `legacyDrawing`).
fn vml_part(s: &mut Sheet, root: &El, sdoc: &str) -> Result<Option<String>> {
    let rels = s.rels()?;
    Ok(root
        .child("legacyDrawing")
        .and_then(|l| prefixed_attr(sdoc, l, "id"))
        .and_then(|id| rels.iter().find(|r| r.0 == id))
        .map(|r| r.2.clone()))
}

/// The VML shapes (`v:shape`) of a VML drawing, anywhere in it.
fn vml_shapes<'a, 'd>(el: &'a El<'d>, out: &mut Vec<&'a El<'d>>) {
    for c in &el.children {
        if c.local() == "shape" {
            out.push(c);
        } else {
            vml_shapes(c, out);
        }
    }
}

fn object_type(doc: &str, shape: &El) -> String {
    shape
        .child("ClientData")
        .and_then(|d| d.attr(doc, "ObjectType"))
        .unwrap_or_default()
}

/// The control of the sheet `sheet` named `name` (its name in `<controls>` or its mirror in the
/// drawing), or whose shape is `shape`: its shape number. None when it is not a control.
pub(super) fn find(
    bytes: &[u8],
    sheet: &str,
    name: &str,
    shape: Option<u32>,
) -> Result<Option<u32>> {
    let mut s = Sheet::open(bytes, sheet)?;
    let sdoc = s.doc.clone();
    let root = tree(&sdoc, &s.part)?;
    if let Some(id) = shape {
        let in_sheet = sheet_controls(&root)
            .is_some_and(|(_, _, list)| list.iter().any(|e| shape_id(&sdoc, e.el) == Some(id)));
        let in_vml = match vml_part(&mut s, &root, &sdoc)? {
            Some(part) => s.text(&part)?.is_some_and(|doc| {
                scan::tree(&doc, &[]).is_some_and(|r| {
                    let mut all = Vec::new();
                    vml_shapes(&r, &mut all);
                    all.iter()
                        .any(|x| spid_of(&doc, x) == Some(id) && object_type(&doc, x) != "Note")
                })
            }),
            None => false,
        };
        return Ok((in_sheet || in_vml).then_some(id));
    }
    if let Some((_, _, list)) = sheet_controls(&root)
        && let Some(e) = list
            .iter()
            .find(|e| e.el.attr(&sdoc, "name").as_deref() == Some(name))
    {
        return Ok(shape_id(&sdoc, e.el));
    }
    // The mirror in the drawing.
    if let Some(d) = drawing_of(&mut s, false)? {
        let droot = scan::tree(&d.doc, &[]).ok_or("描画のパートを読み取れません")?;
        for u in units(&droot) {
            for a in &u.anchors {
                if name_of(&d.doc, a).as_deref() == Some(name)
                    && let Some(spid) = first(a, "compatExt").and_then(|c| c.attr(&d.doc, "spid"))
                {
                    return Ok(spid.strip_prefix("_x0000_s").and_then(|n| n.parse().ok()));
                }
            }
        }
    }
    Ok(None)
}

// ---------------------------------------------------------------------------- moved, deleted

/// A VML anchor: the columns and rows with the offsets in pixels.
fn vml_anchor(sizes: &Sizes, b: (f64, f64, f64, f64)) -> String {
    let (c1, dx1, r1, dy1) = sizes.marker(b.0, b.1);
    let (c2, dx2, r2, dy2) = sizes.marker(b.2, b.3);
    let r = |v: f64| v.round() as i64;
    format!(
        "{c1}, {}, {r1}, {}, {c2}, {}, {r2}, {}",
        r(dx1),
        r(dy1),
        r(dx2),
        r(dy2)
    )
}

fn vml_box(sizes: &Sizes, text: &str) -> Option<(f64, f64, f64, f64)> {
    let v: Vec<f64> = text
        .split(',')
        .map(|n| n.trim().parse::<f64>())
        .collect::<std::result::Result<_, _>>()
        .ok()?;
    (v.len() == 8).then(|| {
        (
            sizes.x(v[0] as u32) + v[1],
            sizes.y(v[2] as u32) + v[3],
            sizes.x(v[4] as u32) + v[5],
            sizes.y(v[6] as u32) + v[7],
        )
    })
}

/// A CSS-like VML style with the position and size of a box (pt) written in it.
fn vml_style(style: &str, b: (f64, f64, f64, f64)) -> String {
    let pt = |px: f64| format!("{}pt", (px * 0.75 * 100.0).round() / 100.0);
    let set = [
        ("margin-left", pt(b.0)),
        ("margin-top", pt(b.1)),
        ("width", pt(b.2 - b.0)),
        ("height", pt(b.3 - b.1)),
    ];
    let mut parts: Vec<String> = style
        .split(';')
        .map(str::trim)
        .filter(|p| !p.is_empty())
        .map(str::to_string)
        .collect();
    for (key, value) in set {
        let item = format!("{key}:{value}");
        match parts
            .iter()
            .position(|p| p.split(':').next().is_some_and(|k| k.trim() == key))
        {
            Some(i) => parts[i] = item,
            None => parts.push(item),
        }
    }
    parts.join(";")
}

/// The markers of an anchor of `<controls>` (`<from><xdr:col>...`), with the prefixes it has.
fn control_marker(tag_name: &str, inner: &str, sizes: &Sizes, x: f64, y: f64) -> String {
    let (col, dx, row, dy) = sizes.marker(x, y);
    format!(
        "<{tag_name}><{inner}col>{col}</{inner}col><{inner}colOff>{}</{inner}colOff><{inner}row>{row}</{inner}row><{inner}rowOff>{}</{inner}rowOff></{tag_name}>",
        emu(dx),
        emu(dy)
    )
}

/// The file with the form control whose shape is `id` on the sheet `sheet` moved and sized to
/// `place` (its size where the place gives none), or (`delete`) taken away with its properties.
pub(super) fn change(
    bytes: &[u8],
    sheet: &str,
    id: u32,
    place: Option<&Place>,
    delete: bool,
) -> Result<Vec<u8>> {
    if !delete && place.is_none() {
        return Ok(bytes.to_vec());
    }
    let sizes = sizes_of(bytes, sheet)?;
    let mut s = Sheet::open(bytes, sheet)?;
    let sdoc = s.doc.clone();
    let root = tree(&sdoc, &s.part)?;
    let rels = s.rels()?;
    let spid = format!("_x0000_s{id}");

    // The sheet's entry: an ActiveX control and an embedded object are not form controls.
    let controls = sheet_controls(&root);
    let entry = controls
        .as_ref()
        .and_then(|(_, _, list)| list.iter().find(|e| shape_id(&sdoc, e.el) == Some(id)));
    let entry_rel = entry
        .and_then(|e| prefixed_attr(&sdoc, e.el, "id"))
        .and_then(|rid| rels.iter().find(|r| r.0 == rid).cloned());
    // An ActiveX control: its part (`xl/activeX/activeX<n>.xml`).
    let activex = entry_rel
        .as_ref()
        .is_some_and(|r| r.1.ends_with("/control"));
    if sdoc.contains(&format!("shapeId=\"{id}\"")) && entry.is_none() {
        return Err(
            "埋め込みオブジェクトは、ここでは移動・削除できません（Excel で行ってください）".into(),
        );
    }

    // Its VML shape.
    let vml = vml_part(&mut s, &root, &sdoc)?;
    let vml_doc = match &vml {
        Some(p) => s.text(p)?,
        None => None,
    };
    let vml_root = vml_doc
        .as_deref()
        .map(|d| scan::tree(d, &[]).ok_or("VML の描画を読み取れません（Excel で行ってください）"));
    let vml_root = vml_root.transpose()?;
    let shape = vml_root.as_ref().and_then(|r| {
        let mut all = Vec::new();
        vml_shapes(r, &mut all);
        let doc = vml_doc.as_deref().unwrap_or_default();
        all.into_iter().find(|x| spid_of(doc, x) == Some(id))
    });
    if entry.is_none() && shape.is_none() {
        return Err(format!(
            "シート `{sheet}` にコントロール `{spid}` がありません"
        ));
    }
    if let (Some(sh), Some(doc)) = (shape, vml_doc.as_deref()) {
        let t = object_type(doc, sh);
        if t == "Note" || (t == "Pict" && !activex) {
            return Err(format!("`{spid}` はフォーム コントロールではありません"));
        }
    }

    // Where it is now: the sheet's anchor (EMU), else VML's (pixels).
    let now = entry
        .and_then(|e| first(e.el, "anchor"))
        .and_then(|a| {
            let f = marker_point(&sdoc, &sizes, a.child("from")?)?;
            let t = marker_point(&sdoc, &sizes, a.child("to")?)?;
            Some((f.0, f.1, t.0, t.1))
        })
        .or_else(|| {
            let doc = vml_doc.as_deref()?;
            let text = shape?.child("ClientData")?.child("Anchor")?.text(doc);
            vml_box(&sizes, &text)
        });

    let mut sheet_edits = Edits::default();
    let mut vml_edits = Edits::default();
    // The parts that go with it: its properties (an ActiveX control's part, with its data), the
    // picture of an ActiveX control; and what its VML shape used.
    let mut parts: Vec<(String, String)> = entry_rel
        .iter()
        .map(|r| (r.0.clone(), r.2.clone()))
        .collect();
    let mut vml_removed = String::new();
    if delete {
        if let Some(sh) = shape {
            vml_edits.replace(sh.span(), "");
            vml_removed = vml_doc.as_deref().unwrap_or_default()[sh.span()].to_string();
        }
        if let Some(rid) = entry
            .and_then(|e| first(e.el, "controlPr"))
            .and_then(|p| prefixed_attr(&sdoc, p, "id"))
            && let Some(r) = rels.iter().find(|r| r.0 == rid)
        {
            parts.push((r.0.clone(), r.2.clone()));
        }
        if let (Some(e), Some((_, outer, list))) = (entry, &controls) {
            // The last control: `<controls>` goes too.
            sheet_edits.replace(
                if list.len() == 1 {
                    outer.clone()
                } else {
                    e.span.clone()
                },
                "",
            );
        }
    } else if let Some(place) = place {
        let now = now.ok_or_else(|| format!("`{spid}` の位置を読み取れません"))?;
        let b = box_of(&sizes, place, (now.2 - now.0, now.3 - now.1));
        let b = (
            b.0.max(0.0),
            b.1.max(0.0),
            b.2.max(b.0.max(0.0)),
            b.3.max(b.1.max(0.0)),
        );
        if let (Some(sh), Some(doc)) = (shape, vml_doc.as_deref()) {
            if let Some(a) = sh.child("ClientData").and_then(|d| d.child("Anchor")) {
                vml_edits.replace(a.inner.clone(), vml_anchor(&sizes, b));
            }
            let style = sh.attr(doc, "style").unwrap_or_default();
            scan::set_attr(
                doc,
                sh.open.clone(),
                "style",
                &vml_style(&style, b),
                &mut vml_edits,
            );
        }
        if let Some(a) = entry.and_then(|e| first(e.el, "anchor")) {
            for (tag, x, y) in [("from", b.0, b.1), ("to", b.2, b.3)] {
                if let Some(m) = a.child(tag) {
                    let inner = m.children.first().map_or("", |c| scan::prefix(c.name));
                    sheet_edits.replace(m.span(), control_marker(m.name, inner, &sizes, x, y));
                }
            }
        }
        // An ActiveX control's own size (HIMETRIC), in its data.
        if activex && let Some(r) = &entry_rel {
            let himetric = |px: f64| (px * 2540.0 / 96.0).round() as i32;
            if let Some((part, mut stream)) = activex_stream(&mut s, &r.2)? {
                stream.set_size(himetric(b.2 - b.0), himetric(b.3 - b.1));
                s.put_bytes(&part, stream.write());
            }
        }
        // (Below: its mirror.)
        return finish(
            s,
            sizes,
            sdoc,
            sheet_edits,
            vml,
            vml_doc,
            vml_edits,
            &spid,
            Some(b),
            Vec::new(),
            String::new(),
        );
    }
    finish(
        s,
        sizes,
        sdoc,
        sheet_edits,
        vml,
        vml_doc,
        vml_edits,
        &spid,
        None,
        parts,
        vml_removed,
    )
}

/// An ActiveX control's part: its class (`CommandButton`, `MorphData`, ...) and display style,
/// how it keeps its properties (`persistStream`, `persistPropertyBag`, ...), its data's part.
pub(super) fn activex_part(
    s: &mut Sheet,
    part: &str,
) -> Result<(String, &'static str, u8, String, Option<String>)> {
    let doc = s
        .text(part)?
        .ok_or_else(|| format!("{part} がありません"))?;
    let root = scan::tree(&doc, &[]).ok_or_else(|| format!("{part} を読み取れません"))?;
    let attr = |local: &str| -> Option<String> {
        scan::attrs(&doc, root.open.clone())
            .into_iter()
            .find(|(n, _)| scan::local(n) == local && !n.starts_with("xmlns"))
            .map(|(_, span)| scan::unescape(&doc[span]))
    };
    let (class, display) = crate::activex::class_of(&attr("classid").unwrap_or_default());
    let persistence = attr("persistence").unwrap_or_default();
    let bin = match attr("id") {
        Some(rid) => s
            .rels_of(part)?
            .into_iter()
            .find(|r| r.0 == rid)
            .map(|r| r.2),
        None => None,
    };
    Ok((doc, class, display, persistence, bin))
}

/// The data of an ActiveX control kept as a stream: its part and the stream read. None when it
/// is kept otherwise or its stream is not read.
fn activex_stream(s: &mut Sheet, part: &str) -> Result<Option<(String, crate::activex::Stream)>> {
    let (_, class, _, persistence, bin) = activex_part(s, part)?;
    let Some(bin) = bin.filter(|_| is_stream(&persistence)) else {
        return Ok(None);
    };
    let Some(data) = s.bytes(&bin)? else {
        return Ok(None);
    };
    Ok(crate::activex::Stream::read(class, &data).map(|st| (bin, st)))
}

/// A cell a control is linked to (`$B$2`, `Sheet2!$B$2`) and what goes in it (as typed).
pub(super) type LinkedValue = (String, String);

/// What each kind of form control (its VML `ObjectType`) has of the settings an edit changes:
/// a text, a linked cell, a range of items, a checked state, a value in a range, a page.
fn has(object: &str, what: &str) -> bool {
    match what {
        "text" => matches!(object, "Button" | "Checkbox" | "Radio" | "Label" | "GBox"),
        "link" => matches!(
            object,
            "Checkbox" | "Radio" | "List" | "Drop" | "Spin" | "Scroll"
        ),
        "items" => matches!(object, "List" | "Drop"),
        "checked" => matches!(object, "Checkbox" | "Radio"),
        "selected" => matches!(object, "List" | "Drop"),
        "value" => matches!(object, "Spin" | "Scroll"),
        "page" => object == "Scroll",
        _ => false,
    }
}

/// What a kind of form control is called in a message.
fn kind_name(object: &str) -> &'static str {
    match object {
        "Button" => "ボタン",
        "Checkbox" => "チェック ボックス",
        "Radio" => "オプション ボタン",
        "Label" => "ラベル",
        "GBox" => "グループ ボックス",
        "List" => "リスト ボックス",
        "Drop" => "コンボ ボックス",
        "Spin" => "スピン ボタン",
        "Scroll" => "スクロール バー",
        _ => "このコントロール",
    }
}

/// The VML object type of a form control's properties' `objectType`.
fn vml_kind(prop: &str) -> &str {
    match prop {
        "CheckBox" => "Checkbox",
        other => other,
    }
}

/// The file with the text and the settings of the form control whose shape is `id` on the sheet
/// `sheet` changed as `change` says: in its VML shape (its text box, `x:ClientData`), its
/// properties (`formControlPr`) and the text of its mirror in the drawing. Also returns the
/// cell it is linked to and what goes in it, when its state or value changed (Excel writes the
/// cell then; the cell's value is what the control shows).
pub(super) fn configure(
    bytes: &[u8],
    sheet: &str,
    id: u32,
    change: &super::ObjectChange,
) -> Result<(Vec<u8>, Option<LinkedValue>)> {
    let mut s = Sheet::open(bytes, sheet)?;
    let sdoc = s.doc.clone();
    let root = tree(&sdoc, &s.part)?;
    let rels = s.rels()?;
    let spid = format!("_x0000_s{id}");

    // Its properties, its VML shape.
    let entry = sheet_controls(&root).and_then(|(_, _, list)| {
        list.into_iter()
            .find(|e| shape_id(&sdoc, e.el) == Some(id))
            .and_then(|e| prefixed_attr(&sdoc, e.el, "id"))
    });
    let rel = entry.and_then(|rid| rels.iter().find(|r| r.0 == rid).cloned());
    if let Some(r) = rel.as_ref().filter(|r| r.1.ends_with("/control")) {
        let part = r.2.clone();
        return configure_activex(s, id, &part, change);
    }
    let props_part = rel.filter(|r| r.1.ends_with("/ctrlProp")).map(|r| r.2);
    let props_doc = match &props_part {
        Some(p) => s.text(p)?,
        None => None,
    };
    let vml = vml_part(&mut s, &root, &sdoc)?;
    let vml_doc = match &vml {
        Some(p) => s.text(p)?,
        None => None,
    };
    let vroot = match vml_doc.as_deref() {
        Some(d) => {
            Some(scan::tree(d, &[]).ok_or("VML の描画を読み取れません（Excel で行ってください）")?)
        }
        None => None,
    };
    let shape = vroot.as_ref().and_then(|r| {
        let mut all = Vec::new();
        vml_shapes(r, &mut all);
        let doc = vml_doc.as_deref().unwrap_or_default();
        all.into_iter().find(|x| spid_of(doc, x) == Some(id))
    });
    let prop = |name: &str| -> Option<String> {
        let doc = props_doc.as_deref()?;
        scan::tree(doc, &[])?.attr(doc, name)
    };
    let client = |name: &str| -> Option<String> {
        let doc = vml_doc.as_deref()?;
        Some(
            shape?
                .child("ClientData")?
                .child(name)?
                .text(doc)
                .trim()
                .to_string(),
        )
    };
    let object = match (shape, vml_doc.as_deref()) {
        (Some(sh), Some(doc)) => object_type(doc, sh),
        _ => prop("objectType")
            .map(|t| vml_kind(&t).to_string())
            .unwrap_or_default(),
    };
    if object.is_empty() {
        return Err(format!(
            "シート `{sheet}` にコントロール `{spid}` がありません"
        ));
    }
    if matches!(object.as_str(), "Pict" | "Note") {
        return Err(format!("`{spid}` はフォーム コントロールではありません"));
    }

    // A form control's numbers: 0 to 30000, as Excel takes them.
    for (key, n) in [
        ("selected", change.selected),
        ("value", change.value),
        ("min", change.min),
        ("max", change.max),
        ("step", change.step),
        ("page", change.page),
    ] {
        if n.is_some_and(|n| !(0..=30000).contains(&n)) {
            return Err(format!("`{key}` は 0〜30000 の整数で指定してください"));
        }
    }

    // What the kind has.
    let kind = kind_name(&object);
    let asked = [
        ("text", change.text.is_some(), "文字"),
        ("link", change.link.is_some(), "リンクするセル"),
        ("items", change.items.is_some(), "入力範囲（items）"),
        ("checked", change.checked.is_some(), "チェックの状態"),
        ("selected", change.selected.is_some(), "選択している項目"),
        (
            "value",
            change.value.is_some()
                || change.min.is_some()
                || change.max.is_some()
                || change.step.is_some(),
            "値・最小値・最大値・変化の増分",
        ),
        ("page", change.page.is_some(), "ページ単位の変化"),
    ];
    for (what, on, label) in asked {
        if on && !has(&object, what) {
            return Err(format!("{kind}には{label}がありません"));
        }
    }

    // The value in its range (Excel's: the least, the greatest, the value between them).
    let now = |p: &str, c: &str, d: i64| -> i64 {
        prop(p)
            .or_else(|| client(c))
            .and_then(|v| v.parse().ok())
            .unwrap_or(d)
    };
    let mut numbers: Vec<(&str, &str, i64)> = Vec::new();
    if has(&object, "value")
        && (change.value.is_some()
            || change.min.is_some()
            || change.max.is_some()
            || change.step.is_some()
            || change.page.is_some())
    {
        let min = change.min.unwrap_or_else(|| now("min", "Min", 0));
        let max = change.max.unwrap_or_else(|| now("max", "Max", 100));
        if min > max {
            return Err(format!(
                "最小値（{min}）は最大値（{max}）以下にしてください"
            ));
        }
        let value = change.value.unwrap_or_else(|| now("val", "Val", min));
        if change.value.is_some() && !(min..=max).contains(&value) {
            return Err(format!(
                "現在値（{value}）は最小値（{min}）から最大値（{max}）の間にしてください"
            ));
        }
        let step = change.step.unwrap_or_else(|| now("inc", "Inc", 1));
        if step < 1 {
            return Err("変化の増分は 1 以上にしてください".into());
        }
        numbers.push(("val", "Val", value.clamp(min, max)));
        numbers.push(("min", "Min", min));
        numbers.push(("max", "Max", max));
        numbers.push(("inc", "Inc", step));
        if let Some(page) = change.page {
            if page < 1 {
                return Err("ページ単位の変化は 1 以上にしてください".into());
            }
            numbers.push(("page", "Page", page));
        }
    }

    // The settings: (the attribute of the properties, the element of VML, the value; None
    // takes it away).
    let mut set: Vec<(&str, &str, Option<String>)> = Vec::new();
    if let Some(link) = &change.link {
        set.push(("fmlaLink", "FmlaLink", link.clone()));
    }
    if let Some(items) = &change.items {
        set.push(("fmlaRange", "FmlaRange", items.clone()));
    }
    if let Some(on) = change.checked {
        set.push(("checked", "Checked", on.then(|| "1".to_string())));
    }
    // The item selected: one of the items (as many as the cells of their range).
    if let Some(sel) = change.selected {
        let items = change
            .items
            .clone()
            .unwrap_or_else(|| prop("fmlaRange").or_else(|| client("FmlaRange")))
            .filter(|r| !r.is_empty());
        let count = items.as_deref().map_or(0, |r| {
            let cells = r.rsplit_once('!').map_or(r, |(_, c)| c);
            gridcraft_core::RangeRef::parse(&cells.replace('$', "")).map_or(0, |g| g.count() as i64)
        });
        if sel > count {
            return Err(if count == 0 {
                format!("{kind}に入力範囲が無いため、項目を選べません")
            } else {
                format!("選択する項目は 0（なし）から {count} までにしてください")
            });
        }
        set.push(("sel", "Sel", Some(sel.to_string())));
    }
    for (p, c, v) in &numbers {
        set.push((p, c, Some(v.to_string())));
    }

    // VML: `x:ClientData`'s elements and the text box.
    if let (Some(sh), Some(doc), Some(part)) = (shape, vml_doc.as_deref(), &vml) {
        let mut edits = Edits::default();
        if let Some(data) = sh.child("ClientData") {
            let x = scan::prefix(data.name);
            for (_, name, value) in &set {
                let el = data.children.iter().find(|c| c.local() == *name);
                let new = value
                    .as_ref()
                    .map(|v| format!("<{x}{name}>{}</{x}{name}>", scan::escape(v)));
                match (el, new) {
                    (Some(el), Some(new)) => edits.replace(el.span(), new),
                    (Some(el), None) => edits.replace(el.span(), ""),
                    (None, Some(new)) => edits.insert(data.inner.end, new),
                    (None, None) => {}
                }
            }
        }
        if let Some(text) = &change.text {
            match sh.child("textbox") {
                Some(tb) => edits.replace(tb.inner.clone(), vml_box_text(doc, tb, text)),
                None => {
                    let at = sh
                        .child("ClientData")
                        .map_or(sh.inner.end, |d| d.open.start);
                    edits.insert(at, vml_textbox(text, "left"));
                }
            }
        }
        if !edits.is_empty() {
            s.put(part, edits.apply(doc));
        }
    }

    // The properties: the attributes of `formControlPr` (the checked state as `Checked`).
    if let (Some(part), Some(doc)) = (&props_part, props_doc.as_deref()) {
        let attrs: Vec<(&str, Option<String>)> = set
            .iter()
            .map(|(p, _, v)| {
                let v = if *p == "checked" {
                    v.as_ref().map(|_| "Checked".to_string())
                } else {
                    v.clone()
                };
                (*p, v)
            })
            .collect();
        if !attrs.is_empty() {
            s.put(part, super::drawing::with_attrs(doc, &attrs));
        }
    }

    // The linked cell: a check box's state, a spin button's or scroll bar's value.
    let link = change
        .link
        .clone()
        .unwrap_or_else(|| prop("fmlaLink").or_else(|| client("FmlaLink")))
        .filter(|l| !l.is_empty());
    let cell = match (&link, object.as_str()) {
        (Some(l), "Checkbox") => change
            .checked
            .map(|on| (l.clone(), if on { "TRUE" } else { "FALSE" }.to_string())),
        (Some(l), "List" | "Drop") => change
            .selected
            .filter(|n| *n > 0)
            .map(|n| (l.clone(), n.to_string())),
        (Some(l), "Spin" | "Scroll") => numbers
            .iter()
            .find(|n| n.0 == "val")
            .map(|n| (l.clone(), n.2.to_string())),
        _ => None,
    };

    // The text of its mirror.
    if let Some(text) = &change.text
        && let Some(d) = drawing_of(&mut s, false)?
    {
        let droot = scan::tree(&d.doc, &[]).ok_or("描画のパートを読み取れません")?;
        let mut edits = Edits::default();
        for u in units(&droot) {
            for a in &u.anchors {
                let mine = first(a, "compatExt")
                    .and_then(|c| c.attr(&d.doc, "spid"))
                    .as_deref()
                    == Some(spid.as_str());
                if let Some(sp) = super::drawing::object_of(a).filter(|o| mine && o.local() == "sp")
                {
                    super::drawing::set_text(&d.doc, sp, text, &mut edits);
                }
            }
        }
        if !edits.is_empty() {
            s.put(&d.part, edits.apply(&d.doc));
        }
    }
    Ok((s.finish(None)?, cell))
}

/// What an ActiveX control (its Microsoft Forms type) has of what an edit changes, and the
/// property of its data each is: its text (`Caption`, or the `Value` of a text, list or combo
/// box), its state (`Value`: "1" or "0"), its linked cell and items (in the sheet's
/// `controlPr`), its value in a range.
fn activex_has(forms: &str, what: &str) -> bool {
    match what {
        "text" => !matches!(forms, "SpinButton" | "ScrollBar" | ""),
        "checked" => matches!(forms, "CheckBox" | "OptionButton" | "ToggleButton"),
        "link" => !matches!(forms, "CommandButton" | "Label" | ""),
        "items" => matches!(forms, "ListBox" | "ComboBox"),
        "value" => matches!(forms, "SpinButton" | "ScrollBar"),
        "page" => forms == "ScrollBar",
        _ => false,
    }
}

/// The file with an ActiveX control's text and settings changed: its data (its stream, or its
/// properties in the part), its linked cell and items in the sheet's `controlPr`. Also returns
/// the linked cell and what goes in it.
fn configure_activex(
    mut s: Sheet,
    id: u32,
    part: &str,
    change: &super::ObjectChange,
) -> Result<(Vec<u8>, Option<LinkedValue>)> {
    let (ocx, class, display, persistence, bin) = activex_part(&mut s, part)?;
    let forms = crate::activex::forms_type(class, display);
    if forms.is_empty() {
        return Err("この ActiveX コントロールは、ここでは変更できません（Microsoft Forms 2.0 のコントロールではありません。Excel で行ってください）".into());
    }
    let asked = [
        ("text", change.text.is_some(), "文字"),
        ("checked", change.checked.is_some(), "オン／オフ"),
        (
            "link",
            change.link.is_some(),
            "リンクするセル（LinkedCell）",
        ),
        ("items", change.items.is_some(), "入力範囲（ListFillRange）"),
        (
            "value",
            change.value.is_some()
                || change.min.is_some()
                || change.max.is_some()
                || change.step.is_some(),
            "値・最小値・最大値・変化の増分",
        ),
        (
            "page",
            change.page.is_some(),
            "ページ単位の変化（LargeChange）",
        ),
    ];
    for (what, on, label) in asked {
        if on && !activex_has(forms, what) {
            return Err(format!("{forms} には{label}がありません"));
        }
    }
    if change.selected.is_some() {
        return Err(format!(
            "ActiveX の {forms} では、選択する項目を `text`（値）で指定してください"
        ));
    }

    // The properties kept in its data: (name, value).
    let mut props: Vec<(&str, String)> = Vec::new();
    if let Some(t) = &change.text {
        let name = if matches!(forms, "TextBox" | "ListBox" | "ComboBox") {
            "Value"
        } else {
            "Caption"
        };
        props.push((name, t.clone()));
    }
    if let Some(on) = change.checked {
        props.push(("Value", if on { "1" } else { "0" }.to_string()));
    }

    // Its data as it is: a stream, or properties in the part.
    let mut stream = match (persistence.as_str(), &bin) {
        (p, Some(b)) if is_stream(p) => {
            let data = s.bytes(b)?.ok_or_else(|| format!("{b} がありません"))?;
            Some(crate::activex::Stream::read(class, &data).ok_or(
                "この ActiveX コントロールの保存データを読み取れません（Excel で行ってください）",
            )?)
        }
        _ => None,
    };
    let bag = |name: &str| -> Option<String> {
        let root = scan::tree(&ocx, &[])?;
        root.children.iter().find_map(|c| {
            let attrs = scan::attrs(&ocx, c.open.clone());
            let get = |k: &str| {
                attrs
                    .iter()
                    .find(|(n, _)| scan::local(n) == k)
                    .map(|(_, sp)| scan::unescape(&ocx[sp.clone()]))
            };
            (c.local() == "ocxPr" && get("name").as_deref() == Some(name)).then(|| get("value"))?
        })
    };
    let numbers_asked = change.value.is_some()
        || change.min.is_some()
        || change.max.is_some()
        || change.step.is_some()
        || change.page.is_some();
    if numbers_asked {
        let now = |name: &str, d: i64| -> i64 {
            match &stream {
                Some(st) => st.number(name).unwrap_or(d),
                None => bag(name).and_then(|v| v.parse().ok()).unwrap_or(d),
            }
        };
        let min = change.min.unwrap_or_else(|| now("Min", 0));
        let max = change
            .max
            .unwrap_or_else(|| now("Max", if forms == "ScrollBar" { 32767 } else { 100 }));
        if min > max {
            return Err(format!(
                "最小値（{min}）は最大値（{max}）以下にしてください"
            ));
        }
        let value = change.value.unwrap_or_else(|| now("Position", min));
        if change.value.is_some() && !(min..=max).contains(&value) {
            return Err(format!(
                "現在値（{value}）は最小値（{min}）から最大値（{max}）の間にしてください"
            ));
        }
        let step = change.step.unwrap_or_else(|| now("SmallChange", 1));
        if step < 1 || change.page.is_some_and(|p| p < 1) {
            return Err("変化の増分は 1 以上にしてください".into());
        }
        props.push(("Min", min.to_string()));
        props.push(("Max", max.to_string()));
        props.push(("Position", value.clamp(min, max).to_string()));
        props.push(("SmallChange", step.to_string()));
        if let Some(p) = change.page {
            props.push(("LargeChange", p.to_string()));
        }
    }

    if !props.is_empty() {
        match (&mut stream, persistence.as_str()) {
            (Some(st), _) => {
                for (name, v) in &props {
                    if matches!(*name, "Caption" | "Value") {
                        st.set_string(name, v);
                    } else {
                        st.set_number(name, v.parse().unwrap_or(0));
                    }
                }
                if let Some(b) = &bin {
                    s.put_bytes(b, st.write());
                }
            }
            (None, "persistPropertyBag") => {
                let root =
                    scan::tree(&ocx, &[]).ok_or_else(|| format!("{part} を読み取れません"))?;
                let p = scan::prefix(root.name).to_string();
                let mut edits = Edits::default();
                for (name, v) in &props {
                    let el = root.children.iter().find(|c| {
                        c.local() == "ocxPr"
                            && scan::attrs(&ocx, c.open.clone()).iter().any(|(n, sp)| {
                                scan::local(n) == "name"
                                    && scan::unescape(&ocx[sp.clone()]) == *name
                            })
                    });
                    let item = format!(
                        "<{p}ocxPr {p}name=\"{name}\" {p}value=\"{}\"/>",
                        scan::escape(v)
                    );
                    match el {
                        Some(el) => edits.replace(el.span(), item),
                        None if root.open.end == root.end => {
                            let head = ocx[root.open.clone()].trim_end_matches("/>");
                            edits.replace(
                                root.open.clone(),
                                format!("{head}>{item}</{}>", root.name),
                            );
                        }
                        None => edits.insert(root.inner.end, item),
                    }
                }
                s.put(part, edits.apply(&ocx));
            }
            _ => {
                return Err(format!(
                    "この ActiveX コントロールの保存データ（{persistence}）は、ここでは変更できません（Excel で行ってください）"
                ));
            }
        }
    }

    // The linked cell and the items: the sheet's `controlPr`.
    let sdoc = s.doc.clone();
    let root = tree(&sdoc, &s.part)?;
    let mut sheet_edits = Edits::default();
    let pr = sheet_controls(&root).and_then(|(_, _, list)| {
        list.into_iter()
            .find(|e| shape_id(&sdoc, e.el) == Some(id))
            .and_then(|e| first(e.el, "controlPr"))
    });
    for (attr, value) in [
        ("linkedCell", &change.link),
        ("listFillRange", &change.items),
    ] {
        let Some(value) = value else { continue };
        let Some(pr) = pr else {
            return Err("この ActiveX コントロールの設定（controlPr）がありません（Excel で行ってください）".into());
        };
        match value {
            Some(v) => scan::set_attr(&sdoc, pr.open.clone(), attr, v, &mut sheet_edits),
            None => {
                // ` name="value"`: from the space before it to the quote after it.
                if let Some(v) = scan::attr_span(&sdoc, pr.open.clone(), attr) {
                    let head = &sdoc[pr.open.start..v.start - 1];
                    if let Some(eq) = head.rfind('=') {
                        let start = pr.open.start + head[..eq].trim_end().len() - attr.len();
                        let start = sdoc[..start].trim_end().len();
                        sheet_edits.replace(start..v.end + 1, "");
                    }
                }
            }
        }
    }
    // The linked cell gets the state, the text or the value.
    let link = change
        .link
        .clone()
        .unwrap_or_else(|| pr.and_then(|p| p.attr(&sdoc, "linkedCell")))
        .filter(|l| !l.is_empty());
    let cell = link.and_then(|l| {
        if let Some(on) = change.checked {
            return Some((l, if on { "TRUE" } else { "FALSE" }.to_string()));
        }
        if numbers_asked {
            return props
                .iter()
                .find(|p| p.0 == "Position")
                .map(|p| (l, p.1.clone()));
        }
        let text = change
            .text
            .as_ref()
            .filter(|_| matches!(forms, "TextBox" | "ListBox" | "ComboBox"))?;
        // (A text typed as it is: not read as a number or a formula.)
        Some((l, format!("'{text}")))
    });
    let doc = (!sheet_edits.is_empty()).then(|| sheet_edits.apply(&sdoc));
    Ok((s.finish(doc)?, cell))
}

/// A linked cell (`$B$2`, `Sheet2!$B$2`, `'売上 4月'!$B$2`): its sheet (`sheet` when it names
/// none) and the cell.
pub(super) fn linked(sheet: &str, link: &str) -> Option<(String, gridcraft_core::RangeRef)> {
    let (name, cell) = match link.rsplit_once('!') {
        Some((n, c)) => {
            let n = n.trim();
            let n = n
                .strip_prefix('\'')
                .and_then(|n| n.strip_suffix('\''))
                .map_or_else(|| n.to_string(), |n| n.replace("''", "'"));
            (n, c)
        }
        None => (sheet.to_string(), link),
    };
    let range = gridcraft_core::RangeRef::parse(&cell.replace('$', ""))?;
    Some((name, gridcraft_core::RangeRef::cell(range.start)))
}

/// The inside of a VML text box with a new text: a paragraph (`div`) a line, each with the
/// first paragraph's style and font.
fn vml_box_text(doc: &str, tb: &El, text: &str) -> String {
    let lines: Vec<String> = text
        .replace("\r\n", "\n")
        .split('\n')
        .map(scan::escape)
        .collect();
    let open = |el: &El| {
        let tag = &doc[el.open.clone()];
        match tag.strip_suffix("/>") {
            Some(head) => format!("{}>", head.trim_end()),
            None => tag.to_string(),
        }
    };
    let (div, font) = match tb.children.iter().find(|c| c.local() == "div") {
        Some(div) => (
            (open(div), format!("</{}>", div.name)),
            first(div, "font").map(|f| (open(f), format!("</{}>", f.name))),
        ),
        None => (
            (
                "<div style=\"text-align:left\">".to_string(),
                "</div>".to_string(),
            ),
            Some((
                format!("<font face=\"{FONT}\" size=\"220\" color=\"#000000\">"),
                "</font>".to_string(),
            )),
        ),
    };
    lines
        .iter()
        .map(|line| match &font {
            Some((fo, fc)) => format!("{}{fo}{line}{fc}{}", div.0, div.1),
            None => format!("{}{line}{}", div.0, div.1),
        })
        .collect()
}

/// The parts of a control's change written: the sheet, its VML, its mirror in the drawing
/// moved to `moved` or (None) taken away, and (`props`: the relationship and the part of its
/// properties) its properties taken away.
#[allow(clippy::too_many_arguments)]
fn finish(
    mut s: Sheet,
    sizes: Sizes,
    sdoc: String,
    sheet_edits: Edits,
    vml: Option<String>,
    vml_doc: Option<String>,
    vml_edits: Edits,
    spid: &str,
    moved: Option<(f64, f64, f64, f64)>,
    parts: Vec<(String, String)>,
    vml_removed: String,
) -> Result<Vec<u8>> {
    if let (Some(part), Some(doc)) = (vml, vml_doc)
        && !vml_edits.is_empty()
    {
        let new = vml_edits.apply(&doc);
        // (The picture of a shape taken away, when nothing else shows it.)
        let new = super::drawing::drop_unused(&mut s, &part, &vml_removed, new)?;
        s.put(&part, new);
    }
    // The mirror in the drawing.
    if let Some(d) = drawing_of(&mut s, false)? {
        let droot = scan::tree(&d.doc, &[]).ok_or("描画のパートを読み取れません")?;
        let mut edits = Edits::default();
        for u in units(&droot) {
            let mine = u.anchors.iter().any(|a| {
                first(a, "compatExt")
                    .and_then(|c| c.attr(&d.doc, "spid"))
                    .as_deref()
                    == Some(spid)
            });
            if !mine {
                continue;
            }
            match moved {
                Some(b) => {
                    for a in &u.anchors {
                        write_anchor(&d.doc, a, &sizes, b, &mut edits);
                        if let Some(obj) = super::drawing::object_of(a) {
                            write_xfrm(&d.doc, obj, b, &mut edits);
                        }
                    }
                }
                None => edits.replace(u.span.clone(), ""),
            }
        }
        if !edits.is_empty() {
            s.put(&d.part, edits.apply(&d.doc));
        }
    }
    if !parts.is_empty() {
        let ids = parts.iter().map(|p| p.0.clone()).collect();
        s.remove_rels_of(&s.part.clone(), &ids)?;
        for (_, part) in parts {
            drop_part(&mut s, &part)?;
        }
    }
    let doc = (!sheet_edits.is_empty()).then(|| sheet_edits.apply(&sdoc));
    s.finish(doc)
}

// ---------------------------------------------------------------------------- put on a sheet

/// The text of a control as VML writes it in its text box.
fn vml_textbox(text: &str, align: &str) -> String {
    let lines: String = text
        .replace("\r\n", "\n")
        .split('\n')
        .map(|l| {
            format!(
                "<div style=\"text-align:{align}\"><font face=\"{FONT}\" size=\"220\" color=\"#000000\">{}</font></div>",
                scan::escape(l)
            )
        })
        .collect();
    format!("<v:textbox style=\"mso-direction-alt:auto\" o:singleclick=\"f\">{lines}</v:textbox>")
}

/// The VML shape of a new control.
fn vml_shape(k: &Kind, id: u32, style: &str, anchor: &str, c: &FormControl, text: &str) -> String {
    let opt = |name: &str, v: &Option<String>| {
        v.as_ref()
            .map(|v| format!("<x:{name}>{}</x:{name}>", scan::escape(v)))
            .unwrap_or_default()
    };
    let num = |name: &str, v: Option<i64>, default: i64| {
        format!("<x:{name}>{}</x:{name}>", v.unwrap_or(default))
    };
    let link = opt("FmlaLink", &c.link);
    let macro_ = opt(
        "FmlaMacro",
        &c.macro_name.as_ref().map(|m| format!("[0]!{m}")),
    );
    let range = opt("FmlaRange", &c.range);
    let (attrs, inner, data) = match k.key {
        "button" => (
            "o:button=\"t\" fillcolor=\"buttonFace [67]\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\"",
            format!(
                "<v:fill color2=\"buttonFace [67]\" o:detectmouseclick=\"t\"/><o:lock v:ext=\"edit\" rotation=\"t\"/>{}",
                vml_textbox(text, "center")
            ),
            format!(
                "<x:PrintObject>False</x:PrintObject><x:AutoFill>False</x:AutoFill>{macro_}<x:TextHAlign>Center</x:TextHAlign><x:TextVAlign>Center</x:TextVAlign>"
            ),
        ),
        "checkBox" | "optionButton" => (
            "filled=\"f\" fillcolor=\"window [65]\" stroked=\"f\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\"",
            format!(
                "<v:path shadowok=\"t\" strokeok=\"t\" fillok=\"t\"/><o:lock v:ext=\"edit\" rotation=\"t\"/>{}",
                vml_textbox(text, "left")
            ),
            format!(
                "<x:AutoFill>False</x:AutoFill><x:AutoLine>False</x:AutoLine>{macro_}<x:TextVAlign>Center</x:TextVAlign>{link}{}<x:NoThreeD/>",
                if c.checked {
                    "<x:Checked>1</x:Checked>"
                } else {
                    ""
                }
            ),
        ),
        "label" => (
            "filled=\"f\" fillcolor=\"window [65]\" stroked=\"f\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\"",
            format!(
                "<o:lock v:ext=\"edit\" rotation=\"t\"/>{}",
                vml_textbox(text, "left")
            ),
            format!("<x:AutoFill>False</x:AutoFill><x:AutoLine>False</x:AutoLine>{macro_}"),
        ),
        "groupBox" => (
            "filled=\"f\" fillcolor=\"window [65]\" stroked=\"f\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\"",
            format!(
                "<o:lock v:ext=\"edit\" rotation=\"t\"/>{}",
                vml_textbox(text, "left")
            ),
            format!(
                "<x:AutoFill>False</x:AutoFill><x:AutoLine>False</x:AutoLine>{macro_}<x:NoThreeD/>"
            ),
        ),
        "listBox" => (
            "stroked=\"f\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\"",
            "<o:lock v:ext=\"edit\" rotation=\"t\"/>".to_string(),
            format!(
                "<x:AutoLine>False</x:AutoLine>{macro_}{range}{link}<x:Sel>0</x:Sel><x:SelType>Single</x:SelType><x:NoThreeD2/>"
            ),
        ),
        "comboBox" => (
            "stroked=\"f\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\"",
            "<o:lock v:ext=\"edit\" rotation=\"t\"/>".to_string(),
            format!(
                "<x:AutoLine>False</x:AutoLine>{macro_}{range}{link}<x:Val>0</x:Val><x:Sel>0</x:Sel><x:DropLines>8</x:DropLines><x:DropStyle>Combo</x:DropStyle><x:NoThreeD2/>"
            ),
        ),
        _ => (
            // Spin buttons and scroll bars.
            "stroked=\"f\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\"",
            "<o:lock v:ext=\"edit\" rotation=\"t\"/>".to_string(),
            format!(
                "<x:AutoLine>False</x:AutoLine>{macro_}{}{}{}{}{}{link}{}",
                num("Val", c.value, c.min.unwrap_or(0)),
                num("Min", c.min, 0),
                num("Max", c.max, default_max(k)),
                num("Inc", c.step, 1),
                num("Page", c.page, 10),
                if k.key == "scrollBar" && c.horizontal {
                    "<x:Horiz/>"
                } else {
                    ""
                }
            ),
        ),
    };
    format!(
        "<v:shape id=\"_x0000_s{id}\" type=\"#_x0000_t201\" style=\"{style}\" {attrs}>{inner}<x:ClientData ObjectType=\"{}\"><x:Anchor>{anchor}</x:Anchor>{data}</x:ClientData></v:shape>",
        k.vml
    )
}

/// The properties of a new control (`formControlPr`).
fn props_xml(k: &Kind, c: &FormControl) -> String {
    let attr = |name: &str, v: &Option<String>| {
        v.as_ref()
            .map(|v| format!(" {name}=\"{}\"", scan::escape(v)))
            .unwrap_or_default()
    };
    let (link, range) = (attr("fmlaLink", &c.link), attr("fmlaRange", &c.range));
    let checked = if c.checked {
        " checked=\"Checked\""
    } else {
        ""
    };
    let rest = match k.key {
        "button" | "label" => " lockText=\"1\"".to_string(),
        "checkBox" | "optionButton" => format!("{checked}{link} lockText=\"1\" noThreeD=\"1\""),
        "groupBox" => " noThreeD=\"1\"".to_string(),
        "listBox" => format!(" dx=\"16\"{range}{link} noThreeD=\"1\" sel=\"0\""),
        "comboBox" => format!(
            " dropStyle=\"combo\" dx=\"16\"{range}{link} noThreeD=\"1\" sel=\"0\" val=\"0\""
        ),
        _ => format!(
            " dx=\"16\"{link}{} inc=\"{}\" max=\"{}\" min=\"{}\" page=\"{}\" val=\"{}\"",
            if k.key == "scrollBar" && c.horizontal {
                " horiz=\"1\""
            } else {
                ""
            },
            c.step.unwrap_or(1),
            c.max.unwrap_or(default_max(k)),
            c.min.unwrap_or(0),
            c.page.unwrap_or(10),
            c.value.unwrap_or(c.min.unwrap_or(0))
        ),
    };
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<formControlPr xmlns=\"{X14_NS}\" objectType=\"{}\"{rest}/>",
        k.prop
    )
}

/// The hidden mirror of a new control in the drawing, as Excel 2010 and later write it.
#[allow(clippy::too_many_arguments)]
fn mirror_xml(
    x: &str,
    a_decl: &str,
    id: u32,
    name: &str,
    k: &Kind,
    text: &str,
    from: &str,
    to: &str,
    b: (f64, f64, f64, f64),
) -> String {
    let body = if k.text {
        let algn = if k.key == "button" { "ctr" } else { "l" };
        format!(
            "<{x}txBody><a:bodyPr vertOverflow=\"clip\" wrap=\"square\" lIns=\"27432\" tIns=\"27432\" rIns=\"27432\" bIns=\"27432\" anchor=\"ctr\" upright=\"1\"/><a:lstStyle/><a:p><a:pPr algn=\"{algn}\" rtl=\"0\"><a:defRPr sz=\"1000\"/></a:pPr><a:r><a:rPr lang=\"ja-JP\" altLang=\"en-US\" sz=\"1100\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" baseline=\"0\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill><a:latin typeface=\"{FONT}\"/><a:ea typeface=\"{FONT}\"/></a:rPr><a:t>{}</a:t></a:r></a:p></{x}txBody>",
            scan::escape(text)
        )
    } else {
        String::new()
    };
    format!(
        "<mc:AlternateContent xmlns:mc=\"{MC_NS}\"{a_decl}><mc:Choice xmlns:a14=\"{A14_NS}\" Requires=\"a14\"><{x}twoCellAnchor editAs=\"oneCell\">{from}{to}<{x}sp macro=\"\" textlink=\"\"><{x}nvSpPr><{x}cNvPr id=\"{id}\" name=\"{}\" hidden=\"1\"><a:extLst><a:ext uri=\"{{63B3BB69-23CF-44E3-9099-C40C66FF867C}}\"><a14:compatExt spid=\"_x0000_s{id}\"/></a:ext></a:extLst></{x}cNvPr><{x}cNvSpPr/></{x}nvSpPr><{x}spPr bwMode=\"auto\"><a:xfrm><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:noFill/></{x}spPr>{body}</{x}sp><{x}clientData fPrintsWithSheet=\"0\"/></{x}twoCellAnchor></mc:Choice><mc:Fallback/></mc:AlternateContent>",
        scan::escape(name),
        emu(b.0),
        emu(b.1),
        emu(b.2 - b.0),
        emu(b.3 - b.1)
    )
}

/// The file with a new form control on the sheet `sheet`: its VML shape (the VML drawing made
/// when the sheet has none), and where the sheet keeps its controls as Excel 2010 does, its
/// entry, its properties and its mirror. Returns the file and the control's name.
pub(super) fn add(bytes: &[u8], sheet: &str, c: &FormControl) -> Result<Vec<u8>> {
    let k = KINDS
        .iter()
        .find(|k| k.key == c.kind)
        .ok_or_else(|| format!("コントロール `{}` はありません", c.kind))?;
    let sizes = sizes_of(bytes, sheet)?;
    let b = box_of(&sizes, &c.place, k.size);
    let b = (b.0, b.1, b.2.max(b.0 + 1.0), b.3.max(b.1 + 1.0));
    let mut s = Sheet::open(bytes, sheet)?;

    // The VML drawing (made, with a block of shape numbers of its own, when there is none).
    let sdoc = s.doc.clone();
    let root = tree(&sdoc, &s.part)?;
    let full = sheet_controls(&root).is_some();
    let (vml, vml_doc) = match vml_part(&mut s, &root, &sdoc)? {
        Some(part) => {
            let doc = s
                .text(&part)?
                .ok_or_else(|| format!("{part} がありません"))?;
            (part, doc)
        }
        None => {
            let mut block = 0;
            let names: Vec<String> = s
                .pack
                .names()
                .iter()
                .filter(|n| n.to_ascii_lowercase().ends_with(".vml"))
                .cloned()
                .collect();
            for n in names {
                if let Some(doc) = s.pack.text(&n)? {
                    block = idmap(&doc).into_iter().fold(block, u32::max);
                }
            }
            let part = s.free("xl/drawings/vmlDrawing", ".vml");
            let doc = format!(
                "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"\n xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\n <o:shapelayout v:ext=\"edit\">\n  <o:idmap v:ext=\"edit\" data=\"{}\"/>\n </o:shapelayout>{CONTROL_SHAPETYPE}</xml>",
                block + 1
            );
            let rid = s.add_rel("vmlDrawing", &relative(&s.part.clone(), &part), false)?;
            s.content_type(None, Some("vml"), VML_TYPE)?;
            let p = scan::prefix(root.name);
            let mut edits = Edits::default();
            edits.insert(
                place(&root, "legacyDrawing"),
                format!("<{p}legacyDrawing {}/>", rid_attr(&sdoc, &root, &rid)),
            );
            s.doc = edits.apply(&sdoc);
            (part, doc)
        }
    };
    let vroot =
        scan::tree(&vml_doc, &[]).ok_or("VML の描画を読み取れません（Excel で行ってください）")?;
    let mut shapes = Vec::new();
    vml_shapes(&vroot, &mut shapes);

    // Its number: after the others in the drawing's last block (the sheet's controls and the
    // objects of its drawing do not have it either).
    let block = idmap(&vml_doc).into_iter().max().unwrap_or(1);
    let (low, high) = (block * 1024, block * 1024 + 1023);
    let mut used: Vec<u32> = shapes.iter().filter_map(|x| spid_of(&vml_doc, x)).collect();
    {
        let root = tree(&s.doc, &s.part)?;
        if let Some((_, _, list)) = sheet_controls(&root) {
            used.extend(list.iter().filter_map(|e| shape_id(&s.doc, e.el)));
        }
    }
    if let Some(d) = drawing_of(&mut s, false)? {
        used.extend(
            scan::tags(&d.doc, 0)
                .filter(|t| !t.is_close() && t.local() == "cNvPr")
                .filter_map(|t| {
                    scan::attr(&d.doc, t.start..t.end, "id")?
                        .parse::<u32>()
                        .ok()
                }),
        );
    }
    let id = used
        .iter()
        .filter(|n| (low..=high).contains(*n))
        .max()
        .map_or(low + 1, |n| n + 1);
    if id > high {
        return Err("このシートの VML の描画には、これ以上コントロールを追加できません".into());
    }
    let name = format!("{} {}", k.name, id - low);
    let text = c.text.clone().unwrap_or_else(|| name.clone());

    // The VML shape, after the others (the shape type of controls added where it is missing).
    let z = shapes.len() + 1;
    let style = vml_style(
        &format!("position:absolute;z-index:{z};mso-wrap-style:tight"),
        b,
    );
    let shape = vml_shape(k, id, &style, &vml_anchor(&sizes, b), c, &text);
    let mut edits = Edits::default();
    let shapetype = if vml_doc.contains("\"_x0000_t201\"") {
        ""
    } else {
        CONTROL_SHAPETYPE
    };
    edits.insert(vroot.inner.end, format!("{shapetype}{shape}"));
    s.put(&vml, edits.apply(&vml_doc));

    if full {
        // Its properties.
        let part = s.free("xl/ctrlProps/ctrlProp", ".xml");
        s.put(&part, props_xml(k, c));
        s.content_type(Some(&part), None, CTRL_PROP_TYPE)?;
        let rid = s.add_rel("ctrlProp", &relative(&s.part.clone(), &part), false)?;

        // Its mirror in the drawing (made when the sheet has none).
        let d = drawing_of(&mut s, true)?.ok_or("描画のパートを作れません")?;
        if let Some(doc) = d.sheet_doc {
            s.doc = doc;
        }
        let droot = scan::tree(&d.doc, &[]).ok_or("描画のパートを読み取れません")?;
        let x = ns_prefix(&d.doc, &droot, XDR_NS)
            .unwrap_or_else(|| scan::prefix(droot.name).to_string());
        let a_decl = if ns_prefix(&d.doc, &droot, A_NS).as_deref() == Some("a:") {
            String::new()
        } else {
            format!(" xmlns:a=\"{A_NS}\"")
        };
        let marker =
            |tag: &str, px: f64, py: f64| super::drawing::marker_xml(&x, tag, &sizes, px, py);
        let mirror = mirror_xml(
            &x,
            &a_decl,
            id,
            &name,
            k,
            &text,
            &marker("from", b.0, b.1),
            &marker("to", b.2, b.3),
            b,
        );
        let mut dedits = Edits::default();
        if droot.open.end == droot.end {
            let head = d.doc[droot.open.clone()].trim_end_matches("/>");
            dedits.replace(
                droot.open.clone(),
                format!("{head}>{mirror}</{}>", droot.name),
            );
        } else {
            dedits.insert(droot.inner.end, mirror);
        }
        s.put(&d.part, dedits.apply(&d.doc));

        // Its entry in the sheet's `<controls>`.
        let sdoc = s.doc.clone();
        let root = tree(&sdoc, &s.part)?;
        let (controls, _, _) =
            sheet_controls(&root).ok_or("シートのコントロールを読み取れません")?;
        let p = scan::prefix(controls.name);
        let xdr = match ns_prefix(&sdoc, &root, XDR_NS) {
            Some(x) => (x, String::new()),
            None => ("xdr:".to_string(), format!(" xmlns:xdr=\"{XDR_NS}\"")),
        };
        let x14 = if ns_prefix(&sdoc, &root, X14_NS).as_deref() == Some("x14:") {
            String::new()
        } else {
            format!(" xmlns:x14=\"{X14_NS}\"")
        };
        let r = rid_attr(&sdoc, &root, &rid);
        let print = if k.key == "button" {
            " print=\"0\""
        } else {
            ""
        };
        // (As Excel 2021 puts them: a button, a label and a spin button sized with the cells too.)
        let size_with = if matches!(k.key, "button" | "label" | "spinner") {
            " sizeWithCells=\"1\""
        } else {
            ""
        };
        let macro_ = c
            .macro_name
            .as_ref()
            .map(|m| format!(" macro=\"[0]!{}\"", scan::escape(m)))
            .unwrap_or_default();
        let (fx, fy) = (b.0, b.1);
        let entry = format!(
            "<mc:AlternateContent xmlns:mc=\"{MC_NS}\"{x14}><mc:Choice Requires=\"x14\"><{p}control shapeId=\"{id}\" {r} name=\"{}\"><{p}controlPr defaultSize=\"0\"{print} autoFill=\"0\" autoLine=\"0\" autoPict=\"0\"{macro_}><{p}anchor moveWithCells=\"1\"{size_with}{}>{}{}</{p}anchor></{p}controlPr></{p}control></mc:Choice><mc:Fallback><{p}control shapeId=\"{id}\" {r} name=\"{}\"/></mc:Fallback></mc:AlternateContent>",
            scan::escape(&name),
            xdr.1,
            control_marker(&format!("{p}from"), &xdr.0, &sizes, fx, fy),
            control_marker(&format!("{p}to"), &xdr.0, &sizes, b.2, b.3),
            scan::escape(&name),
        );
        let mut sedits = Edits::default();
        if controls.open.end == controls.end {
            let head = sdoc[controls.open.clone()].trim_end_matches("/>");
            sedits.replace(
                controls.open.clone(),
                format!("{head}>{entry}</{}>", controls.name),
            );
        } else {
            sedits.insert(controls.inner.end, entry);
        }
        s.doc = sedits.apply(&sdoc);
    }
    let doc = s.doc.clone();
    s.finish(Some(doc))
}

/// Whether an ActiveX control keeps its data as its CLSID and a stream: `persistStream`, or
/// `persistStreamInit` as Excel saves the controls of Microsoft Forms 2.0.
fn is_stream(persistence: &str) -> bool {
    matches!(persistence, "persistStream" | "persistStreamInit")
}

/// The greatest value of a new spin button or scroll bar, as Excel puts one: 30000 and 100.
fn default_max(k: &Kind) -> i64 {
    if k.key == "spinner" { 30000 } else { 100 }
}
