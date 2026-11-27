//! Form controls (開発 > 挿入 > フォーム コントロール): put on a sheet, moved, sized and deleted.
//!
//! A form control is kept in up to four places, which must agree: its shape in the sheet's VML
//! drawing (`x:ClientData`: its kind, anchor, macro, linked cell, ...: what Excel 2007 and
//! LibreOffice read), and, as Excel 2010 and later save it, an entry of the sheet's `<controls>`
//! (its name and anchor), its properties (`xl/ctrlProps/ctrlProp<n>.xml`) and a hidden mirror in
//! the sheet's drawing (`a14:compatExt`). A control is changed in all the places it is in. A new
//! one is written as the sheet keeps the others: in VML alone on a sheet without `<controls>`
//! (as Excel 2007 and XlsxWriter write them), in all four on one with them.
//!
//! An ActiveX control (開発 > 挿入 > ActiveX コントロール) is kept in its part
//! (`xl/activeX/activeX<n>.xml`) and data (`activeX<n>.bin`), an entry of `<controls>` and a
//! VML shape (`ObjectType="Pict"`); one is changed, moved and deleted here too, and a new one
//! (of Microsoft Forms 2.0) put on a sheet as Excel 2010 and later save it.

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
    // Its format: a button's font, 配置 and 余白 (Excel's コントロールの書式設定 has them for a
    // button only); 文字列のロック, of those with a text. What an ActiveX control has besides.
    let fmt = change.control_format.clone().unwrap_or_default();
    let font_asked = fmt.font_name.is_some()
        || fmt.font_size.is_some()
        || fmt.bold.is_some()
        || fmt.italic.is_some()
        || fmt.underline.is_some()
        || fmt.strike.is_some()
        || fmt.fore_color.is_some();
    let button = object == "Button";
    for (on, ok, label) in [
        (font_asked, button, "フォント"),
        (
            change.text_align.is_some() || change.text_valign.is_some(),
            button,
            "文字の配置",
        ),
        (
            change.inset.is_some() || change.inset_auto.is_some(),
            button,
            "余白",
        ),
        (
            change.lock_text.is_some(),
            has(&object, "text"),
            "文字列のロック",
        ),
        (
            fmt.back_color.is_some()
                || fmt.opaque.is_some()
                || fmt.enabled.is_some()
                || fmt.word_wrap.is_some()
                || fmt.auto_size.is_some(),
            false,
            "背景の色・BackStyle・Enabled・WordWrap・AutoSize（ActiveX コントロールのプロパティ）",
        ),
    ] {
        if on && !ok {
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
    // 配置: VML's `Left`, `Center`, ...; the properties' `left`, `center`, ...
    let capital = |v: &str| {
        let mut c = v.chars();
        c.next()
            .map(|f| f.to_ascii_uppercase().to_string() + c.as_str())
            .unwrap_or_default()
    };
    let mut props_extra: Vec<(&str, Option<String>)> = Vec::new();
    if let Some(a) = &change.text_align {
        set.push(("", "TextHAlign", Some(capital(a))));
        props_extra.push(("textHAlign", Some(a.clone())));
    }
    if let Some(a) = &change.text_valign {
        set.push(("", "TextVAlign", Some(capital(a))));
        props_extra.push(("textVAlign", Some(a.clone())));
    }
    // 文字列のロック: the properties' `lockText` "1" when on.
    // (VML's below, with the other protections.)
    if let Some(on) = change.lock_text {
        props_extra.push(("lockText", on.then(|| "1".to_string())));
    }
    let styled = BoxStyle {
        face: fmt.font_name.clone(),
        size: fmt.font_size,
        color: fmt.fore_color.clone(),
        bold: fmt.bold,
        italic: fmt.italic,
        underline: fmt.underline,
        strike: fmt.strike,
        align: change.text_align.clone(),
    };

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
        if change.text.is_some() || styled.any() {
            match sh.child("textbox") {
                Some(tb) if styled.any() => edits.replace(
                    tb.inner.clone(),
                    vml_box_styled(doc, tb, change.text.as_deref(), &styled),
                ),
                Some(tb) => edits.replace(
                    tb.inner.clone(),
                    vml_box_text(doc, tb, change.text.as_deref().unwrap_or_default()),
                ),
                None => {
                    let at = sh
                        .child("ClientData")
                        .map_or(sh.inner.end, |d| d.open.start);
                    let text = change.text.clone().unwrap_or_default();
                    let new = vml_textbox(&text, "left");
                    // (Styled as it is asked.)
                    let new = match scan::tree(&new, &[]) {
                        Some(tb) if styled.any() => {
                            let inner = vml_box_styled(&new, &tb, Some(&text), &styled);
                            format!("{}{inner}</v:textbox>", &new[tb.open.clone()])
                        }
                        _ => new,
                    };
                    edits.insert(at, new);
                }
            }
        }
        // 余白: the text box's `inset` (points), its shape's `o:insetmode` (`auto`, `custom`).
        if change.inset.is_some() || change.inset_auto.is_some() {
            let auto = change.inset_auto.unwrap_or(false) && change.inset.is_none();
            scan::set_attr(
                doc,
                sh.open.clone(),
                "o:insetmode",
                if auto { "auto" } else { "custom" },
                &mut edits,
            );
            if let Some(tb) = sh.child("textbox") {
                if auto {
                    put_attr(doc, tb.open.clone(), "inset", None, &mut edits);
                } else {
                    let now = vml_insets(tb.attr(doc, "inset").as_deref());
                    let four = change.inset.unwrap_or([None; 4]);
                    // (VML's order: left, top, right, bottom, as `inset`'s.)
                    let v: Vec<String> = (0..4)
                        .map(|i| {
                            format!("{}pt", (four[i].unwrap_or(now[i]) * 100.0).round() / 100.0)
                        })
                        .collect();
                    scan::set_attr(doc, tb.open.clone(), "inset", &v.join(","), &mut edits);
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
        let attrs: Vec<(&str, Option<String>)> = attrs
            .into_iter()
            .filter(|a| !a.0.is_empty())
            .chain(props_extra.iter().cloned())
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
    // Its mirror's text formatted, its 配置 and 余白.
    if (styled.any()
        || change.text_valign.is_some()
        || change.inset.is_some()
        || change.inset_auto.is_some())
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
                    mirror_format(&d.doc, sp, &styled, change, &mut edits);
                }
            }
        }
        if !edits.is_empty() {
            s.put(&d.part, edits.apply(&d.doc));
        }
    }
    // 保護, プロパティ, 代替テキスト: the sheet's `controlPr`, VML, the mirror.
    let mut sheet_edits = Edits::default();
    let pr = sheet_controls(&root).and_then(|(_, _, list)| {
        list.into_iter()
            .find(|e| shape_id(&sdoc, e.el) == Some(id))
            .and_then(|e| first(e.el, "controlPr"))
    });
    sheet_settings(&sdoc, pr, change, &mut sheet_edits)?;
    if change.placement.is_some()
        || change.print.is_some()
        || change.locked.is_some()
        || change.lock_text.is_some()
    {
        placed(&mut s, &root, &sdoc, id, change)?;
    }
    let doc = (!sheet_edits.is_empty()).then(|| sheet_edits.apply(&sdoc));
    Ok((s.finish(doc)?, cell))
}

/// The sheet's settings of a control (its `controlPr`): オブジェクトを印刷する, ロック (`print`,
/// `locked`: "0" when off), 代替テキスト (`altText`), オブジェクトの位置関係 (its `anchor`'s
/// `moveWithCells`, `sizeWithCells`).
fn sheet_settings(
    sdoc: &str,
    pr: Option<&El>,
    change: &super::ObjectChange,
    edits: &mut Edits,
) -> Result<()> {
    let off = |on: Option<bool>| on.map(|on| (!on).then(|| "0".to_string()));
    for (attr, value) in [
        ("print", off(change.print)),
        ("locked", off(change.locked)),
        (
            "altText",
            change
                .alt_text
                .clone()
                .map(|t| Some(t).filter(|t| !t.is_empty())),
        ),
    ] {
        let Some(value) = value else { continue };
        let Some(pr) = pr else {
            if matches!(attr, "print" | "locked") {
                // (Its VML shape has them too.)
                continue;
            }
            return Err(
                "このコントロールの設定（controlPr）がありません（Excel で行ってください）".into(),
            );
        };
        put_attr(sdoc, pr.open.clone(), attr, value.as_deref(), edits);
    }
    if let Some(p) = &change.placement
        && let Some(a) = pr.and_then(|p| p.child("anchor"))
    {
        let one = |on: bool| on.then_some("1");
        put_attr(
            sdoc,
            a.open.clone(),
            "moveWithCells",
            one(p != "absolute"),
            edits,
        );
        put_attr(
            sdoc,
            a.open.clone(),
            "sizeWithCells",
            one(p == "twoCell"),
            edits,
        );
    }
    Ok(())
}

/// The font and the alignment a form button's text is given (None: what it has stays). `color`
/// None in it: 自動 (black).
#[derive(Default)]
struct BoxStyle {
    face: Option<String>,
    size: Option<f64>,
    color: Option<Option<String>>,
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    strike: Option<bool>,
    align: Option<String>,
}

impl BoxStyle {
    fn any(&self) -> bool {
        self.face.is_some()
            || self.size.is_some()
            || self.color.is_some()
            || self.bold.is_some()
            || self.italic.is_some()
            || self.underline.is_some()
            || self.strike.is_some()
            || self.align.is_some()
    }
}

/// The text of an element (its tags left out; `<br>` a new line).
fn plain(doc: &str, el: &El) -> String {
    let mut out = String::new();
    let mut rest = &doc[el.inner.clone()];
    while let Some(i) = rest.find('<') {
        out.push_str(&rest[..i]);
        let Some(j) = rest[i..].find('>') else {
            break;
        };
        let tag = rest[i + 1..i + j]
            .trim_start_matches('/')
            .to_ascii_lowercase();
        if tag.starts_with("br") {
            out.push('\n');
        }
        rest = &rest[i + j + 1..];
    }
    out.push_str(rest);
    scan::unescape(&out)
}

/// The insets of a VML text box (`inset`: left, top, right, bottom) in points; those it does
/// not give: VML's (0.1in, 0.05in).
fn vml_insets(inset: Option<&str>) -> [f64; 4] {
    let mut out = [7.2, 3.6, 7.2, 3.6];
    for (i, v) in inset.unwrap_or("").split(',').take(4).enumerate() {
        let v = v.trim();
        if v.is_empty() {
            continue;
        }
        let (num, unit) = v
            .find(|c: char| c.is_ascii_alphabetic())
            .map_or((v, ""), |k| (&v[..k], &v[k..]));
        let Ok(n) = num.trim().parse::<f64>() else {
            continue;
        };
        out[i] = match unit {
            "pt" => n,
            "in" => n * 72.0,
            "mm" => n * 72.0 / 25.4,
            "cm" => n * 72.0 / 2.54,
            "px" => n * 0.75,
            // (A number alone: EMU.)
            _ => n / 12700.0,
        };
    }
    out
}

/// A form button's VML text box with its text (`text`, else the one it has) in the font and the
/// alignment asked for, the others as it has them: a `div` a line.
fn vml_box_styled(doc: &str, tb: &El, text: Option<&str>, st: &BoxStyle) -> String {
    let divs: Vec<&El> = tb.children.iter().filter(|c| c.local() == "div").collect();
    let lines: Vec<String> = match text {
        Some(t) => t
            .replace("\r\n", "\n")
            .split('\n')
            .map(str::to_string)
            .collect(),
        None if divs.is_empty() => vec![plain(doc, tb).trim().to_string()],
        None => divs.iter().map(|d| plain(doc, d)).collect(),
    };
    let font = divs.first().and_then(|d| first(d, "font"));
    let attr = |k: &str| font.and_then(|f| f.attr(doc, k));
    let has = |t: &str| font.is_some_and(|f| first(f, t).is_some());
    let face = st
        .face
        .clone()
        .or_else(|| attr("face"))
        .unwrap_or_else(|| FONT.to_string());
    let size = st
        .size
        .map(|p| (p * 20.0).round() as i64)
        .or_else(|| attr("size").and_then(|v| v.parse().ok()))
        .unwrap_or(220);
    let color = match &st.color {
        Some(Some(c)) => format!("#{c}"),
        Some(None) => "#000000".to_string(),
        None => attr("color").unwrap_or_else(|| "#000000".to_string()),
    };
    let align = st.align.clone().unwrap_or_else(|| {
        divs.first()
            .and_then(|d| d.attr(doc, "style"))
            .and_then(|s| {
                s.split(';').find_map(|p| {
                    p.trim()
                        .strip_prefix("text-align:")
                        .map(|a| a.trim().to_string())
                })
            })
            .unwrap_or_else(|| "left".to_string())
    });
    let tags: Vec<&str> = [
        (st.bold.unwrap_or_else(|| has("b")), "b"),
        (st.italic.unwrap_or_else(|| has("i")), "i"),
        (st.underline.unwrap_or_else(|| has("u")), "u"),
        (st.strike.unwrap_or_else(|| has("s")), "s"),
    ]
    .into_iter()
    .filter(|t| t.0)
    .map(|t| t.1)
    .collect();
    let open: String = tags.iter().map(|t| format!("<{t}>")).collect();
    let close: String = tags.iter().rev().map(|t| format!("</{t}>")).collect();
    lines
        .iter()
        .map(|line| {
            format!(
                "<div style='text-align:{align}'><font face=\"{}\" size=\"{size}\" color=\"{color}\">{open}{}{close}</font></div>",
                scan::escape(&face),
                scan::escape(line)
            )
        })
        .collect()
}

/// An element's text with attributes set or taken away, the children named in `drop` taken
/// away, and children put in: `early` after the last of the names given it has (else at its
/// start), `late` before the first of the names given it has (else at its end).
fn reshaped(
    text: &str,
    attrs: &[(&str, Option<String>)],
    drop: &[&str],
    early: (&[&str], String),
    late: (&[&str], String),
) -> String {
    let mut text = super::drawing::with_attrs(text, attrs);
    if early.1.is_empty() && late.1.is_empty() && drop.is_empty() {
        return text;
    }
    // (A tag closed in itself opened.)
    if let Some(el) = scan::tree(&text, &[])
        && el.open.end == el.end
    {
        let head = text[..el.open.end]
            .trim_end_matches("/>")
            .trim_end()
            .to_string();
        text = format!("{head}></{}>", el.name);
    }
    let Some(el) = scan::tree(&text, &[]) else {
        return text;
    };
    let mut edits = Edits::default();
    for c in &el.children {
        if drop.contains(&c.local()) {
            edits.replace(c.span(), "");
        }
    }
    let kept: Vec<&El> = el
        .children
        .iter()
        .filter(|c| !drop.contains(&c.local()))
        .collect();
    if !early.1.is_empty() {
        let at = kept
            .iter()
            .rev()
            .find(|c| early.0.contains(&c.local()))
            .map_or(el.inner.start, |c| c.end);
        edits.insert(at, early.1);
    }
    if !late.1.is_empty() {
        let at = kept
            .iter()
            .find(|c| late.0.contains(&c.local()))
            .map_or(el.inner.end, |c| c.span().start);
        edits.insert(at, late.1);
    }
    edits.apply(&text)
}

/// A form button's mirror (`sp`) formatted as its VML text box: its runs' font (`rPr`), its
/// paragraphs' alignment (`algn`), its body's vertical anchor and insets.
fn mirror_format(
    doc: &str,
    sp: &El,
    st: &BoxStyle,
    change: &super::ObjectChange,
    edits: &mut Edits,
) {
    let Some(body) = sp.child("txBody") else {
        return;
    };
    let a = body
        .children
        .iter()
        .find(|c| c.local() == "p")
        .map_or("a:", |p| scan::prefix(p.name))
        .to_string();
    if let Some(bp) = body.child("bodyPr") {
        let mut attrs: Vec<(&str, Option<String>)> = Vec::new();
        if let Some(v) = &change.text_valign {
            let anchor = match v.as_str() {
                "top" => "t",
                "bottom" => "b",
                "justify" => "just",
                "distributed" => "dist",
                _ => "ctr",
            };
            attrs.push(("anchor", Some(anchor.to_string())));
        }
        if change.inset.is_some() || change.inset_auto.is_some() {
            let auto = change.inset_auto.unwrap_or(false) && change.inset.is_none();
            let four = change.inset.unwrap_or([None; 4]);
            for (i, k) in ["lIns", "tIns", "rIns", "bIns"].into_iter().enumerate() {
                let v = if auto {
                    Some(27432)
                } else {
                    four[i].map(|pt| (pt * 12700.0).round() as i64)
                };
                if let Some(v) = v {
                    attrs.push((k, Some(v.to_string())));
                }
            }
        }
        if !attrs.is_empty() {
            edits.replace(
                bp.span(),
                super::drawing::with_attrs(&doc[bp.span()], &attrs),
            );
        }
    }
    let algn = st.align.as_deref().map(|v| {
        match v {
            "center" => "ctr",
            "right" => "r",
            "justify" => "just",
            "distributed" => "dist",
            _ => "l",
        }
        .to_string()
    });
    let flag = |on: Option<bool>| on.map(|b| if b { "1" } else { "0" }.to_string());
    let rattrs: Vec<(&str, Option<String>)> = [
        (
            "sz",
            st.size.map(|p| ((p * 100.0).round() as i64).to_string()),
        ),
        ("b", flag(st.bold)),
        ("i", flag(st.italic)),
        (
            "u",
            st.underline
                .map(|u| if u { "sng" } else { "none" }.to_string()),
        ),
        (
            "strike",
            st.strike
                .map(|u| if u { "sngStrike" } else { "noStrike" }.to_string()),
        ),
    ]
    .into_iter()
    .filter(|v| v.1.is_some())
    .collect();
    let fill = st.color.as_ref().map(|c| {
        format!(
            "<{a}solidFill><{a}srgbClr val=\"{}\"/></{a}solidFill>",
            c.clone().unwrap_or_else(|| "000000".to_string())
        )
    });
    let faces = st.face.as_ref().map(|f| {
        let f = scan::escape(f);
        format!("<{a}latin typeface=\"{f}\"/><{a}ea typeface=\"{f}\"/>")
    });
    const FILLS: &[&str] = &[
        "noFill",
        "solidFill",
        "gradFill",
        "blipFill",
        "pattFill",
        "grpFill",
    ];
    let mut drop: Vec<&str> = Vec::new();
    if fill.is_some() {
        drop.extend(FILLS);
    }
    if faces.is_some() {
        drop.extend(["latin", "ea", "cs"]);
    }
    let fonted = !rattrs.is_empty() || fill.is_some() || faces.is_some();
    let run_props = |text: &str| {
        reshaped(
            text,
            &rattrs,
            &drop,
            (&["ln"], fill.clone().unwrap_or_default()),
            (
                &["sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"],
                faces.clone().unwrap_or_default(),
            ),
        )
    };
    for p in body.children.iter().filter(|c| c.local() == "p") {
        if let Some(al) = &algn {
            match p.child("pPr") {
                Some(ppr) => edits.replace(
                    ppr.span(),
                    super::drawing::with_attrs(&doc[ppr.span()], &[("algn", Some(al.clone()))]),
                ),
                None => edits.insert(p.inner.start, format!("<{a}pPr algn=\"{al}\"/>")),
            }
        }
        if !fonted {
            continue;
        }
        for r in p.children.iter().filter(|c| c.local() == "r") {
            match r.child("rPr") {
                Some(rpr) => edits.replace(rpr.span(), run_props(&doc[rpr.span()])),
                None => edits.insert(
                    r.inner.start,
                    run_props(&format!("<{a}rPr lang=\"ja-JP\" altLang=\"en-US\"/>")),
                ),
            }
        }
        if let Some(end) = p.child("endParaRPr") {
            edits.replace(end.span(), run_props(&doc[end.span()]));
        }
    }
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
    if change.lock_text.is_some()
        || change.text_align.is_some()
        || change.text_valign.is_some()
        || change.inset.is_some()
        || change.inset_auto.is_some()
    {
        return Err(format!(
            "ActiveX の {forms} の文字列のロック・文字の配置・余白は、ここでは変えられません（フォーム コントロールのボタンにあります）"
        ));
    }
    let fmt = change.control_format.clone().unwrap_or_default();
    let font_asked = fmt.font_name.is_some()
        || fmt.font_size.is_some()
        || fmt.bold.is_some()
        || fmt.italic.is_some()
        || fmt.underline.is_some()
        || fmt.strike.is_some();
    let bar = matches!(forms, "SpinButton" | "ScrollBar");
    for (on, without, label) in [
        (font_asked, bar, "フォント（Font）"),
        (fmt.opaque.is_some(), bar, "背景のスタイル（BackStyle）"),
        (
            fmt.word_wrap.is_some(),
            bar || matches!(forms, "ListBox" | "ComboBox"),
            "折り返し（WordWrap）",
        ),
        (
            fmt.auto_size.is_some(),
            bar || forms == "ListBox",
            "自動サイズ（AutoSize）",
        ),
    ] {
        if on && without {
            return Err(format!("{forms} には{label}がありません"));
        }
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

    if change.control_format.is_some() {
        let Some(st) = &mut stream else {
            return Err(format!(
                "この ActiveX コントロールの保存データ（{persistence}）のフォント・色・プロパティは、ここでは変更できません（Excel で行ってください）"
            ));
        };
        use crate::activex::{
            AUTO_SIZE, BACK_STYLE, BOLD, ENABLED, ITALIC, STRIKE, UNDERLINE, WORD_WRAP,
            default_colors, ole_of,
        };
        let (fore, back) = default_colors(forms);
        if let Some(c) = &fmt.fore_color {
            st.set_raw("ForeColor", c.as_deref().and_then(ole_of).unwrap_or(fore));
        }
        if let Some(c) = &fmt.back_color {
            st.set_raw("BackColor", c.as_deref().and_then(ole_of).unwrap_or(back));
        }
        for (on, bit) in [
            (fmt.opaque, BACK_STYLE),
            (fmt.enabled, ENABLED),
            (fmt.word_wrap, WORD_WRAP),
            (fmt.auto_size, AUTO_SIZE),
        ] {
            if let Some(on) = on {
                st.set_bit(bit, on);
            }
        }
        if font_asked
            && !st.change_font(|f| {
                if let Some(n) = &fmt.font_name {
                    f.set_string("FontName", n);
                }
                if let Some(pt) = fmt.font_size {
                    f.set_number("FontHeight", (pt * 20.0).round() as i64);
                }
                let effects = [
                    (fmt.bold, BOLD),
                    (fmt.italic, ITALIC),
                    (fmt.underline, UNDERLINE),
                    (fmt.strike, STRIKE),
                ];
                if effects.iter().any(|e| e.0.is_some()) {
                    let mut e = f.number("FontEffects").unwrap_or(0);
                    for (on, bit) in effects {
                        match on {
                            Some(true) => e |= bit,
                            Some(false) => e &= !bit,
                            None => {}
                        }
                    }
                    f.set_number("FontEffects", e);
                }
                // (The weight, when it has one, says whether it is bold.)
                if let Some(b) = fmt.bold
                    && f.number("FontWeight").is_some()
                {
                    f.set_number("FontWeight", if b { 700 } else { 400 });
                }
            })
        {
            return Err(
                "この ActiveX コントロールのフォントを読み取れません（Excel で行ってください）"
                    .into(),
            );
        }
    }

    if !props.is_empty() || change.control_format.is_some() {
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
        ("linkedCell", change.link.clone()),
        ("listFillRange", change.items.clone()),
    ] {
        let Some(value) = value else { continue };
        let Some(pr) = pr else {
            return Err("この ActiveX コントロールの設定（controlPr）がありません（Excel で行ってください）".into());
        };
        put_attr(
            &sdoc,
            pr.open.clone(),
            attr,
            value.as_deref(),
            &mut sheet_edits,
        );
    }
    // 保護, プロパティ, 代替テキスト.
    sheet_settings(&sdoc, pr, change, &mut sheet_edits)?;
    if change.placement.is_some() || change.print.is_some() || change.locked.is_some() {
        placed(&mut s, &root, &sdoc, id, change)?;
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

/// An attribute of the start tag at `open` set, or (None) taken away.
fn put_attr(
    doc: &str,
    open: std::ops::Range<usize>,
    attr: &str,
    value: Option<&str>,
    edits: &mut Edits,
) {
    match value {
        Some(v) => scan::set_attr(doc, open, attr, v, edits),
        None => {
            // ` name="value"`: from the space before it to the quote after it.
            if let Some(v) = scan::attr_span(doc, open.clone(), attr) {
                let head = &doc[open.start..v.start - 1];
                if let Some(eq) = head.rfind('=') {
                    let start = open.start + head[..eq].trim_end().len() - attr.len();
                    let start = doc[..start].trim_end().len();
                    edits.replace(start..v.end + 1, "");
                }
            }
        }
    }
}

/// An ActiveX control's オブジェクトの位置関係, オブジェクトを印刷する and ロック where its VML
/// shape keeps them (`x:ClientData`: `x:MoveWithCells` and `x:SizeWithCells` there when it does
/// not move or size with the cells, `x:PrintObject` and `x:Locked` False when off) and its
/// mirror in the drawing (its anchor's `editAs`).
fn placed(
    s: &mut Sheet,
    root: &El,
    sdoc: &str,
    id: u32,
    change: &super::ObjectChange,
) -> Result<()> {
    if let Some(part) = vml_part(s, root, sdoc)?
        && let Some(doc) = s.text(&part)?
    {
        let vroot =
            scan::tree(&doc, &[]).ok_or("VML の描画を読み取れません（Excel で行ってください）")?;
        let mut all = Vec::new();
        vml_shapes(&vroot, &mut all);
        if let Some(data) = all
            .into_iter()
            .find(|x| spid_of(&doc, x) == Some(id))
            .and_then(|sh| sh.child("ClientData"))
        {
            let x = scan::prefix(data.name);
            let mut edits = Edits::default();
            let mut drop = |names: &[&str]| {
                for c in &data.children {
                    if names.contains(&c.local()) {
                        edits.replace(c.span(), "");
                    }
                }
            };
            if change.placement.is_some() {
                drop(&["MoveWithCells", "SizeWithCells"]);
            }
            if change.print.is_some() {
                drop(&["PrintObject"]);
            }
            if change.locked.is_some() {
                drop(&["Locked"]);
            }
            if change.lock_text.is_some() {
                drop(&["LockText"]);
            }
            let at = data.open.end;
            if data.open.end == data.end {
                return Err(
                    "この ActiveX コントロールの VML を読み取れません（Excel で行ってください）"
                        .into(),
                );
            }
            let mut head = String::new();
            match change.placement.as_deref() {
                Some("oneCell") => head.push_str(&format!("<{x}SizeWithCells/>")),
                Some("absolute") => {
                    head.push_str(&format!("<{x}MoveWithCells/><{x}SizeWithCells/>"));
                }
                _ => {}
            }
            if change.locked == Some(false) {
                head.push_str(&format!("<{x}Locked>False</{x}Locked>"));
            }
            if change.print == Some(false) {
                head.push_str(&format!("<{x}PrintObject>False</{x}PrintObject>"));
            }
            if change.lock_text == Some(false) {
                head.push_str(&format!("<{x}LockText>False</{x}LockText>"));
            }
            if !head.is_empty() {
                edits.insert(at, head);
            }
            if !edits.is_empty() {
                s.put(&part, edits.apply(&doc));
            }
        }
    }
    // Its mirror in the drawing: its anchor's `editAs`, its `clientData`'s `fPrintsWithSheet`
    // and `fLocksWithSheet`, its `fLocksText` ("0" when off).
    if let Some(d) = drawing_of(s, false)? {
        let droot = scan::tree(&d.doc, &[]).ok_or("描画のパートを読み取れません")?;
        let spid = format!("_x0000_s{id}");
        let mut edits = Edits::default();
        let off = |on: bool| (!on).then(|| "0".to_string());
        for u in units(&droot) {
            for a in &u.anchors {
                let mine = first(a, "compatExt")
                    .and_then(|c| c.attr(&d.doc, "spid"))
                    .is_some_and(|v| v == spid);
                if !mine {
                    continue;
                }
                if let Some(p) = &change.placement
                    && a.local() == "twoCellAnchor"
                {
                    edits.replace(
                        a.open.clone(),
                        super::drawing::tag_with(
                            &d.doc,
                            a,
                            &[("editAs", (p != "twoCell").then(|| p.clone()))],
                        ),
                    );
                }
                let mut data: Vec<(&str, Option<String>)> = Vec::new();
                if let Some(on) = change.print {
                    data.push(("fPrintsWithSheet", off(on)));
                }
                if let Some(on) = change.locked {
                    data.push(("fLocksWithSheet", off(on)));
                }
                if !data.is_empty() {
                    match a.child("clientData") {
                        Some(c) => edits.replace(
                            c.span(),
                            super::drawing::with_attrs(&d.doc[c.span()], &data),
                        ),
                        None => edits.insert(
                            a.inner.end,
                            super::drawing::with_attrs(
                                &format!("<{}clientData/>", scan::prefix(a.name)),
                                &data,
                            ),
                        ),
                    }
                }
                if let Some(on) = change.lock_text
                    && let Some(sp) = super::drawing::object_of(a).filter(|o| o.local() == "sp")
                {
                    edits.replace(
                        sp.open.clone(),
                        super::drawing::tag_with(&d.doc, sp, &[("fLocksText", off(on))]),
                    );
                }
            }
        }
        if !edits.is_empty() {
            s.put(&d.part, edits.apply(&d.doc));
        }
    }
    Ok(())
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
    let full = sheet_controls(&tree(&s.doc, &s.part)?).is_some();

    // Its VML shape, after the others.
    let n = new_shape(&mut s)?;
    let id = n.id;
    let name = format!("{} {}", k.name, id - n.low);
    let text = c.text.clone().unwrap_or_else(|| name.clone());
    let style = vml_style(
        &format!("position:absolute;z-index:{};mso-wrap-style:tight", n.z),
        b,
    );
    let shape = vml_shape(k, id, &style, &vml_anchor(&sizes, b), c, &text);
    put_shape(&mut s, &n, &shape);

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
        append(&d.doc, &droot, &mirror, &mut dedits);
        s.put(&d.part, dedits.apply(&d.doc));

        // Its entry in the sheet's `<controls>`.
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
        put_entry(&mut s, &rid, |e| {
            let (p, r) = (&e.p, &e.r);
            let entry = format!(
                "<mc:AlternateContent xmlns:mc=\"{MC_NS}\"{}><mc:Choice Requires=\"x14\"><{p}control shapeId=\"{id}\" {r} name=\"{}\"><{p}controlPr defaultSize=\"0\"{print} autoFill=\"0\" autoLine=\"0\" autoPict=\"0\"{macro_}><{p}anchor moveWithCells=\"1\"{size_with}{}>{}{}</{p}anchor></{p}controlPr></{p}control></mc:Choice><mc:Fallback><{p}control shapeId=\"{id}\" {r} name=\"{}\"/></mc:Fallback></mc:AlternateContent>",
                e.x14_decl,
                scan::escape(&name),
                e.xdr_decl,
                control_marker(&format!("{p}from"), &e.xdr, &sizes, b.0, b.1),
                control_marker(&format!("{p}to"), &e.xdr, &sizes, b.2, b.3),
                scan::escape(&name),
            );
            (entry, None)
        })?;
    }
    let doc = s.doc.clone();
    s.finish(Some(doc))
}

/// A new control's place in the sheet's VML drawing: the drawing and its text, its shape number
/// and the first of its block, its `z-index`, where its shape goes.
struct NewShape {
    vml: String,
    vml_doc: String,
    id: u32,
    low: u32,
    z: usize,
    at: usize,
}

/// Where a new control goes in the sheet's VML drawing (made, with a block of shape numbers of
/// its own and the sheet's `<legacyDrawing>`, when there is none): its number after the others
/// in the drawing's last block (the sheet's controls and the objects of its drawing do not have
/// it either), after the drawing's shapes.
fn new_shape(s: &mut Sheet) -> Result<NewShape> {
    let sdoc = s.doc.clone();
    let root = tree(&sdoc, &s.part)?;
    let (vml, vml_doc) = match vml_part(s, &root, &sdoc)? {
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

    let block = idmap(&vml_doc).into_iter().max().unwrap_or(1);
    let (low, high) = (block * 1024, block * 1024 + 1023);
    let mut used: Vec<u32> = shapes.iter().filter_map(|x| spid_of(&vml_doc, x)).collect();
    {
        let root = tree(&s.doc, &s.part)?;
        if let Some((_, _, list)) = sheet_controls(&root) {
            used.extend(list.iter().filter_map(|e| shape_id(&s.doc, e.el)));
        }
    }
    if let Some(d) = drawing_of(s, false)? {
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
    let (z, at) = (shapes.len() + 1, vroot.inner.end);
    Ok(NewShape {
        vml,
        vml_doc,
        id,
        low,
        z,
        at,
    })
}

/// The VML drawing with a new control's shape (and the shape type of controls, where it is
/// missing).
fn put_shape(s: &mut Sheet, n: &NewShape, shape: &str) {
    let shapetype = if n.vml_doc.contains("\"_x0000_t201\"") {
        ""
    } else {
        CONTROL_SHAPETYPE
    };
    let mut edits = Edits::default();
    edits.insert(n.at, format!("{shapetype}{shape}"));
    s.put(&n.vml, edits.apply(&n.vml_doc));
}

/// What a new control's entry in the sheet's `<controls>` is written with: the prefix of the
/// elements of `<controls>`, the relationship's attribute, the prefix of the anchor's markers
/// and its declaration (where the sheet has none), the declaration of `x14` (for
/// `Requires="x14"`, where the sheet has none).
struct EntryNames {
    p: String,
    r: String,
    xdr: String,
    xdr_decl: String,
    x14_decl: String,
}

/// An element with a new last child (an empty element opened).
fn append(doc: &str, el: &El, child: &str, edits: &mut Edits) {
    if el.open.end == el.end {
        let head = doc[el.open.clone()].trim_end_matches("/>");
        edits.replace(el.open.clone(), format!("{head}>{child}</{}>", el.name));
    } else {
        edits.insert(el.inner.end, child.to_string());
    }
}

/// The sheet with a new control's entry in its `<controls>` (made where the sheet has none, in
/// the place the schema has it). `entry` writes the entry, and the plain `<control>` for the
/// fallback of an `mc:AlternateContent` that `<controls>` is in (as Excel 2010 and later save a
/// sheet's form controls), when it has one.
fn put_entry(
    s: &mut Sheet,
    rid: &str,
    entry: impl FnOnce(&EntryNames) -> (String, Option<String>),
) -> Result<()> {
    let sdoc = s.doc.clone();
    let root = tree(&sdoc, &s.part)?;
    let found = sheet_controls(&root);
    let p = match &found {
        Some((controls, _, _)) => scan::prefix(controls.name),
        None => scan::prefix(root.name),
    }
    .to_string();
    let (xdr, xdr_decl) = match ns_prefix(&sdoc, &root, XDR_NS) {
        Some(x) => (x, String::new()),
        None => ("xdr:".to_string(), format!(" xmlns:xdr=\"{XDR_NS}\"")),
    };
    let x14_decl = if ns_prefix(&sdoc, &root, X14_NS).as_deref() == Some("x14:") {
        String::new()
    } else {
        format!(" xmlns:x14=\"{X14_NS}\"")
    };
    let names = EntryNames {
        r: rid_attr(&sdoc, &root, rid),
        p,
        xdr,
        xdr_decl,
        x14_decl,
    };
    let (choice, fallback) = entry(&names);
    let mut edits = Edits::default();
    match found {
        Some((controls, outer, _)) => {
            append(&sdoc, controls, &choice, &mut edits);
            // `<controls>` in an `mc:AlternateContent`: its fallback's `<controls>` too.
            let fallback_controls = root
                .children
                .iter()
                .filter(|c| c.local() == "AlternateContent" && c.span() == outer)
                .flat_map(|ac| ac.children.iter().filter(|c| c.local() == "Fallback"))
                .find_map(|f| f.children.iter().find(|c| c.local() == "controls"));
            if let (Some(el), Some(plain)) = (fallback_controls, &fallback) {
                append(&sdoc, el, plain, &mut edits);
            }
        }
        None => {
            let p = &names.p;
            edits.insert(
                place(&root, "controls"),
                format!("<{p}controls>{choice}</{p}controls>"),
            );
        }
    }
    s.doc = edits.apply(&sdoc);
    Ok(())
}

// ---------------------------------------------------------------------------- ActiveX controls

/// The relationship of an ActiveX control's part to its data.
const ACTIVEX_BINARY: &str =
    "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary";
const ACTIVEX_TYPE: &str = "application/vnd.ms-office.activeX+xml";
const ACTIVEX_BIN_TYPE: &str = "application/vnd.ms-office.activeX";

/// A kind of ActiveX control (開発 > 挿入 > ActiveX コントロール, of Microsoft Forms 2.0): its
/// key in the edit, its Microsoft Forms type (the name Excel gives a new one, followed by its
/// number), its size when it is put with a click (px), and whether it shows a caption (else
/// the text of an edit is its value).
pub(super) struct ActiveXKind {
    pub key: &'static str,
    forms: &'static str,
    size: (f64, f64),
    caption: bool,
}

pub(super) const ACTIVEX_KINDS: [ActiveXKind; 10] = [
    ActiveXKind {
        key: "commandButton",
        forms: "CommandButton",
        size: (96.0, 32.0),
        caption: true,
    },
    ActiveXKind {
        key: "comboBox",
        forms: "ComboBox",
        size: (96.0, 24.0),
        caption: false,
    },
    ActiveXKind {
        key: "checkBox",
        forms: "CheckBox",
        size: (96.0, 24.0),
        caption: true,
    },
    ActiveXKind {
        key: "listBox",
        forms: "ListBox",
        size: (96.0, 72.0),
        caption: false,
    },
    ActiveXKind {
        key: "textBox",
        forms: "TextBox",
        size: (96.0, 24.0),
        caption: false,
    },
    ActiveXKind {
        key: "scrollBar",
        forms: "ScrollBar",
        size: (17.0, 96.0),
        caption: false,
    },
    ActiveXKind {
        key: "spinButton",
        forms: "SpinButton",
        size: (17.0, 32.0),
        caption: false,
    },
    ActiveXKind {
        key: "optionButton",
        forms: "OptionButton",
        size: (96.0, 24.0),
        caption: true,
    },
    ActiveXKind {
        key: "label",
        forms: "Label",
        size: (96.0, 24.0),
        caption: true,
    },
    ActiveXKind {
        key: "toggleButton",
        forms: "ToggleButton",
        size: (96.0, 32.0),
        caption: true,
    },
];

/// Whether an ActiveX control's kind has a text an edit can give it: its caption, or the value
/// of a text box or a combo box.
pub(super) fn activex_text(key: &str) -> bool {
    ACTIVEX_KINDS
        .iter()
        .any(|k| k.key == key && (k.caption || matches!(k.forms, "TextBox" | "ComboBox")))
}

/// The names the objects of a sheet have: its controls', its VML shapes' (an ActiveX control's
/// shape is named after it) and its drawing's.
fn names_on(s: &mut Sheet, vml_doc: &str) -> Result<Vec<String>> {
    let mut out = Vec::new();
    let root = tree(&s.doc, &s.part)?;
    if let Some((_, _, list)) = sheet_controls(&root) {
        out.extend(list.iter().filter_map(|e| e.el.attr(&s.doc, "name")));
    }
    if let Some(vroot) = scan::tree(vml_doc, &[]) {
        let mut shapes = Vec::new();
        vml_shapes(&vroot, &mut shapes);
        out.extend(shapes.iter().filter_map(|x| x.attr(vml_doc, "id")));
    }
    if let Some(d) = drawing_of(s, false)? {
        out.extend(
            scan::tags(&d.doc, 0)
                .filter(|t| !t.is_close() && t.local() == "cNvPr")
                .filter_map(|t| scan::attr(&d.doc, t.start..t.end, "name")),
        );
    }
    Ok(out)
}

/// The file with a new ActiveX control on the sheet `sheet`, written as Excel writes one it
/// puts there: its part (`xl/activeX/activeX<n>.xml`, `persistStreamInit`) and its data
/// (`activeX<n>.bin`: its CLSID and the stream of a new control of its kind, MS-OFORMS), its
/// entry in the sheet's `<controls>` (made when the sheet has none) with the relationship to
/// its part, and its VML shape (`ObjectType="Pict"`, named after it; the VML drawing made when
/// the sheet has none). Its name is `name`, or as Excel names it: its type and the first number
/// no object of the sheet has (`CommandButton1`). Its caption is `text` or its name (a text
/// box's or a combo box's value is `text`).
pub(super) fn add_activex(
    bytes: &[u8],
    sheet: &str,
    kind: &str,
    place: &Place,
    name: Option<&str>,
    text: Option<&str>,
) -> Result<Vec<u8>> {
    let k = ACTIVEX_KINDS
        .iter()
        .find(|k| k.key == kind)
        .ok_or_else(|| format!("ActiveX コントロール `{kind}` はありません"))?;
    let sizes = sizes_of(bytes, sheet)?;
    let b = box_of(&sizes, place, k.size);
    let b = (b.0, b.1, b.2.max(b.0 + 1.0), b.3.max(b.1 + 1.0));
    let mut s = Sheet::open(bytes, sheet)?;
    let n = new_shape(&mut s)?;

    // Its name: one no other object of the sheet has (VBA names them so).
    let taken: Vec<String> = names_on(&mut s, &n.vml_doc)?
        .into_iter()
        .map(|t| t.to_lowercase())
        .collect();
    let name = match name {
        Some(name) => {
            if taken.contains(&name.to_lowercase()) {
                return Err(format!(
                    "シート `{sheet}` には、もう `{name}` という名前のオブジェクトがあります"
                ));
            }
            name.to_string()
        }
        None => (1..)
            .map(|i| format!("{}{i}", k.forms))
            .find(|n| !taken.contains(&n.to_lowercase()))
            .unwrap_or_default(),
    };

    // Its part and its data.
    let number = (1..)
        .find(|i| {
            !s.has(&format!("xl/activeX/activeX{i}.xml"))
                && !s.has(&format!("xl/activeX/activeX{i}.bin"))
        })
        .unwrap_or(1);
    let part = format!("xl/activeX/activeX{number}.xml");
    let bin = format!("xl/activeX/activeX{number}.bin");
    let clsid = crate::activex::clsid_of(k.forms).ok_or("ActiveX コントロールの種類が不明です")?;
    s.put(
        &part,
        format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?>\n<ax:ocx ax:classid=\"{{{clsid}}}\" ax:persistence=\"persistStreamInit\" r:id=\"rId1\" xmlns:ax=\"http://schemas.microsoft.com/office/2006/activeX\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>"
        ),
    );
    s.put(
        &super::insert::rels_part_of(&part),
        format!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{ACTIVEX_BINARY}\" Target=\"activeX{number}.bin\"/></Relationships>"
        ),
    );
    let himetric = |px: f64| (px * 2540.0 / 96.0).round() as i32;
    let (caption, value) = match (k.caption, k.forms) {
        (true, "CheckBox" | "OptionButton" | "ToggleButton") => {
            (Some(text.unwrap_or(&name)), Some("0"))
        }
        (true, _) => (Some(text.unwrap_or(&name)), None),
        (false, _) => (None, text),
    };
    let data = crate::activex::fresh(
        k.forms,
        caption,
        value,
        (himetric(b.2 - b.0), himetric(b.3 - b.1)),
    )
    .ok_or("ActiveX コントロールのデータを作れません")?;
    s.put_bytes(&bin, data);
    s.content_type(Some(&part), None, ACTIVEX_TYPE)?;
    bin_content_type(&mut s, &bin)?;
    let rid = s.add_rel("control", &relative(&s.part.clone(), &part), false)?;

    // Its VML shape: named after it, its picture made by Excel (`x:AutoPict`).
    let id = n.id;
    let style = vml_style(&format!("position:absolute;z-index:{}", n.z), b);
    let shape = format!(
        "<v:shape id=\"{}\" o:spid=\"_x0000_s{id}\" type=\"#_x0000_t201\" style=\"{style}\" stroked=\"f\" strokecolor=\"windowText [64]\" o:insetmode=\"auto\"><o:lock v:ext=\"edit\" rotation=\"t\" text=\"t\"/><x:ClientData ObjectType=\"Pict\"><x:SizeWithCells/><x:Anchor>{}</x:Anchor><x:CF>Pict</x:CF><x:AutoPict/></x:ClientData></v:shape>",
        scan::escape(&name),
        vml_anchor(&sizes, b)
    );
    put_shape(&mut s, &n, &shape);

    // Its entry in `<controls>`, with the fallback Excel 2007 reads.
    put_entry(&mut s, &rid, |e| {
        let (p, r) = (&e.p, &e.r);
        let plain = format!(
            "<{p}control shapeId=\"{id}\" {r} name=\"{}\"/>",
            scan::escape(&name)
        );
        let entry = format!(
            "<mc:AlternateContent xmlns:mc=\"{MC_NS}\"{}><mc:Choice Requires=\"x14\"><{p}control shapeId=\"{id}\" {r} name=\"{}\"><{p}controlPr defaultSize=\"0\" autoLine=\"0\" autoPict=\"0\"><{p}anchor moveWithCells=\"1\"{}>{}{}</{p}anchor></{p}controlPr></{p}control></mc:Choice><mc:Fallback>{plain}</mc:Fallback></mc:AlternateContent>",
            e.x14_decl,
            scan::escape(&name),
            e.xdr_decl,
            control_marker(&format!("{p}from"), &e.xdr, &sizes, b.0, b.1),
            control_marker(&format!("{p}to"), &e.xdr, &sizes, b.2, b.3),
        );
        (entry, Some(plain))
    })?;
    let doc = s.doc.clone();
    s.finish(Some(doc))
}

/// The content type of an ActiveX control's data: the files' `bin` (`Default`) when they have
/// none, else (an `.xlsm` whose `bin` is its VBA project's) its own (`Override`).
fn bin_content_type(s: &mut Sheet, bin: &str) -> Result<()> {
    let doc = s
        .text("[Content_Types].xml")?
        .ok_or("[Content_Types].xml がありません")?;
    let default = scan::tags(&doc, 0)
        .filter(|t| !t.is_close() && t.local() == "Default")
        .find(|t| {
            scan::attr(&doc, t.start..t.end, "Extension")
                .is_some_and(|e| e.eq_ignore_ascii_case("bin"))
        })
        .map(|t| scan::attr(&doc, t.start..t.end, "ContentType").unwrap_or_default());
    match default {
        None => s.content_type(None, Some("bin"), ACTIVEX_BIN_TYPE),
        Some(t) if t == ACTIVEX_BIN_TYPE => Ok(()),
        Some(_) => s.content_type(Some(bin), None, ACTIVEX_BIN_TYPE),
    }
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
