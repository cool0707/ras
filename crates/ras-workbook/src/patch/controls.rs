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
/// click (px), and whether it shows a text.
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
        size: (96.0, 32.0),
        text: true,
    },
    Kind {
        key: "checkBox",
        vml: "Checkbox",
        prop: "CheckBox",
        name: "チェック",
        size: (112.0, 22.0),
        text: true,
    },
    Kind {
        key: "optionButton",
        vml: "Radio",
        prop: "Radio",
        name: "オプション",
        size: (112.0, 22.0),
        text: true,
    },
    Kind {
        key: "label",
        vml: "Label",
        prop: "Label",
        name: "ラベル",
        size: (72.0, 20.0),
        text: true,
    },
    Kind {
        key: "groupBox",
        vml: "GBox",
        prop: "GBox",
        name: "グループ",
        size: (144.0, 96.0),
        text: true,
    },
    Kind {
        key: "listBox",
        vml: "List",
        prop: "List",
        name: "リスト",
        size: (96.0, 80.0),
        text: false,
    },
    Kind {
        key: "comboBox",
        vml: "Drop",
        prop: "Drop",
        name: "ドロップ",
        size: (96.0, 22.0),
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
        size: (20.0, 96.0),
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
    if entry_rel
        .as_ref()
        .is_some_and(|r| r.1.ends_with("/control"))
    {
        return Err(
            "ActiveX コントロールは、ここでは移動・削除できません（Excel で行ってください）".into(),
        );
    }
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
    if let (Some(sh), Some(doc)) = (shape, vml_doc.as_deref())
        && matches!(object_type(doc, sh).as_str(), "Pict" | "Note")
    {
        return Err(format!("`{spid}` はフォーム コントロールではありません"));
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
    if delete {
        if let Some(sh) = shape {
            vml_edits.replace(sh.span(), "");
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
            None,
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
        entry_rel.map(|r| (r.0, r.2)),
    )
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
    props: Option<(String, String)>,
) -> Result<Vec<u8>> {
    if let (Some(part), Some(doc)) = (vml, vml_doc)
        && !vml_edits.is_empty()
    {
        s.put(&part, vml_edits.apply(&doc));
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
    if let Some((rid, part)) = props {
        s.remove_rels_of(&s.part.clone(), &[rid].into_iter().collect())?;
        drop_part(&mut s, &part)?;
    }
    let doc = (!sheet_edits.is_empty()).then(|| sheet_edits.apply(&sdoc));
    s.finish(doc)
}

// ---------------------------------------------------------------------------- put on a sheet

/// The text of a control as VML writes it in its text box.
fn vml_textbox(text: &str, align: &str) -> String {
    let lines: Vec<String> = text.split('\n').map(scan::escape).collect();
    format!(
        "<v:textbox style=\"mso-direction-alt:auto\" o:singleclick=\"f\"><div style=\"text-align:{align}\"><font face=\"{FONT}\" size=\"220\" color=\"#000000\">{}</font></div></v:textbox>",
        lines.join("<br/>")
    )
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
                num("Max", c.max, 100),
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
            c.max.unwrap_or(100),
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
        let size_with = if k.key == "button" {
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
