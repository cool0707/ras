//! The macro a shape or a form control runs when it is clicked (`OnAction`, "マクロの登録"),
//! written where Excel keeps it: the `macro` attribute of the object in the sheet's drawing, and
//! for a form control also `controlPr` in the sheet and `x:FmlaMacro` in its VML shape.

use std::collections::{BTreeMap, HashSet};

use super::scan::{self, Edits, El};
use super::{Pack, Result, WORKBOOK, sheet_parts};

/// The elements of a drawing that have a macro.
const WITH_MACRO: [&str; 4] = ["sp", "cxnSp", "graphicFrame", "pic"];

/// The file with the object `object` (its name, as the name box shows it; or the form control
/// whose VML shape is `_x0000_s<shape>`) of the sheet `sheet` running the macro `name`
/// (`Sheet1.button1Click`) when it is clicked.
pub(super) fn assign(
    bytes: &[u8],
    sheet: &str,
    object: &str,
    shape: Option<u32>,
    name: &str,
) -> Result<Vec<u8>> {
    let value = format!("[0]!{name}");
    let mut pack = Pack::open(bytes)?;
    let workbook = pack.text(WORKBOOK)?.ok_or("xl/workbook.xml がありません")?;
    let parts = sheet_parts(&mut pack, &workbook)?;
    let sheet_part = parts
        .get(sheet)
        .or_else(|| {
            parts
                .iter()
                .find(|(n, _)| n.to_lowercase() == sheet.to_lowercase())
                .map(|(_, p)| p)
        })
        .cloned()
        .ok_or_else(|| format!("シート `{sheet}` がありません"))?;
    let rels = pack.rels(&sheet_part)?;
    let mut replace: BTreeMap<String, Vec<u8>> = BTreeMap::new();
    // The shapes of the form controls named so (their VML shapes: `_x0000_s1025`).
    let mut spids: HashSet<String> = shape
        .map(|id| format!("_x0000_s{id}"))
        .into_iter()
        .collect();
    let mut found = false;

    // The drawing: the object itself, or the mirror of a form control.
    for (_, _, part) in rels.iter().filter(|r| r.1.ends_with("/drawing")) {
        let Some(doc) = pack.text(part)? else {
            continue;
        };
        let Some(root) = scan::tree(&doc, &["txBody", "graphic", "spPr", "style"]) else {
            continue;
        };
        let mut edits = Edits::default();
        drawing(
            &doc, &root, object, &value, &mut edits, &mut spids, &mut found,
        );
        if !edits.is_empty() {
            replace.insert(entry(&pack, part)?, edits.apply(&doc).into_bytes());
        }
    }

    // The sheet: `<control shapeId="1025" name="ボタン 1"><controlPr macro="...">`.
    if let Some(doc) = pack.text(&sheet_part)? {
        let mut edits = Edits::default();
        let mut tags = scan::tags(&doc, 0).peekable();
        while let Some(tag) = tags.next() {
            if tag.is_close() || tag.local() != "control" {
                continue;
            }
            let span = tag.start..tag.end;
            let Some(id) = scan::attr(&doc, span.clone(), "shapeId") else {
                continue;
            };
            let named =
                !object.is_empty() && scan::attr(&doc, span, "name").as_deref() == Some(object);
            if !named && shape.map(|s| s.to_string()).as_deref() != Some(id.as_str()) {
                continue;
            }
            found = true;
            spids.insert(format!("_x0000_s{id}"));
            if tag.is_empty() {
                continue;
            }
            // Its controlPr (the anchor is inside it).
            for inner in tags.by_ref() {
                if inner.local() == "control" && inner.is_close() {
                    break;
                }
                if inner.local() == "controlPr" && !inner.is_close() {
                    scan::set_attr(&doc, inner.start..inner.end, "macro", &value, &mut edits);
                }
            }
        }
        if !edits.is_empty() {
            replace.insert(entry(&pack, &sheet_part)?, edits.apply(&doc).into_bytes());
        }
    }

    // The VML shapes of the form controls.
    if !spids.is_empty() {
        for (_, _, part) in rels.iter().filter(|r| r.1.ends_with("/vmlDrawing")) {
            let Some(doc) = pack.text(part)? else {
                continue;
            };
            let edits = vml(&doc, &spids, &value);
            if !edits.is_empty() {
                found = true;
                replace.insert(entry(&pack, part)?, edits.apply(&doc).into_bytes());
            }
        }
    }

    if !found {
        return Err(format!(
            "シート `{sheet}` に、マクロを登録できるオブジェクト `{}` がありません",
            if object.is_empty() {
                shape.map(|s| format!("_x0000_s{s}")).unwrap_or_default()
            } else {
                object.to_string()
            }
        ));
    }
    pack.rewrite(&replace, &HashSet::new())
}

fn entry(pack: &Pack, part: &str) -> Result<String> {
    pack.find(part)
        .ok_or_else(|| format!("{part} がありません"))
}

/// The objects of a drawing named `object`, in groups too.
fn drawing(
    doc: &str,
    el: &El,
    object: &str,
    value: &str,
    edits: &mut Edits,
    spids: &mut HashSet<String>,
    found: &mut bool,
) {
    for child in &el.children {
        if WITH_MACRO.contains(&child.local()) {
            let nv = child
                .children
                .iter()
                .find(|c| c.local().starts_with("nv"))
                .and_then(|nv| nv.child("cNvPr"));
            // A form control's mirror: its VML shape.
            let spid = nv
                .and_then(|c| find(c, "compatExt"))
                .and_then(|c| c.attr(doc, "spid"));
            let named = !object.is_empty()
                && nv.and_then(|c| c.attr(doc, "name")).as_deref() == Some(object);
            if named || spid.as_ref().is_some_and(|s| spids.contains(s)) {
                *found = true;
                scan::set_attr(doc, child.open.clone(), "macro", value, edits);
                spids.extend(spid);
            }
        }
        drawing(doc, child, object, value, edits, spids, found);
    }
}

fn find<'a, 'b>(el: &'a El<'b>, name: &str) -> Option<&'a El<'b>> {
    el.children.iter().find_map(|c| {
        if c.local() == name {
            Some(c)
        } else {
            find(c, name)
        }
    })
}

/// `x:FmlaMacro` of the VML shapes `spids` set to `value` (added when they have none).
fn vml(doc: &str, spids: &HashSet<String>, value: &str) -> Edits {
    let mut edits = Edits::default();
    // (VML is not always well-formed XML: its tags are read one by one.)
    let mut tags = scan::tags(doc, 0);
    while let Some(tag) = tags.next() {
        if tag.is_close() || tag.local() != "shape" {
            continue;
        }
        let span = tag.start..tag.end;
        let ids = [
            scan::attr(doc, span.clone(), "id"),
            scan::attr(doc, span, "o:spid"),
        ];
        if !ids.iter().flatten().any(|id| spids.contains(id)) || tag.is_empty() {
            continue;
        }
        let mut client: Option<(usize, String)> = None;
        let mut after: Option<usize> = None;
        let mut done = false;
        while let Some(inner) = tags.next() {
            match (inner.local(), inner.is_close()) {
                ("shape", true) => break,
                ("ClientData", false) if !inner.is_empty() => {
                    client = Some((inner.end, scan::prefix(inner.name).to_string()));
                    after = Some(inner.end);
                }
                // Excel writes it after these.
                ("Anchor" | "PrintObject" | "AutoFill" | "Locked" | "Visible", true)
                    if client.is_some() =>
                {
                    after = Some(inner.end);
                }
                ("FmlaMacro", false) if client.is_some() && !done => {
                    if inner.is_empty() {
                        let prefix = scan::prefix(inner.name);
                        edits.replace(
                            inner.start..inner.end,
                            format!(
                                "<{prefix}FmlaMacro>{}</{prefix}FmlaMacro>",
                                scan::escape(value)
                            ),
                        );
                    } else if let Some(close) = tags
                        .by_ref()
                        .find(|t| t.local() == "FmlaMacro" && t.is_close())
                    {
                        edits.replace(inner.end..close.start, scan::escape(value));
                    }
                    done = true;
                }
                ("ClientData", true) => {
                    if !done && let (Some((_, prefix)), Some(at)) = (&client, after) {
                        edits.insert(
                            at,
                            format!(
                                "<{prefix}FmlaMacro>{}</{prefix}FmlaMacro>",
                                scan::escape(value)
                            ),
                        );
                    }
                    done = true;
                }
                _ => {}
            }
        }
    }
    edits
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn vml_macro_is_set_or_added() {
        let doc = "<xml><v:shape id=\"_x0000_s1025\" o:spid=\"_x0000_s1025\"><x:ClientData ObjectType=\"Button\"><x:Anchor>1, 2</x:Anchor><x:AutoFill>False</x:AutoFill><x:TextHAlign>Center</x:TextHAlign></x:ClientData></v:shape><v:shape id=\"_x0000_s1026\"><x:ClientData ObjectType=\"Button\"><x:FmlaMacro>[0]!Old</x:FmlaMacro></x:ClientData></v:shape><br></xml>";
        let ids: HashSet<String> = ["_x0000_s1025".to_string(), "_x0000_s1026".to_string()].into();
        let out = vml(doc, &ids, "[0]!Sheet1.go").apply(doc);
        assert!(out.contains("<x:AutoFill>False</x:AutoFill><x:FmlaMacro>[0]!Sheet1.go</x:FmlaMacro><x:TextHAlign>"), "{out}");
        assert!(
            out.contains("<x:FmlaMacro>[0]!Sheet1.go</x:FmlaMacro></x:ClientData></v:shape><br>"),
            "{out}"
        );
        assert!(!out.contains("Old"));
    }

    #[test]
    fn drawing_macro_is_set_by_name() {
        let doc = "<xdr:wsDr><xdr:twoCellAnchor><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"2\" name=\"楕円 1\"/></xdr:nvSpPr></xdr:sp></xdr:twoCellAnchor><xdr:twoCellAnchor><xdr:grpSp><xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"3\" name=\"図 2\"/></xdr:nvPicPr></xdr:pic></xdr:grpSp></xdr:twoCellAnchor></xdr:wsDr>";
        let root = scan::tree(doc, &[]).unwrap();
        for (name, want) in [
            ("楕円 1", "<xdr:sp macro=\"[0]!Sheet1.a\" textlink"),
            ("図 2", "<xdr:pic macro=\"[0]!Sheet1.a\">"),
        ] {
            let (mut edits, mut spids, mut found) = (Edits::default(), HashSet::new(), false);
            drawing(
                doc,
                &root,
                name,
                "[0]!Sheet1.a",
                &mut edits,
                &mut spids,
                &mut found,
            );
            assert!(found);
            let out = edits.apply(doc);
            assert!(out.contains(want), "{out}");
        }
    }
}
