//! The `anchor` of a control (RAS language specification §14.11.16): the edges
//! of its container it keeps its distance to when the container is resized, as
//! the `Anchor` property of Windows Forms.
//!
//! MSForms has no such property, so the form's code module gets a procedure that
//! `UserForm_Resize` calls (see [`crate::codebehind`]). It moves and sizes the
//! anchored controls by how much the inside of the form changed since the last
//! time:
//!
//! ```text
//! Private Sub RAS_FormResize(Optional ByVal mode As Long)
//!     Static w As Single, h As Single, ready As Boolean
//!     Dim dw As Single, dh As Single
//!     If mode = 1 Then ready = False: Exit Sub
//!     If mode = 2 Then w = Me.InsideWidth: h = Me.InsideHeight: ready = True: Exit Sub
//!     If Not ready Then Exit Sub
//!     dw = Me.InsideWidth - w: dh = Me.InsideHeight - h
//!     If dw = 0 And dh = 0 Then Exit Sub
//!     w = Me.InsideWidth: h = Me.InsideHeight
//!     With Me.txtNote
//!         If .Width + dw > 0 Then .Width = .Width + dw Else .Width = 0
//!     End With
//!     With Me.btnOK
//!         .Left = .Left + dw
//!         .Top = .Top + dh
//!     End With
//! End Sub
//! ```
//!
//! `RAS_FormInit` calls it with `1` before it fixes the size of the form for the
//! display scaling and with `2` after, so that only what happens after the form
//! has loaded moves the controls. The rules are those of the designer
//! (`editors/vscode/src/userForm/anchor.ts`) and of the debug run
//! (`ras-compiler/src/js/forms.js`):
//!
//! * left and right: the control is stretched; right only: it moves with the
//!   right edge; neither: it moves by half the change; left only (the
//!   default): it stays. The same for top and bottom.
//! * A Frame or MultiPage that is stretched passes the change on to the
//!   controls in it (the controls of its pages).

use crate::spec::{FormSpec, FormSpecControl};

/// Name of the generated procedure that moves the anchored controls.
pub const RESIZE_PROCEDURE: &str = "RAS_FormResize";

/// The edges a control keeps its distance to.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Anchor {
    /// The top edge.
    pub top: bool,
    /// The bottom edge.
    pub bottom: bool,
    /// The left edge.
    pub left: bool,
    /// The right edge.
    pub right: bool,
}

impl Default for Anchor {
    fn default() -> Anchor {
        Anchor {
            top: true,
            bottom: false,
            left: true,
            right: false,
        }
    }
}

/// Parses an `anchor` value: `top,left`, `Top, Bottom, Left, Right`, `none`.
/// `None` when it is not such a list.
pub fn parse(text: &str) -> Option<Anchor> {
    let words: Vec<String> = text
        .split(|c: char| c == ',' || c == '|' || c.is_whitespace())
        .filter(|w| !w.is_empty())
        .map(str::to_lowercase)
        .collect();
    if words.len() == 1 && words[0] == "none" {
        return Some(Anchor {
            top: false,
            bottom: false,
            left: false,
            right: false,
        });
    }
    if words.is_empty() {
        return None;
    }
    let mut anchor = Anchor {
        top: false,
        bottom: false,
        left: false,
        right: false,
    };
    for word in words {
        let side = match word.as_str() {
            "top" => &mut anchor.top,
            "bottom" => &mut anchor.bottom,
            "left" => &mut anchor.left,
            "right" => &mut anchor.right,
            _ => return None,
        };
        if *side {
            return None;
        }
        *side = true;
    }
    Some(anchor)
}

/// Whether a control of `control_type` can be anchored (a Page fills its
/// MultiPage; a progress bar lays out its labels itself).
pub fn takes_anchor(control_type: &str) -> bool {
    let t = control_type.trim();
    !t.eq_ignore_ascii_case("Page") && !crate::progress::is_progress_bar(t)
}

/// The anchor of `control` (the default when it has none or an invalid one).
pub fn of(control: &FormSpecControl) -> Anchor {
    if !takes_anchor(&control.control_type) {
        return Anchor::default();
    }
    control
        .scalar("anchor")
        .and_then(serde_json::Value::as_str)
        .and_then(parse)
        .unwrap_or_default()
}

/// What happens to one axis of a control: how it moves and grows with the
/// change of its container.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Follow {
    Stay,
    Move,
    Half,
    Stretch,
}

fn follow(near: bool, far: bool) -> Follow {
    match (near, far) {
        (true, true) => Follow::Stretch,
        (false, true) => Follow::Move,
        (false, false) => Follow::Half,
        (true, false) => Follow::Stay,
    }
}

/// The statements for one axis of a control (`Left` / `Width`), inside its `With`.
fn axis_lines(how: Follow, start: &str, size: &str, delta: &str, out: &mut Vec<String>) {
    match how {
        Follow::Stay => {}
        Follow::Move => out.push(format!(".{start} = .{start} + {delta}")),
        Follow::Half => out.push(format!(".{start} = .{start} + {delta} / 2")),
        Follow::Stretch => out.push(format!(
            "If .{size} + {delta} > 0 Then .{size} = .{size} + {delta} Else .{size} = 0"
        )),
    }
}

/// The generated procedure that moves the anchored controls of `spec`, or
/// `None` when no control follows the size of the form.
pub fn resize_procedure(spec: &FormSpec) -> Option<String> {
    let mut body = String::new();
    // Controls in the order of the list; a container passes the change on to its
    // children only when it is stretched (`dx` / `dy` say whether it changes).
    fn visit(spec: &FormSpec, parent: &str, dx: bool, dy: bool, body: &mut String, depth: usize) {
        if (!dx && !dy) || depth > 64 {
            return;
        }
        for control in spec.controls.iter().filter(|c| c.parent_id == parent) {
            let kind = control.control_type.trim();
            if kind.eq_ignore_ascii_case("Page") {
                visit(spec, &control.id, dx, dy, body, depth + 1);
                continue;
            }
            if !takes_anchor(kind) || control.name.is_empty() {
                continue;
            }
            let anchor = of(control);
            let horizontal = if dx {
                follow(anchor.left, anchor.right)
            } else {
                Follow::Stay
            };
            let vertical = if dy {
                follow(anchor.top, anchor.bottom)
            } else {
                Follow::Stay
            };
            let picker = crate::picker::is_file_picker(kind);
            // A file picker is a text box and a button: the box stretches, the button moves.
            let parts: Vec<(String, Follow)> = if picker {
                let (text, button) = match horizontal {
                    Follow::Stretch => (Follow::Stretch, Follow::Move),
                    other => (other, other),
                };
                vec![
                    (format!("{}_Text", control.name), text),
                    (format!("{}_Button", control.name), button),
                ]
            } else {
                vec![(control.name.clone(), horizontal)]
            };
            for (name, horizontal) in parts {
                let mut lines = Vec::new();
                axis_lines(horizontal, "Left", "Width", "dw", &mut lines);
                axis_lines(vertical, "Top", "Height", "dh", &mut lines);
                if lines.is_empty() {
                    continue;
                }
                body.push_str(&format!("    With Me.{name}\n"));
                for line in lines {
                    body.push_str(&format!("        {line}\n"));
                }
                body.push_str("    End With\n");
            }
            if !picker {
                visit(
                    spec,
                    &control.id,
                    horizontal == Follow::Stretch,
                    vertical == Follow::Stretch,
                    body,
                    depth + 1,
                );
            }
        }
    }
    visit(spec, "", true, true, &mut body, 0);
    if body.is_empty() {
        return None;
    }
    Some(format!(
        "Private Sub {RESIZE_PROCEDURE}(Optional ByVal mode As Long)\n\
         \x20   Static w As Single, h As Single, ready As Boolean\n\
         \x20   Dim dw As Single, dh As Single\n\
         \x20   If mode = 1 Then ready = False: Exit Sub\n\
         \x20   If mode = 2 Then w = Me.InsideWidth: h = Me.InsideHeight: ready = True: Exit Sub\n\
         \x20   If Not ready Then Exit Sub\n\
         \x20   dw = Me.InsideWidth - w: dh = Me.InsideHeight - h\n\
         \x20   If dw = 0 And dh = 0 Then Exit Sub\n\
         \x20   w = Me.InsideWidth: h = Me.InsideHeight\n\
         {body}End Sub\n"
    ))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_lists_of_edges() {
        assert_eq!(parse("top,left"), Some(Anchor::default()));
        assert_eq!(
            parse("Top, Bottom, Left, Right"),
            Some(Anchor {
                top: true,
                bottom: true,
                left: true,
                right: true
            })
        );
        assert_eq!(
            parse("none"),
            Some(Anchor {
                top: false,
                bottom: false,
                left: false,
                right: false
            })
        );
        assert_eq!(parse(""), None);
        assert_eq!(parse("top,top"), None);
        assert_eq!(parse("middle"), None);
    }
}
