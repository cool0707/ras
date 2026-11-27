//! The `placeholder` of a TextBox (RAS language specification §14.11.18): a gray hint shown
//! in the box while it is empty and does not have the focus.
//!
//! MSForms has no such property, so the form gets a Label `<name>_Placeholder` over the box
//! when it is written ([`expand`]); `ras form pull` turns it back into the property
//! ([`collapse`]). The label is transparent, shows the I-beam pointer and has the font and
//! the alignment of the box. The code module of the form shows and hides it
//! ([`handlers`], called from [`crate::codebehind`]):
//!
//! ```text
//! Private Sub RAS_Placeholder(ByVal box As Object, ByVal mark As Object, ByVal mode As Long)
//!     ' mode: 0 the box lost the focus, 1 it got it, 2 by the active control
//!     ...
//! End Sub
//!
//! Private Sub txtName_Enter()
//!     RAS_Placeholder Me.txtName, Me.txtName_Placeholder, 1
//! End Sub
//!
//! Private Sub txtName_Placeholder_MouseUp(ByVal Button As Integer, ...)
//!     RAS_PlaceholderClick Me.txtName, Me.txtName_Placeholder
//! End Sub
//! ```
//!
//! * Releasing a mouse button over the label hides it and gives the box the focus.
//! * The box getting the focus (Enter, and Enter of the Frame / MultiPage it is in) hides it.
//! * The box losing the focus (Exit, and Exit of its Frame / MultiPage) shows it again when
//!   the box is empty; so does a change of the text by code while the box has no focus.
//!
//! The debug run (`ras-compiler/src/js/forms.js`, `webview/formRunner`) and the designer draw
//! the same hint without the label.

use serde_json::{json, Value};

use crate::spec::{FormSpec, FormSpecColor, FormSpecControl};

/// The suffix of the name of the label.
pub const SUFFIX: &str = "_Placeholder";
/// The `Tag` of the label, by which `ras form pull` knows it.
pub const TAG: &str = "RAS.Placeholder";
/// The generated procedure that shows or hides a label.
pub const PROCEDURE: &str = "RAS_Placeholder";
/// The generated procedure run when a mouse button is released over a label.
pub const CLICK_PROCEDURE: &str = "RAS_PlaceholderClick";
/// Distance in points from the edge of the box to the text (the sunken border and the
/// margin MSForms leaves in front of the first character).
pub const INSET: f64 = 3.0;
/// `fmMousePointerIBeam`.
const IBEAM: i64 = 3;

/// The placeholder of `control`, when it is a TextBox with one that is not empty.
pub fn of(control: &FormSpecControl) -> Option<&str> {
    if !control.control_type.trim().eq_ignore_ascii_case("TextBox") {
        return None;
    }
    control
        .scalar("placeholder")
        .and_then(Value::as_str)
        .filter(|text| !text.is_empty())
}

/// Whether the hint of `control` shows when the form opens (the box is empty).
fn shown_at_first(control: &FormSpecControl) -> bool {
    let text = control.text.clone().or_else(|| match &control.value {
        Some(Value::String(s)) => Some(s.clone()),
        Some(Value::Null) | None => None,
        Some(other) => Some(other.to_string()),
    });
    text.unwrap_or_default().is_empty() && control.visible != Some(false)
}

/// The label written over `control` (a TextBox with a placeholder).
pub fn label(control: &FormSpecControl, text: &str) -> FormSpecControl {
    let left = control.left.unwrap_or(0.0);
    let top = control.top.unwrap_or(0.0);
    let width = control.width.unwrap_or(72.0);
    let height = control.height.unwrap_or(18.0);
    let mut scalars = serde_json::Map::new();
    scalars.insert("backStyle".into(), json!(0));
    scalars.insert("mousePointer".into(), json!(IBEAM));
    let multi_line = control
        .scalar("multiLine")
        .and_then(Value::as_bool)
        .unwrap_or(false);
    scalars.insert("wordWrap".into(), json!(multi_line));
    FormSpecControl {
        id: format!("{}{SUFFIX}", control.id),
        parent_id: control.parent_id.clone(),
        z_index: control.z_index,
        control_type: "Label".into(),
        name: format!("{}{SUFFIX}", control.name),
        caption: Some(text.to_string()),
        tag: Some(TAG.into()),
        control_tip_text: control.control_tip_text.clone(),
        left: Some(left + INSET),
        top: Some(top + INSET),
        width: Some((width - 2.0 * INSET).max(0.75)),
        height: Some((height - 2.0 * INSET).max(0.75)),
        visible: Some(shown_at_first(control)),
        font: control.font.clone(),
        fore_color: Some(FormSpecColor {
            kind: "system".into(),
            value: "grayText".into(),
        }),
        text_align: control.text_align,
        scalars,
        ..FormSpecControl::default()
    }
}

/// Puts the label of each TextBox with a placeholder right after the box (over it).
pub fn expand(spec: &FormSpec) -> FormSpec {
    if !spec.controls.iter().any(|c| of(c).is_some()) {
        return spec.clone();
    }
    let mut out = spec.clone();
    out.controls = spec
        .controls
        .iter()
        .flat_map(|c| {
            let mut box_ = c.clone();
            box_.scalars
                .retain(|k, _| !k.eq_ignore_ascii_case("placeholder"));
            match of(c) {
                Some(text) => vec![box_, label(c, text)],
                None => vec![box_],
            }
        })
        .collect();
    out
}

/// Turns the labels of a form read from a workbook (`ras form pull`) back into the
/// `placeholder` of their boxes.
pub fn collapse(spec: &mut FormSpec) {
    loop {
        let found = spec.controls.iter().enumerate().find_map(|(i, c)| {
            let base = c.name.strip_suffix(SUFFIX)?;
            if !c.control_type.eq_ignore_ascii_case("Label") || c.tag.as_deref() != Some(TAG) {
                return None;
            }
            let at = spec.controls.iter().position(|o| {
                o.name.eq_ignore_ascii_case(base)
                    && o.control_type.eq_ignore_ascii_case("TextBox")
                    && o.parent_id == c.parent_id
            })?;
            Some((i, at, c.caption.clone().unwrap_or_default()))
        });
        let Some((label_at, box_at, text)) = found else {
            return;
        };
        spec.controls[box_at]
            .scalars
            .insert("placeholder".into(), json!(text));
        spec.controls.remove(label_at);
    }
}

/// One line the code module runs in the event procedure `handler`.
pub struct Hook {
    /// `txtName_Enter`.
    pub handler: String,
    /// The parameters of the event procedure (`ByVal Cancel As MSForms.ReturnBoolean`).
    pub params: &'static str,
    /// The statement (`RAS_Placeholder Me.txtName, Me.txtName_Placeholder, 1`).
    pub line: String,
}

const EXIT_PARAMS: &str = "ByVal Cancel As MSForms.ReturnBoolean";
const MOUSE_PARAMS: &str =
    "ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single";

/// The generated procedures and the lines for the event procedures, or `None` when no
/// TextBox has a placeholder.
pub fn handlers(spec: &FormSpec) -> Option<(String, Vec<Hook>)> {
    let boxes: Vec<&FormSpecControl> = spec
        .controls
        .iter()
        .filter(|c| of(c).is_some() && !c.name.is_empty())
        .collect();
    if boxes.is_empty() {
        return None;
    }
    let mut hooks = Vec::new();
    let mut hook = |handler: String, params: &'static str, line: String| {
        hooks.push(Hook {
            handler,
            params,
            line,
        })
    };
    for c in &boxes {
        let name = &c.name;
        let call = |mode: u8| format!("{PROCEDURE} Me.{name}, Me.{name}{SUFFIX}, {mode}");
        hook("UserForm_Activate".into(), "", call(2));
        hook(format!("{name}_Enter"), "", call(1));
        hook(format!("{name}_Exit"), EXIT_PARAMS, call(0));
        hook(format!("{name}_Change"), "", call(2));
        hook(
            format!("{name}{SUFFIX}_MouseUp"),
            MOUSE_PARAMS,
            format!("{CLICK_PROCEDURE} Me.{name}, Me.{name}{SUFFIX}"),
        );
        // MSForms does not run Enter / Exit of a control when the focus comes into or leaves
        // the Frame or MultiPage it is in, only those of the Frame / MultiPage.
        let mut parent = c.parent_id.as_str();
        for _ in 0..64 {
            let Some(container) = spec
                .controls
                .iter()
                .find(|o| o.id == parent && !parent.is_empty())
            else {
                break;
            };
            let kind = container.control_type.trim();
            if kind.eq_ignore_ascii_case("Frame") || kind.eq_ignore_ascii_case("MultiPage") {
                hook(format!("{}_Enter", container.name), "", call(2));
                hook(format!("{}_Exit", container.name), EXIT_PARAMS, call(0));
            }
            parent = container.parent_id.as_str();
        }
    }
    let procedures = format!(
        "Private Sub {PROCEDURE}(ByVal box As Object, ByVal mark As Object, ByVal mode As Long)\n\
         \x20   Dim shown As Boolean, c As Object, n As Object, i As Long\n\
         \x20   On Error Resume Next\n\
         \x20   shown = mode <> 1 And box.Visible And Len(box.Text) = 0\n\
         \x20   If shown And mode = 2 Then\n\
         \x20       Set c = Me.ActiveControl\n\
         \x20       For i = 1 To 64\n\
         \x20           If c Is Nothing Then Exit For\n\
         \x20           If c.Name = box.Name Then shown = False: Exit For\n\
         \x20           Set n = Nothing\n\
         \x20           Select Case TypeName(c)\n\
         \x20           Case \"Frame\": Set n = c.ActiveControl\n\
         \x20           Case \"MultiPage\": Set n = c.SelectedItem.ActiveControl\n\
         \x20           End Select\n\
         \x20           Set c = n\n\
         \x20       Next\n\
         \x20   End If\n\
         \x20   If shown Then\n\
         \x20       If mark.Left <> box.Left + {INSET} Then mark.Left = box.Left + {INSET}\n\
         \x20       If mark.Top <> box.Top + {INSET} Then mark.Top = box.Top + {INSET}\n\
         \x20   End If\n\
         \x20   If mark.Visible <> shown Then mark.Visible = shown\n\
         End Sub\n\
         \n\
         Private Sub {CLICK_PROCEDURE}(ByVal box As Object, ByVal mark As Object)\n\
         \x20   On Error Resume Next\n\
         \x20   If Not box.Enabled Then Exit Sub\n\
         \x20   mark.Visible = False\n\
         \x20   box.SetFocus\n\
         End Sub\n"
    );
    Some((procedures, hooks))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn spec(controls: &str) -> FormSpec {
        let text = format!(
            r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"F"}},"controls":[{controls}]}}"#
        );
        let parsed = crate::parse(&text);
        assert!(!parsed.has_errors(), "{:?}", parsed.issues);
        parsed.spec.unwrap()
    }

    #[test]
    fn a_placeholder_is_a_label_over_the_box() {
        let spec = spec(
            r#"{"id":"t","name":"txtName","type":"TextBox","left":10,"top":20,"width":120,"height":18,"placeholder":"名前を入力","textAlign":3},
               {"id":"u","name":"txtFull","type":"TextBox","text":"x","placeholder":"hint"},
               {"id":"v","name":"txtPlain","type":"TextBox","placeholder":""}"#,
        );
        let out = expand(&spec);
        let names: Vec<&str> = out.controls.iter().map(|c| c.name.as_str()).collect();
        assert_eq!(
            names,
            [
                "txtName",
                "txtName_Placeholder",
                "txtFull",
                "txtFull_Placeholder",
                "txtPlain"
            ]
        );
        let label = &out.controls[1];
        assert_eq!(label.control_type, "Label");
        assert_eq!(label.caption.as_deref(), Some("名前を入力"));
        assert_eq!(label.tag.as_deref(), Some(TAG));
        assert_eq!(
            (label.left, label.top, label.width, label.height),
            (Some(13.0), Some(23.0), Some(114.0), Some(12.0))
        );
        assert_eq!(label.text_align, Some(3));
        assert_eq!(label.scalar("mousePointer"), Some(&json!(3)));
        assert_eq!(label.visible, Some(true));
        // A box with text hides its hint until it is emptied.
        assert_eq!(out.controls[3].visible, Some(false));
        assert!(out.controls[0].scalar("placeholder").is_none());
        // The whole form can be written.
        crate::compile(&spec, 932).unwrap();

        let mut back = out.clone();
        collapse(&mut back);
        let names: Vec<&str> = back.controls.iter().map(|c| c.name.as_str()).collect();
        assert_eq!(names, ["txtName", "txtFull", "txtPlain"]);
        assert_eq!(
            back.controls[0].scalar("placeholder"),
            Some(&json!("名前を入力"))
        );
    }

    #[test]
    fn the_box_and_its_containers_show_and_hide_the_hint() {
        let spec = spec(
            r#"{"id":"m","name":"mpMain","type":"MultiPage"},
               {"id":"p","parentId":"m","name":"pgA","type":"Page"},
               {"id":"f","parentId":"p","name":"fraIn","type":"Frame"},
               {"id":"t","parentId":"f","name":"txtName","type":"TextBox","placeholder":"x"}"#,
        );
        let (_, hooks) = handlers(&spec).unwrap();
        let lines: Vec<String> = hooks
            .iter()
            .map(|h| format!("{}: {}", h.handler, h.line))
            .collect();
        let call = "RAS_Placeholder Me.txtName, Me.txtName_Placeholder";
        assert_eq!(
            lines,
            [
                format!("UserForm_Activate: {call}, 2"),
                format!("txtName_Enter: {call}, 1"),
                format!("txtName_Exit: {call}, 0"),
                format!("txtName_Change: {call}, 2"),
                "txtName_Placeholder_MouseUp: RAS_PlaceholderClick Me.txtName, Me.txtName_Placeholder".into(),
                format!("fraIn_Enter: {call}, 2"),
                format!("fraIn_Exit: {call}, 0"),
                format!("mpMain_Enter: {call}, 2"),
                format!("mpMain_Exit: {call}, 0"),
            ]
        );
        assert!(handlers(&self::spec(r#"{"id":"t","name":"txt","type":"TextBox"}"#)).is_none());
    }
}
