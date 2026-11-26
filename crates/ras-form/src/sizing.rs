//! The size of a form at run time (RAS language specification §14.11.17):
//! whether the user can resize it by its border (`form.resizable`), and the
//! least and greatest size of its inside (`form.minimumWidth`,
//! `form.minimumHeight`, `form.maximumWidth`, `form.maximumHeight`), as
//! `FormBorderStyle.Sizable`, `MinimumSize` and `MaximumSize` of Windows Forms.
//!
//! MSForms has none of these, so the designer storage does not hold them:
//!
//! * `RAS_FormInit` ([`crate::codebehind`]) calls `RAS_FormSizable Me` of the
//!   runtime (`ras-compiler/src/runtime.vba`), which gives the window of the
//!   form a sizing border through the Windows API, before it fixes the inside
//!   size for the display scaling.
//! * `RAS_FormResize` ([`crate::anchor`]), which `UserForm_Resize` calls, puts
//!   a size outside the limits back to the nearest limit before it moves the
//!   anchored controls:
//!
//! ```text
//!     If busy Then Exit Sub
//!     busy = True
//!     If 200 - Me.InsideWidth >= 1 Then Me.Width = Me.Width + 200 - Me.InsideWidth
//!     If Me.InsideWidth - 480 >= 1 Then Me.Width = Me.Width + 480 - Me.InsideWidth
//!     If 120 - Me.InsideHeight >= 1 Then Me.Height = Me.Height + 120 - Me.InsideHeight
//!     busy = False
//! ```
//!
//! The limits apply to the user's resizing and to the form's code setting
//! `Width` / `Height` alike. The debug run (`ras-compiler/src/js/forms.js`) and
//! the designer (`editors/vscode/src/userForm/engine.ts`) follow the same rules.

use serde_json::Value;

use crate::spec::{FormSpec, FormSpecForm};

/// Name of the runtime procedure that gives a form its sizing border.
pub const SIZABLE_PROCEDURE: &str = "RAS_FormSizable";

/// The fields of `form` that this module reads (none is in the designer storage).
pub const FIELDS: [&str; 5] = [
    "resizable",
    "minimumWidth",
    "minimumHeight",
    "maximumWidth",
    "maximumHeight",
];

/// Whether the user can resize the form by its border.
pub fn resizable(form: &FormSpecForm) -> bool {
    form.scalar("resizable").and_then(Value::as_bool) == Some(true)
}

/// The least and greatest inside size of a form, in points. `0` (or a field
/// not written) is no limit.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub struct Limits {
    /// `minimumWidth`.
    pub min_width: f64,
    /// `minimumHeight`.
    pub min_height: f64,
    /// `maximumWidth`.
    pub max_width: f64,
    /// `maximumHeight`.
    pub max_height: f64,
}

impl Limits {
    /// No size is limited.
    pub fn is_empty(&self) -> bool {
        self.min_width <= 0.0
            && self.min_height <= 0.0
            && self.max_width <= 0.0
            && self.max_height <= 0.0
    }
}

/// A limit of `form` (`0` when it is not written or not a number).
fn limit(form: &FormSpecForm, name: &str) -> f64 {
    form.scalar(name)
        .and_then(Value::as_f64)
        .filter(|v| v.is_finite() && *v > 0.0)
        .unwrap_or(0.0)
}

/// The limits of the size of `form`.
pub fn limits(form: &FormSpecForm) -> Limits {
    Limits {
        min_width: limit(form, "minimumWidth"),
        min_height: limit(form, "minimumHeight"),
        max_width: limit(form, "maximumWidth"),
        max_height: limit(form, "maximumHeight"),
    }
}

/// The lines of `RAS_FormResize` that keep the inside of the form within
/// `limits` (empty without limits). `busy` (a `Static` of the procedure)
/// keeps the `UserForm_Resize` of the size put back from running it again.
pub(crate) fn clamp_lines(limits: &Limits) -> String {
    if limits.is_empty() {
        return String::new();
    }
    let mut out = String::from("    If busy Then Exit Sub\n    busy = True\n");
    for (inside, size, min, max) in [
        ("InsideWidth", "Width", limits.min_width, limits.max_width),
        (
            "InsideHeight",
            "Height",
            limits.min_height,
            limits.max_height,
        ),
    ] {
        // Within a point: the window is sized in whole pixels.
        if min > 0.0 {
            let n = crate::codebehind::points_literal(min);
            out.push_str(&format!(
                "    If {n} - Me.{inside} >= 1 Then Me.{size} = Me.{size} + {n} - Me.{inside}\n"
            ));
        }
        if max > 0.0 {
            let n = crate::codebehind::points_literal(max);
            out.push_str(&format!(
                "    If Me.{inside} - {n} >= 1 Then Me.{size} = Me.{size} + {n} - Me.{inside}\n"
            ));
        }
    }
    out.push_str("    busy = False\n");
    out
}

/// Whether the generated code of `spec` changes the size of the form at run
/// time in any way (a sizing border or limits).
pub fn applies(spec: &FormSpec) -> bool {
    resizable(&spec.form) || !limits(&spec.form).is_empty()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn form(fields: &str) -> FormSpecForm {
        serde_json::from_str(&format!(r#"{{"name":"F"{fields}}}"#)).unwrap()
    }

    #[test]
    fn limits_of_zero_are_none() {
        assert!(limits(&form("")).is_empty());
        assert!(limits(&form(r#","minimumWidth":0,"maximumHeight":0"#)).is_empty());
        let l = limits(&form(r#","minimumWidth":200,"maximumHeight":300.5"#));
        assert_eq!(
            l,
            Limits {
                min_width: 200.0,
                min_height: 0.0,
                max_width: 0.0,
                max_height: 300.5
            }
        );
        assert!(!resizable(&form("")));
        assert!(resizable(&form(r#","Resizable":true"#)));
    }

    #[test]
    fn limits_are_validated() {
        let codes = |fields: &str| -> Vec<String> {
            let text = format!(
                r#"{{"schemaVersion":1,"kind":"ras.userform","basis":"designer","form":{{"name":"F","width":240,"height":180{fields}}},"controls":[]}}"#
            );
            crate::parse(&text)
                .issues
                .iter()
                .filter(|i| i.severity == crate::Severity::Error)
                .map(|i| format!("{} {}", i.code, i.field))
                .collect()
        };
        assert!(codes(
            r#","resizable":true,"minimumWidth":200,"maximumWidth":480,"maximumHeight":0"#
        )
        .is_empty());
        assert_eq!(
            codes(r#","minimumWidth":500,"maximumWidth":480"#),
            ["UFV027 form.minimumWidth"]
        );
        assert_eq!(codes(r#","minimumHeight":200"#), ["UFV027 form.height"]);
        assert_eq!(codes(r#","maximumWidth":200"#), ["UFV027 form.width"]);
        assert_eq!(codes(r#","minimumWidth":-1"#), ["UFV016 form.minimumWidth"]);
        assert_eq!(codes(r#","resizable":1"#).len(), 1);
    }

    #[test]
    fn sizes_outside_the_limits_are_put_back() {
        let l = limits(&form(
            r#","minimumWidth":200,"maximumWidth":480,"minimumHeight":120"#,
        ));
        assert_eq!(
            clamp_lines(&l),
            "    If busy Then Exit Sub\n    busy = True\n    If 200 - Me.InsideWidth >= 1 Then Me.Width = Me.Width + 200 - Me.InsideWidth\n    If Me.InsideWidth - 480 >= 1 Then Me.Width = Me.Width + 480 - Me.InsideWidth\n    If 120 - Me.InsideHeight >= 1 Then Me.Height = Me.Height + 120 - Me.InsideHeight\n    busy = False\n"
        );
        assert_eq!(clamp_lines(&Limits::default()), "");
    }
}
