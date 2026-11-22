//! Events of user forms and their controls (spec §14.11.2).
//!
//! The catalogue (`userform_events.json`) is xlflow's VBE-observed list of event
//! declarations (`internal/vba/userforms/events/catalogue.json`, MIT License,
//! Copyright (c) 2026 harumiWeb); the VS Code extension uses the same file.

use std::sync::OnceLock;

/// One parameter of an event procedure.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EventParam {
    /// Name used by the VBE (`Cancel`, `KeyAscii`).
    pub name: String,
    /// VBA type (`MSForms.ReturnBoolean`, `Integer`, `Single`).
    pub vba_type: String,
    /// Passed `ByRef` (no `ByVal` in the declaration).
    pub byref: bool,
}

/// An event of a host type.
#[derive(Clone, Debug)]
pub struct EventDef {
    /// `UserForm` or a control type (`TextBox`).
    pub host_type: String,
    pub name: String,
    pub params: Vec<EventParam>,
    /// The event the VBE opens on double click.
    pub is_default: bool,
}

impl EventDef {
    /// RAS spelling of the parameters, for messages and generated handlers:
    /// `cancel: MSForms.ReturnBoolean`, `byref closeMode: short`.
    pub fn ras_params(&self) -> String {
        self.params
            .iter()
            .map(|p| {
                format!(
                    "{}{}: {}",
                    if p.byref { "byref " } else { "" },
                    crate::names::js_name(&p.name),
                    ras_type(&p.vba_type)
                )
            })
            .collect::<Vec<_>>()
            .join(", ")
    }
}

/// RAS type of a VBA type that appears in an event declaration.
pub fn ras_type(vba: &str) -> String {
    match vba {
        "Integer" => "short".into(),
        "Long" => "int".into(),
        "Single" => "float".into(),
        "Double" => "double".into(),
        "String" => "string".into(),
        "Boolean" => "boolean".into(),
        "Variant" => "any".into(),
        "Object" => "object".into(),
        other => other.to_string(),
    }
}

fn parse_params(signature: &str) -> Vec<EventParam> {
    let Some(open) = signature.find('(') else {
        return Vec::new();
    };
    let inner = signature[open + 1..].trim_end().trim_end_matches(')');
    inner
        .split(',')
        .map(str::trim)
        .filter(|p| !p.is_empty())
        .filter_map(|p| {
            let (byref, rest) = match p.strip_prefix("ByVal ") {
                Some(rest) => (false, rest),
                None => (true, p.strip_prefix("ByRef ").unwrap_or(p)),
            };
            let (name, ty) = rest.split_once(" As ")?;
            Some(EventParam {
                name: name.trim().to_string(),
                vba_type: ty.trim().to_string(),
                byref,
            })
        })
        .collect()
}

/// Every event of the catalogue.
pub fn catalogue() -> &'static [EventDef] {
    static CATALOGUE: OnceLock<Vec<EventDef>> = OnceLock::new();
    CATALOGUE.get_or_init(|| {
        let value: serde_json::Value =
            serde_json::from_str(include_str!("userform_events.json")).unwrap_or_default();
        value
            .as_array()
            .into_iter()
            .flatten()
            .filter_map(|e| {
                Some(EventDef {
                    host_type: e.get("type")?.as_str()?.to_string(),
                    name: e.get("name")?.as_str()?.to_string(),
                    params: parse_params(e.get("signature")?.as_str()?),
                    is_default: e
                        .get("isDefault")
                        .and_then(|v| v.as_bool())
                        .unwrap_or(false),
                })
            })
            .collect()
    })
}

/// The host type whose events a control raises: a `Page` raises none of its own (its
/// events go to the `MultiPage`), custom controls have no catalogue entry.
pub fn events_of(host_type: &str) -> Vec<&'static EventDef> {
    catalogue()
        .iter()
        .filter(|e| e.host_type.eq_ignore_ascii_case(host_type))
        .collect()
}

/// The event `event` of `host_type` (names compared without case).
pub fn find_event(host_type: &str, event: &str) -> Option<&'static EventDef> {
    catalogue()
        .iter()
        .find(|e| e.host_type.eq_ignore_ascii_case(host_type) && e.name.eq_ignore_ascii_case(event))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn catalogue_signatures() {
        let click = find_event("CommandButton", "click").unwrap();
        assert!(click.params.is_empty());
        let exit = find_event("TextBox", "Exit").unwrap();
        assert_eq!(
            exit.params,
            vec![EventParam {
                name: "Cancel".into(),
                vba_type: "MSForms.ReturnBoolean".into(),
                byref: false
            }]
        );
        let close = find_event("UserForm", "QueryClose").unwrap();
        assert!(close
            .params
            .iter()
            .all(|p| p.byref && p.vba_type == "Integer"));
        assert_eq!(
            close.ras_params(),
            "byref cancel: short, byref closeMode: short"
        );
        assert!(events_of("Page").is_empty());
        assert!(events_of("UserForm").iter().any(|e| e.name == "Initialize"));
    }

    #[test]
    fn same_catalogue_as_the_vscode_extension() {
        let ours: serde_json::Value =
            serde_json::from_str(include_str!("userform_events.json")).unwrap();
        let theirs: serde_json::Value = serde_json::from_str(include_str!(
            "../../../editors/vscode/src/userForm/eventCatalogue.json"
        ))
        .unwrap();
        assert_eq!(ours, theirs);
    }
}
