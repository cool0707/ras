//! FormSpec validation (diagnostic codes `UFY001`, `UFV001`–`UFV021`).
//!
//! Port of the validation in xlflow's `internal/vba/userforms/spec`
//! (`spec.go`, `appearance.go`; Copyright (c) 2026 harumiWeb, MIT License;
//! see `NOTICE.md`). Two passes run, like in xlflow: one over the raw JSON
//! value (unknown fields, value types, enums) and one over the decoded,
//! normalised model (ids, parents, typed appearance values).

use std::collections::{BTreeMap, HashMap, HashSet};

use serde_json::{Map, Value};

use crate::contract::{self, lookup, Prop, Props, Support, ValueType};
use crate::spec::{
    FormSpec, FormSpecColor, FormSpecControl, FormSpecFont, FormSpecForm, FormSpecTab, KIND,
    XLFLOW_KIND,
};

/// Severity of an [`Issue`].
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Severity {
    /// The form cannot be generated.
    Error,
    /// The form can be generated, but something deserves attention.
    Warning,
}

/// One validation finding.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Issue {
    /// Stable code (`UFV001`, ...).
    pub code: String,
    /// Error or warning.
    pub severity: Severity,
    /// Message in Japanese.
    pub message: String,
    /// JSON path of the offending field (`controls[2].caption`), empty for
    /// the document.
    pub field: String,
}

fn issue(code: &str, severity: Severity, message: String, field: &str) -> Issue {
    Issue {
        code: code.to_string(),
        severity,
        message,
        field: field.to_string(),
    }
}

fn error(code: &str, message: String, field: &str) -> Issue {
    issue(code, Severity::Error, message, field)
}

fn join(base: &str, key: &str) -> String {
    if base.is_empty() {
        key.to_string()
    } else {
        format!("{base}.{key}")
    }
}

fn lookup_raw<'a>(root: &'a Map<String, Value>, key: &str) -> Option<&'a Value> {
    root.get(key).or_else(|| {
        root.iter()
            .find(|(k, _)| k.trim().eq_ignore_ascii_case(key))
            .map(|(_, v)| v)
    })
}

fn str_field<'a>(root: &'a Map<String, Value>, key: &str) -> &'a str {
    lookup_raw(root, key).and_then(Value::as_str).unwrap_or("")
}

fn is_integer(value: &Value) -> bool {
    value.as_i64().is_some() || value.as_u64().is_some()
}

fn matches_type(value: &Value, value_type: ValueType) -> bool {
    match (value_type, value) {
        (_, Value::Null) | (ValueType::Any, _) => true,
        (ValueType::String, v) => v.is_string(),
        (ValueType::Number, v) => v.is_number(),
        (ValueType::Integer, v) => is_integer(v),
        (ValueType::Boolean, v) => v.is_boolean(),
        (ValueType::StringArray, Value::Array(items)) => items.iter().all(Value::is_string),
        (ValueType::Object, v) => v.is_object(),
        (ValueType::ObjectArray, Value::Array(items)) => items.iter().all(Value::is_object),
        _ => false,
    }
}

fn value_text(value: &Value) -> String {
    match value {
        Value::String(s) => s.clone(),
        other => other.to_string(),
    }
}

/// `validateObjectProperties`: unknown fields, types, enums, nested objects,
/// support-level warnings and required fields.
fn validate_object(root: &Map<String, Value>, props: &Props, path: &str) -> Vec<Issue> {
    let mut issues = Vec::new();
    let mut keys: Vec<&String> = root.keys().collect();
    keys.sort();
    for key in keys {
        let value = &root[key];
        let field = join(path, key);
        let Some(prop) = lookup(props, key) else {
            if path == "form" && key.eq_ignore_ascii_case("build") {
                issues.push(error("UFV001", "form.build は廃止されました。form.caption / form.width / form.height を使ってください。".into(), &field));
                continue;
            }
            issues.push(error(
                "UFV001",
                format!("{field} は UserForm の仕様にないフィールドです。"),
                &field,
            ));
            continue;
        };
        if !matches_type(value, prop.value_type) {
            issues.push(error(
                "UFV002",
                format!("{field} は{}にしてください。", prop.value_type.label()),
                &field,
            ));
            continue;
        }
        if !value.is_null()
            && !prop.allowed.is_empty()
            && !prop.allowed.contains(&value_text(value).as_str())
        {
            issues.push(error(
                "UFV003",
                format!(
                    "{field} は次のいずれかにしてください: {}。",
                    prop.allowed.join(", ")
                ),
                &field,
            ));
        }
        if let (Some(children), Some(object)) = (&prop.children, value.as_object()) {
            issues.extend(validate_object(object, children, &field));
        }
        if !matches!(prop.support, Support::Supported | Support::SnapshotOnly) {
            issues.push(issue(
                "UFV013",
                Severity::Warning,
                format!("{field} のサポートレベルは {} です。", prop.support.label()),
                &field,
            ));
        }
    }
    let mut required: Vec<&(&str, Prop)> = props.iter().filter(|(_, p)| p.required).collect();
    required.sort_by_key(|(k, _)| *k);
    for (key, _) in required {
        if lookup_raw(root, key).is_none() {
            issues.push(required_issue(&join(path, key)));
        }
    }
    issues
}

fn required_issue(field: &str) -> Issue {
    error("UFV004", format!("{field} は必須です。"), field)
}

fn invalid_property_issue(field: &str, control_type: &str) -> Issue {
    error(
        "UFV005",
        format!("{field} はコントロールの種類 {control_type} では使えません。"),
        field,
    )
}

fn selected_index_issue(field: &str) -> Issue {
    error(
        "UFV020",
        format!("{field} は -1 か、範囲内の 0 から始まる番号にしてください。"),
        field,
    )
}

fn dimension_issue(field: &str) -> Issue {
    error(
        "UFV016",
        format!("{field} は有限で 0 以上、HIMETRIC に変換できる範囲の値にしてください。"),
        field,
    )
}

fn appearance_issue(field: &str, message: String) -> Issue {
    error("UFV021", message, field)
}

pub(crate) fn valid_dimension(value: f64) -> bool {
    value.is_finite() && value >= 0.0 && (value * 2540.0 / 72.0).round() <= i32::MAX as f64
}

// ---------------------------------------------------------------------------
// Raw pass

struct RawRef {
    path: String,
    id: String,
    parent_id: String,
    name: String,
    control_type: String,
    prog_id: String,
    selected_index: Option<f64>,
    observed_selected_index: Option<f64>,
    children_known: bool,
}

/// Validates the raw JSON document (`ValidateFormSpecSource`).
pub(crate) fn validate_raw(root: &Value) -> Vec<Issue> {
    let Some(root) = root.as_object() else {
        return vec![error(
            "UFV002",
            "UserForm の仕様はオブジェクトにしてください。".into(),
            "",
        )];
    };
    let mut issues = validate_object(root, &contract::document_properties(), "");
    if let Some(form) = root.get("form").and_then(Value::as_object) {
        issues.extend(validate_object(form, &contract::form_properties(), "form"));
        if let Some(observed) = form.get("observed").and_then(Value::as_object) {
            issues.extend(validate_object(
                observed,
                &contract::form_observed_properties(),
                "form.observed",
            ));
        }
        for key in ["width", "height"] {
            if let Some(n) = lookup_raw(form, key).and_then(Value::as_f64) {
                if !valid_dimension(n) {
                    issues.push(dimension_issue(&format!("form.{key}")));
                }
            }
        }
        if let Some(observed) = lookup_raw(form, "observed").and_then(Value::as_object) {
            for key in [
                "width",
                "height",
                "insideWidth",
                "insideHeight",
                "clientWidth",
                "clientHeight",
            ] {
                if let Some(n) = lookup_raw(observed, key).and_then(Value::as_f64) {
                    if !valid_dimension(n) {
                        issues.push(dimension_issue(&format!("form.observed.{key}")));
                    }
                }
            }
        }
    }
    if let Some(controls) = root.get("controls") {
        let mut refs = Vec::new();
        validate_raw_controls(controls, "controls", "", &mut refs, &mut issues);
        issues.extend(validate_raw_structure(&refs));
    }
    if let Some(Value::Array(items)) = root.get("warnings") {
        let props: Props = ["code", "message", "control"]
            .into_iter()
            .map(|k| (k, contract_string()))
            .collect();
        for (index, item) in items.iter().enumerate() {
            let path = format!("warnings[{index}]");
            let Some(entry) = item.as_object() else {
                continue;
            };
            let entry_issues = validate_object(entry, &props, &path);
            let failed = entry_issues.iter().any(|i| i.severity == Severity::Error);
            issues.extend(entry_issues);
            if failed {
                continue;
            }
            let message = str_field(entry, "message");
            if message.trim().is_empty() {
                continue;
            }
            issues.push(issue(
                str_field(entry, "code"),
                Severity::Warning,
                message.to_string(),
                &format!("{path}.message"),
            ));
        }
    }
    issues
}

fn contract_string() -> Prop {
    Prop {
        value_type: ValueType::String,
        required: false,
        support: Support::Supported,
        allowed: &[],
        children: None,
    }
}

fn validate_raw_controls(
    value: &Value,
    path: &str,
    inherited: &str,
    refs: &mut Vec<RawRef>,
    issues: &mut Vec<Issue>,
) {
    let Value::Array(items) = value else { return };
    for (i, item) in items.iter().enumerate() {
        let item_path = format!("{path}[{i}]");
        let Some(control) = item.as_object() else {
            issues.push(error(
                "UFV002",
                format!("{item_path} はオブジェクトにしてください。"),
                &item_path,
            ));
            continue;
        };
        let control_type = str_field(control, "type");
        let prog_id = str_field(control, "progId");
        issues.extend(validate_raw_control_properties(
            control,
            control_type,
            prog_id,
            &item_path,
        ));
        let mut parent_id = str_field(control, "parentId").trim().to_string();
        if parent_id.is_empty() {
            parent_id = inherited.to_string();
        }
        let selected = |object: &Map<String, Value>| {
            lookup_raw(object, "selectedIndex")
                .filter(|v| is_integer(v))
                .and_then(Value::as_f64)
        };
        let mut reference = RawRef {
            path: item_path.clone(),
            id: str_field(control, "id").trim().to_string(),
            parent_id,
            name: str_field(control, "name").trim().to_string(),
            control_type: control_type.trim().to_string(),
            prog_id: prog_id.trim().to_string(),
            selected_index: selected(control),
            observed_selected_index: None,
            children_known: false,
        };
        if let Some(observed) = control.get("observed").and_then(Value::as_object) {
            issues.extend(validate_raw_observed(
                observed,
                control_type,
                prog_id,
                &format!("{item_path}.observed"),
            ));
            reference.observed_selected_index = observed
                .get("selectedIndex")
                .filter(|v| is_integer(v))
                .and_then(Value::as_f64);
        }
        let id = reference.id.clone();
        let children = control.get("controls");
        reference.children_known = children.is_some();
        refs.push(reference);
        if let Some(children) = children {
            issues.push(issue(
                "UFV013",
                Severity::Warning,
                format!("{item_path}.controls は旧形式の入れ子構造です。parentId を使った平らな controls にしてください。"),
                &format!("{item_path}.controls"),
            ));
            validate_raw_controls(
                children,
                &format!("{item_path}.controls"),
                &id,
                refs,
                issues,
            );
        }
    }
}

fn is_custom_prog_id(prog_id: &str) -> bool {
    !prog_id.trim().is_empty() && contract::control_type_by_prog_id(prog_id).is_none()
}

fn add_generic_custom_properties(props: &mut Props) {
    for (name, value_type) in [
        ("caption", ValueType::String),
        ("text", ValueType::String),
        ("value", ValueType::Any),
    ] {
        props.push((
            name,
            Prop {
                value_type,
                required: false,
                support: Support::CustomUnchecked,
                allowed: &[],
                children: None,
            },
        ));
    }
}

fn mark_unsupported(issues: &mut [Issue], control_type: &str) {
    for issue in issues.iter_mut().filter(|i| i.code == "UFV001") {
        let name = issue.field.rsplit('.').next().unwrap_or("").to_string();
        if !known_control_property_name(&name) {
            continue;
        }
        issue.code = "UFV005".into();
        issue.message = format!(
            "{} はコントロールの種類 {control_type} では使えません。",
            issue.field
        );
    }
}

fn known_control_property_name(name: &str) -> bool {
    lookup(&contract::common_control_properties(), name).is_some()
        || contract::CONTROL_TYPES.iter().any(|t| {
            contract::type_properties(t.name).is_some_and(|props| lookup(&props, name).is_some())
        })
}

fn is_placeholder(control_type: &str, prog_id: &str, unsupported: &[String]) -> bool {
    control_type.trim().eq_ignore_ascii_case("Control")
        && prog_id.trim().is_empty()
        && unsupported.iter().any(|u| u == "controlType")
}

fn placeholder_issue(path: &str) -> Issue {
    issue(
        "UFV015",
        Severity::Warning,
        format!("{path} は種類を復元できなかったコントロールの、スナップショット専用のプレースホルダーです。"),
        &format!("{path}.type"),
    )
}

fn validate_raw_control_properties(
    control: &Map<String, Value>,
    control_type: &str,
    prog_id: &str,
    path: &str,
) -> Vec<Issue> {
    let built_in = contract::control_type(control_type).is_some();
    let mut allowed = contract::common_control_properties();
    if built_in {
        allowed = contract::control_properties(control_type);
    } else if is_custom_prog_id(prog_id) {
        add_generic_custom_properties(&mut allowed);
    }
    let mut issues = validate_object(control, &allowed, path);
    if let Some(value) = lookup_raw(control, "picture") {
        issues.extend(validate_raw_picture(
            value,
            control_type,
            &format!("{path}.picture"),
        ));
    }
    if let Some(tabs) = lookup_raw(control, "tabs") {
        issues.extend(validate_raw_tabs(tabs, &format!("{path}.tabs")));
    }
    issues.extend(validate_raw_selected_index(control, control_type, path));
    if built_in {
        mark_unsupported(&mut issues, control_type);
    }
    if !control_type.trim().is_empty() && !prog_id.trim().is_empty() {
        if let Some(by_prog_id) = contract::control_type_by_prog_id(prog_id) {
            if !by_prog_id.name.eq_ignore_ascii_case(control_type.trim()) {
                issues.push(error(
                    "UFV012",
                    format!(
                        "{path}.progId {prog_id:?} は {} 用で、{control_type} 用ではありません。",
                        by_prog_id.name
                    ),
                    &format!("{path}.progId"),
                ));
            }
        }
    }
    if !control_type.trim().is_empty() && !built_in {
        let unsupported: Vec<String> = lookup_raw(control, "unsupported")
            .and_then(Value::as_array)
            .map(|a| {
                a.iter()
                    .filter_map(|v| v.as_str().map(str::to_string))
                    .collect()
            })
            .unwrap_or_default();
        if is_placeholder(control_type, prog_id, &unsupported) {
            issues.push(placeholder_issue(path));
        } else if prog_id.trim().is_empty() {
            issues.push(error(
                "UFV006",
                format!(
                    "{path}.type {control_type:?} は組み込みのコントロールの種類ではありません。"
                ),
                &format!("{path}.type"),
            ));
        } else if contract::control_type_by_prog_id(prog_id).is_none() {
            issues.push(issue(
                "UFV014",
                Severity::Warning,
                format!("{path} はカスタムコントロール {control_type:?}（ProgID {prog_id:?}）で、内容は検査されません。"),
                &format!("{path}.progId"),
            ));
        }
    }
    issues
}

fn validate_raw_picture(value: &Value, control_type: &str, path: &str) -> Vec<Issue> {
    if !control_type.trim().eq_ignore_ascii_case("Image") {
        return vec![invalid_property_issue(path, control_type)];
    }
    let Some(object) = value.as_object() else {
        return vec![error(
            "UFV005",
            format!("{path} はオブジェクトにしてください。"),
            path,
        )];
    };
    let mut issues = Vec::new();
    let mut fields: HashMap<String, &Value> = HashMap::new();
    for (key, item) in object {
        let name = key.trim().to_lowercase();
        if name != "path" && name != "remove" {
            issues.push(error(
                "UFV001",
                format!("{path}.{key} は picture にないフィールドです。"),
                &format!("{path}.{key}"),
            ));
            continue;
        }
        if fields.contains_key(&name) {
            issues.push(error(
                "UFV001",
                format!("{path}.{key} が重複しています。"),
                &format!("{path}.{key}"),
            ));
            continue;
        }
        fields.insert(name, item);
    }
    match (fields.get("path"), fields.get("remove")) {
        (Some(_), Some(_)) => issues.push(error(
            "UFV005",
            format!("{path} には path と remove のどちらか一方だけを書いてください。"),
            path,
        )),
        (Some(p), None) => {
            if p.as_str().is_none_or(|s| s.trim().is_empty()) {
                issues.push(error(
                    "UFV005",
                    format!("{path}.path は空でない文字列にしてください。"),
                    &format!("{path}.path"),
                ));
            }
        }
        (None, Some(r)) => {
            if r.as_bool() != Some(true) {
                issues.push(error(
                    "UFV005",
                    format!("{path}.remove は true にしてください。"),
                    &format!("{path}.remove"),
                ));
            }
        }
        (None, None) => issues.push(error(
            "UFV005",
            format!("{path} には path か remove: true を書いてください。"),
            path,
        )),
    }
    issues
}

fn validate_raw_observed(
    control: &Map<String, Value>,
    control_type: &str,
    prog_id: &str,
    path: &str,
) -> Vec<Issue> {
    let common = contract::common_control_properties();
    let mut allowed: Props = Vec::new();
    for key in [
        "left",
        "top",
        "width",
        "height",
        "tabIndex",
        "enabled",
        "visible",
        "unsupported",
        "properties",
    ] {
        if let Some(prop) = lookup(&common, key) {
            allowed.push((
                common
                    .iter()
                    .find(|(k, _)| *k == key)
                    .map(|(k, _)| *k)
                    .unwrap_or(key),
                prop.clone(),
            ));
        }
    }
    if let Some(specific) = contract::type_properties(control_type) {
        allowed.extend(specific);
    } else if is_custom_prog_id(prog_id) {
        add_generic_custom_properties(&mut allowed);
    }
    let mut issues = validate_object(control, &allowed, path);
    if let Some(tabs) = lookup_raw(control, "tabs") {
        issues.extend(validate_raw_tabs(tabs, &format!("{path}.tabs")));
    }
    issues.extend(validate_raw_selected_index(control, control_type, path));
    if contract::control_type(control_type).is_some() {
        mark_unsupported(&mut issues, control_type);
    }
    issues
}

fn validate_raw_tabs(value: &Value, path: &str) -> Vec<Issue> {
    let Value::Array(items) = value else {
        return Vec::new();
    };
    let props = contract::tab_properties();
    let mut issues = Vec::new();
    let mut seen: Vec<String> = Vec::new();
    for (i, item) in items.iter().enumerate() {
        let item_path = format!("{path}[{i}]");
        let Some(object) = item.as_object() else {
            continue;
        };
        issues.extend(validate_object(object, &props, &item_path));
        let name = str_field(object, "name").trim().to_string();
        if name.is_empty() {
            continue;
        }
        if let Some(first) = seen
            .iter()
            .find(|s| s.to_lowercase() == name.to_lowercase())
        {
            let field = format!("{item_path}.name");
            issues.push(error(
                "UFV019",
                format!(
                    "{field} はタブ名 {first:?} と重複しています（大文字小文字は区別しません）。"
                ),
                &field,
            ));
            continue;
        }
        seen.push(name);
    }
    issues
}

fn is_page_selection_control(control_type: &str) -> bool {
    let t = control_type.trim();
    t.eq_ignore_ascii_case("MultiPage") || t.eq_ignore_ascii_case("TabStrip")
}

fn validate_raw_selected_index(
    values: &Map<String, Value>,
    control_type: &str,
    path: &str,
) -> Vec<Issue> {
    if !is_page_selection_control(control_type) {
        return Vec::new();
    }
    let Some(index) = lookup_raw(values, "selectedIndex")
        .filter(|v| is_integer(v))
        .and_then(Value::as_f64)
    else {
        return Vec::new();
    };
    let mut bad = index < -1.0;
    if control_type.trim().eq_ignore_ascii_case("TabStrip") {
        if let Some(Value::Array(tabs)) = lookup_raw(values, "tabs") {
            bad = (tabs.is_empty() && index != -1.0)
                || (!tabs.is_empty() && (index < 0.0 || index >= tabs.len() as f64));
        }
    }
    if bad {
        vec![selected_index_issue(&format!("{path}.selectedIndex"))]
    } else {
        Vec::new()
    }
}

fn parent_issue(field: &str, parent: (&str, &str, &str), child_type: &str) -> Issue {
    let (parent_id, parent_name, parent_type) = parent;
    let parent_contract = contract::control_type(parent_type);
    if parent_contract.is_some_and(|c| !c.container) {
        return error("UFV011", format!("{field} {parent_id:?} はコンテナーではないコントロール {parent_name:?} を指しています。"), field);
    }
    error(
        "UFV011",
        format!("{field} の親 {parent_name:?} には {child_type} を入れられません（MultiPage には Page だけ、Page は MultiPage の中だけに置けます）。"),
        field,
    )
}

fn known_type(control_type: &str, prog_id: &str) -> Option<&'static contract::ControlType> {
    if !prog_id.trim().is_empty() {
        if let Some(t) = contract::control_type_by_prog_id(prog_id) {
            return Some(t);
        }
    }
    contract::control_type(control_type)
}

fn validate_raw_structure(controls: &[RawRef]) -> Vec<Issue> {
    let mut issues = Vec::new();
    let mut ids: HashMap<&str, &RawRef> = HashMap::new();
    for control in controls {
        if control.id.is_empty() {
            continue;
        }
        if let Some(existing) = ids.get(control.id.as_str()) {
            issues.push(error(
                "UFV007",
                format!(
                    "{}.id {:?} が重複しています（最初の定義は {}.id）。",
                    control.path, control.id, existing.path
                ),
                &format!("{}.id", control.path),
            ));
            continue;
        }
        ids.insert(&control.id, control);
    }
    let mut parent_by_id = BTreeMap::new();
    let mut parent_field_by_id = HashMap::new();
    for control in controls {
        if control.parent_id.is_empty() {
            if control.control_type.eq_ignore_ascii_case("Page") {
                issues.push(page_parent_issue(&format!("{}.parentId", control.path)));
            }
            continue;
        }
        let field = format!("{}.parentId", control.path);
        if !control.id.is_empty() && control.parent_id == control.id {
            issues.push(error(
                "UFV009",
                format!("{field} が自分自身を指しています。"),
                &field,
            ));
            continue;
        }
        let Some(parent) = ids.get(control.parent_id.as_str()) else {
            issues.push(error(
                "UFV008",
                format!("{field} {:?} が見つかりません。", control.parent_id),
                &field,
            ));
            continue;
        };
        let allowed = contract::parent_allows_child(
            known_type(&parent.control_type, &parent.prog_id),
            known_type(&control.control_type, &control.prog_id),
        );
        if allowed == Some(false) {
            issues.push(parent_issue(
                &field,
                (&parent.id, &parent.name, &parent.control_type),
                &control.control_type,
            ));
        }
        if !control.id.is_empty() {
            parent_by_id.insert(control.id.clone(), control.parent_id.clone());
            parent_field_by_id.insert(control.id.clone(), field);
        }
    }
    let mut page_count: HashMap<&str, usize> = HashMap::new();
    let mut known: HashSet<&str> = HashSet::new();
    for control in controls {
        if control.control_type.eq_ignore_ascii_case("MultiPage") && control.children_known {
            known.insert(&control.id);
        }
        if control.control_type.eq_ignore_ascii_case("Page") && !control.parent_id.is_empty() {
            *page_count.entry(&control.parent_id).or_default() += 1;
            known.insert(&control.parent_id);
        }
    }
    for control in controls {
        if !control.control_type.eq_ignore_ascii_case("MultiPage")
            || !known.contains(control.id.as_str())
        {
            continue;
        }
        let count = page_count.get(control.id.as_str()).copied().unwrap_or(0) as f64;
        for (path, index) in [
            (
                format!("{}.selectedIndex", control.path),
                control.selected_index,
            ),
            (
                format!("{}.observed.selectedIndex", control.path),
                control.observed_selected_index,
            ),
        ] {
            let Some(index) = index else { continue };
            if (count == 0.0 && index != -1.0) || (count > 0.0 && (index < 0.0 || index >= count)) {
                issues.push(selected_index_issue(&path));
            }
        }
    }
    issues.extend(cycle_issues(&parent_by_id, &parent_field_by_id));
    issues
}

fn page_parent_issue(field: &str) -> Issue {
    error(
        "UFV011",
        format!("{field} は Page に必須です。Page は MultiPage の中に置いてください。"),
        field,
    )
}

fn cycle_issues(
    parent_by_id: &BTreeMap<String, String>,
    field_by_id: &HashMap<String, String>,
) -> Vec<Issue> {
    let mut issues = Vec::new();
    let mut reported: HashSet<String> = HashSet::new();
    for id in parent_by_id.keys() {
        if reported.contains(id) {
            continue;
        }
        let mut path: Vec<String> = Vec::new();
        let mut seen: HashMap<String, usize> = HashMap::new();
        let mut current = id.clone();
        while !current.is_empty() {
            if let Some(start) = seen.get(&current) {
                let cycle = &path[*start..];
                for cycle_id in cycle {
                    reported.insert(cycle_id.clone());
                }
                let field = field_by_id
                    .get(&cycle[0])
                    .cloned()
                    .unwrap_or_else(|| "controls".into());
                issues.push(error(
                    "UFV010",
                    format!("{:?} の parentId をたどると循環しています。", cycle[0]),
                    &field,
                ));
                break;
            }
            if reported.contains(&current) {
                break;
            }
            seen.insert(current.clone(), path.len());
            path.push(current.clone());
            current = parent_by_id.get(&current).cloned().unwrap_or_default();
        }
    }
    issues
}

// ---------------------------------------------------------------------------
// Strict pass (on the normalised model)

fn validate_font(font: Option<&FormSpecFont>, path: &str) -> Vec<Issue> {
    let mut issues = Vec::new();
    let Some(font) = font else { return issues };
    if let Some(size) = font.size {
        if !size.is_finite() || size <= 0.0 || size > 3276.75 {
            issues.push(appearance_issue(
                &format!("{path}.size"),
                "フォントサイズは 0 より大きく 3276.75 ポイント以下にしてください。".into(),
            ));
        }
    }
    if let Some(charset) = font.charset {
        if !(0..=255).contains(&charset) {
            issues.push(appearance_issue(
                &format!("{path}.charset"),
                "charset は 0〜255 の整数にしてください。".into(),
            ));
        }
    }
    if let Some(name) = &font.name {
        if name.len() > 255 || name.contains('\0') {
            issues.push(appearance_issue(
                &format!("{path}.name"),
                "フォント名は 255 バイト以下で、NUL を含めないでください。".into(),
            ));
        }
    }
    issues
}

fn validate_typed_form(form: &FormSpecForm) -> Vec<Issue> {
    let mut issues = validate_font(form.font.as_ref(), "form.font");
    for (name, color) in [
        ("backColor", &form.back_color),
        ("foreColor", &form.fore_color),
    ] {
        if let Some(color) = color {
            if let Err(message) = color.ole() {
                issues.push(appearance_issue(&format!("form.{name}"), message));
            }
        }
    }
    if form.border_style.is_some_and(|b| !(0..=1).contains(&b)) {
        issues.push(appearance_issue(
            "form.borderStyle",
            "borderStyle は 0 か 1 にしてください。".into(),
        ));
    }
    issues
}

/// The typed appearance fields of a control, in `sortedMapKeys` order.
fn typed_fields(control: &FormSpecControl) -> Vec<(&'static str, Option<Value>)> {
    let json = |v: Option<Value>| v;
    vec![
        (
            "backColor",
            json(control.back_color.as_ref().map(color_json)),
        ),
        ("borderStyle", control.border_style.map(Value::from)),
        (
            "font",
            control
                .font
                .as_ref()
                .map(|f| serde_json::to_value(f).expect("font")),
        ),
        ("foreColor", control.fore_color.as_ref().map(color_json)),
        (
            "pictureAlignment",
            control.picture_alignment.map(Value::from),
        ),
        (
            "pictureSizeMode",
            control.picture_size_mode.map(Value::from),
        ),
        ("textAlign", control.text_align.map(Value::from)),
        ("style", control.style.map(Value::from)),
        ("multiSelect", control.multi_select.map(Value::from)),
        ("listStyle", control.list_style.map(Value::from)),
        ("matchEntry", control.match_entry.map(Value::from)),
        (
            "showDropButtonWhen",
            control.show_drop_button_when.map(Value::from),
        ),
        ("listRows", control.list_rows.map(Value::from)),
    ]
}

fn color_json(color: &FormSpecColor) -> Value {
    serde_json::to_value(color).expect("color")
}

fn prog_id_matches_type(control_type: &str, prog_id: &str) -> bool {
    contract::control_type(control_type)
        .is_some_and(|t| t.prog_id.eq_ignore_ascii_case(prog_id.trim()))
}

fn legacy_equal(name: &str, typed: &Value, legacy: &Value) -> bool {
    if name == "foreColor" || name == "backColor" {
        if let Ok(color) = serde_json::from_value::<FormSpecColor>(typed.clone()) {
            let ole = color.ole();
            let n = legacy.as_f64();
            return match (ole, n) {
                (Ok(ole), Some(n)) => n == ole as f64 || n == (ole as i32) as f64,
                _ => false,
            };
        }
    }
    json_equal(typed, legacy)
}

/// JSON equality where `12` equals `12.0` (Go compares the marshalled text,
/// in which both are `12`).
fn json_equal(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Number(x), Value::Number(y)) => x.as_f64() == y.as_f64(),
        (Value::Array(x), Value::Array(y)) => {
            x.len() == y.len() && x.iter().zip(y).all(|(a, b)| json_equal(a, b))
        }
        (Value::Object(x), Value::Object(y)) => {
            x.len() == y.len()
                && x.iter()
                    .all(|(k, v)| y.get(k).is_some_and(|w| json_equal(v, w)))
        }
        _ => a == b,
    }
}

fn validate_typed_control(control: &FormSpecControl, path: &str) -> Vec<Issue> {
    let mut issues = validate_font(control.font.as_ref(), &format!("{path}.font"));
    let props = contract::control_properties(&control.control_type);
    for (name, value) in typed_fields(control) {
        let Some(value) = value else { continue };
        let field = format!("{path}.{name}");
        let prop = if contract::control_type(&control.control_type).is_some() {
            lookup(&props, name)
        } else {
            None
        };
        let Some(prop) = prop.filter(|_| {
            control.prog_id.is_empty()
                || prog_id_matches_type(&control.control_type, &control.prog_id)
        }) else {
            issues.push(invalid_property_issue(&field, &control.control_type));
            continue;
        };
        if !prop.allowed.is_empty() && !prop.allowed.contains(&value_text(&value).as_str()) {
            issues.push(appearance_issue(
                &field,
                "値がプロパティの選択肢の範囲外です。".into(),
            ));
        }
        if name == "listRows" && !value.as_i64().is_some_and(|n| (0..=i16::MAX as i64).contains(&n)) {
            issues.push(appearance_issue(
                &field,
                "listRows は 0〜32767 の整数にしてください。".into(),
            ));
        }
        if name == "foreColor" || name == "backColor" {
            if let Ok(color) = serde_json::from_value::<FormSpecColor>(value.clone()) {
                if let Err(message) = color.ole() {
                    issues.push(appearance_issue(&field, message));
                }
            }
        }
        if let Some(legacy) = control
            .properties
            .as_ref()
            .and_then(|bag| lookup_raw(bag, name))
        {
            if !legacy_equal(name, &value, legacy) {
                issues.push(appearance_issue(
                    &field,
                    "型付きのプロパティと properties の値が食い違っています。".into(),
                ));
            }
        }
    }
    issues
}

fn validate_tabs(tabs: &[FormSpecTab], path: &str) -> Vec<Issue> {
    let mut issues = Vec::new();
    let mut seen: Vec<String> = Vec::new();
    for (i, tab) in tabs.iter().enumerate() {
        let field = format!("{path}[{i}].name");
        let name = tab.name.trim();
        if name.is_empty() {
            issues.push(required_issue(&field));
            continue;
        }
        if let Some(first) = seen
            .iter()
            .find(|s| s.to_lowercase() == name.to_lowercase())
        {
            issues.push(error(
                "UFV019",
                format!("{field} はタブ名 {first:?} と重複しています。"),
                &field,
            ));
            continue;
        }
        seen.push(name.to_string());
    }
    issues
}

fn validate_selected_index(
    index: Option<i64>,
    count: usize,
    path: &str,
    bounded: bool,
) -> Vec<Issue> {
    let Some(index) = index else {
        return Vec::new();
    };
    if index < -1 || (bounded && ((index == -1 && count > 0) || index >= count as i64)) {
        return vec![selected_index_issue(path)];
    }
    Vec::new()
}

/// The least height of a single-line TextBox / ComboBox for its font, in points
/// (`editHeight.ts` of the VS Code extension computes the same). MSForms does not
/// grow an edit field to its font: a lower one clips the text and leaves parts of
/// glyphs and of the caret as stray lines along the edge.
fn minimum_edit_height(control: &FormSpecControl) -> Option<f64> {
    if !matches!(control.control_type.to_lowercase().as_str(), "textbox" | "combobox") {
        return None;
    }
    let bag = control.properties.as_ref();
    if bag.and_then(|b| lookup_raw(b, "multiLine")).and_then(Value::as_bool) == Some(true) {
        return None;
    }
    let bag_font = bag
        .and_then(|b| lookup_raw(b, "font"))
        .and_then(|f| serde_json::from_value::<crate::spec::FormSpecFont>(f.clone()).ok());
    let font = control.font.clone().or(bag_font).unwrap_or_default();
    let name = font
        .name
        .filter(|n| !n.trim().is_empty())
        .unwrap_or_else(|| "MS UI Gothic".into())
        .trim()
        .to_lowercase();
    let size = font.size.filter(|s| *s > 0.0).unwrap_or(9.0);
    // Line height (ascent + descent) per point of size.
    let starts = |prefixes: &[&str]| prefixes.iter().any(|p| name.starts_with(p));
    let has = |parts: &[&str]| parts.iter().any(|p| name.contains(p));
    let factor = if starts(&["ms", "ｍｓ"]) && has(&["gothic", "ゴシック", "mincho", "明朝"]) {
        1.0
    } else if starts(&["meiryo ui", "メイリオ ui"]) {
        1.33
    } else if starts(&["meiryo", "メイリオ"]) {
        1.6
    } else if starts(&["yu gothic ui", "游ゴシック ui"]) {
        1.33
    } else if starts(&["yu gothic", "yu mincho", "游ゴシック", "游明朝"]) {
        1.71
    } else if starts(&["tahoma"]) {
        1.21
    } else if starts(&["segoe ui"]) {
        1.33
    } else if starts(&["arial", "times"]) {
        1.15
    } else if starts(&["calibri", "verdana"]) {
        1.22
    } else {
        1.3
    };
    // The sunken edge and the inner margin, then whole pixels at 96 DPI.
    let points = size * factor + 4.5;
    Some((points / 0.75 - 1e-9).ceil() * 0.75)
}

fn validate_control(control: &FormSpecControl, path: &str) -> Vec<Issue> {
    let mut issues = validate_typed_control(control, path);
    if let (Some(minimum), Some(height)) = (minimum_edit_height(control), control.height) {
        if height < minimum {
            issues.push(issue(
                "UFV022",
                Severity::Warning,
                format!(
                    "{path}.height（{height} pt）はフォントに対して低すぎます。Excel では文字が欠け、端に線が残ります。{minimum} pt 以上にしてください。"
                ),
                &format!("{path}.height"),
            ));
        }
    }
    if control.id.trim().is_empty() {
        issues.push(required_issue(&format!("{path}.id")));
    }
    if control.name.trim().is_empty() {
        issues.push(required_issue(&format!("{path}.name")));
    }
    if control.control_type.trim().is_empty() {
        issues.push(required_issue(&format!("{path}.type")));
    }
    let unsupported = control.unsupported.clone().unwrap_or_default();
    if is_placeholder(&control.control_type, &control.prog_id, &unsupported) {
        issues.push(placeholder_issue(path));
    } else if control.prog_id.trim().is_empty()
        && contract::control_type(&control.control_type).is_none()
    {
        issues.push(error(
            "UFV006",
            format!(
                "{path}: コントロールの種類 {:?} には対応していません。",
                control.control_type
            ),
            &format!("{path}.type"),
        ));
    }
    if !control.control_type.trim().is_empty() && !control.prog_id.trim().is_empty() {
        if let Some(t) = contract::control_type_by_prog_id(&control.prog_id) {
            if !t.name.eq_ignore_ascii_case(control.control_type.trim()) {
                issues.push(error(
                    "UFV012",
                    format!(
                        "{path}.progId {:?} は {} 用で、{} 用ではありません。",
                        control.prog_id, t.name, control.control_type
                    ),
                    &format!("{path}.progId"),
                ));
            }
        }
    }
    let is = |name: &str| control.control_type.trim().eq_ignore_ascii_case(name);
    if control.picture.is_some() && !is("Image") {
        issues.push(invalid_property_issue(
            &format!("{path}.picture"),
            &control.control_type,
        ));
    }
    if is("Page") {
        for (name, value) in [
            ("left", control.left),
            ("top", control.top),
            ("width", control.width),
            ("height", control.height),
        ] {
            if value.is_some() {
                issues.push(invalid_property_issue(
                    &format!("{path}.{name}"),
                    &control.control_type,
                ));
            }
        }
    }
    if let Some(tabs) = &control.tabs {
        if !is("TabStrip") {
            issues.push(invalid_property_issue(
                &format!("{path}.tabs"),
                &control.control_type,
            ));
        } else {
            issues.extend(validate_tabs(tabs, &format!("{path}.tabs")));
        }
    }
    if is_page_selection_control(&control.control_type) {
        let count = control.tabs.as_ref().map_or(0, Vec::len);
        issues.extend(validate_selected_index(
            control.selected_index,
            count,
            &format!("{path}.selectedIndex"),
            is("TabStrip") && control.tabs.is_some(),
        ));
    }
    if is("ComboBox") || is("ListBox") {
        // The selection is set by generated code after the items were
        // added: it has to name one of them (or none, -1).
        if let Some(index) = control.selected_index {
            let count = control.list.as_ref().map_or(0, Vec::len) as i64;
            if index < -1 || index >= count {
                issues.push(selected_index_issue(&format!("{path}.selectedIndex")));
            }
        }
    }
    issues
}

/// `ValidateFormSpecStrict`.
pub(crate) fn validate_strict(spec: &FormSpec) -> Vec<Issue> {
    let mut issues = Vec::new();
    if spec.schema_version != 1 {
        issues.push(error(
            "UFV003",
            "schemaVersion は 1 にしてください。".into(),
            "schemaVersion",
        ));
    }
    if spec.kind != KIND && spec.kind != XLFLOW_KIND {
        issues.push(error(
            "UFV003",
            format!("kind は {KIND:?} にしてください。"),
            "kind",
        ));
    }
    if spec.basis.trim() != "designer" {
        issues.push(error(
            "UFV003",
            "basis は \"designer\" にしてください。".into(),
            "basis",
        ));
    }
    if spec.form.name.trim().is_empty() {
        issues.push(required_issue("form.name"));
    }
    for (name, value) in [("width", spec.form.width), ("height", spec.form.height)] {
        if value.is_some_and(|v| !valid_dimension(v)) {
            issues.push(dimension_issue(&format!("form.{name}")));
        }
    }
    issues.extend(validate_typed_form(&spec.form));
    let mut by_id: HashMap<&str, &FormSpecControl> = HashMap::new();
    for (i, control) in spec.controls.iter().enumerate() {
        let path = format!("controls[{i}]");
        issues.extend(validate_control(control, &path));
        if control.id.trim().is_empty() {
            continue;
        }
        if by_id.contains_key(control.id.as_str()) {
            issues.push(error(
                "UFV007",
                format!("{path}.id {:?} が重複しています。", control.id),
                &format!("{path}.id"),
            ));
            continue;
        }
        by_id.insert(&control.id, control);
    }
    let mut parent_by_id = BTreeMap::new();
    let mut field_by_id = HashMap::new();
    let mut page_count: HashMap<&str, usize> = HashMap::new();
    let mut topology_known: HashSet<&str> = HashSet::new();
    for (i, control) in spec.controls.iter().enumerate() {
        let is_page = control.control_type.trim().eq_ignore_ascii_case("Page");
        if is_page && !control.parent_id.trim().is_empty() {
            *page_count.entry(&control.parent_id).or_default() += 1;
            topology_known.insert(&control.parent_id);
        }
        let field = format!("controls[{i}].parentId");
        if control.parent_id.trim().is_empty() {
            if is_page {
                issues.push(page_parent_issue(&field));
            }
            continue;
        }
        if control.parent_id == control.id {
            issues.push(error(
                "UFV009",
                format!("{field} が自分自身を指しています。"),
                &field,
            ));
            continue;
        }
        let Some(parent) = by_id.get(control.parent_id.as_str()) else {
            issues.push(error(
                "UFV008",
                format!("{field} {:?} が見つかりません。", control.parent_id),
                &field,
            ));
            continue;
        };
        let allowed = contract::parent_allows_child(
            known_type(&parent.control_type, &parent.prog_id),
            known_type(&control.control_type, &control.prog_id),
        );
        if allowed == Some(false) {
            issues.push(parent_issue(
                &field,
                (&parent.id, &parent.name, &parent.control_type),
                &control.control_type,
            ));
        }
        parent_by_id.insert(control.id.clone(), control.parent_id.clone());
        field_by_id.insert(control.id.clone(), field);
    }
    for (i, control) in spec.controls.iter().enumerate() {
        if !control
            .control_type
            .trim()
            .eq_ignore_ascii_case("MultiPage")
            || !topology_known.contains(control.id.as_str())
        {
            continue;
        }
        let count = page_count.get(control.id.as_str()).copied().unwrap_or(0);
        issues.extend(validate_selected_index(
            control.selected_index,
            count,
            &format!("controls[{i}].selectedIndex"),
            true,
        ));
    }
    issues.extend(cycle_issues(&parent_by_id, &field_by_id));
    issues
}
