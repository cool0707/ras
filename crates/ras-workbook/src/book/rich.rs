//! Text whose characters have formats of their own (Excel's rich text: a word of a cell in red,
//! a size bigger, struck through): the runs of a cell, as the file has them, as the grid shows
//! and types them (JSON), and as they are kept in the workbook while it is edited.

use gridcraft_core::{CellRef, RangeRef};
use serde_json::{Value as Json, json};
use umya_spreadsheet::{
    CellRawValue, Font, RichText, TextElement, VerticalAlignmentRunValues, VerticalTextAlignment,
};

use super::{Book, Result, at_of, coord};

/// The font of a run (`rPr`). What it does not give is the font's default (not bold, ...); a
/// name or a size it does not give is the cell's.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct RunFont {
    pub name: Option<String>,
    pub size: Option<f64>,
    pub bold: bool,
    pub italic: bool,
    /// `single`, `double`, `singleAccounting`, `doubleAccounting`.
    pub underline: Option<String>,
    pub strike: bool,
    /// `RRGGBB` (a theme's color as it comes to).
    pub color: Option<String>,
    /// `superscript` or `subscript`.
    pub va: Option<String>,
    pub charset: Option<i32>,
    pub family: Option<i32>,
    /// `minor` / `major`: the theme's font in place of the name.
    pub scheme: Option<String>,
}

/// A run of the text: None for the font, the cell's own.
#[derive(Clone, Debug, PartialEq)]
pub struct RichRun {
    pub text: String,
    pub font: Option<RunFont>,
}

/// The text of the runs.
pub fn text_of(runs: &[RichRun]) -> String {
    runs.iter().map(|r| r.text.as_str()).collect()
}

/// The font of a run as umya-spreadsheet read it.
fn font_of(f: &Font, theme: &umya_spreadsheet::drawing::Theme) -> RunFont {
    let c = f.color();
    let color = (c.has_value() && !c.automatic())
        .then(|| c.argb_with_theme(theme).to_string())
        .filter(|s| s.len() >= 6 && s.bytes().all(|b| b.is_ascii_hexdigit()))
        .map(|s| s[s.len() - 6..].to_ascii_uppercase());
    RunFont {
        name: Some(f.name().to_string()).filter(|n| !n.is_empty()),
        size: Some(f.size()).filter(|s| *s > 0.0),
        bold: f.bold(),
        italic: f.italic(),
        underline: f
            .underline_value()
            .filter(|u| *u != "none")
            .map(str::to_string),
        strike: f.strikethrough(),
        color,
        va: match f.vertical_text_alignment().val() {
            VerticalAlignmentRunValues::Superscript => Some("superscript".into()),
            VerticalAlignmentRunValues::Subscript => Some("subscript".into()),
            _ => None,
        },
        charset: Some(f.charset()).filter(|c| *c != 0),
        family: Some(f.family()).filter(|c| *c != 0),
        scheme: Some(f.scheme().to_string()).filter(|s| s == "minor" || s == "major"),
    }
}

/// The runs of rich text as umya-spreadsheet read it (runs without text dropped).
pub fn runs_of(rt: &RichText, theme: &umya_spreadsheet::drawing::Theme) -> Vec<RichRun> {
    rt.rich_text_elements()
        .iter()
        .filter(|e| !e.text().is_empty())
        .map(|e| RichRun {
            text: e.text().to_string(),
            font: e.run_properties().map(|f| font_of(f, theme)),
        })
        .collect()
}

/// The runs as umya-spreadsheet keeps them, for the cell of the workbook.
pub fn to_rich_text(runs: &[RichRun]) -> RichText {
    let mut rt = RichText::default();
    for run in runs {
        let mut e = TextElement::default();
        e.set_text(run.text.clone());
        if let Some(rf) = &run.font {
            let mut f = Font::default();
            if let Some(n) = &rf.name {
                f.set_name(n.clone());
            }
            if let Some(s) = &rf.scheme {
                f.set_scheme(s.clone());
            }
            if let Some(s) = rf.size {
                f.set_size(s);
            }
            if rf.bold {
                f.set_bold(true);
            }
            if rf.italic {
                f.set_italic(true);
            }
            if let Some(u) = &rf.underline {
                f.set_underline(u.clone());
            }
            if rf.strike {
                f.set_strikethrough(true);
            }
            if let Some(c) = &rf.color {
                f.color_mut().set_argb_str(format!("FF{c}"));
            }
            if let Some(va) = &rf.va {
                let mut v = VerticalTextAlignment::default();
                v.set_val(if va == "subscript" {
                    VerticalAlignmentRunValues::Subscript
                } else {
                    VerticalAlignmentRunValues::Superscript
                });
                f.set_vertical_text_alignment(v);
            }
            if let Some(c) = rf.charset {
                f.set_charset(c);
            }
            if let Some(c) = rf.family {
                f.set_family(c);
            }
            e.set_run_properties(f);
        }
        rt.add_rich_text_elements(e);
    }
    rt
}

/// The runs for the grid: `[{ t, f: { name, size, b, i, u, s, color, va } }]` (`f` left out
/// for the cell's font; `color` `#RRGGBB`).
pub fn runs_json(runs: &[RichRun]) -> Json {
    Json::Array(
        runs.iter()
            .map(|r| {
                let mut o = json!({ "t": r.text });
                if let Some(f) = &r.font {
                    let mut j = serde_json::Map::new();
                    if let Some(n) = &f.name {
                        j.insert("name".into(), json!(n));
                    }
                    if let Some(s) = f.size {
                        j.insert("size".into(), json!(s));
                    }
                    if f.bold {
                        j.insert("b".into(), json!(true));
                    }
                    if f.italic {
                        j.insert("i".into(), json!(true));
                    }
                    if let Some(u) = &f.underline {
                        j.insert("u".into(), json!(u));
                    }
                    if f.strike {
                        j.insert("s".into(), json!(true));
                    }
                    if let Some(c) = &f.color {
                        j.insert("color".into(), json!(format!("#{c}")));
                    }
                    if let Some(v) = &f.va {
                        j.insert("va".into(), json!(v));
                    }
                    o["f"] = Json::Object(j);
                }
                o
            })
            .collect(),
    )
}

/// The runs of an edit (as [`runs_json`] gives them).
pub fn runs_from_json(j: &Json) -> Result<Vec<RichRun>> {
    let list = j
        .as_array()
        .ok_or("`runs` は `{ t, f }` の配列で指定してください")?;
    let mut out: Vec<RichRun> = Vec::new();
    for r in list {
        let text = r
            .get("t")
            .and_then(Json::as_str)
            .ok_or("`runs` の要素には文字列 `t` を指定してください")?
            .to_string();
        let font = match r.get("f") {
            None | Some(Json::Null) => None,
            Some(Json::Object(f)) => {
                let s = |k: &str| f.get(k).and_then(Json::as_str).map(str::to_string);
                let b = |k: &str| f.get(k).and_then(Json::as_bool).unwrap_or(false);
                let color = match s("color") {
                    None => None,
                    Some(c) => {
                        let hex = c.trim_start_matches('#');
                        if hex.len() != 6 || !hex.bytes().all(|b| b.is_ascii_hexdigit()) {
                            return Err(format!(
                                "`runs` の `color` の `{c}` は #RRGGBB ではありません"
                            ));
                        }
                        Some(hex.to_ascii_uppercase())
                    }
                };
                let underline = s("u").filter(|u| u != "none");
                if let Some(u) = &underline
                    && !["single", "double", "singleAccounting", "doubleAccounting"]
                        .contains(&u.as_str())
                {
                    return Err(format!(
                        "`runs` の `u` の `{u}` は指定できません（single / double / singleAccounting / doubleAccounting）"
                    ));
                }
                let va = s("va").filter(|v| v != "baseline");
                if let Some(v) = &va
                    && v != "superscript"
                    && v != "subscript"
                {
                    return Err(format!(
                        "`runs` の `va` の `{v}` は指定できません（superscript / subscript）"
                    ));
                }
                let size = f.get("size").and_then(Json::as_f64);
                if size.is_some_and(|s| !(1.0..=409.0).contains(&s)) {
                    return Err("`runs` の `size` は 1〜409 で指定してください".into());
                }
                let name = s("name")
                    .map(|n| n.trim().to_string())
                    .filter(|n| !n.is_empty());
                Some(RunFont {
                    charset: name.as_ref().filter(|n| !n.is_ascii()).map(|_| 128),
                    name,
                    size,
                    bold: b("b"),
                    italic: b("i"),
                    underline,
                    strike: b("s"),
                    color,
                    va,
                    family: None,
                    scheme: None,
                })
            }
            Some(_) => return Err("`runs` の `f` はオブジェクトで指定してください".into()),
        };
        if text.is_empty() {
            continue;
        }
        // Runs of the same font are one.
        match out.last_mut() {
            Some(last) if last.font == font => last.text.push_str(&text),
            _ => out.push(RichRun { text, font }),
        }
    }
    Ok(out)
}

/// Whether the runs are worth keeping as rich text: more than one font, or one of its own.
pub fn is_rich(runs: &[RichRun]) -> bool {
    runs.len() > 1 || runs.iter().any(|r| r.font.is_some())
}

// ---- the runs of the cells of the workbook: kept on the file's cells (umya-spreadsheet's rich
// text), where the cells' formats are; `store` leaves them while the cell has their text.

impl Book {
    /// The rich text of the cells of `r`.
    pub(super) fn rich_in(&self, si: usize, r: RangeRef) -> Vec<(CellRef, RichText)> {
        let Ok(ws) = self.file.sheet(si) else {
            return Vec::new();
        };
        let rich = |c: &umya_spreadsheet::Cell| match c.raw_value() {
            CellRawValue::RichText(rt) => Some(rt.clone()),
            _ => None,
        };
        if r.count() <= 4096 {
            return r
                .iter()
                .filter_map(|at| ws.cell(coord(at)).and_then(rich).map(|rt| (at, rt)))
                .collect();
        }
        ws.cells()
            .iter()
            .filter(|c| r.contains(at_of(c)))
            .filter_map(|c| rich(c).map(|rt| (at_of(c), rt)))
            .collect()
    }

    /// The cells of `r` lose their rich text (their text stays, plain).
    pub(crate) fn clear_rich(&mut self, si: usize, r: RangeRef) {
        let list = self.rich_in(si, r);
        if let Ok(ws) = self.file.sheet_mut(si) {
            for (at, rt) in list {
                ws.cell_mut(coord(at))
                    .cell_value_mut()
                    .set_value_string(rt.text().into_owned());
            }
        }
    }

    /// The cell `at` gets the rich text `rt` (`store` keeps it while the cell has its text).
    pub(super) fn put_rich(&mut self, si: usize, at: CellRef, rt: RichText) {
        if let Ok(ws) = self.file.sheet_mut(si) {
            ws.cell_mut(coord(at)).cell_value_mut().set_rich_text(rt);
        }
    }

    /// `Range.Characters(...).Font` as the grid gives it: the text of each cell of `r` laid out
    /// in `runs` — a cell whose text is not theirs (a formula, a number, other text) is left as
    /// it is. Runs of one font that is the cell's make plain text.
    pub fn set_rich(&mut self, si: usize, r: RangeRef, runs: &[RichRun]) -> Result<()> {
        self.check_sheet(si)?;
        let text = text_of(runs);
        let cells: Vec<(CellRef, String)> = match self.model.sheet(si) {
            Some(sh) => sh
                .cells
                .iter_range(r)
                .filter(|(_, c)| c.formula.is_none())
                .filter_map(|(at, c)| match &c.value {
                    gridcraft_core::Value::Text(t) => Some((at, t.to_string())),
                    _ => None,
                })
                .filter(|(_, t)| *t == text || format!("'{t}") == text)
                .collect(),
            None => Vec::new(),
        };
        // (`'` typed before the text is not of it.)
        let unquoted = |runs: &[RichRun]| -> Vec<RichRun> {
            let mut out = runs.to_vec();
            if let Some(first) = out.first_mut() {
                first.text.remove(0);
            }
            out.retain(|r| !r.text.is_empty());
            out
        };
        for (at, shown) in cells {
            let runs = if shown == text {
                runs.to_vec()
            } else {
                unquoted(runs)
            };
            if is_rich(&runs) {
                self.put_rich(si, at, to_rich_text(&runs));
            } else {
                self.clear_rich(si, RangeRef::cell(at));
            }
        }
        Ok(())
    }
}
