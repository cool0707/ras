//! The formats the program sets: `Font`, `Interior`, `Borders`, alignment, protection, the sizes
//! of rows and columns, merged cells. They are kept in the file (umya-spreadsheet's styles), so
//! they move with the cells and are saved with the workbook. Nothing is drawn.
//!
//! A property is named as the stand-in's object model names it, in lower case:
//! `font.bold`, `interior.color`, `borders.9.linestyle` (the number is `XlBordersIndex`, `all`
//! for every side of every cell), `horizontalalignment`, `columnwidth`, ...

use gridcraft_core::{CellRef, RangeRef};
use umya_spreadsheet::{
    Border, BorderStyleValues, Color, EnumTrait, HorizontalAlignmentValues, PatternValues, Style,
    UnderlineValues, VerticalAlignmentRunValues, VerticalAlignmentValues,
};

use super::{Book, MAX_COLS, MAX_ROWS, Result, Val, at_of, coord};

pub(super) const DEFAULT_COLUMN_WIDTH: f64 = 8.38;
/// The width of a digit of the default font in pixels, which the widths of columns are counted
/// in. The default font is taken to be 游ゴシック 11pt, as everywhere here (Calibri 11pt has 7).
const DIGIT_WIDTH: f64 = 8.0;
/// The padding of a cell in pixels, which the width of a column has in the file.
const CELL_PADDING: f64 = 5.0;

/// The width of a column as the file has it, of a `ColumnWidth` (ECMA-376 18.3.1.13).
pub(super) fn width_to_file(chars: f64) -> f64 {
    ((chars * DIGIT_WIDTH + CELL_PADDING) / DIGIT_WIDTH * 256.0).trunc() / 256.0
}

/// The `ColumnWidth` of a width of the file.
pub(super) fn width_from_file(width: f64) -> f64 {
    (((width * DIGIT_WIDTH - CELL_PADDING) / DIGIT_WIDTH * 100.0 + 0.5).trunc() / 100.0).max(0.0)
}
const DEFAULT_ROW_HEIGHT: f64 = 18.75;
const DEFAULT_FONT: &str = "游ゴシック";
const NONE: f64 = -4142.0;
const AUTOMATIC: f64 = -4105.0;

/// Excel's default palette (`ColorIndex` 1 to 56), as RGB.
const PALETTE: [u32; 56] = [
    0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF, 0x800000,
    0x008000, 0x000080, 0x808000, 0x800080, 0x008080, 0xC0C0C0, 0x808080, 0x9999FF, 0x993366,
    0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF, 0x000080, 0xFF00FF, 0xFFFF00,
    0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF, 0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99,
    0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99, 0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900,
    0xFF6600, 0x666699, 0x969696, 0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366,
    0x333399, 0x333333,
];

/// VBA's colour number (`RGB(r, g, b)`: red in the low byte) of RGB.
fn bgr(rgb: u32) -> f64 {
    (((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF)) as f64
}

fn rgb_of(bgr: f64) -> u32 {
    let n = bgr.max(0.0) as u32 & 0xFF_FFFF;
    ((n & 0xFF) << 16) | (n & 0xFF00) | ((n >> 16) & 0xFF)
}

fn num(v: &Val) -> Result<f64> {
    match v {
        Val::Num(n) | Val::Date(n) => Ok(*n),
        Val::Bool(b) => Ok(if *b { -1.0 } else { 0.0 }),
        Val::Empty => Ok(0.0),
        Val::Text(s) => s
            .trim()
            .parse()
            .map_err(|_| "型が一致しません。".to_string()),
        Val::Err(_) => Err("型が一致しません。".into()),
    }
}

fn flag(v: &Val) -> Result<bool> {
    match v {
        Val::Bool(b) => Ok(*b),
        Val::Text(s) if s.eq_ignore_ascii_case("true") => Ok(true),
        Val::Text(s) if s.eq_ignore_ascii_case("false") => Ok(false),
        other => Ok(num(other)? != 0.0),
    }
}

fn text(v: &Val) -> String {
    match v {
        Val::Text(s) => s.clone(),
        Val::Num(n) | Val::Date(n) => gridcraft_core::number_to_text(*n),
        Val::Bool(b) => if *b { "True" } else { "False" }.into(),
        _ => String::new(),
    }
}

/// `LineStyle` and `Weight` of a border style of the file.
fn line_of(style: &BorderStyleValues) -> (f64, f64) {
    use BorderStyleValues as B;
    const THIN: f64 = 2.0;
    const MEDIUM: f64 = -4138.0;
    match style {
        B::None => (NONE, THIN),
        B::Hair => (1.0, 1.0),
        B::Thin => (1.0, THIN),
        B::Medium => (1.0, MEDIUM),
        B::Thick => (1.0, 4.0),
        B::Dashed => (-4115.0, THIN),
        B::MediumDashed => (-4115.0, MEDIUM),
        B::DashDot => (4.0, THIN),
        B::MediumDashDot => (4.0, MEDIUM),
        B::DashDotDot => (5.0, THIN),
        B::MediumDashDotDot => (5.0, MEDIUM),
        B::Dotted => (-4118.0, THIN),
        B::Double => (-4119.0, 4.0),
        B::SlantDashDot => (13.0, MEDIUM),
    }
}

/// The border style of the file for a `LineStyle` and `Weight`.
fn style_of_line(line: f64, weight: f64) -> BorderStyleValues {
    use BorderStyleValues as B;
    let medium = weight == -4138.0 || weight == 4.0;
    match line as i32 {
        -4142 => B::None,
        -4115 => {
            if medium {
                B::MediumDashed
            } else {
                B::Dashed
            }
        }
        4 => {
            if medium {
                B::MediumDashDot
            } else {
                B::DashDot
            }
        }
        5 => {
            if medium {
                B::MediumDashDotDot
            } else {
                B::DashDotDot
            }
        }
        -4118 => B::Dotted,
        -4119 => B::Double,
        13 => B::SlantDashDot,
        _ => match weight as i32 {
            1 => B::Hair,
            -4138 => B::Medium,
            4 => B::Thick,
            _ => B::Thin,
        },
    }
}

/// A side of a cell.
#[derive(Clone, Copy, PartialEq)]
enum Side {
    Left,
    Right,
    Top,
    Bottom,
    DiagonalDown,
    DiagonalUp,
}

fn border(style: &Style, side: Side) -> Option<&Border> {
    let b = style.borders()?;
    match side {
        Side::Left => Some(b.left()),
        Side::Right => Some(b.right()),
        Side::Top => Some(b.top()),
        Side::Bottom => Some(b.bottom()),
        Side::DiagonalDown => b.diagonal_down().then(|| b.diagonal()),
        Side::DiagonalUp => b.diagonal_up().then(|| b.diagonal()),
    }
}

fn border_mut(style: &mut Style, side: Side) -> &mut Border {
    let b = style.borders_mut();
    match side {
        Side::Left => b.left_mut(),
        Side::Right => b.right_mut(),
        Side::Top => b.top_mut(),
        Side::Bottom => b.bottom_mut(),
        Side::DiagonalDown => {
            b.set_diagonal_down(true);
            b.diagonal_mut()
        }
        Side::DiagonalUp => {
            b.set_diagonal_up(true);
            b.diagonal_mut()
        }
    }
}

impl Book {
    fn colour(&self, c: &Color) -> u32 {
        let hex = c.argb_with_theme(self.file.theme());
        let hex = hex.trim_start_matches('#');
        let hex = if hex.len() > 6 {
            &hex[hex.len() - 6..]
        } else {
            hex
        };
        u32::from_str_radix(hex, 16).unwrap_or(0)
    }

    fn colour_index(&self, c: Option<&Color>, none: f64) -> f64 {
        let Some(c) = c.filter(|c| c.has_value()) else {
            return none;
        };
        let rgb = self.colour(c);
        PALETTE
            .iter()
            .position(|p| *p == rgb)
            .map_or(none, |i| (i + 1) as f64)
    }

    /// The style of a cell: its own, or its row's or column's when the cell is not stored.
    pub(super) fn style_at(&self, si: usize, at: CellRef) -> Style {
        let Ok(ws) = self.file.sheet(si) else {
            return Style::default();
        };
        if let Some(c) = ws.cell(coord(at)) {
            return c.style().clone();
        }
        let plain = Style::default();
        if let Some(r) = ws.row_dimension(at.row + 1).filter(|r| *r.style() != plain) {
            return r.style().clone();
        }
        ws.column_dimension_by_number(at.col + 1)
            .map(|c| c.style().clone())
            .unwrap_or(plain)
    }

    /// The styles of the cells of `r` that may differ: every cell of a small range; of a large
    /// one its first cell and the cells the file stores.
    fn styles_of(&self, si: usize, r: RangeRef) -> Vec<Style> {
        if r.count() <= 4096 {
            return r.iter().map(|at| self.style_at(si, at)).collect();
        }
        let mut out = vec![self.style_at(si, r.start)];
        if let Ok(ws) = self.file.sheet(si) {
            out.extend(
                ws.cells()
                    .iter()
                    .filter(|c| r.contains(at_of(c)))
                    .map(|c| c.style().clone()),
            );
        }
        out
    }

    /// Changes the style of every cell of `r`. Whole rows and columns are formatted as such, and
    /// where cells are stored.
    fn restyle(&mut self, si: usize, r: RangeRef, f: &dyn Fn(&mut Style)) -> Result<()> {
        let whole_cols = r.is_full_cols();
        let whole_rows = r.is_full_rows() && !whole_cols;
        if !whole_cols && !whole_rows && r.count() > 5_000_000 {
            return Err("範囲が大きすぎます".into());
        }
        let cells: Vec<CellRef> = if whole_cols || whole_rows {
            let mut cells: Vec<CellRef> = self
                .model
                .sheet(si)
                .map(|s| s.cells.iter_range(r).map(|(at, _)| at).collect())
                .unwrap_or_default();
            if let Ok(ws) = self.file.sheet(si) {
                cells.extend(
                    ws.cells()
                        .iter()
                        .map(|c| at_of(c))
                        .filter(|at| r.contains(*at)),
                );
            }
            cells.sort_by_key(|at| (at.row, at.col));
            cells.dedup();
            cells
        } else {
            r.iter().collect()
        };
        let Ok(ws) = self.file.sheet_mut(si) else {
            return Ok(());
        };
        if whole_cols {
            for col in r.start.col..=r.end.col {
                f(ws.column_dimension_by_number_mut(col + 1).style_mut());
            }
        } else if whole_rows {
            for row in r.start.row..=r.end.row {
                f(ws.row_dimension_mut(row + 1).style_mut());
            }
        }
        for at in cells {
            f(ws.style_mut(coord(at)));
        }
        Ok(())
    }

    /// A cell the file does not store yet gets the format of its row or column, as a cell
    /// typed into has in Excel.
    pub(super) fn inherit_style(&mut self, si: usize, at: CellRef) {
        let Ok(ws) = self.file.sheet(si) else { return };
        if ws.cell(coord(at)).is_some() {
            return;
        }
        let style = self.style_at(si, at);
        if style != Style::default()
            && let Ok(ws) = self.file.sheet_mut(si)
        {
            ws.cell_mut(coord(at)).set_style(style);
        }
    }

    fn style_prop(&self, style: &Style, prop: &str) -> Result<Val> {
        let n = Val::Num;
        Ok(match prop {
            "font.name" => Val::Text(
                style
                    .font()
                    .map(|f| f.name().to_string())
                    .filter(|s| !s.is_empty())
                    .unwrap_or_else(|| DEFAULT_FONT.into()),
            ),
            "font.size" => n(style
                .font()
                .map(|f| f.size())
                .filter(|s| *s > 0.0)
                .unwrap_or(11.0)),
            "font.bold" => Val::Bool(style.font().is_some_and(|f| f.bold())),
            "font.italic" => Val::Bool(style.font().is_some_and(|f| f.italic())),
            "font.strikethrough" => Val::Bool(style.font().is_some_and(|f| f.strikethrough())),
            "font.underline" => n(
                match style.font().map(|f| f.underline()).unwrap_or("none") {
                    "single" => 2.0,
                    "double" => -4119.0,
                    "singleAccounting" => 4.0,
                    "doubleAccounting" => 5.0,
                    _ => NONE,
                },
            ),
            "font.superscript" => Val::Bool(style.font().is_some_and(|f| {
                *f.vertical_text_alignment().val() == VerticalAlignmentRunValues::Superscript
            })),
            "font.subscript" => Val::Bool(style.font().is_some_and(|f| {
                *f.vertical_text_alignment().val() == VerticalAlignmentRunValues::Subscript
            })),
            "font.color" => n(bgr(style
                .font()
                .map(|f| f.color())
                .filter(|c| c.has_value())
                .map_or(0, |c| self.colour(c)))),
            "font.colorindex" => n(self.colour_index(style.font().map(|f| f.color()), AUTOMATIC)),
            "interior.color" => n(bgr(self.fill_colour(style).unwrap_or(0xFFFFFF))),
            "interior.colorindex" => n(match self.fill_colour(style) {
                Some(rgb) => PALETTE
                    .iter()
                    .position(|p| *p == rgb)
                    .map_or(NONE, |i| (i + 1) as f64),
                None => NONE,
            }),
            "interior.pattern" => n(
                match style
                    .fill()
                    .and_then(|f| f.pattern_fill())
                    .map(|p| p.pattern_type().value_string().to_string())
                    .as_deref()
                {
                    None | Some("none") => NONE,
                    Some("solid") => 1.0,
                    Some("mediumGray") => -4124.0,
                    Some("darkGray") => -4125.0,
                    Some("lightGray") => -4126.0,
                    Some("gray125") => 17.0,
                    Some("gray0625") => 18.0,
                    Some("darkHorizontal") => -4128.0,
                    Some("darkVertical") => -4166.0,
                    Some("darkDown") => -4121.0,
                    Some("darkUp") => -4162.0,
                    Some("darkGrid") => 9.0,
                    Some("darkTrellis") => 10.0,
                    Some("lightHorizontal") => 11.0,
                    Some("lightVertical") => 12.0,
                    Some("lightDown") => 13.0,
                    Some("lightUp") => 14.0,
                    Some("lightGrid") => 15.0,
                    Some(_) => 16.0,
                },
            ),
            "interior.patterncolor" => n(bgr(style
                .fill()
                .and_then(|f| f.pattern_fill())
                .and_then(|p| p.background_color())
                .filter(|c| c.has_value())
                .map_or(0, |c| self.colour(c)))),
            "horizontalalignment" => n(
                match style
                    .alignment()
                    .map(|a| a.horizontal().value_string().to_string())
                    .as_deref()
                {
                    Some("left") => -4131.0,
                    Some("center") => -4108.0,
                    Some("right") => -4152.0,
                    Some("fill") => 5.0,
                    Some("justify") => -4130.0,
                    Some("centerContinuous") => 7.0,
                    Some("distributed") => -4117.0,
                    _ => 1.0,
                },
            ),
            "verticalalignment" => n(
                match style
                    .alignment()
                    .map(|a| a.vertical().value_string().to_string())
                    .as_deref()
                {
                    Some("top") => -4160.0,
                    Some("center") => -4108.0,
                    Some("justify") => -4130.0,
                    Some("distributed") => -4117.0,
                    // What Japanese Excel gives new workbooks; files say when it is the bottom.
                    None => -4108.0,
                    _ => -4107.0,
                },
            ),
            "wraptext" => Val::Bool(style.alignment().is_some_and(|a| a.wrap_text())),
            "shrinktofit" => Val::Bool(style.alignment().is_some_and(|a| a.shrink_to_fit())),
            "indentlevel" => n(style.alignment().map_or(0.0, |a| a.indent() as f64)),
            "orientation" => n(match style.alignment().map_or(0, |a| a.text_rotation()) {
                0 => -4128.0,
                255 => -4166.0,
                d if d <= 90 => d as f64,
                d => 90.0 - d as f64,
            }),
            "locked" => Val::Bool(style.protection().is_none_or(|p| p.locked())),
            "formulahidden" => {
                Val::Bool(style.protection().cloned().is_some_and(|mut p| p.hidden()))
            }
            _ => return Err(format!("書式 {prop} には対応していません")),
        })
    }

    fn fill_colour(&self, style: &Style) -> Option<u32> {
        let p = style.fill()?.pattern_fill()?;
        if *p.pattern_type() == PatternValues::None {
            return None;
        }
        p.foreground_color()
            .filter(|c| c.has_value())
            .map(|c| self.colour(c))
    }

    fn set_style_prop(style: &mut Style, prop: &str, v: &Val) -> Result<()> {
        let colour = |c: &mut Color, v: &Val| -> Result<()> {
            c.set_argb_str(format!("FF{:06X}", rgb_of(num(v)?)));
            Ok(())
        };
        let indexed = |v: &Val| -> Result<Option<u32>> {
            let i = num(v)?;
            Ok(if (1.0..=56.0).contains(&i) {
                Some(PALETTE[i as usize - 1])
            } else {
                None
            })
        };
        match prop {
            "font.name" => {
                style.font_mut().set_name(text(v));
            }
            "font.size" => {
                style.font_mut().set_size(num(v)?);
            }
            "font.bold" => {
                style.font_mut().set_bold(flag(v)?);
            }
            "font.italic" => {
                style.font_mut().set_italic(flag(v)?);
            }
            "font.strikethrough" => {
                style.font_mut().set_strikethrough(flag(v)?);
            }
            "font.underline" => {
                let u = match v {
                    Val::Bool(true) => UnderlineValues::Single,
                    Val::Bool(false) => UnderlineValues::None,
                    other => match num(other)? as i32 {
                        2 => UnderlineValues::Single,
                        -4119 => UnderlineValues::Double,
                        4 => UnderlineValues::SingleAccounting,
                        5 => UnderlineValues::DoubleAccounting,
                        _ => UnderlineValues::None,
                    },
                };
                style.font_mut().set_underline(u.value_string());
            }
            "font.superscript" | "font.subscript" => {
                let on = if prop == "font.superscript" {
                    VerticalAlignmentRunValues::Superscript
                } else {
                    VerticalAlignmentRunValues::Subscript
                };
                let value = if flag(v)? {
                    on
                } else {
                    VerticalAlignmentRunValues::Baseline
                };
                style
                    .font_mut()
                    .vertical_text_alignment_mut()
                    .set_val(value);
            }
            "font.color" => colour(style.font_mut().color_mut(), v)?,
            "font.colorindex" => {
                let rgb = indexed(v)?.unwrap_or(0);
                style
                    .font_mut()
                    .color_mut()
                    .set_argb_str(format!("FF{rgb:06X}"));
            }
            "interior.color" => {
                let p = style.fill_mut().pattern_fill_mut();
                if matches!(*p.pattern_type(), PatternValues::None) {
                    p.set_pattern_type(PatternValues::Solid);
                }
                colour(p.foreground_color_mut(), v)?;
            }
            "interior.colorindex" => match indexed(v)? {
                Some(rgb) => {
                    let p = style.fill_mut().pattern_fill_mut();
                    if matches!(*p.pattern_type(), PatternValues::None) {
                        p.set_pattern_type(PatternValues::Solid);
                    }
                    p.foreground_color_mut()
                        .set_argb_str(format!("FF{rgb:06X}"));
                }
                None => {
                    style.remove_fill();
                }
            },
            "interior.pattern" => {
                let pattern = match num(v)? as i32 {
                    -4142 => None,
                    1 => Some(PatternValues::Solid),
                    -4124 => Some(PatternValues::MediumGray),
                    -4125 => Some(PatternValues::DarkGray),
                    -4126 => Some(PatternValues::LightGray),
                    17 => Some(PatternValues::Gray125),
                    18 => Some(PatternValues::Gray0625),
                    -4128 => Some(PatternValues::DarkHorizontal),
                    -4166 => Some(PatternValues::DarkVertical),
                    -4121 => Some(PatternValues::DarkDown),
                    -4162 => Some(PatternValues::DarkUp),
                    9 => Some(PatternValues::DarkGrid),
                    10 => Some(PatternValues::DarkTrellis),
                    11 => Some(PatternValues::LightHorizontal),
                    12 => Some(PatternValues::LightVertical),
                    13 => Some(PatternValues::LightDown),
                    14 => Some(PatternValues::LightUp),
                    15 => Some(PatternValues::LightGrid),
                    16 => Some(PatternValues::LightTrellis),
                    _ => Some(PatternValues::Solid),
                };
                match pattern {
                    Some(p) => {
                        style.fill_mut().pattern_fill_mut().set_pattern_type(p);
                    }
                    None => {
                        style.remove_fill();
                    }
                }
            }
            "interior.patterncolor" => colour(
                style.fill_mut().pattern_fill_mut().background_color_mut(),
                v,
            )?,
            "horizontalalignment" => {
                use HorizontalAlignmentValues as H;
                style.alignment_mut().set_horizontal(match num(v)? as i32 {
                    -4131 => H::Left,
                    -4108 => H::Center,
                    -4152 => H::Right,
                    5 => H::Fill,
                    -4130 => H::Justify,
                    7 => H::CenterContinuous,
                    -4117 => H::Distributed,
                    _ => H::General,
                });
            }
            "verticalalignment" => {
                use VerticalAlignmentValues as V;
                style.alignment_mut().set_vertical(match num(v)? as i32 {
                    -4160 => V::Top,
                    -4108 => V::Center,
                    -4130 => V::Justify,
                    -4117 => V::Distributed,
                    _ => V::Bottom,
                });
            }
            "wraptext" => style.alignment_mut().set_wrap_text(flag(v)?),
            "shrinktofit" => style.alignment_mut().set_shrink_to_fit(flag(v)?),
            "indentlevel" => style
                .alignment_mut()
                .set_indent(num(v)?.clamp(0.0, 250.0) as u32),
            "orientation" => {
                let d = num(v)?;
                style.alignment_mut().set_text_rotation(match d as i32 {
                    -4128 => 0,
                    -4166 => 255,
                    -4171 => 90,
                    -4170 => 180,
                    x if (0..=90).contains(&x) => x as u32,
                    x if (-90..0).contains(&x) => (90 - x) as u32,
                    _ => return Err("Orientation の値が正しくありません。".into()),
                });
            }
            "locked" => style.protection_mut().set_locked(flag(v)?),
            "formulahidden" => style.protection_mut().set_hidden(flag(v)?),
            _ => return Err(format!("書式 {prop} には対応していません")),
        }
        Ok(())
    }

    /// The sides of the cells of `r` a `Borders(index)` is: (cells, side) pairs. `index` is
    /// `XlBordersIndex`, or the side of every cell (`xlLeft` and the like), or `all`.
    fn border_parts(r: RangeRef, index: &str) -> Result<Vec<(RangeRef, Side)>> {
        let row =
            |row: u32| RangeRef::new(CellRef::new(row, r.start.col), CellRef::new(row, r.end.col));
        let col =
            |col: u32| RangeRef::new(CellRef::new(r.start.row, col), CellRef::new(r.end.row, col));
        Ok(match index {
            "all" => vec![
                (r, Side::Left),
                (r, Side::Right),
                (r, Side::Top),
                (r, Side::Bottom),
            ],
            "7" => vec![(col(r.start.col), Side::Left)],
            "10" => vec![(col(r.end.col), Side::Right)],
            "8" => vec![(row(r.start.row), Side::Top)],
            "9" => vec![(row(r.end.row), Side::Bottom)],
            "11" if r.width() > 1 => {
                vec![
                    (
                        RangeRef::new(CellRef::new(r.start.row, r.start.col + 1), r.end),
                        Side::Left,
                    ),
                    (
                        RangeRef::new(r.start, CellRef::new(r.end.row, r.end.col - 1)),
                        Side::Right,
                    ),
                ]
            }
            "12" if r.height() > 1 => {
                vec![
                    (
                        RangeRef::new(CellRef::new(r.start.row + 1, r.start.col), r.end),
                        Side::Top,
                    ),
                    (
                        RangeRef::new(r.start, CellRef::new(r.end.row - 1, r.end.col)),
                        Side::Bottom,
                    ),
                ]
            }
            "11" | "12" => Vec::new(),
            "5" => vec![(r, Side::DiagonalDown)],
            "6" => vec![(r, Side::DiagonalUp)],
            "1" | "-4131" => vec![(r, Side::Left)],
            "2" | "-4152" => vec![(r, Side::Right)],
            "3" | "-4160" => vec![(r, Side::Top)],
            "4" | "-4107" => vec![(r, Side::Bottom)],
            _ => return Err("Borders のインデックスが正しくありません。".into()),
        })
    }

    fn border_prop(&self, b: Option<&Border>, what: &str) -> Result<Val> {
        let (line, weight) = b.map_or((NONE, 2.0), |b| line_of(b.style()));
        Ok(Val::Num(match what {
            "linestyle" => line,
            "weight" => weight,
            "color" => bgr(b
                .and_then(|b| b.color())
                .filter(|c| c.has_value())
                .map_or(0, |c| self.colour(&c))),
            "colorindex" => match b {
                Some(b) if *b.style() != BorderStyleValues::None => {
                    self.colour_index(b.color().as_ref(), AUTOMATIC)
                }
                _ => NONE,
            },
            _ => return Err(format!("罫線の {what} には対応していません")),
        }))
    }

    fn set_border_prop(b: &mut Border, what: &str, v: &Val) -> Result<()> {
        let (line, weight) = line_of(b.style());
        let rgb = |rgb: u32| {
            let mut c = Color::default();
            c.set_argb_str(format!("FF{rgb:06X}"));
            c
        };
        match what {
            "linestyle" => {
                b.set_style(style_of_line(num(v)?, weight));
            }
            "weight" => {
                b.set_style(style_of_line(
                    if line == NONE { 1.0 } else { line },
                    num(v)?,
                ));
            }
            "color" => {
                if line == NONE {
                    b.set_style(BorderStyleValues::Thin);
                }
                b.set_color(rgb(rgb_of(num(v)?)));
            }
            "colorindex" => {
                let i = num(v)?;
                if i == NONE {
                    b.set_style(BorderStyleValues::None);
                } else {
                    if line == NONE {
                        b.set_style(BorderStyleValues::Thin);
                    }
                    b.set_color(rgb(if (1.0..=56.0).contains(&i) {
                        PALETTE[i as usize - 1]
                    } else {
                        0
                    }));
                }
            }
            _ => return Err(format!("罫線の {what} には対応していません")),
        }
        Ok(())
    }

    /// Reads a format of the cells of `r`: None when they differ (`Null` in VBA).
    pub fn format_of(&mut self, si: usize, r: RangeRef, prop: &str) -> Result<Option<Val>> {
        self.check_sheet(si)?;
        let mut values: Vec<Val> = Vec::new();
        match prop {
            "columnwidth" => {
                let ws = self.file.sheet(si).map_err(|e| e.to_string())?;
                let default = Some(ws.sheet_format_properties().default_column_width())
                    .filter(|w| *w > 0.0)
                    .unwrap_or(DEFAULT_COLUMN_WIDTH);
                let stored: Vec<f64> = ws
                    .column_dimensions()
                    .iter()
                    .filter(|c| (r.start.col + 1..=r.end.col + 1).contains(&c.col_num()))
                    .map(|c| c.width())
                    .collect();
                if (stored.len() as u32) < r.width() {
                    values.push(Val::Num(default));
                }
                values.extend(stored.into_iter().map(Val::Num));
            }
            "rowheight" => {
                let ws = self.file.sheet(si).map_err(|e| e.to_string())?;
                let default = Some(ws.sheet_format_properties().default_row_height())
                    .filter(|h| *h > 0.0)
                    .unwrap_or(DEFAULT_ROW_HEIGHT);
                let stored: Vec<f64> = ws
                    .row_dimensions()
                    .iter()
                    .filter(|x| (r.start.row + 1..=r.end.row + 1).contains(&x.row_num()))
                    .map(|x| {
                        if x.height() > 0.0 {
                            x.height()
                        } else {
                            default
                        }
                    })
                    .collect();
                if (stored.len() as u32) < r.height() {
                    values.push(Val::Num(default));
                }
                values.extend(stored.into_iter().map(Val::Num));
            }
            "mergecells" => {
                let sh = self.model.sheet(si).expect("checked");
                return Ok(if sh.merges.iter().any(|m| m.contains_range(&r)) {
                    Some(Val::Bool(true))
                } else if sh.merges.iter().any(|m| m.intersects(&r)) {
                    None
                } else {
                    Some(Val::Bool(false))
                });
            }
            _ => {
                if let Some(rest) = prop.strip_prefix("borders.") {
                    let (index, what) =
                        rest.split_once('.').ok_or("罫線の指定が正しくありません")?;
                    for (cells, side) in Self::border_parts(r, index)? {
                        for style in self.styles_of(si, cells) {
                            values.push(self.border_prop(border(&style, side), what)?);
                        }
                    }
                    if values.is_empty() {
                        values.push(self.border_prop(None, what)?);
                    }
                } else {
                    for style in self.styles_of(si, r) {
                        values.push(self.style_prop(&style, prop)?);
                    }
                }
            }
        }
        let first = values.first().cloned();
        Ok(first.filter(|f| values.iter().all(|v| v == f)))
    }

    /// Sets a format of the cells of `r`.
    pub fn set_format(&mut self, si: usize, r: RangeRef, prop: &str, v: &Val) -> Result<()> {
        self.check_sheet(si)?;
        match prop {
            "columnwidth" => {
                let w = num(v)?;
                if !(0.0..=255.0).contains(&w) {
                    return Err("Range クラスの ColumnWidth プロパティを設定できません。".into());
                }
                let ws = self.file.sheet_mut(si).map_err(|e| e.to_string())?;
                if r.width() == MAX_COLS {
                    ws.sheet_format_properties_mut().set_default_column_width(w);
                    let stored: Vec<u32> =
                        ws.column_dimensions().iter().map(|c| c.col_num()).collect();
                    for col in stored {
                        ws.column_dimension_by_number_mut(col).set_width(w);
                    }
                } else {
                    for col in r.start.col..=r.end.col {
                        ws.column_dimension_by_number_mut(col + 1).set_width(w);
                    }
                }
            }
            "rowheight" => {
                let h = num(v)?;
                if !(0.0..=409.5).contains(&h) {
                    return Err("Range クラスの RowHeight プロパティを設定できません。".into());
                }
                let ws = self.file.sheet_mut(si).map_err(|e| e.to_string())?;
                if r.height() == MAX_ROWS {
                    ws.sheet_format_properties_mut().set_default_row_height(h);
                    let stored: Vec<u32> =
                        ws.row_dimensions().iter().map(|x| x.row_num()).collect();
                    for row in stored {
                        ws.row_dimension_mut(row)
                            .set_height(h)
                            .set_custom_height(true);
                    }
                } else {
                    for row in r.start.row..=r.end.row {
                        ws.row_dimension_mut(row + 1)
                            .set_height(h)
                            .set_custom_height(true);
                    }
                }
            }
            "mergecells" => {
                if flag(v)? {
                    self.merge(si, r);
                } else {
                    self.unmerge(si, r);
                }
            }
            _ => {
                if let Some(rest) = prop.strip_prefix("borders.") {
                    let (index, what) =
                        rest.split_once('.').ok_or("罫線の指定が正しくありません")?;
                    let what = what.to_string();
                    for (cells, side) in Self::border_parts(r, index)? {
                        // Checked on a scratch border first: `restyle` cannot fail half way.
                        Self::set_border_prop(&mut Border::default(), &what, v)?;
                        self.restyle(si, cells, &|style| {
                            let _ = Self::set_border_prop(border_mut(style, side), &what, v);
                        })?;
                    }
                } else {
                    Self::set_style_prop(&mut Style::default(), prop, v)?;
                    self.restyle(si, r, &|style| {
                        let _ = Self::set_style_prop(style, prop, v);
                    })?;
                }
            }
        }
        Ok(())
    }

    /// `Range.ClearFormats`: the cells of `r` get the default format, and are no longer merged.
    pub fn clear_formats(&mut self, si: usize, r: RangeRef) -> Result<()> {
        self.check_sheet(si)?;
        self.reset_styles(si, r);
        self.unmerge(si, r);
        if let Ok(ws) = self.file.sheet_mut(si) {
            if r.is_full_cols() {
                for col in r.start.col..=r.end.col {
                    if ws.column_dimension_by_number(col + 1).is_some() {
                        ws.column_dimension_by_number_mut(col + 1)
                            .set_style(Style::default());
                    }
                }
            } else if r.is_full_rows() {
                for row in r.start.row..=r.end.row {
                    if ws.row_dimension(row + 1).is_some() {
                        ws.row_dimension_mut(row + 1).set_style(Style::default());
                    }
                }
            }
        }
        Ok(())
    }

    /// `Range.MergeArea`: the merged cell a cell is part of (the cell itself when it is not).
    pub fn merge_area(&self, si: usize, at: CellRef) -> RangeRef {
        self.model
            .sheet(si)
            .and_then(|s| s.merge_at(at))
            .unwrap_or(RangeRef::cell(at))
    }

    /// `Range.Merge`: `across` merges each row of `r` on its own.
    pub fn merge_cells(&mut self, si: usize, r: RangeRef, across: bool) -> Result<()> {
        self.check_sheet(si)?;
        if across {
            for row in r.start.row..=r.end.row {
                self.merge(
                    si,
                    RangeRef::new(CellRef::new(row, r.start.col), CellRef::new(row, r.end.col)),
                );
            }
        } else {
            self.merge(si, r);
        }
        Ok(())
    }

    /// `Worksheet.Tab.Color` as VBA's colour number, None when the tab has no colour.
    pub fn tab_color(&self, si: usize) -> Option<f64> {
        let ws = self.file.sheet(si).ok()?;
        ws.tab_color()
            .filter(|c| c.has_value())
            .map(|c| bgr(self.colour(c)))
    }

    pub fn set_tab_color(&mut self, si: usize, color: Option<f64>) {
        let Ok(ws) = self.file.sheet_mut(si) else {
            return;
        };
        match color {
            Some(c) => {
                ws.tab_color_mut()
                    .set_argb_str(format!("FF{:06X}", rgb_of(c)));
            }
            None => {
                ws.remove_tab_color();
            }
        }
    }
}
