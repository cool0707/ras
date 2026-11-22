//! EMF pictures as SVG, for the grid: what Excel stores for a range copied as a picture, a linked
//! picture, the preview of an embedded object. Browsers do not draw EMF.
//!
//! The GDI records are played (MS-EMF): lines, polygons, curves, rectangles and ellipses, paths,
//! text (with the advances of its characters), bitmaps (as PNG), with the pens, brushes and fonts
//! selected, the window, viewport and world transforms, the saved states and clipping
//! rectangles. EMF+ records (in comments) are skipped: a "dual" EMF has the same drawing in GDI
//! records too. What is not drawn: hatched and pattern brushes (filled with their color), regions
//! other than rectangles, raster operations other than copying.

use std::collections::HashMap;
use std::fmt::Write as _;

use super::drawing::{esc, font_family, n};

#[derive(Clone, Debug)]
struct Pen {
    width: f64,
    color: String,
    dash: u32,
    null: bool,
    /// A cosmetic pen: one pixel wide, whatever the transform.
    cosmetic: bool,
}

#[derive(Clone, Debug)]
struct Brush {
    color: String,
    null: bool,
}

#[derive(Clone, Debug)]
struct Font {
    height: f64,
    weight: i32,
    italic: bool,
    underline: bool,
    strike: bool,
    escapement: f64,
    face: String,
}

#[derive(Clone, Debug)]
enum Obj {
    Pen(Pen),
    Brush(Brush),
    Font(Font),
}

/// A clipping area in device units: a rectangle, or a path.
#[derive(Clone, Debug, PartialEq)]
enum Clip {
    Rect(f64, f64, f64, f64),
    Path(String),
}

#[derive(Clone, Debug)]
struct Dc {
    window_org: (f64, f64),
    window_ext: (f64, f64),
    viewport_org: (f64, f64),
    viewport_ext: (f64, f64),
    map_mode: u32,
    /// The world transform: x' = a x + c y + e, y' = b x + d y + f.
    world: [f64; 6],
    pen: Pen,
    brush: Brush,
    font: Font,
    text_color: String,
    bk_color: String,
    opaque_bk: bool,
    text_align: u32,
    winding: bool,
    cur: (f64, f64),
    clip: Option<Clip>,
}

impl Default for Dc {
    fn default() -> Dc {
        Dc {
            window_org: (0.0, 0.0),
            window_ext: (1.0, 1.0),
            viewport_org: (0.0, 0.0),
            viewport_ext: (1.0, 1.0),
            map_mode: 1,
            world: [1.0, 0.0, 0.0, 1.0, 0.0, 0.0],
            pen: Pen {
                width: 1.0,
                color: "#000000".into(),
                dash: 0,
                null: false,
                cosmetic: true,
            },
            brush: Brush {
                color: "#ffffff".into(),
                null: false,
            },
            font: Font {
                height: 12.0,
                weight: 400,
                italic: false,
                underline: false,
                strike: false,
                escapement: 0.0,
                face: "Arial".into(),
            },
            text_color: "#000000".into(),
            bk_color: "#ffffff".into(),
            opaque_bk: true,
            text_align: 0,
            winding: false,
            cur: (0.0, 0.0),
            clip: None,
        }
    }
}

impl Dc {
    /// Logical units to device units: the world transform, then the window to the viewport.
    fn matrix(&self) -> [f64; 6] {
        let (sx, sy) = if matches!(self.map_mode, 7 | 8)
            && self.window_ext.0 != 0.0
            && self.window_ext.1 != 0.0
        {
            (
                self.viewport_ext.0 / self.window_ext.0,
                self.viewport_ext.1 / self.window_ext.1,
            )
        } else {
            (1.0, 1.0)
        };
        let [a, b, c, d, e, f] = self.world;
        [
            sx * a,
            sy * b,
            sx * c,
            sy * d,
            sx * (e - self.window_org.0) + self.viewport_org.0,
            sy * (f - self.window_org.1) + self.viewport_org.1,
        ]
    }

    fn at(&self, p: (f64, f64)) -> (f64, f64) {
        let [a, b, c, d, e, f] = self.matrix();
        (a * p.0 + c * p.1 + e, b * p.0 + d * p.1 + f)
    }

    /// How much a length grows from logical to device units.
    fn scale(&self) -> f64 {
        let [a, b, c, d, ..] = self.matrix();
        (a * d - b * c).abs().sqrt()
    }
}

fn colorref(v: u32) -> String {
    format!(
        "#{:02x}{:02x}{:02x}",
        v & 0xff,
        (v >> 8) & 0xff,
        (v >> 16) & 0xff
    )
}

/// The stock objects (GetStockObject).
fn stock(i: u32) -> Option<Obj> {
    let brush = |c: &str| {
        Some(Obj::Brush(Brush {
            color: c.into(),
            null: false,
        }))
    };
    let pen = |c: &str, null: bool| {
        Some(Obj::Pen(Pen {
            width: 1.0,
            color: c.into(),
            dash: 0,
            null,
            cosmetic: true,
        }))
    };
    match i & 0x7fff_ffff {
        0 => brush("#ffffff"),
        1 => brush("#c0c0c0"),
        2 => brush("#808080"),
        3 => brush("#404040"),
        4 => brush("#000000"),
        5 => Some(Obj::Brush(Brush {
            color: "#000000".into(),
            null: true,
        })),
        6 => pen("#ffffff", false),
        7 => pen("#000000", false),
        8 => pen("#000000", true),
        10..=17 => Some(Obj::Font(Dc::default().font)),
        18 => brush("#ffffff"),
        19 => pen("#000000", false),
        _ => None,
    }
}

struct Rd<'a>(&'a [u8]);

impl Rd<'_> {
    fn u32(&self, at: usize) -> u32 {
        self.0
            .get(at..at + 4)
            .map(|b| u32::from_le_bytes([b[0], b[1], b[2], b[3]]))
            .unwrap_or(0)
    }
    fn i32(&self, at: usize) -> f64 {
        self.u32(at) as i32 as f64
    }
    fn i16(&self, at: usize) -> f64 {
        self.0
            .get(at..at + 2)
            .map(|b| i16::from_le_bytes([b[0], b[1]]) as f64)
            .unwrap_or(0.0)
    }
    fn f32(&self, at: usize) -> f64 {
        f32::from_bits(self.u32(at)) as f64
    }
    /// `count` points from `at`, of 16 or 32 bits.
    fn points(&self, at: usize, count: usize, short: bool) -> Vec<(f64, f64)> {
        let size = if short { 4 } else { 8 };
        let count = count.min(self.0.len().saturating_sub(at) / size);
        (0..count)
            .map(|i| {
                let p = at + i * size;
                if short {
                    (self.i16(p), self.i16(p + 2))
                } else {
                    (self.i32(p), self.i32(p + 4))
                }
            })
            .collect()
    }
}

struct Player {
    dc: Dc,
    saved: Vec<Dc>,
    objects: HashMap<u32, Obj>,
    out: String,
    defs: String,
    clips: Vec<Clip>,
    /// The path being made (BEGINPATH .. ENDPATH), in device units.
    path: Option<String>,
    last_path: String,
    drawn: usize,
    /// The mask of an icon (a bitmap copied with SRCAND), for the colors that follow it.
    mask: Option<((f64, f64, f64, f64), Dib)>,
}

const KAPPA: f64 = 0.552_284_75;

impl Player {
    fn xy(p: (f64, f64)) -> String {
        format!("{},{}", n(p.0), n(p.1))
    }

    /// The attribute that clips to the current clipping area.
    fn clip_attr(&mut self) -> String {
        let Some(clip) = self.dc.clip.clone() else {
            return String::new();
        };
        let i = match self.clips.iter().position(|c| *c == clip) {
            Some(i) => i,
            None => {
                let shape = match &clip {
                    Clip::Rect(x, y, w, h) => format!(
                        "<rect x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\"/>",
                        n(*x),
                        n(*y),
                        n(w.max(0.0)),
                        n(h.max(0.0))
                    ),
                    Clip::Path(d) => format!("<path d=\"{d}\"/>"),
                };
                let _ = write!(
                    self.defs,
                    "<clipPath id=\"c{}\">{shape}</clipPath>",
                    self.clips.len()
                );
                self.clips.push(clip);
                self.clips.len() - 1
            }
        };
        format!(" clip-path=\"url(#c{i})\"")
    }

    fn stroke_attrs(&self) -> String {
        let p = &self.dc.pen;
        if p.null {
            return " stroke=\"none\"".into();
        }
        let w = if p.cosmetic {
            1.0
        } else {
            (p.width * self.dc.scale()).max(1.0)
        };
        let dash = match p.dash {
            1 => format!(" stroke-dasharray=\"{} {}\"", n(w * 6.0), n(w * 2.0)),
            2 => format!(" stroke-dasharray=\"{} {}\"", n(w), n(w * 2.0)),
            3 => format!(
                " stroke-dasharray=\"{} {} {} {}\"",
                n(w * 6.0),
                n(w * 2.0),
                n(w),
                n(w * 2.0)
            ),
            4 => format!(
                " stroke-dasharray=\"{} {} {} {} {} {}\"",
                n(w * 6.0),
                n(w * 2.0),
                n(w),
                n(w * 2.0),
                n(w),
                n(w * 2.0)
            ),
            _ => String::new(),
        };
        format!(
            " stroke=\"{}\" stroke-width=\"{}\"{dash} stroke-linejoin=\"round\"",
            p.color,
            n(w)
        )
    }

    fn fill_attrs(&self) -> String {
        let b = &self.dc.brush;
        if b.null {
            " fill=\"none\"".into()
        } else {
            format!(
                " fill=\"{}\"{}",
                b.color,
                if self.dc.winding {
                    ""
                } else {
                    " fill-rule=\"evenodd\""
                }
            )
        }
    }

    /// Draws (or adds to the path being made) the device path `d`.
    fn shape(&mut self, d: String, fill: bool, stroke: bool) {
        if d.is_empty() {
            return;
        }
        if let Some(path) = &mut self.path {
            path.push_str(&d);
            path.push(' ');
            return;
        }
        let fill = if fill {
            self.fill_attrs()
        } else {
            " fill=\"none\"".into()
        };
        let stroke = if stroke {
            self.stroke_attrs()
        } else {
            " stroke=\"none\"".into()
        };
        if fill.contains("none") && stroke.contains("none") {
            return;
        }
        let clip = self.clip_attr();
        let _ = write!(self.out, "<path d=\"{d}\"{fill}{stroke}{clip}/>");
        self.drawn += 1;
    }

    fn poly(&self, pts: &[(f64, f64)], close: bool) -> String {
        let mut d = String::new();
        for (i, p) in pts.iter().enumerate() {
            let _ = write!(
                d,
                "{}{} ",
                if i == 0 { "M" } else { "L" },
                Self::xy(self.dc.at(*p))
            );
        }
        if close && !d.is_empty() {
            d.push('Z');
        }
        d.trim_end().to_string()
    }

    /// Bézier curves from `start` (the first point given, or the current position).
    fn bezier(&self, start: Option<(f64, f64)>, pts: &[(f64, f64)]) -> String {
        let mut d = format!("M{} ", Self::xy(self.dc.at(start.unwrap_or(self.dc.cur))));
        for c in pts.chunks(3).filter(|c| c.len() == 3) {
            let _ = write!(
                d,
                "C{} {} {} ",
                Self::xy(self.dc.at(c[0])),
                Self::xy(self.dc.at(c[1])),
                Self::xy(self.dc.at(c[2]))
            );
        }
        d.trim_end().to_string()
    }

    fn rect(&self, l: f64, t: f64, r: f64, b: f64) -> String {
        self.poly(&[(l, t), (r, t), (r, b), (l, b)], true)
    }

    /// An ellipse in the box, as four curves.
    fn ellipse(&self, l: f64, t: f64, r: f64, b: f64) -> String {
        let (cx, cy, rx, ry) = (
            (l + r) / 2.0,
            (t + b) / 2.0,
            (r - l).abs() / 2.0,
            (b - t).abs() / 2.0,
        );
        let (kx, ky) = (rx * KAPPA, ry * KAPPA);
        let p = |x: f64, y: f64| Self::xy(self.dc.at((x, y)));
        format!(
            "M{} C{} {} {} C{} {} {} C{} {} {} C{} {} {} Z",
            p(cx + rx, cy),
            p(cx + rx, cy + ky),
            p(cx + kx, cy + ry),
            p(cx, cy + ry),
            p(cx - kx, cy + ry),
            p(cx - rx, cy + ky),
            p(cx - rx, cy),
            p(cx - rx, cy - ky),
            p(cx - kx, cy - ry),
            p(cx, cy - ry),
            p(cx + kx, cy - ry),
            p(cx + rx, cy - ky),
            p(cx + rx, cy)
        )
    }

    /// The points of an arc of the ellipse in the box, counterclockwise from the radial through
    /// `start` to the one through `end` (as GDI draws them).
    fn arc_points(
        &self,
        b: (f64, f64, f64, f64),
        start: (f64, f64),
        end: (f64, f64),
    ) -> Vec<(f64, f64)> {
        let (cx, cy, rx, ry) = (
            (b.0 + b.2) / 2.0,
            (b.1 + b.3) / 2.0,
            (b.2 - b.0).abs() / 2.0,
            (b.3 - b.1).abs() / 2.0,
        );
        let a1 = (-(start.1 - cy)).atan2(start.0 - cx);
        let mut a2 = (-(end.1 - cy)).atan2(end.0 - cx);
        if a2 <= a1 {
            a2 += std::f64::consts::TAU;
        }
        (0..=48)
            .map(|i| a1 + (a2 - a1) * i as f64 / 48.0)
            .map(|a| (cx + rx * a.cos(), cy - ry * a.sin()))
            .collect()
    }

    fn select(&mut self, id: u32) {
        let obj = if id & 0x8000_0000 != 0 {
            stock(id)
        } else {
            self.objects.get(&id).cloned()
        };
        match obj {
            Some(Obj::Pen(p)) => self.dc.pen = p,
            Some(Obj::Brush(b)) => self.dc.brush = b,
            Some(Obj::Font(f)) => self.dc.font = f,
            None => {}
        }
    }

    fn intersect_clip(&mut self, c: (f64, f64, f64, f64)) {
        let next = match self.dc.clip.clone() {
            Some(Clip::Rect(x, y, w, h)) => {
                let (l, t) = (x.max(c.0), y.max(c.1));
                let (r, b) = ((x + w).min(c.0 + c.2), (y + h).min(c.1 + c.3));
                Clip::Rect(l, t, (r - l).max(0.0), (b - t).max(0.0))
            }
            _ => Clip::Rect(c.0, c.1, c.2, c.3),
        };
        self.dc.clip = Some(next);
    }

    /// A logical rectangle as a device one: (x, y, width, height).
    fn device_rect(&self, l: f64, t: f64, r: f64, b: f64) -> (f64, f64, f64, f64) {
        let pts = [
            self.dc.at((l, t)),
            self.dc.at((r, t)),
            self.dc.at((r, b)),
            self.dc.at((l, b)),
        ];
        let (x0, x1) = pts
            .iter()
            .fold((f64::MAX, f64::MIN), |(a, z), p| (a.min(p.0), z.max(p.0)));
        let (y0, y1) = pts
            .iter()
            .fold((f64::MAX, f64::MIN), |(a, z), p| (a.min(p.1), z.max(p.1)));
        (x0, y0, x1 - x0, y1 - y0)
    }

    fn text(&mut self, rd: &Rd, at: usize, wide: bool) {
        // EMREXTTEXTOUT: bounds, graphics mode, scales, then EMRTEXT.
        let t = at + 8 + 16 + 12;
        let reference = (rd.i32(t), rd.i32(t + 4));
        let count = (rd.u32(t + 8) as usize).min(8192);
        let off = rd.u32(t + 12) as usize;
        let options = rd.u32(t + 16);
        let rect = (
            rd.i32(t + 20),
            rd.i32(t + 24),
            rd.i32(t + 28),
            rd.i32(t + 32),
        );
        let off_dx = rd.u32(t + 36) as usize;
        let Some(raw) =
            rd.0.get(at + off..at + off + count * if wide { 2 } else { 1 })
        else {
            return;
        };
        let text: String = if wide {
            String::from_utf16_lossy(
                &raw.chunks_exact(2)
                    .map(|c| u16::from_le_bytes([c[0], c[1]]))
                    .collect::<Vec<_>>(),
            )
        } else {
            raw.iter().map(|b| *b as char).collect()
        };
        // ETO_OPAQUE: the background of the text's rectangle.
        if options & 0x2 != 0 && rect.2 > rect.0 && rect.3 > rect.1 && self.path.is_none() {
            let d = self.rect(rect.0, rect.1, rect.2, rect.3);
            let clip = self.clip_attr();
            let _ = write!(
                self.out,
                "<path d=\"{d}\" fill=\"{}\"{clip}/>",
                self.dc.bk_color
            );
        }
        let text = text.trim_end_matches('\0');
        if text.trim().is_empty() {
            return;
        }
        let f = self.dc.font.clone();
        let origin = if self.dc.text_align & 1 != 0 {
            self.dc.cur
        } else {
            reference
        };
        let p = self.dc.at(origin);
        let scale = self.dc.scale();
        // A negative height is the characters' (em) height; a positive one the cell's.
        let size = (if f.height < 0.0 {
            -f.height
        } else {
            f.height * 0.82
        })
        .max(1.0)
            * scale;
        // The advances of the characters, when given: each character where GDI put it.
        let mut xs = String::new();
        let chars: Vec<char> = text.chars().collect();
        if off_dx > 0 && f.escapement == 0.0 {
            let step = if options & 0x2000 != 0 { 8 } else { 4 };
            let mut x = origin.0;
            let mut list = Vec::new();
            let mut units = 0usize;
            for ch in &chars {
                list.push(n(self.dc.at((x, origin.1)).0));
                // A character outside the BMP has two advances (one per UTF-16 unit).
                for _ in 0..ch.len_utf16() {
                    x += rd.i32(at + off_dx + units * step);
                    units += 1;
                }
            }
            if list.len() > 1 {
                xs = list.join(" ");
            }
        }
        let anchor = match self.dc.text_align & 6 {
            6 => " text-anchor=\"middle\"",
            2 => " text-anchor=\"end\"",
            _ => "",
        };
        let baseline = match self.dc.text_align & 24 {
            24 => "",
            8 => " dominant-baseline=\"text-after-edge\"",
            _ => " dominant-baseline=\"text-before-edge\"",
        };
        let mut style = format!(
            "font-family:{};font-size:{}px;",
            esc(&font_family(&f.face)),
            n(size)
        );
        if f.weight > 550 {
            style.push_str("font-weight:bold;");
        }
        if f.italic {
            style.push_str("font-style:italic;");
        }
        match (f.underline, f.strike) {
            (true, true) => style.push_str("text-decoration:underline line-through;"),
            (true, false) => style.push_str("text-decoration:underline;"),
            (false, true) => style.push_str("text-decoration:line-through;"),
            _ => {}
        }
        let rotate = if f.escapement != 0.0 {
            format!(
                " transform=\"rotate({} {} {})\"",
                n(-f.escapement / 10.0),
                n(p.0),
                n(p.1)
            )
        } else {
            String::new()
        };
        let x = if xs.is_empty() { n(p.0) } else { xs };
        let clip = self.clip_attr();
        let _ = write!(
            self.out,
            "<text x=\"{x}\" y=\"{}\" fill=\"{}\" style=\"{style}white-space:pre\"{anchor}{baseline}{rotate}{clip}>{}</text>",
            n(p.1),
            self.dc.text_color,
            esc(text)
        );
        self.drawn += 1;
    }

    /// A bitmap of a BITBLT, STRETCHBLT, STRETCHDIBITS, ALPHABLEND record, in the logical box,
    /// copied with the raster operation `rop`.
    fn bitmap(
        &mut self,
        rd: &Rd,
        at: usize,
        (off_bmi, cb_bmi, off_bits, cb_bits): (usize, usize, usize, usize),
        dest: (f64, f64, f64, f64),
        rop: u32,
    ) {
        if cb_bmi == 0 {
            // No bitmap: a fill with the brush (PATCOPY), white (WHITENESS) or black (BLACKNESS).
            let d = self.rect(dest.0, dest.1, dest.0 + dest.2, dest.1 + dest.3);
            let brush = self.dc.brush.clone();
            match rop {
                0x00FF_0062 => {
                    self.dc.brush = Brush {
                        color: "#ffffff".into(),
                        null: false,
                    }
                }
                0x0000_0042 => {
                    self.dc.brush = Brush {
                        color: "#000000".into(),
                        null: false,
                    }
                }
                _ => {}
            }
            self.shape(d, true, false);
            self.dc.brush = brush;
            return;
        }
        let (Some(bmi), Some(bits)) = (
            rd.0.get(at + off_bmi..at + off_bmi + cb_bmi),
            rd.0.get(at + off_bits..at + off_bits + cb_bits),
        ) else {
            return;
        };
        let Some(mut image) = dib(bmi, bits) else {
            return;
        };
        // An icon: its mask (SRCAND), then its colors (SRCINVERT) where the mask is black.
        if rop == 0x0088_00C6 {
            self.mask = Some((dest, image));
            return;
        }
        if let Some((d, Dib::Pixels(mw, mh, mask))) = self.mask.take()
            && rop == 0x0066_0046
            && d == dest
            && let Dib::Pixels(w, h, rgba) = &mut image
        {
            // (The mask may be of another size: stretched to the colors, as GDI stretches both.)
            let (w, h) = (*w, *h);
            for y in 0..h {
                for x in 0..w {
                    let m = ((y * mh / h) * mw + x * mw / w) * 4;
                    if mask.get(m).is_some_and(|v| *v > 127) {
                        rgba[(y * w + x) * 4 + 3] = 0;
                    }
                }
            }
        }
        let uri = image.uri();
        let (x, y, w, h) = self.device_rect(dest.0, dest.1, dest.0 + dest.2, dest.1 + dest.3);
        if !(w > 0.0 && h > 0.0) {
            return;
        }
        let clip = self.clip_attr();
        let _ = write!(
            self.out,
            "<image x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\" preserveAspectRatio=\"none\" href=\"{uri}\"{clip}/>",
            n(x),
            n(y),
            n(w),
            n(h)
        );
        self.drawn += 1;
    }

    fn record(&mut self, kind: u32, rd: &Rd, at: usize, size: usize) {
        let i32_at = |k: usize| rd.i32(at + 8 + k * 4);
        let u32_at = |k: usize| rd.u32(at + 8 + k * 4);
        match kind {
            // Polygons, lines and curves (32 and 16 bits).
            2..=6 | 85..=89 => {
                let short = kind >= 85;
                let pts = rd.points(at + 28, u32_at(4) as usize, short);
                match kind % 83 {
                    2 => {
                        let d = self.bezier(pts.first().copied(), pts.get(1..).unwrap_or_default());
                        self.shape(d, false, true);
                    }
                    3 => {
                        let d = self.poly(&pts, true);
                        self.shape(d, true, true);
                    }
                    4 => {
                        let d = self.poly(&pts, false);
                        self.shape(d, false, true);
                    }
                    5 => {
                        let d = self.bezier(None, &pts);
                        self.shape(d, false, true);
                        if let Some(p) = pts.last() {
                            self.dc.cur = *p;
                        }
                    }
                    _ => {
                        let mut all = vec![self.dc.cur];
                        all.extend(&pts);
                        let d = self.poly(&all, false);
                        self.shape(d, false, true);
                        if let Some(p) = pts.last() {
                            self.dc.cur = *p;
                        }
                    }
                }
            }
            // Polypolylines and polypolygons.
            7 | 8 | 90 | 91 => {
                let short = kind >= 90;
                let polys = (u32_at(4) as usize).min(65536);
                let count = u32_at(5) as usize;
                let counts: Vec<usize> = (0..polys)
                    .map(|k| rd.u32(at + 32 + k * 4) as usize)
                    .collect();
                let pts = rd.points(at + 32 + polys * 4, count, short);
                let close = kind == 8 || kind == 91;
                let mut d = String::new();
                let mut k = 0;
                for c in counts {
                    let Some(part) = pts.get(k..k + c) else { break };
                    d.push_str(&self.poly(part, close));
                    d.push(' ');
                    k += c;
                }
                self.shape(d.trim_end().to_string(), close, true);
            }
            9 => self.dc.window_ext = (i32_at(0), i32_at(1)),
            10 => self.dc.window_org = (i32_at(0), i32_at(1)),
            11 => self.dc.viewport_ext = (i32_at(0), i32_at(1)),
            12 => self.dc.viewport_org = (i32_at(0), i32_at(1)),
            17 => self.dc.map_mode = u32_at(0),
            18 => self.dc.opaque_bk = u32_at(0) == 2,
            19 => self.dc.winding = u32_at(0) == 2,
            22 => self.dc.text_align = u32_at(0),
            24 => self.dc.text_color = colorref(u32_at(0)),
            25 => self.dc.bk_color = colorref(u32_at(0)),
            27 => self.dc.cur = (i32_at(0), i32_at(1)),
            // INTERSECTCLIPRECT
            30 => {
                let r = self.device_rect(i32_at(0), i32_at(1), i32_at(2), i32_at(3));
                self.intersect_clip(r);
            }
            33 => self.saved.push(self.dc.clone()),
            34 => {
                let rel = i32_at(0) as i64;
                let keep = if rel < 0 {
                    self.saved.len() as i64 + rel
                } else {
                    rel - 1
                };
                if keep >= 0 && (keep as usize) < self.saved.len() {
                    self.dc = self.saved[keep as usize].clone();
                    self.saved.truncate(keep as usize);
                }
            }
            35 => {
                self.dc.world = [
                    rd.f32(at + 8),
                    rd.f32(at + 12),
                    rd.f32(at + 16),
                    rd.f32(at + 20),
                    rd.f32(at + 24),
                    rd.f32(at + 28),
                ]
            }
            36 => {
                let x = [
                    rd.f32(at + 8),
                    rd.f32(at + 12),
                    rd.f32(at + 16),
                    rd.f32(at + 20),
                    rd.f32(at + 24),
                    rd.f32(at + 28),
                ];
                // Matrices as (a b c d e f) applied to column points: m(n(p)) is mul(m, n).
                let mul = |m: [f64; 6], k: [f64; 6]| {
                    [
                        m[0] * k[0] + m[2] * k[1],
                        m[1] * k[0] + m[3] * k[1],
                        m[0] * k[2] + m[2] * k[3],
                        m[1] * k[2] + m[3] * k[3],
                        m[0] * k[4] + m[2] * k[5] + m[4],
                        m[1] * k[4] + m[3] * k[5] + m[5],
                    ]
                };
                let w = self.dc.world;
                self.dc.world = match rd.u32(at + 32) {
                    1 => [1.0, 0.0, 0.0, 1.0, 0.0, 0.0],
                    2 => mul(w, x),
                    3 => mul(x, w),
                    4 => x,
                    _ => w,
                };
            }
            37 => self.select(u32_at(0)),
            // CREATEPEN
            38 => {
                let style = u32_at(1);
                let pen = Pen {
                    width: i32_at(2).abs(),
                    color: colorref(u32_at(4)),
                    dash: style & 0xf,
                    null: style & 0xf == 5,
                    cosmetic: i32_at(2).abs() <= 1.0,
                };
                self.objects.insert(u32_at(0), Obj::Pen(pen));
            }
            // CREATEBRUSHINDIRECT
            39 => {
                let style = u32_at(1);
                self.objects.insert(
                    u32_at(0),
                    Obj::Brush(Brush {
                        color: colorref(u32_at(2)),
                        null: style == 1,
                    }),
                );
            }
            40 => {
                self.objects.remove(&u32_at(0));
            }
            // ELLIPSE, RECTANGLE, ROUNDRECT
            42 => {
                let d = self.ellipse(i32_at(0), i32_at(1), i32_at(2), i32_at(3));
                self.shape(d, true, true);
            }
            43 => {
                let d = self.rect(i32_at(0), i32_at(1), i32_at(2), i32_at(3));
                self.shape(d, true, true);
            }
            44 => {
                let (l, t, r, b) = (i32_at(0), i32_at(1), i32_at(2), i32_at(3));
                let (rx, ry) = (
                    (i32_at(4) / 2.0).min((r - l).abs() / 2.0),
                    (i32_at(5) / 2.0).min((b - t).abs() / 2.0),
                );
                let p = |x: f64, y: f64| Self::xy(self.dc.at((x, y)));
                let (kx, ky) = (rx * (1.0 - KAPPA), ry * (1.0 - KAPPA));
                let d = format!(
                    "M{} L{} C{} {} {} L{} C{} {} {} L{} C{} {} {} L{} C{} {} {} Z",
                    p(l + rx, t),
                    p(r - rx, t),
                    p(r - kx, t),
                    p(r, t + ky),
                    p(r, t + ry),
                    p(r, b - ry),
                    p(r, b - ky),
                    p(r - kx, b),
                    p(r - rx, b),
                    p(l + rx, b),
                    p(l + kx, b),
                    p(l, b - ky),
                    p(l, b - ry),
                    p(l, t + ry),
                    p(l, t + ky),
                    p(l + kx, t),
                    p(l + rx, t)
                );
                self.shape(d, true, true);
            }
            // ARC, CHORD, PIE, ARCTO
            45..=47 | 55 => {
                let b = (i32_at(0), i32_at(1), i32_at(2), i32_at(3));
                let pts = self.arc_points(b, (i32_at(4), i32_at(5)), (i32_at(6), i32_at(7)));
                match kind {
                    46 => {
                        let d = self.poly(&pts, true);
                        self.shape(d, true, true);
                    }
                    47 => {
                        let mut all = vec![((b.0 + b.2) / 2.0, (b.1 + b.3) / 2.0)];
                        all.extend(&pts);
                        let d = self.poly(&all, true);
                        self.shape(d, true, true);
                    }
                    55 => {
                        let mut all = vec![self.dc.cur];
                        all.extend(&pts);
                        let d = self.poly(&all, false);
                        self.shape(d, false, true);
                        if let Some(p) = pts.last() {
                            self.dc.cur = *p;
                        }
                    }
                    _ => {
                        let d = self.poly(&pts, false);
                        self.shape(d, false, true);
                    }
                }
            }
            // LINETO
            54 => {
                let to = (i32_at(0), i32_at(1));
                let d = self.poly(&[self.dc.cur, to], false);
                self.shape(d, false, true);
                self.dc.cur = to;
            }
            // Paths.
            59 => self.path = Some(String::new()),
            60 => self.last_path = self.path.take().unwrap_or_default().trim().to_string(),
            61 => {
                if let Some(p) = &mut self.path {
                    p.push_str("Z ");
                }
            }
            62..=64 => {
                let d = std::mem::take(&mut self.last_path);
                self.shape(d, kind != 64, kind != 62);
            }
            // SELECTCLIPPATH
            67 => {
                let d = std::mem::take(&mut self.last_path);
                if !d.is_empty() {
                    self.dc.clip = Some(Clip::Path(d));
                }
            }
            68 => self.path = None,
            // EXTSELECTCLIPRGN: no region (reset), or the region's bounds.
            75 => {
                let (cb, mode) = (u32_at(0), u32_at(1));
                if cb == 0 {
                    if mode == 5 {
                        self.dc.clip = None;
                    }
                } else {
                    let b = (
                        rd.i32(at + 16 + 16),
                        rd.i32(at + 16 + 20),
                        rd.i32(at + 16 + 24),
                        rd.i32(at + 16 + 28),
                    );
                    let r = (b.0, b.1, b.2 - b.0, b.3 - b.1);
                    match mode {
                        1 => self.intersect_clip(r),
                        5 => self.dc.clip = Some(Clip::Rect(r.0, r.1, r.2, r.3)),
                        _ => {}
                    }
                }
            }
            // BITBLT, STRETCHBLT
            76 | 77 => {
                let dest = (i32_at(4), i32_at(5), i32_at(6), i32_at(7));
                let base = 8 + 16 + 16 + 4 + 8 + 24 + 4 + 4;
                let parts = (
                    rd.u32(at + base) as usize,
                    rd.u32(at + base + 4) as usize,
                    rd.u32(at + base + 8) as usize,
                    rd.u32(at + base + 12) as usize,
                );
                self.bitmap(rd, at, parts, dest, u32_at(8));
            }
            // STRETCHDIBITS
            81 => {
                let (x, y) = (i32_at(4), i32_at(5));
                let parts = (
                    u32_at(10) as usize,
                    u32_at(11) as usize,
                    u32_at(12) as usize,
                    u32_at(13) as usize,
                );
                let (w, h) = (i32_at(16), i32_at(17));
                self.bitmap(rd, at, parts, (x, y, w, h), u32_at(15));
            }
            // SETDIBITSTODEVICE
            80 => {
                let (x, y) = (i32_at(4), i32_at(5));
                let (w, h) = (i32_at(8), i32_at(9));
                let parts = (
                    u32_at(10) as usize,
                    u32_at(11) as usize,
                    u32_at(12) as usize,
                    u32_at(13) as usize,
                );
                self.bitmap(rd, at, parts, (x, y, w, h), 0x00CC_0020);
            }
            // ALPHABLEND
            114 => {
                let dest = (i32_at(4), i32_at(5), i32_at(6), i32_at(7));
                let base = 8 + 16 + 16 + 4 + 8 + 24 + 4 + 4;
                let parts = (
                    rd.u32(at + base) as usize,
                    rd.u32(at + base + 4) as usize,
                    rd.u32(at + base + 8) as usize,
                    rd.u32(at + base + 12) as usize,
                );
                self.bitmap(rd, at, parts, dest, 0x00CC_0020);
            }
            // EXTCREATEFONTINDIRECTW
            82 => {
                let lf = at + 12;
                let face: Vec<u16> = (0..32)
                    .map(|k| {
                        rd.0.get(lf + 28 + k * 2..lf + 30 + k * 2)
                            .map(|b| u16::from_le_bytes([b[0], b[1]]))
                            .unwrap_or(0)
                    })
                    .take_while(|c| *c != 0)
                    .collect();
                let font = Font {
                    height: rd.i32(lf),
                    escapement: rd.i32(lf + 8),
                    weight: rd.i32(lf + 16) as i32,
                    italic: rd.0.get(lf + 20).is_some_and(|b| *b != 0),
                    underline: rd.0.get(lf + 21).is_some_and(|b| *b != 0),
                    strike: rd.0.get(lf + 22).is_some_and(|b| *b != 0),
                    face: String::from_utf16_lossy(&face),
                };
                self.objects.insert(u32_at(0), Obj::Font(font));
            }
            83 | 84 => self.text(rd, at, kind == 84),
            // EXTCREATEPEN
            95 => {
                let style = u32_at(5);
                let width = u32_at(6) as f64;
                let brush_style = u32_at(7);
                let geometric = style & 0x0001_0000 != 0;
                let pen = Pen {
                    width,
                    color: colorref(u32_at(8)),
                    dash: style & 0xf,
                    null: style & 0xf == 5 || brush_style == 1,
                    cosmetic: !geometric || width <= 1.0 && !geometric,
                };
                self.objects.insert(u32_at(0), Obj::Pen(pen));
            }
            // CREATEMONOBRUSH, CREATEDIBPATTERNBRUSHPT: a gray fill.
            93 | 94 => {
                self.objects.insert(
                    u32_at(0),
                    Obj::Brush(Brush {
                        color: "#c0c0c0".into(),
                        null: false,
                    }),
                );
            }
            _ => {
                let _ = size;
            }
        }
    }
}

/// A bitmap: decoded pixels (width, height, RGBA), or a JPEG / PNG file as it is.
enum Dib {
    Pixels(usize, usize, Vec<u8>),
    File(&'static str, Vec<u8>),
}

impl Dib {
    fn uri(&self) -> String {
        match self {
            Dib::Pixels(w, h, rgba) => format!(
                "data:image/png;base64,{}",
                super::drawing::base64(&png(*w, *h, rgba))
            ),
            Dib::File(mime, data) => format!("data:{mime};base64,{}", super::drawing::base64(data)),
        }
    }
}

/// A DIB (BITMAPINFO and bits).
fn dib(bmi: &[u8], bits: &[u8]) -> Option<Dib> {
    let rd = Rd(bmi);
    let header = rd.u32(0) as usize;
    let (w, h, count, compression, used) = if header == 12 {
        (
            rd.0.get(4..6)
                .map(|b| u16::from_le_bytes([b[0], b[1]]) as i64)?,
            rd.0.get(6..8)
                .map(|b| u16::from_le_bytes([b[0], b[1]]) as i64)?,
            rd.0.get(10..12).map(|b| u16::from_le_bytes([b[0], b[1]]))?,
            0,
            0,
        )
    } else {
        (
            rd.i32(4) as i64,
            rd.i32(8) as i64,
            bmi.get(14..16).map(|b| u16::from_le_bytes([b[0], b[1]]))?,
            rd.u32(16),
            rd.u32(32) as usize,
        )
    };
    match compression {
        4 => return Some(Dib::File("image/jpeg", bits.to_vec())),
        5 => return Some(Dib::File("image/png", bits.to_vec())),
        0 | 3 => {}
        _ => return None,
    }
    let (width, height) = (w.unsigned_abs() as usize, h.unsigned_abs() as usize);
    if width == 0 || height == 0 || width > 16384 || height > 16384 {
        return None;
    }
    let bottom_up = h > 0;
    let palette_at = header
        + if compression == 3 && header == 40 {
            12
        } else {
            0
        };
    let entries = if count <= 8 {
        if used > 0 { used } else { 1 << count }
    } else {
        0
    };
    let palette: Vec<[u8; 3]> = (0..entries)
        .map(|i| {
            bmi.get(palette_at + i * if header == 12 { 3 } else { 4 }..)
                .and_then(|b| b.get(..3))
                .map(|b| [b[2], b[1], b[0]])
                .unwrap_or([0, 0, 0])
        })
        .collect();
    // The masks of 16 and 32 bits per pixel (BI_BITFIELDS), else 5-5-5 and 8-8-8.
    let masks = if compression == 3 {
        [rd.u32(40), rd.u32(44), rd.u32(48)]
    } else if count == 16 {
        [0x7c00, 0x03e0, 0x001f]
    } else {
        [0xff0000, 0xff00, 0xff]
    };
    let channel = |v: u32, m: u32| -> u8 {
        if m == 0 {
            return 0;
        }
        let shift = m.trailing_zeros();
        let max = m >> shift;
        (((v & m) >> shift) * 255 / max.max(1)) as u8
    };
    let stride = (width * count as usize).div_ceil(32) * 4;
    let mut rgba = vec![0u8; width * height * 4];
    let mut any_alpha = false;
    for y in 0..height {
        let row = if bottom_up { height - 1 - y } else { y };
        let src = bits.get(row * stride..row * stride + stride)?;
        for x in 0..width {
            let px: [u8; 4] = match count {
                1 | 4 | 8 => {
                    let bit = x * count as usize;
                    let byte = src.get(bit / 8).copied().unwrap_or(0);
                    let idx =
                        (byte >> (8 - count as usize - bit % 8)) & ((1u16 << count) - 1) as u8;
                    let c = palette.get(idx as usize).copied().unwrap_or([0, 0, 0]);
                    [c[0], c[1], c[2], 255]
                }
                16 => {
                    let v = u16::from_le_bytes([src[x * 2], src[x * 2 + 1]]) as u32;
                    [
                        channel(v, masks[0]),
                        channel(v, masks[1]),
                        channel(v, masks[2]),
                        255,
                    ]
                }
                24 => [src[x * 3 + 2], src[x * 3 + 1], src[x * 3], 255],
                32 => {
                    let v = u32::from_le_bytes([
                        src[x * 4],
                        src[x * 4 + 1],
                        src[x * 4 + 2],
                        src[x * 4 + 3],
                    ]);
                    let a = src[x * 4 + 3];
                    any_alpha |= a != 0;
                    [
                        channel(v, masks[0]),
                        channel(v, masks[1]),
                        channel(v, masks[2]),
                        a,
                    ]
                }
                _ => return None,
            };
            rgba[(y * width + x) * 4..(y * width + x) * 4 + 4].copy_from_slice(&px);
        }
    }
    // 32 bits: the alpha when there is any (GDI+ writes transparent pictures so), else opaque.
    // Colors premultiplied by it (none above its alpha) are divided back.
    if count == 32 && !any_alpha {
        for p in rgba.chunks_exact_mut(4) {
            p[3] = 255;
        }
    } else if count == 32
        && rgba
            .chunks_exact(4)
            .all(|p| p[0] <= p[3] && p[1] <= p[3] && p[2] <= p[3])
    {
        for p in rgba.chunks_exact_mut(4) {
            if p[3] > 0 {
                for k in 0..3 {
                    p[k] = ((p[k] as u32 * 255 + p[3] as u32 / 2) / p[3] as u32).min(255) as u8;
                }
            }
        }
    }
    Some(Dib::Pixels(width, height, rgba))
}

/// RGBA pixels as a PNG file.
fn png(width: usize, height: usize, rgba: &[u8]) -> Vec<u8> {
    use std::io::Write as _;
    let mut raw = Vec::with_capacity(height * (width * 4 + 1));
    for row in rgba.chunks_exact(width * 4) {
        raw.push(0);
        raw.extend_from_slice(row);
    }
    let mut z = flate2::write::ZlibEncoder::new(Vec::new(), flate2::Compression::fast());
    let _ = z.write_all(&raw);
    let data = z.finish().unwrap_or_default();
    let mut out = b"\x89PNG\r\n\x1a\n".to_vec();
    let mut chunk = |kind: &[u8], body: &[u8]| {
        out.extend((body.len() as u32).to_be_bytes());
        let mut crc = flate2::Crc::new();
        crc.update(kind);
        crc.update(body);
        out.extend(kind);
        out.extend(body);
        out.extend(crc.sum().to_be_bytes());
    };
    let mut ihdr = Vec::new();
    ihdr.extend((width as u32).to_be_bytes());
    ihdr.extend((height as u32).to_be_bytes());
    ihdr.extend([8, 6, 0, 0, 0]);
    chunk(b"IHDR", &ihdr);
    chunk(b"IDAT", &data);
    chunk(b"IEND", &[]);
    out
}

/// An EMF picture as an SVG document filling the picture's box; None when it is not EMF or
/// draws nothing this player draws (an EMF+ only picture).
pub(crate) fn to_svg(data: &[u8]) -> Option<String> {
    let rd = Rd(data);
    if rd.u32(0) != 1 || rd.u32(40) != 0x464D_4520 {
        return None;
    }
    // The bounds (device units, inclusive); else the frame (0.01 mm) at the reference device.
    let (l, t, r, b) = (rd.i32(8), rd.i32(12), rd.i32(16), rd.i32(20));
    let (vx, vy, vw, vh) = if r > l && b > t {
        (l, t, r - l + 1.0, b - t + 1.0)
    } else {
        let (fl, ft, fr, fb) = (rd.i32(24), rd.i32(28), rd.i32(32), rd.i32(36));
        let (dx, dy, mx, my) = (rd.i32(72), rd.i32(76), rd.i32(80), rd.i32(84));
        if !(mx > 0.0 && my > 0.0 && fr > fl && fb > ft) {
            return None;
        }
        let (kx, ky) = (dx / mx / 100.0, dy / my / 100.0);
        (fl * kx, ft * ky, (fr - fl) * kx, (fb - ft) * ky)
    };
    let mut p = Player {
        dc: Dc::default(),
        saved: Vec::new(),
        objects: HashMap::new(),
        out: String::new(),
        defs: String::new(),
        clips: Vec::new(),
        path: None,
        last_path: String::new(),
        drawn: 0,
        mask: None,
    };
    let mut at = 0usize;
    let mut records = 0;
    while at + 8 <= data.len() && records < 2_000_000 {
        let (kind, size) = (rd.u32(at), rd.u32(at + 4) as usize);
        if size < 8 || at + size > data.len() {
            break;
        }
        if kind == 14 {
            break;
        }
        p.record(kind, &Rd(&data[..at + size]), at, size);
        at += size;
        records += 1;
    }
    if p.drawn == 0 {
        return None;
    }
    Some(format!(
        "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"{} {} {} {}\" preserveAspectRatio=\"none\"><defs>{}</defs>{}</svg>",
        n(vx),
        n(vy),
        n(vw),
        n(vh),
        p.defs,
        p.out
    ))
}

#[cfg(test)]
pub(crate) mod tests {
    use super::*;

    /// An EMF made of records (type, body), with its header and end.
    pub(crate) fn emf(bounds: (i32, i32, i32, i32), records: &[(u32, Vec<u8>)]) -> Vec<u8> {
        let mut body = Vec::new();
        for (kind, data) in records {
            let mut d = data.clone();
            while d.len() % 4 != 0 {
                d.push(0);
            }
            body.extend(kind.to_le_bytes());
            body.extend(((d.len() + 8) as u32).to_le_bytes());
            body.extend(d);
        }
        body.extend(14u32.to_le_bytes());
        body.extend(20u32.to_le_bytes());
        body.extend([0u8; 12]);
        let mut h = vec![0u8; 108];
        h[0..4].copy_from_slice(&1u32.to_le_bytes());
        h[4..8].copy_from_slice(&108u32.to_le_bytes());
        for (k, v) in [bounds.0, bounds.1, bounds.2, bounds.3].iter().enumerate() {
            h[8 + k * 4..12 + k * 4].copy_from_slice(&v.to_le_bytes());
        }
        h[40..44].copy_from_slice(&0x464D_4520u32.to_le_bytes());
        h.extend(body);
        h
    }

    pub(crate) fn ints(v: &[i32]) -> Vec<u8> {
        v.iter().flat_map(|x| x.to_le_bytes()).collect()
    }

    /// A rectangle filled with a red brush, text in a font, and a small bitmap.
    pub(crate) fn sample() -> Vec<u8> {
        let mut text = Vec::new();
        // bounds, graphics mode, scales
        text.extend(ints(&[0, 0, 0, 0, 1]));
        text.extend(1f32.to_le_bytes());
        text.extend(1f32.to_le_bytes());
        // EMRTEXT: reference, chars, offString, options, rect, offDx
        let s: Vec<u8> = "売上表"
            .encode_utf16()
            .flat_map(|u| u.to_le_bytes())
            .collect();
        let off_string = 8 + 28 + 40;
        text.extend(ints(&[10, 60, 3, off_string, 0, 0, 0, 0, 0, 0]));
        text.extend(&s);
        let mut font = ints(&[1, -16, 0, 0, 0, 700]);
        font.extend([0u8, 0, 0, 128, 0, 0, 0, 0]);
        let face: Vec<u8> = "ＭＳ Ｐゴシック"
            .encode_utf16()
            .flat_map(|u| u.to_le_bytes())
            .collect();
        font.extend(&face);
        font.resize(4 + 92, 0);
        // A 2 × 1 bitmap, 24 bits (blue, green), bottom up.
        let mut bmi = ints(&[40, 2, 1]);
        bmi.extend(1u16.to_le_bytes());
        bmi.extend(24u16.to_le_bytes());
        bmi.extend(ints(&[0, 8, 0, 0, 0, 0]));
        let bits = vec![255u8, 0, 0, 0, 255, 0, 0, 0];
        let mut dib = ints(&[0, 0, 0, 0, 100, 0, 0, 0, 2, 1]);
        let off_bmi = 8 + 72;
        dib.extend(ints(&[
            off_bmi,
            bmi.len() as i32,
            off_bmi + bmi.len() as i32,
            bits.len() as i32,
            0,
            0x00CC_0020,
            40,
            20,
        ]));
        dib.extend(&bmi);
        dib.extend(&bits);
        emf(
            (0, 0, 199, 99),
            &[
                (39, ints(&[1, 0, 0x0000_00FF, 0])),
                (37, ints(&[1])),
                (37, ints(&[0x8000_0008u32 as i32])),
                (43, ints(&[10, 10, 90, 40])),
                (82, font),
                (37, ints(&[1])),
                (24, ints(&[0x0000_8000])),
                (84, text),
                (81, dib),
            ],
        )
    }

    #[test]
    fn records_as_svg() {
        let svg = to_svg(&sample()).unwrap();
        assert!(
            svg.starts_with("<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 200 100\""),
            "{svg}"
        );
        assert!(svg.contains("<path d=\"M10,10 L90,10 L90,40 L10,40 Z\" fill=\"#ff0000\" fill-rule=\"evenodd\" stroke=\"none\"/>"), "{svg}");
        assert!(
            svg.contains("fill=\"#008000\"")
                && svg.contains("font-size:16px;font-weight:bold;")
                && svg.contains(">売上表</text>"),
            "{svg}"
        );
        assert!(svg.contains("<image x=\"100\" y=\"0\" width=\"40\" height=\"20\" preserveAspectRatio=\"none\" href=\"data:image/png;base64,"), "{svg}");
        // Not EMF.
        assert!(to_svg(b"GIF89a").is_none());
    }

    #[test]
    fn transforms() {
        // MM_ANISOTROPIC: a window of 1000 × 1000 onto a viewport of 100 × 100.
        let data = emf(
            (0, 0, 99, 99),
            &[
                (17, ints(&[8])),
                (9, ints(&[1000, 1000])),
                (11, ints(&[100, 100])),
                (43, ints(&[0, 0, 500, 250])),
            ],
        );
        let svg = to_svg(&data).unwrap();
        assert!(svg.contains("d=\"M0,0 L50,0 L50,25 L0,25 Z\""), "{svg}");
    }
}
