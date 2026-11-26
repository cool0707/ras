// The height a single-line TextBox or ComboBox needs for its font.
//
// MSForms does not grow an edit field to its font: when the field is lower than a
// line of text, the text is clipped and parts of glyphs and of the caret are left
// as stray lines along the edge. The Designer raises such a height (see engine.ts)
// and validation warns about one written by hand.

/** Line height (ascent + descent) per point of size, for fonts common in Excel forms. */
const lineFactors: [RegExp, number][] = [
	[/^(ms|ｍｓ)\s*(ui\s*)?(p\s*)?(gothic|ゴシック|mincho|明朝)/i, 1.0],
	[/^meiryo\s*ui|^メイリオ\s*ui/i, 1.33],
	[/^meiryo|^メイリオ/i, 1.6],
	[/^(yu\s*gothic|游ゴシック)\s*ui/i, 1.33],
	[/^(yu\s*(gothic|mincho)|游(ゴシック|明朝))/i, 1.71],
	[/^tahoma/i, 1.21],
	[/^segoe\s*ui/i, 1.33],
	[/^arial|^times/i, 1.15],
	[/^calibri/i, 1.22],
	[/^verdana/i, 1.22],
];

/** The font the run time and Excel's Japanese forms use when none is written. */
const defaultFont = { name: "MS UI Gothic", size: 9 };

/** The sunken edge (2px each side) and the field's inner margin (1px each side), in points. */
const chrome = 4.5;

const isEditField = (type: string) => ["textbox", "combobox", "filepicker"].includes(type.toLowerCase());

/** Points rounded up to whole pixels at 96 DPI (0.75 pt), as Excel stores them. */
const toPixels = (points: number) => Math.ceil(points / 0.75 - 1e-9) * 0.75;

/** The least height of a single-line TextBox / ComboBox with `font`, or undefined for other controls. */
export function minimumEditHeight(type: string, font: unknown, multiLine = false): number | undefined {
	if (!isEditField(type) || multiLine) return undefined;
	const f = typeof font === "object" && font !== null ? (font as { name?: unknown; size?: unknown }) : {};
	const name = typeof f.name === "string" && f.name.trim() ? f.name.trim() : defaultFont.name;
	const size = typeof f.size === "number" && f.size > 0 ? f.size : defaultFont.size;
	const factor = lineFactors.find(([pattern]) => pattern.test(name))?.[1] ?? 1.3;
	return toPixels(size * factor + chrome);
}
