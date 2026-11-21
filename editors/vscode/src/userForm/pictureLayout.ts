// How MSForms lays out a picture: beside the caption of a Label, a button or a toggle
// (PicturePosition), and as the background of a form, Frame or Page (PictureAlignment,
// PictureSizeMode, PictureTiling). Shared by the Designer canvas and the debug-run panel.

/** Picture and caption: side by side (the caption aligned to the picture), or the caption over a centred picture. */
export interface PictureCaptionLayout {
	/** CSS flex-direction of picture then caption, or "stack" for Center. */
	direction: "row" | "row-reverse" | "column" | "column-reverse" | "stack";
	/** CSS align-items: where the caption is along the picture. */
	align: "flex-start" | "center" | "flex-end";
}

/** fmPicturePosition 0 (LeftTop) … 12 (Center); omitted, AboveCenter (7). */
export function picturePositionLayout(position: unknown): PictureCaptionLayout {
	const n = typeof position === "number" && position >= 0 && position <= 12 ? position : 7;
	if (n === 12) return { direction: "stack", align: "center" };
	const direction = (["row", "row-reverse", "column", "column-reverse"] as const)[Math.floor(n / 3)];
	const align = (["flex-start", "center", "flex-end"] as const)[n % 3];
	return { direction, align };
}

/** CSS of a background picture: where it is, how it fits, whether it repeats. */
export function backgroundPicture(dataUri: string | undefined, alignment: unknown, sizeMode: unknown, tiling: unknown): Record<string, string> {
	if (!dataUri) return {};
	const position = ["left top", "right top", "center", "left bottom", "right bottom"][typeof alignment === "number" ? alignment : 2] ?? "center";
	const size = sizeMode === 1 ? "100% 100%" : sizeMode === 3 ? "contain" : "auto";
	return {
		backgroundImage: `url("${dataUri}")`,
		backgroundPosition: position,
		backgroundSize: size,
		backgroundRepeat: tiling === true ? "repeat" : "no-repeat",
	};
}
