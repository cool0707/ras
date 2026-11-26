// The progress bar (`"type": "ProgressBar"`), a control RAS writes as four labels. The layout is
// the one of ras-form's progress.rs (what is written into the workbook), of the bundled class
// `ProgressBar` and of the debug run (js/forms.js `progressParts`).

export interface ProgressBarLayout {
	/** The track (under the bar), the bar, the caption and the percentage: left, top, width, height in points, relative to the control. */
	track: [number, number, number, number];
	bar: [number, number, number, number];
	caption: [number, number, number, number];
	value: [number, number, number, number];
	/** The text of the percentage (`56%`). */
	text: string;
	showValue: boolean;
}

const num = (v: unknown, d: number) => (typeof v === "number" && Number.isFinite(v) ? v : d);

export function progressBarLayout(shown: { width: number; height: number }): ProgressBarLayout {
	const control = shown as { width: number; height: number } & Record<string, unknown>;
	const min = num(control.min, 0);
	const max = num(control.max, 100);
	const value = num(control.value, 0);
	const showValue = control.showValue !== false;
	const { width, height } = control;
	const fraction = max <= min ? 0 : Math.min(1, Math.max(0, (value - min) / (max - min)));
	const barHeight = Math.min(Math.max(num(control.barHeight, 4.5), 0.75), Math.max(height, 0.75));
	const barTop = height - barHeight;
	const textHeight = Math.max(height - barHeight - 3, 0);
	const valueWidth = showValue ? Math.min(36, width) : 0;
	return {
		track: [0, barTop, width, barHeight],
		bar: [0, barTop, Math.round((width * fraction) / 0.75) * 0.75, barHeight],
		caption: [0, 0, Math.max(width - valueWidth, 0), textHeight],
		value: [width - valueWidth, 0, valueWidth, textHeight],
		text: `${Math.floor(fraction * 100 + 0.5)}%`,
		showValue,
	};
}
