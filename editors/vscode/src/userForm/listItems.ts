// The `list` of a ComboBox / ListBox as the text of the Property Grid's
// multi-line editor: one item per line.

/** Text shown for a `list` value; anything that is not a list of strings is empty. */
export function listToText(value: unknown): string {
	return Array.isArray(value) && value.every((item) => typeof item === "string") ? value.join("\n") : "";
}

/**
 * Items entered in the editor. Blank lines at the end are dropped, blank
 * lines between items are kept as empty items. Text without any item means
 * "no list" and yields null.
 */
export function textToList(text: string): string[] | null {
	const lines = text.split(/\r\n|\r|\n/);
	while (lines.length && lines[lines.length - 1] === "") lines.pop();
	return lines.length ? lines : null;
}
