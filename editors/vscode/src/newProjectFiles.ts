// The parts of "New RAS project" that do not need VS Code (unit-tested in test/newProjectFiles.test.ts).

/** The workbooks a project can use (an .xlsx is converted to .xlsm). */
export const WORKBOOK_EXTENSIONS = [".xlsm", ".xlsx", ".xlam", ".xlsb"];

/**
 * A folder to start a project in: empty, or holding only hidden entries (`.git`, `.vscode` …)
 * and workbooks (Excel's lock files `~$…` included).
 */
export function isNewProjectFolder(entries: readonly string[]): boolean {
	return entries.every((name) => name.startsWith(".") || name.startsWith("~$") || WORKBOOK_EXTENSIONS.includes(extension(name)));
}

/** The default name of a project's new workbook: the folder's name, as `ras init` names the project. */
export function projectBookName(folder: string): string {
	const stem = folder.replace(/[\\/:*?"<>|]/g, "").trim();
	return `${stem === "" || stem.startsWith(".") ? "Book" : stem}.xlsm`;
}

/** A file name usable for a project's new workbook. */
export function workbookFileNameError(name: string): string | undefined {
	if (name === "" || name.startsWith(".")) return "名前を入力してください。";
	if (/[\\/:*?"<>|]/.test(name)) return "ファイル名に使えない文字が含まれています。";
	const ext = extension(name);
	if (ext !== ".xlsm" && ext !== ".xlam") return "拡張子は .xlsm（マクロ有効ブック）か .xlam（アドイン）にしてください。";
	return undefined;
}

/** `ras.toml` text with `project.workbook` set to `book` (the line `ras init` writes, commented out or not). */
export function setWorkbook(toml: string, book: string): string {
	const line = `workbook = ${JSON.stringify(book)}`;
	const lines = toml.split("\n");
	const existing = lines.findIndex((l) => /^\s*#?\s*workbook\s*=/.test(l));
	if (existing >= 0) {
		lines[existing] = line;
		return lines.join("\n");
	}
	const project = lines.findIndex((l) => l.trim() === "[project]");
	if (project >= 0) {
		lines.splice(project + 1, 0, line);
		return lines.join("\n");
	}
	return `[project]\n${line}\n\n${toml}`;
}

function extension(name: string): string {
	const dot = name.lastIndexOf(".");
	return dot <= 0 ? "" : name.slice(dot).toLowerCase();
}
