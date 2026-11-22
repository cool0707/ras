// The sheet panel: a workbook drawn as a grid (`ras book view`), for working without Excel.
//
// - "RAS: ブックを表示" shows a workbook file: the project's workbook, or one a run saved
//   (build/debug/workbooks, the workbooks of failed tests), with what the run changed marked.
// - While a debug run on the stand-in for Excel is stopped (a breakpoint, a step), the panel
//   shows the workbook as it is at that moment, with the cells changed since the last stop
//   marked: the run time writes the workbooks out when the editor asks (`$rasBookSnapshot`).
// - A workbook file shown as it is (not one a run saved, not the workbook of a debug run) is
//   edited in the panel as in Excel: the page posts what was done (`edit`: cells typed into,
//   cleared, inserted, deleted, pasted, ...), `ras book edit` writes the workbook file and makes
//   the page again. Undo and redo put back the bytes the file had.
// - On the project's workbook, the shapes and controls of a sheet get their events as the form
//   designer's controls do (the page's menu and button, a double click): an ActiveX control's
//   handler is an `@on("CommandButton1.Click")` function of the sheet's file; a shape or a form
//   control gets an `export function` of the sheet's file as its macro (`ras book edit`'s `macro`
//   writes `OnAction` into the workbook file).
// - "RAS: シートのグリッドを表示" on the RAS file of a worksheet (`objects/Sheet1.ras`) shows the
//   project's workbook at that sheet: from the editor's title bar and context menu, and from
//   the Explorer's context menu.

import * as cp from "child_process";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import * as vscode from "vscode";
import { sheetOfRasFile } from "./sheetFile";
import { findServerExecutable } from "./serverPath";
import { macrosOf, moduleName, newMacroName, sheetControlEvents } from "./sheetEvents";
import { bodyIndent, findHandler, handlerText, rasParams } from "./userForm/events";
import type { EventDescriptor } from "./userForm/events";

let panel: vscode.WebviewPanel | undefined;

/** What the panel shows: a workbook file edited in place has `editable`. */
interface Shown {
	root: string | undefined;
	file: string;
	/** The sheet the page was asked to show first. */
	sheet?: string;
	editable: boolean;
}
let shown: Shown | undefined;

function setting(name: string): string {
	return (vscode.workspace.getConfiguration("ras").get<string>(name) ?? "").trim();
}

/** The project (the folder of `ras.toml`) a file belongs to. */
function projectRoot(file: string): string | undefined {
	for (let dir = path.dirname(file); ; dir = path.dirname(dir)) {
		if (fs.existsSync(path.join(dir, "ras.toml"))) {
			return dir;
		}
		if (path.dirname(dir) === dir) {
			return undefined;
		}
	}
}

/** `workbook` of the project's `ras.toml`, as a path. */
function projectWorkbook(root: string): string | undefined {
	try {
		const text = fs.readFileSync(path.join(root, "ras.toml"), "utf8");
		const m = /^\s*workbook\s*=\s*"([^"]+)"/m.exec(text);
		return m ? path.resolve(root, m[1].replace(/\\\\/g, "\\")) : undefined;
	} catch {
		return undefined;
	}
}

/**
 * The page of `ras book view` for a workbook. `drawings`: the workbook whose shapes, pictures and
 * controls are drawn (a workbook a run saved keeps its cells, not all of them).
 */
function page(root: string | undefined, file: string, before: string | undefined, title: string | undefined, drawings?: string, sheet?: string, editable = false): Promise<string> {
	const exe = findServerExecutable(setting("path"), root ?? path.dirname(file));
	if (exe === undefined) {
		return Promise.reject(new Error("ras 実行ファイルが見つかりません。設定 ras.path を確認してください。"));
	}
	const args = ["book", "view", file, "--stdout"];
	if (before !== undefined) {
		args.push("--diff", before);
	}
	if (title !== undefined) {
		args.push("--title", title);
	}
	if (drawings !== undefined) {
		args.push("--drawings", drawings);
	}
	if (sheet !== undefined) {
		args.push("--sheet", sheet);
	}
	if (editable) {
		args.push("--editable");
	}
	return ras(exe, args, root ?? path.dirname(file));
}

/** Runs `ras` and gives its standard output; `input` goes to its standard input. */
function ras(exe: string, args: string[], cwd: string, input?: string): Promise<string> {
	return new Promise((resolve, reject) => {
		const child = cp.execFile(exe, args, { cwd, windowsHide: true, maxBuffer: 256 * 1024 * 1024 }, (error, stdout, stderr) => {
			if (error) {
				// (`ras` writes "error: ..." for people reading a terminal.)
				reject(new Error(stderr.trim().replace(/^error:\s*/, "") || error.message));
			} else {
				resolve(stdout);
			}
		});
		if (input !== undefined) {
			child.stdin?.on("error", () => undefined);
			child.stdin?.end(input);
		}
	});
}

function show(html: string, title: string, reveal: boolean, what?: Shown): void {
	if (panel === undefined) {
		panel = vscode.window.createWebviewPanel("ras.sheetView", title, { viewColumn: vscode.ViewColumn.Beside, preserveFocus: true }, { enableScripts: true, retainContextWhenHidden: true });
		panel.onDidDispose(() => {
			panel = undefined;
			shown = undefined;
		});
		panel.webview.onDidReceiveMessage((message: unknown) => receive(message));
	} else if (reveal) {
		panel.reveal(undefined, true);
	}
	shown = what;
	panel.title = title;
	panel.webview.html = html;
}

// ---- editing: what the page posts

/** The bytes a workbook file had before the edits made here, and after those undone. */
interface History {
	undo: Buffer[];
	redo: Buffer[];
	/** The file as it was left here: changed by anything else, the history no longer fits it. */
	stamp: { mtimeMs: number; size: number } | undefined;
}
const histories = new Map<string, History>();
const HISTORY_ENTRIES = 50;
const HISTORY_BYTES = 200 * 1024 * 1024;

function stampOf(file: string): { mtimeMs: number; size: number } | undefined {
	try {
		const stat = fs.statSync(file);
		return { mtimeMs: stat.mtimeMs, size: stat.size };
	} catch {
		return undefined;
	}
}

function historyOf(file: string): History {
	const key = process.platform === "win32" ? file.toLowerCase() : file;
	let history = histories.get(key);
	if (history === undefined) {
		history = { undo: [], redo: [], stamp: undefined };
		histories.set(key, history);
	}
	const now = stampOf(file);
	if (history.stamp !== undefined && (now === undefined || now.mtimeMs !== history.stamp.mtimeMs || now.size !== history.stamp.size)) {
		history.undo = [];
		history.redo = [];
	}
	return history;
}

/** Keeps the history within its limits: the oldest states go first. */
function trim(history: History): void {
	const total = () => history.undo.reduce((n, b) => n + b.length, 0) + history.redo.reduce((n, b) => n + b.length, 0);
	while (history.undo.length > HISTORY_ENTRIES || (history.undo.length > 0 && total() > HISTORY_BYTES)) {
		history.undo.shift();
	}
	while (history.redo.length > 0 && total() > HISTORY_BYTES) {
		history.redo.shift();
	}
}

/** Puts `bytes` in place of the file: written next to it, then renamed over it. */
function writeBack(file: string, bytes: Buffer): void {
	const temp = path.join(path.dirname(file), `.${path.basename(file)}.${process.pid}-${Date.now().toString(16)}.ras-tmp`);
	try {
		fs.writeFileSync(temp, bytes);
		fs.renameSync(temp, file);
	} catch (error) {
		fs.rmSync(temp, { force: true });
		throw error;
	}
}

function executable(what: Shown): string {
	const exe = findServerExecutable(setting("path"), what.root ?? path.dirname(what.file));
	if (exe === undefined) {
		throw new Error("ras 実行ファイルが見つかりません。設定 ras.path を確認してください。");
	}
	return exe;
}

/** Applies what the page posted to the workbook file and gives the page made from it. */
async function edit(what: Shown, ops: unknown): Promise<string> {
	const history = historyOf(what.file);
	const before = fs.readFileSync(what.file);
	const args = ["book", "edit", what.file, "--ops-file", "-", "--view", "--editable"];
	if (what.sheet !== undefined) {
		args.push("--sheet", what.sheet);
	}
	const html = await ras(executable(what), args, what.root ?? path.dirname(what.file), JSON.stringify(ops));
	// An edit that changed nothing (the same value typed again) is not a step to undo.
	if (!before.equals(fs.readFileSync(what.file))) {
		history.undo.push(before);
		history.redo = [];
		trim(history);
	}
	history.stamp = stampOf(what.file);
	return html;
}

/** Puts back the bytes the file had before the last edit (`redo`: after the edit last undone). */
async function travel(what: Shown, redo: boolean): Promise<string> {
	const history = historyOf(what.file);
	const [from, to] = redo ? [history.redo, history.undo] : [history.undo, history.redo];
	const bytes = from.pop();
	if (bytes === undefined) {
		throw new Error(redo ? "やり直す操作はありません。" : "元に戻す操作はありません。");
	}
	try {
		const current = fs.readFileSync(what.file);
		writeBack(what.file, bytes);
		to.push(current);
	} catch (error) {
		from.push(bytes);
		throw new Error(`${path.basename(what.file)} を書き換えられませんでした（Excel で開いている場合は閉じてください）: ${error instanceof Error ? error.message : String(error)}`);
	}
	trim(history);
	history.stamp = stampOf(what.file);
	return page(what.root, what.file, undefined, undefined, undefined, what.sheet, true);
}

/** One edit at a time, in the order the page posted them. */
let queue: Promise<void> = Promise.resolve();

function receive(message: unknown): void {
	const type = typeof message === "object" && message !== null ? (message as { type?: unknown }).type : undefined;
	if (type === "ready") {
		// The page asks whether its objects have events: those of the project's workbook do.
		if (panel !== undefined && shown !== undefined && shown.editable && sheetFolder(shown) !== undefined) {
			void panel.webview.postMessage({ type: "events", on: true });
		}
		return;
	}
	if (type === "link") {
		// A hyperlink of a cell: web pages and mail only.
		const url = (message as { url?: unknown }).url;
		if (typeof url === "string" && /^(https?:\/\/|mailto:)/i.test(url)) {
			void vscode.env.openExternal(vscode.Uri.parse(url, true));
		}
		return;
	}
	if (type !== "edit" && type !== "undo" && type !== "redo" && type !== "event") {
		return;
	}
	queue = queue.then(async () => {
		const what = shown;
		if (panel === undefined || what === undefined) {
			return;
		}
		try {
			if (!what.editable) {
				throw new Error("このブックは、この表示では編集できません。");
			}
			const html =
				type === "event"
					? await openObjectEvent(what, message as ObjectEvent)
					: type === "edit"
						? await edit(what, (message as { ops?: unknown }).ops ?? [])
						: await travel(what, type === "redo");
			// (Another workbook may be shown by now.)
			if (html !== undefined && panel !== undefined && shown === what) {
				panel.webview.html = html;
			}
		} catch (error) {
			if (panel !== undefined && shown === what) {
				void panel.webview.postMessage({ type: "error", message: error instanceof Error ? error.message : String(error) });
			}
		}
	});
}

// ---- the events of the shapes and controls of a sheet

/** What the page posts for the events of an object (`choose`: the menu's "…", not the default). */
interface ObjectEvent {
	sheet: string;
	codeName: string;
	/** Its name, as the name box shows it (none for some form controls). */
	object: string;
	/** A form control's VML shape (1025 for `_x0000_s1025`). */
	shape?: number;
	/** An ActiveX control: its Microsoft Forms type (`CommandButton`; empty for another control). */
	activex?: string;
	/** The macro assigned to it (`Sheet1.button1Click`). */
	macro: string;
	choose: boolean;
}

const samePath = (a: string, b: string): boolean => (process.platform === "win32" ? a.toLowerCase() === b.toLowerCase() : a === b);

/** The `objects/` folder of the project whose workbook is shown (its sheets' files), if it is. */
function sheetFolder(what: Shown): string | undefined {
	if (what.root === undefined) {
		return undefined;
	}
	const book = projectWorkbook(what.root);
	return book !== undefined && samePath(path.resolve(book), path.resolve(what.file)) ? path.join(what.root, "objects") : undefined;
}

/** The RAS files that hold macros: the sheets' files and the standard modules, with their modules. */
async function macroFiles(root: string): Promise<{ file: string; module: string }[]> {
	const base = new vscode.RelativePattern(root, "{objects/*.ras,src/**/*.ras}");
	const found = await vscode.workspace.findFiles(base, "**/{node_modules,.git,build,vba}/**", 2000);
	return found
		.map((uri) => uri.fsPath)
		.filter((file) => !/\.d\.ras$/i.test(file) && !/^ThisWorkbook\.ras$/i.test(path.basename(file)))
		.map((file) => ({ file, module: sheetOfRasFile(file) ?? moduleName(file) }))
		.sort((a, b) => a.module.localeCompare(b.module));
}

async function textOf(file: string): Promise<string> {
	const open = vscode.workspace.textDocuments.find((d) => samePath(d.uri.fsPath, file));
	if (open !== undefined) {
		return open.getText();
	}
	try {
		return await fs.promises.readFile(file, "utf8");
	} catch {
		return "";
	}
}

/** The document of a RAS file, made when there is none. */
async function documentOf(file: string): Promise<vscode.TextDocument> {
	const uri = vscode.Uri.file(file);
	try {
		return await vscode.workspace.openTextDocument(uri);
	} catch {
		const create = new vscode.WorkspaceEdit();
		create.createFile(uri, { ignoreIfExists: true });
		await vscode.workspace.applyEdit(create);
		return vscode.workspace.openTextDocument(uri);
	}
}

/** Adds text at the end of a document: `make` is given its text, the indentation and the newline. */
async function append(code: vscode.TextDocument, make: (text: string, indent: string, newline: string) => string): Promise<void> {
	const text = code.getText();
	const config = vscode.workspace.getConfiguration("editor", code.uri);
	const tabSize = config.get<number>("tabSize", 4);
	const fallback = config.get<boolean>("insertSpaces", true) ? " ".repeat(tabSize > 0 ? tabSize : 4) : "\t";
	const newline = code.eol === vscode.EndOfLine.CRLF ? "\r\n" : "\n";
	const edit = new vscode.WorkspaceEdit();
	edit.insert(code.uri, code.positionAt(text.length), make(text, bodyIndent(text, fallback), newline));
	await vscode.workspace.applyEdit(edit);
}

/** Shows a line of a document beside the grid, the cursor at its end. */
async function reveal(code: vscode.TextDocument, line: number): Promise<void> {
	const at = Math.min(Math.max(line, 0), code.lineCount - 1);
	const position = new vscode.Position(at, code.lineAt(at).text.length);
	await vscode.window.showTextDocument(code, { viewColumn: vscode.ViewColumn.Beside, selection: new vscode.Range(position, position) });
}

/** The first body line of the function `name` of a document (its declaration when it has none). */
function bodyLineOf(code: vscode.TextDocument, name: string): number | undefined {
	const found = macrosOf(code.getText()).find((m) => m.name.toLowerCase() === name.toLowerCase());
	if (found === undefined) {
		return undefined;
	}
	for (let line = found.line; line < code.lineCount; line++) {
		if (/\{\s*(\/\/.*)?$/.test(code.lineAt(line).text)) {
			return Math.min(line + 1, code.lineCount - 1);
		}
	}
	return found.line;
}

/** Opens (or adds) the handler of an event of an object; the page made again when its macro changed. */
async function openObjectEvent(what: Shown, m: ObjectEvent): Promise<string | undefined> {
	const folder = sheetFolder(what);
	if (folder === undefined || what.root === undefined) {
		throw new Error("イベントは、ras.toml の project.workbook のブックで追加できます。");
	}
	if (!m.codeName) {
		throw new Error(`シート ${m.sheet} のコード名が分かりません。`);
	}
	const sheetFile = path.join(folder, `${m.codeName}.ras`);
	if (m.activex !== undefined) {
		await openControlEvent(sheetFile, m);
		return undefined;
	}
	return assignMacro(what, what.root, sheetFile, m);
}

/** An ActiveX control's event: its `@on("<control>.<event>")` function in the sheet's file. */
async function openControlEvent(sheetFile: string, m: ObjectEvent): Promise<void> {
	const events = m.activex && m.object ? sheetControlEvents(m.activex, m.object) : [];
	if (events.length === 0) {
		throw new Error(`${m.object || "このコントロール"} は Microsoft Forms 2.0 のコントロールではないため、イベントを選べません。ハンドラー（@on("${m.object || "<名前>"}.<イベント>")）は、VBE が作るイベント プロシージャの宣言どおりの引数で手で書いてください。`);
	}
	let event: EventDescriptor | undefined = events.find((e) => e.isDefault) ?? events[0];
	if (m.choose) {
		const text = await textOf(sheetFile);
		const items = [...events]
			.sort((a, b) => Number(b.isDefault) - Number(a.isDefault))
			.map((e) => ({ label: e.name, description: `(${rasParams(e.signature)})`, detail: findHandler(text, m.object, e.name) ? "ハンドラーがあります" : undefined, event: e }));
		const picked = await vscode.window.showQuickPick(items, { placeHolder: `${m.object} の開くイベントを選んでください` });
		event = picked?.event;
	}
	if (event === undefined) {
		return;
	}
	const chosen = event;
	const code = await documentOf(sheetFile);
	let handler = findHandler(code.getText(), m.object, chosen.name);
	if (handler === undefined) {
		await append(code, (text, indent, newline) => handlerText(text, m.object, chosen, indent, newline));
		handler = findHandler(code.getText(), m.object, chosen.name);
	}
	await reveal(code, handler?.bodyLine ?? 0);
}

/** The file and line of a macro (`Sheet1.go`, `Report.reload`, `go`), if it is in a RAS file. */
async function findMacro(root: string, macro: string): Promise<{ file: string; name: string } | undefined> {
	// `'Book.xlsm'!Module.name`: the workbook's own.
	const bare = macro.replace(/^.*!/, "");
	const dot = bare.lastIndexOf(".");
	const [module, name] = dot < 0 ? [undefined, bare] : [bare.slice(0, dot), bare.slice(dot + 1)];
	for (const f of await macroFiles(root)) {
		if (module !== undefined && f.module.toLowerCase() !== module.toLowerCase()) {
			continue;
		}
		if (macrosOf(await textOf(f.file)).some((x) => x.name.toLowerCase() === name.toLowerCase())) {
			return { file: f.file, name };
		}
	}
	return undefined;
}

/** A shape's or a form control's macro: opened, or a new one (or one chosen) assigned to it. */
async function assignMacro(what: Shown, root: string, sheetFile: string, m: ObjectEvent): Promise<string | undefined> {
	const label = m.object || "このコントロール";
	const open = async (file: string, name: string) => {
		const code = await documentOf(file);
		await reveal(code, bodyLineOf(code, name) ?? 0);
	};
	if (!m.choose && m.macro) {
		const found = await findMacro(root, m.macro);
		if (found === undefined) {
			throw new Error(`${label} のマクロ ${m.macro} は、プロジェクトの RAS のファイルにありません（ブックの VBA だけにあるマクロかもしれません）。`);
		}
		await open(found.file, found.name);
		return undefined;
	}
	const fresh = newMacroName(await textOf(sheetFile), m.object, m.shape);
	let chosen: { file: string; module: string; name: string } | undefined = { file: sheetFile, module: m.codeName, name: fresh };
	let create = true;
	if (m.choose) {
		type Item = vscode.QuickPickItem & { target?: { file: string; module: string; name: string }; create?: boolean };
		const items: Item[] = [{ label: `$(add) 新しいマクロ: ${m.codeName}.${fresh}`, description: path.basename(sheetFile), target: chosen, create: true }];
		for (const f of await macroFiles(root)) {
			const macros = macrosOf(await textOf(f.file));
			if (macros.length > 0) {
				items.push({ label: f.module, kind: vscode.QuickPickItemKind.Separator });
			}
			for (const x of macros) {
				const full = `${f.module}.${x.name}`;
				items.push({ label: full, description: vscode.workspace.asRelativePath(f.file), detail: m.macro.toLowerCase() === full.toLowerCase() ? "登録されています" : undefined, target: { file: f.file, module: f.module, name: x.name } });
			}
		}
		const picked = await vscode.window.showQuickPick(items, { placeHolder: `${label} に登録するマクロを選んでください（クリックすると実行されます）` });
		if (picked?.target === undefined) {
			return undefined;
		}
		chosen = picked.target;
		create = picked.create === true;
	}
	const target = chosen;
	if (create) {
		const code = await documentOf(target.file);
		await append(code, (text, indent, newline) => {
			const separator = !text.trim() ? "" : /(\r\n|\r|\n)$/.test(text) ? newline : newline + newline;
			return `${separator}export function ${target.name}() {${newline}${indent}${newline}}${newline}`;
		});
		await code.save();
	}
	const full = `${target.module}.${target.name}`;
	const html = full.toLowerCase() === m.macro.toLowerCase() ? undefined : await edit(what, [{ op: "macro", sheet: m.sheet, object: m.object, shape: m.shape, macro: full }]);
	await open(target.file, target.name);
	return html;
}

/** Shows a workbook file. One a run saved is compared with the project's workbook it was read from. */
export async function showWorkbook(file: string): Promise<void> {
	const root = projectRoot(file);
	const source = root === undefined ? undefined : projectWorkbook(root);
	const saved = root !== undefined && path.relative(root, file).split(path.sep)[0] === "build";
	const sameBook = source !== undefined && path.parse(source).name.toLowerCase() === path.parse(file).name.toLowerCase();
	const before = saved && sameBook && source !== undefined && fs.existsSync(source) ? source : undefined;
	try {
		// A workbook a run saved is shown with what the run changed; the others are edited in place.
		const editable = before === undefined && !saved;
		show(await page(root, file, before, undefined, before, undefined, editable), `シート: ${path.basename(file)}`, true, { root, file, editable });
	} catch (error) {
		void vscode.window.showErrorMessage(`RAS: ${path.basename(file)} を表示できませんでした: ${error instanceof Error ? error.message : String(error)}`);
	}
}

/** Shows the project's workbook at the worksheet the RAS file `uri` is the module of. */
export async function showSheetOf(uri: vscode.Uri): Promise<void> {
	const file = uri.fsPath;
	const sheet = sheetOfRasFile(file);
	if (sheet === undefined) {
		void vscode.window.showInformationMessage(`RAS: ${path.basename(file)} はワークシートのファイルではありません（シートのファイルは objects/Sheet1.ras のように、objects/ にシートのコード名で置きます）。`);
		return;
	}
	const root = projectRoot(file);
	const book = root === undefined ? undefined : projectWorkbook(root);
	if (root === undefined || book === undefined || !fs.existsSync(book)) {
		void vscode.window.showErrorMessage("RAS: ras.toml の project.workbook のブックが見つかりません。");
		return;
	}
	try {
		show(await page(root, book, undefined, undefined, undefined, sheet, true), `シート: ${sheet}（${path.basename(book)}）`, true, { root, file: book, sheet, editable: true });
	} catch (error) {
		void vscode.window.showErrorMessage(`RAS: シート ${sheet} を表示できませんでした: ${error instanceof Error ? error.message : String(error)}`);
	}
}

/**
 * Keeps the `ras.sheetFiles` context key current: the paths of the RAS files of the workspace
 * that are the module of a worksheet (`objects/<CodeName>.ras`). The editor's title bar and the Explorer offer the sheet's grid on them.
 */
function trackSheetFiles(): vscode.Disposable {
	let timer: NodeJS.Timeout | undefined;
	const update = async () => {
		const found = await vscode.workspace.findFiles("**/objects/*.ras", "**/{node_modules,.git,build,vba}/**", 5000);
		const files = found.map((uri) => uri.fsPath).filter((file) => sheetOfRasFile(file) !== undefined);
		await vscode.commands.executeCommand("setContext", "ras.sheetFiles", files);
	};
	const refresh = () => {
		clearTimeout(timer);
		timer = setTimeout(() => void update(), 300);
	};
	const watcher = vscode.workspace.createFileSystemWatcher("**/objects/*.ras");
	void update();
	return vscode.Disposable.from(
		watcher,
		watcher.onDidCreate(refresh),
		watcher.onDidDelete(refresh),
		vscode.workspace.onDidChangeWorkspaceFolders(refresh),
		{ dispose: () => clearTimeout(timer) },
	);
}

/** The workbooks of the workspace a command without an argument offers: the project's and the saved ones. */
async function pickWorkbook(): Promise<string | undefined> {
	const found = await vscode.workspace.findFiles("**/*.{xlsx,xlsm}", "**/{node_modules,.git}/**", 200);
	const files = found.map((u) => u.fsPath).filter((f) => !path.basename(f).startsWith("~$"));
	if (files.length === 0) {
		void vscode.window.showInformationMessage("RAS: ワークスペースにブック（.xlsx / .xlsm）がありません。");
		return undefined;
	}
	const time = (f: string): number => {
		try {
			return fs.statSync(f).mtimeMs;
		} catch {
			return 0;
		}
	};
	// The newest first: what the last run saved.
	files.sort((a, b) => time(b) - time(a));
	const picked = await vscode.window.showQuickPick(
		files.map((f) => ({ label: path.basename(f), description: vscode.workspace.asRelativePath(path.dirname(f)), file: f })),
		{ placeHolder: "表示するブックを選択" },
	);
	return picked?.file;
}

interface Snapshot {
	books: { name: string; file: string; project: boolean; rev: number; source: string | null; active: boolean }[];
}

/** Follows a debug run: at every stop, the workbook of the stand-in as it is then. */
class Follower implements vscode.DebugAdapterTracker {
	private readonly dir = fs.mkdtempSync(path.join(os.tmpdir(), "ras-sheets-"));
	private shown: string | undefined;
	private busy = false;

	constructor(private readonly session: vscode.DebugSession) {}

	onDidSendMessage(message: { type?: string; event?: string; body?: { threadId?: number } }): void {
		if (message.type === "event" && message.event === "stopped" && message.body?.threadId !== undefined) {
			void this.refresh(message.body.threadId);
		}
	}

	onExit(): void {
		fs.rm(this.dir, { recursive: true, force: true }, () => undefined);
	}

	private async refresh(threadId: number): Promise<void> {
		if (this.busy || setting("debug.sheetView") === "off") {
			return;
		}
		this.busy = true;
		try {
			const current = path.join(this.dir, "current");
			const previous = path.join(this.dir, "previous");
			// What was shown at the last stop is what this stop is compared with.
			fs.rmSync(previous, { recursive: true, force: true });
			if (fs.existsSync(current)) {
				fs.renameSync(current, previous);
			}
			const frames = (await this.session.customRequest("stackTrace", { threadId, startFrame: 0, levels: 1 })) as { stackFrames?: { id: number }[] };
			const frameId = frames.stackFrames?.[0]?.id;
			const expression = `typeof $rasBookSnapshot === "function" ? $rasBookSnapshot(${JSON.stringify(current)}) : 0`;
			await this.session.customRequest("evaluate", { expression, frameId, context: "repl" });
			const list = path.join(current, "snapshot.json");
			if (!fs.existsSync(list)) {
				return;
			}
			const snapshot = JSON.parse(fs.readFileSync(list, "utf8")) as Snapshot;
			const book = snapshot.books.find((b) => b.active) ?? snapshot.books.find((b) => b.project) ?? snapshot.books[0];
			if (book === undefined) {
				return;
			}
			const key = `${book.name}#${book.rev}`;
			if (key === this.shown && panel !== undefined) {
				return;
			}
			const root = (this.session.configuration.cwd as string | undefined) ?? undefined;
			const last = path.join(previous, book.file);
			// The first time: against the file the workbook was read from.
			const before = fs.existsSync(last) ? last : book.source !== null && fs.existsSync(book.source) ? book.source : undefined;
			const source = book.source !== null && fs.existsSync(book.source) ? book.source : undefined;
			const html = await page(root, path.join(current, book.file), before, book.name, source);
			show(html, `シート: ${book.name}（デバッグ中）`, false, undefined);
			this.shown = key;
		} catch {
			// The program went on, or the session ended, before the answer: nothing to show.
		} finally {
			this.busy = false;
		}
	}
}

export function registerSheetView(): vscode.Disposable {
	return vscode.Disposable.from(
		vscode.commands.registerCommand("ras.book.show", async (resource?: unknown) => {
			const file = resource instanceof vscode.Uri ? resource.fsPath : await pickWorkbook();
			if (file !== undefined) {
				await showWorkbook(file);
			}
		}),
		vscode.commands.registerCommand("ras.sheet.show", async (resource?: unknown) => {
			const uri = resource instanceof vscode.Uri ? resource : vscode.window.activeTextEditor?.document.uri;
			if (uri !== undefined) {
				await showSheetOf(uri);
			}
		}),
		// After a sheet was added, copied or renamed in the workbook file: the grid, when it is open.
		vscode.commands.registerCommand("ras.sheet.refresh", async (resource?: unknown) => {
			if (panel !== undefined && resource instanceof vscode.Uri) {
				await showSheetOf(resource);
			}
		}),
		// The keys the grid takes (insert, delete, fill, ...) do nothing else while it has the focus.
		vscode.commands.registerCommand("ras.sheet.gridKey", () => undefined),
		trackSheetFiles(),
		vscode.debug.registerDebugAdapterTrackerFactory("*", {
			createDebugAdapterTracker(session) {
				// The debug runs this extension started (js-debug's sessions inherit the configuration).
				return session.configuration.__rasProcess === undefined ? undefined : new Follower(session);
			},
		}),
	);
}
