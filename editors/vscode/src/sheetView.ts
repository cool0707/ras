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
	if (type !== "edit" && type !== "undo" && type !== "redo") {
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
			const html = type === "edit" ? await edit(what, (message as { ops?: unknown }).ops ?? []) : await travel(what, type === "redo");
			// (Another workbook may be shown by now.)
			if (panel !== undefined && shown === what) {
				panel.webview.html = html;
			}
		} catch (error) {
			if (panel !== undefined && shown === what) {
				void panel.webview.postMessage({ type: "error", message: error instanceof Error ? error.message : String(error) });
			}
		}
	});
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
