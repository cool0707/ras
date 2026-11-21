// The sheet panel: a workbook drawn as a grid (`ras book view`), for working without Excel.
//
// - "RAS: ブックを表示" shows a workbook file: the project's workbook, or one a run saved
//   (build/debug/workbooks, the workbooks of failed tests), with what the run changed marked.
// - While a debug run on the stand-in for Excel is stopped (a breakpoint, a step), the panel
//   shows the workbook as it is at that moment, with the cells changed since the last stop
//   marked: the run time writes the workbooks out when the editor asks (`$rasBookSnapshot`).
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
function page(root: string | undefined, file: string, before: string | undefined, title: string | undefined, drawings?: string, sheet?: string): Promise<string> {
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
	return new Promise((resolve, reject) => {
		cp.execFile(exe, args, { cwd: root ?? path.dirname(file), windowsHide: true, maxBuffer: 256 * 1024 * 1024 }, (error, stdout, stderr) => {
			if (error) {
				reject(new Error(stderr.trim() || error.message));
			} else {
				resolve(stdout);
			}
		});
	});
}

function show(html: string, title: string, reveal: boolean): void {
	if (panel === undefined) {
		panel = vscode.window.createWebviewPanel("ras.sheetView", title, { viewColumn: vscode.ViewColumn.Beside, preserveFocus: true }, { enableScripts: true, retainContextWhenHidden: true });
		panel.onDidDispose(() => {
			panel = undefined;
		});
	} else if (reveal) {
		panel.reveal(undefined, true);
	}
	panel.title = title;
	panel.webview.html = html;
}

/** Shows a workbook file. One a run saved is compared with the project's workbook it was read from. */
export async function showWorkbook(file: string): Promise<void> {
	const root = projectRoot(file);
	const source = root === undefined ? undefined : projectWorkbook(root);
	const saved = root !== undefined && path.relative(root, file).split(path.sep)[0] === "build";
	const sameBook = source !== undefined && path.parse(source).name.toLowerCase() === path.parse(file).name.toLowerCase();
	const before = saved && sameBook && source !== undefined && fs.existsSync(source) ? source : undefined;
	try {
		show(await page(root, file, before, undefined, before), `シート: ${path.basename(file)}`, true);
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
		show(await page(root, book, undefined, undefined, undefined, sheet), `シート: ${sheet}（${path.basename(book)}）`, true);
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
			show(html, `シート: ${book.name}（デバッグ中）`, false);
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
		trackSheetFiles(),
		vscode.debug.registerDebugAdapterTrackerFactory("*", {
			createDebugAdapterTracker(session) {
				// The debug runs this extension started (js-debug's sessions inherit the configuration).
				return session.configuration.__rasProcess === undefined ? undefined : new Follower(session);
			},
		}),
	);
}
