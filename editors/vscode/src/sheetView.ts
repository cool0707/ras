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
//   control's is an `@on("ボタン 1.Click")` function too, whose generated macro
//   (`Sheet1.ボタン1_click`) `ras book edit`'s `macro` assigns to it (`OnAction`).
// - On the project's workbook, the switch at the right end of the ribbon's tabs puts the page in
//   run mode: the program is built and runs on the stand-in for Excel (`ras-main.js --serve`,
//   a `RunSession`), and what is done on the page raises its events there (`Workbook.Open` when
//   the mode starts; `Activate`, `SelectionChange`, `Change`, `BeforeDoubleClick`,
//   `BeforeRightClick`; an ActiveX control's `Click` and the others; a shape's or a form
//   control's macro). The page is drawn from the workbook as the program left it; the workbook
//   file is not changed (the run saves what it changed under build/debug/workbooks/, as a debug
//   run does). Back in edit mode, the page shows the workbook file again.
// - "RAS: シートプレビューを表示" on the RAS file of a worksheet (`objects/Sheet1.ras`) shows the
//   project's workbook at that sheet: from the editor's title bar and context menu, and from
//   the Explorer's context menu.

import * as cp from "child_process";
import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import * as vscode from "vscode";
import { build as buildJs, channel as runOutput, standInRun } from "./debug";
import { FormHost } from "./formRunner/host";
import type { SheetResult } from "./sheetCommands";
import { sheetOfRasFile } from "./sheetFile";
import { findServerExecutable } from "./serverPath";
import { isShapeMacro, macrosOf, moduleName, newMacroName, shapeHandlerText, shapeMacro, sheetControlEvents } from "./sheetEvents";
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
	/** The run mode (the project's workbook): the program the page's events go to. */
	run?: RunSession;
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
 * controls are drawn (a workbook a run saved may not keep all of what they look like).
 */
function page(root: string | undefined, file: string, before: string | undefined, title: string | undefined, drawings?: string, sheet?: string, editable = false, run = false): Promise<string> {
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
	if (run) {
		args.push("--run");
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
			void shown?.run?.stop();
			panel = undefined;
			shown = undefined;
		});
		panel.webview.onDidReceiveMessage((message: unknown) => receive(message));
	} else if (reveal) {
		panel.reveal(undefined, true);
	}
	// Another workbook (or the same one shown afresh): the run mode ends.
	if (shown?.run !== undefined && shown.run !== what?.run) {
		void shown.run.stop();
		shown.run = undefined;
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

/** The families of the fonts installed (asked once a session; none when they cannot be listed). */
let fonts: Promise<string[]> | undefined;

function installedFonts(): Promise<string[]> {
	fonts ??= listFonts().catch(() => []);
	return fonts;
}

function listFonts(): Promise<string[]> {
	// Windows: the names in the user's language ("游ゴシック"), as Excel shows them.
	const [command, args, split] =
		process.platform === "win32"
			? [
					"powershell.exe",
					[
						"-NoProfile",
						"-NonInteractive",
						"-Command",
						"[Console]::OutputEncoding = [Text.Encoding]::UTF8; Add-Type -AssemblyName System.Drawing; (New-Object System.Drawing.Text.InstalledFontCollection).Families | ForEach-Object { $_.Name }",
					],
					false,
				]
			: process.platform === "darwin"
				? [
						"osascript",
						["-l", "JavaScript", "-e", 'ObjC.import("AppKit"); ObjC.deepUnwrap($.NSFontManager.sharedFontManager.availableFontFamilies).join("\\n")'],
						false,
					]
				: // fontconfig: a family's names in each language, separated by commas.
					["fc-list", ["--format", "%{family}\\n"], true];
	return new Promise((resolve, reject) => {
		cp.execFile(command, args as string[], { encoding: "utf8", timeout: 20000, maxBuffer: 16 * 1024 * 1024, windowsHide: true }, (error, stdout) => {
			if (error) {
				reject(error);
				return;
			}
			const names = new Set<string>();
			for (const line of stdout.split(/\r?\n/)) {
				for (const name of split ? line.split(",") : [line]) {
					const n = name.trim();
					if (n !== "") {
						names.add(n);
					}
				}
			}
			resolve([...names]);
		});
	});
}

/** One edit at a time, in the order the page posted them. */
let queue: Promise<void> = Promise.resolve();

function receive(message: unknown): void {
	const type = typeof message === "object" && message !== null ? (message as { type?: unknown }).type : undefined;
	if (type === "ready") {
		// The page asks whether its objects have events: those of the project's workbook do.
		if (panel !== undefined && shown !== undefined && shown.editable && sheetFolder(shown) !== undefined) {
			void panel.webview.postMessage({ type: "events", on: true, run: shown.run !== undefined, session: shown.run?.id });
		}
		return;
	}
	if (type === "fonts") {
		// The fonts installed on the system, for the font box of the ribbon.
		void installedFonts().then((names) => {
			if (panel !== undefined) {
				void panel.webview.postMessage({ type: "fonts", names });
			}
		});
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
	if (type !== "edit" && type !== "undo" && type !== "redo" && type !== "event" && type !== "sheet" && type !== "mode" && type !== "run") {
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
			if (what.run !== undefined && type !== "mode" && type !== "run") {
				throw new Error("実行モードでは編集できません。リボンの右端のスイッチで「編集」モードに切り替えてください。");
			}
			const html =
				type === "mode"
					? await switchMode(what, (message as { run?: unknown }).run === true)
					: type === "run"
						? await runEvent(what, message as RunMessage)
						: type === "event"
							? await openObjectEvent(what, message as ObjectEvent)
							: type === "sheet"
								? await sheetAction(what, message as SheetMessage)
								: type === "edit"
									? await edit(what, (message as { ops?: unknown }).ops ?? [])
									: await travel(what, type === "redo");
			// (Another workbook may be shown by now.)
			if (html !== undefined && panel !== undefined && shown === what) {
				panel.webview.html = html;
			} else if (html === undefined && (type === "sheet" || type === "mode") && panel !== undefined && shown === what) {
				// Nothing was done: the page takes edits again.
				void panel.webview.postMessage({ type: "idle" });
			}
		} catch (error) {
			if (panel !== undefined && shown === what) {
				void panel.webview.postMessage({ type: "error", message: error instanceof Error ? error.message : String(error) });
			}
		}
	});
}

// ---- the run mode: the program, run on the stand-in for Excel, gets the page's events

let extensionContext: vscode.ExtensionContext | undefined;

/** What the page posts in run mode (`$rt.serve` of the run time says what each request does). */
interface RunMessage {
	id: number;
	op: string;
	[key: string]: unknown;
}

/** What the program answers a request with. */
interface RunReply {
	id: number;
	ok: boolean;
	error?: { code: number; description: string; file: string | null; line: number };
	/** The errors the event handlers did not catch (the program goes on, as in Excel). */
	errors?: { handler: string; code: number; description: string; file: string | null; line: number }[];
	/** The sheet active and its selection, as the program left them. */
	where?: { sheet: string | null; cell: string | null; selection: string | null };
	/** The revision of the project's workbook (it changes with each change of its cells). */
	rev?: number;
	cancel?: boolean;
	snapshotError?: string;
}

/** What starts an answer of `ras-main.js --serve` on its standard output. */
const SERVE_MARK = "\u001eRAS ";
let runIds = 0;

/** `ras-main.js --serve`: what is done on the page goes to it, and it says what changed. */
class RunSession {
	readonly id = ++runIds;
	/** Where the program writes the workbooks after each request (for the page). */
	readonly dir = fs.mkdtempSync(path.join(os.tmpdir(), "ras-run-"));
	rev = -1;
	sheet: string | undefined;
	/** Called when the program ends by itself (not by `stop`). */
	onExit: ((code: number | null) => void) | undefined;
	private child: cp.ChildProcess | undefined;
	private readonly waiting = new Map<number, (reply: RunReply) => void>();
	private next = 1;
	private stopping = false;
	private formHost: FormHost | undefined;

	private constructor(readonly root: string) {}

	/** Builds the program and starts it. */
	static async start(context: vscode.ExtensionContext, root: string): Promise<RunSession> {
		await buildJs(root);
		const session = new RunSession(root);
		try {
			await session.spawn(context);
		} catch (error) {
			await session.stop();
			throw error;
		}
		return session;
	}

	private async spawn(context: vscode.ExtensionContext): Promise<void> {
		const { runtime, env, book } = standInRun(context);
		if (book === undefined) {
			throw new Error("Excel の代わりにブックを読み書きするアドオン（ras_book.node）が見つからないため、実行モードにできません。設定 ras.debug.bookAddon で場所を指定できます。");
		}
		const debugDir = path.join(this.root, "build", "debug");
		let forms: string[] = [];
		try {
			forms = (JSON.parse(fs.readFileSync(path.join(debugDir, "ras-debug.json"), "utf8")) as { forms?: string[] }).forms ?? [];
		} catch {
			// (No forms.)
		}
		// The forms the handlers show are drawn in panels, as in a debug run.
		if (forms.length > 0) {
			this.formHost = await FormHost.start(context);
			env.RAS_FORM_HOST = String(this.formHost.port);
		}
		const out = runOutput();
		out.appendLine("> 実行モード（シートプレビュー）");
		const child = cp.spawn(runtime, ["--enable-source-maps", path.join(debugDir, "ras-main.js"), "--serve", this.dir], { cwd: this.root, env, windowsHide: true });
		this.child = child;
		let pending = "";
		child.stdout?.setEncoding("utf8");
		child.stdout?.on("data", (chunk: string) => {
			pending += chunk;
			const lines = pending.split(/\r?\n/);
			pending = lines.pop() ?? "";
			for (const line of lines) {
				this.line(line, out);
			}
		});
		child.stderr?.setEncoding("utf8");
		child.stderr?.on("data", (chunk: string) => out.append(chunk));
		child.stdin?.on("error", () => undefined);
		child.on("error", (error) => out.appendLine(`[ras] 実行モードのプロセスを起動できません: ${error.message}`));
		child.on("exit", (code) => {
			this.child = undefined;
			for (const done of this.waiting.values()) {
				done({ id: 0, ok: false, error: { code: 0, description: `プログラムが終了しました（終了コード ${code ?? "?"}）。`, file: null, line: 0 } });
			}
			this.waiting.clear();
			this.formHost?.dispose();
			this.formHost = undefined;
			if (!this.stopping) {
				this.onExit?.(code);
			}
		});
	}

	/** A line of the program's output: an answer, or what it printed. */
	private line(line: string, out: vscode.OutputChannel): void {
		const at = line.indexOf(SERVE_MARK);
		if (at >= 0) {
			if (at > 0) {
				out.appendLine(line.slice(0, at));
			}
			try {
				const reply = JSON.parse(line.slice(at + SERVE_MARK.length)) as RunReply;
				const done = this.waiting.get(reply.id);
				this.waiting.delete(reply.id);
				done?.(reply);
			} catch {
				// (Not an answer after all.)
			}
			return;
		}
		out.appendLine(line);
		// The stand-in answers a message box with OK and prints what it showed: shown here too.
		const box = /^\[MsgBox\] (.*)$/.exec(line);
		if (box !== null) {
			void vscode.window.showInformationMessage(box[1]);
		}
	}

	request(m: Record<string, unknown>): Promise<RunReply> {
		const stdin = this.child?.stdin;
		if (stdin === null || stdin === undefined || !stdin.writable) {
			return Promise.resolve({ id: 0, ok: false, error: { code: 0, description: "実行モードのプログラムは終了しています。", file: null, line: 0 } });
		}
		const id = this.next++;
		return new Promise((resolve) => {
			this.waiting.set(id, resolve);
			stdin.write(`${JSON.stringify({ ...m, id })}\n`);
		});
	}

	/** Ends the program: it saves the workbooks it changed, as a run does. */
	async stop(): Promise<void> {
		this.stopping = true;
		const child = this.child;
		if (child !== undefined) {
			await new Promise<void>((resolve) => {
				const timer = setTimeout(() => {
					child.kill();
					resolve();
				}, 5000);
				child.once("exit", () => {
					clearTimeout(timer);
					resolve();
				});
				child.stdin?.end();
			});
		}
		this.formHost?.dispose();
		this.formHost = undefined;
		fs.rm(this.dir, { recursive: true, force: true }, () => undefined);
	}

	/** The project's workbook as the program has it now (the last answer wrote it). */
	snapshot(): string | undefined {
		try {
			const list = JSON.parse(fs.readFileSync(path.join(this.dir, "snapshot.json"), "utf8")) as { books: { file: string; project: boolean }[] };
			const own = list.books.find((b) => b.project);
			return own === undefined ? undefined : path.join(this.dir, own.file);
		} catch {
			return undefined;
		}
	}
}

/** What the program reported: the error that ended a request, those of the handlers. */
function reportRun(reply: RunReply): void {
	const at = (file: string | null, line: number) => (file ? `（${file}:${line}）` : "");
	if (!reply.ok && reply.error !== undefined) {
		void vscode.window.showErrorMessage(`RAS: 実行時エラー ${reply.error.code ? `${reply.error.code}: ` : ""}${reply.error.description}${at(reply.error.file, reply.error.line)}`);
	}
	for (const e of reply.errors ?? []) {
		void vscode.window.showErrorMessage(`RAS: ${e.handler}: 実行時エラー ${e.code}: ${e.description}${at(e.file, e.line)}`);
	}
	if (reply.snapshotError !== undefined) {
		runOutput().appendLine(`[ras] ブックを書き出せませんでした: ${reply.snapshotError}`);
	}
}

/** The page of the run mode: the project's workbook as the program left it, with its drawings. */
async function runPage(what: Shown, session: RunSession): Promise<string> {
	const file = session.snapshot();
	if (file === undefined) {
		throw new Error("実行中のブックを読み取れませんでした。");
	}
	return page(what.root, file, undefined, path.basename(what.file), what.file, session.sheet, true, true);
}

/** The page's switch: run mode on (the program started) or off (the workbook file again). */
async function switchMode(what: Shown, run: boolean): Promise<string | undefined> {
	if (run) {
		if (what.run !== undefined) {
			return undefined;
		}
		if (sheetFolder(what) === undefined || what.root === undefined) {
			throw new Error("実行モードは、ras.toml の project.workbook のブックで使えます。");
		}
		if (!vscode.workspace.isTrusted) {
			throw new Error("ワークスペースを信頼するまで、プログラムは実行できません。");
		}
		if (extensionContext === undefined) {
			throw new Error("拡張機能の準備ができていません。");
		}
		const context = extensionContext;
		const root = what.root;
		await vscode.workspace.saveAll(false);
		const session = await vscode.window.withProgress({ location: vscode.ProgressLocation.Window, title: "RAS: 実行モードを開始しています…" }, () => RunSession.start(context, root));
		const reply = await session.request({ op: "start" });
		reportRun(reply);
		if (!reply.ok) {
			await session.stop();
			return undefined;
		}
		session.rev = reply.rev ?? 0;
		session.sheet = reply.where?.sheet ?? undefined;
		what.run = session;
		session.onExit = (code) => {
			if (what.run !== session) {
				return;
			}
			what.run = undefined;
			void vscode.window.showWarningMessage(`RAS: 実行モードのプログラムが終了しました（終了コード ${code ?? "?"}）。編集モードに戻ります。`);
			queue = queue
				.then(async () => {
					if (panel !== undefined && shown === what) {
						panel.webview.html = await page(what.root, what.file, undefined, undefined, undefined, session.sheet, true);
					}
				})
				.catch(() => undefined);
		};
		return runPage(what, session);
	}
	const session = what.run;
	if (session === undefined) {
		return undefined;
	}
	what.run = undefined;
	await session.stop();
	runOutput().appendLine("（実行モードを終了しました。プログラムが変えたブックは build/debug/workbooks/ に保存されます）");
	// The page shows the sheet it showed last.
	what.sheet = undefined;
	return page(what.root, what.file, undefined, undefined, undefined, undefined, true);
}

/** What was done on the page in run mode: its events raised, the page drawn again when the workbook changed. */
async function runEvent(what: Shown, m: RunMessage): Promise<string | undefined> {
	const session = what.run;
	if (session === undefined) {
		throw new Error("実行モードではありません。");
	}
	const { type: _type, id, ...request } = m;
	const reply = await session.request(request);
	reportRun(reply);
	const sheet = reply.where?.sheet ?? undefined;
	if (reply.rev !== undefined && (reply.rev !== session.rev || (sheet !== undefined && sheet !== session.sheet))) {
		session.rev = reply.rev;
		session.sheet = sheet;
		return runPage(what, session);
	}
	if (panel !== undefined && shown === what) {
		void panel.webview.postMessage({ type: "runDone", id, cancel: reply.cancel === true });
	}
	return undefined;
}

// ---- the sheets: inserted, removed, renamed and copied from the tabs of the page

/** What the tabs of the page post. */
interface SheetMessage {
	action: "add" | "copy" | "remove" | "rename";
	sheet: string;
	name?: string;
	codeName?: string;
	color?: string | null;
	/** A copy's place ("シートの移動またはコピー"): before this sheet, or after it. */
	before?: string;
	after?: string;
}

/**
 * Changes the sheets as the page asked, and gives the page made again (undefined: nothing was
 * done). The project's workbook changes with the files of its sheets in objects/ (the commands
 * "RAS: シートを追加 / ..."), another workbook as the other edits change it (`ras book edit`).
 */
async function sheetAction(what: Shown, m: SheetMessage): Promise<string | undefined> {
	if (!["add", "copy", "remove", "rename"].includes(m.action)) {
		return undefined;
	}
	// The page shows the sheet it remembers (the new one, the one renamed), not the one asked
	// for first, which may be gone.
	what.sheet = undefined;
	const color = typeof m.color === "string" && m.color !== "" ? m.color : undefined;
	if (sheetFolder(what) === undefined) {
		if (m.action === "remove") {
			const remove = "削除する";
			const choice = await vscode.window.showWarningMessage(`シート「${m.sheet}」を ${path.basename(what.file)} から削除します。`, { modal: true }, remove);
			if (choice !== remove) {
				return undefined;
			}
		}
		const ops: unknown[] = [{ op: "sheet", action: m.action, sheet: m.sheet, name: m.name ?? "", codeName: m.action === "add" ? (m.codeName ?? "") : "", before: m.before ?? "", after: m.after ?? "" }];
		if (m.action === "add" && m.name !== undefined && color !== undefined) {
			ops.push({ op: "tabColor", sheet: m.name, color });
		}
		return edit(what, ops);
	}
	const result = await vscode.commands.executeCommand<SheetResult | undefined>(`ras.sheet.${m.action}`, { sheet: m.sheet, name: m.name, codeName: m.codeName, before: m.before, after: m.after, fromGrid: true });
	if (result === undefined || !result.ok) {
		return undefined;
	}
	// (The command may have shown the page again, at the new sheet.)
	const now = shown !== undefined && shown !== what && samePath(shown.file, what.file) ? shown : what;
	let html: string | undefined;
	if (m.action === "add" && color !== undefined && result.name !== undefined) {
		if (result.live === true) {
			void vscode.window.showWarningMessage("RAS: Excel で開いているブックのシート見出しの色は、Excel で設定してください。");
		} else {
			html = await edit(now, [{ op: "tabColor", sheet: result.name, color }]);
		}
	}
	if (now !== what) {
		if (html !== undefined && panel !== undefined && shown === now) {
			panel.webview.html = html;
		}
		return undefined;
	}
	return html ?? page(what.root, what.file, undefined, undefined, undefined, undefined, true);
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
	/** What runs it, as the page names it: "クリック", "選択" (a list or combo box), "変更" (a spin button, a scroll bar). */
	eventName?: string;
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
	await vscode.window.showTextDocument(code, { viewColumn: codeColumn(code), selection: new vscode.Range(position, position) });
}

/**
 * Where the code opens: the group already showing it, else the group of the code next to the
 * panel (the one left of it first), and only without another group a new one beside the panel
 * (`Beside` from the panel would split the window into a third group).
 */
function codeColumn(code: vscode.TextDocument): vscode.ViewColumn {
	const shown = vscode.window.visibleTextEditors.find((e) => e.document.uri.toString() === code.uri.toString());
	if (shown?.viewColumn !== undefined) {
		return shown.viewColumn;
	}
	const own = panel?.viewColumn;
	const others = vscode.window.tabGroups.all.map((g) => g.viewColumn).filter((c) => c !== own);
	if (others.length === 0) {
		return vscode.ViewColumn.Beside;
	}
	if (own === undefined) {
		return others[0];
	}
	const left = others.filter((c) => c < own);
	return left.length > 0 ? Math.max(...left) : Math.min(...others);
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

/**
 * A shape's or a form control's event: its `@on("<name>.Click")` handler in the sheet's file,
 * added when there is none, whose macro (`Sheet1.ボタン1_click`, spec §12.2.1) is assigned to it
 * (`OnAction`); or the macro assigned to it, or one chosen from the RAS files.
 */
async function assignMacro(what: Shown, root: string, sheetFile: string, m: ObjectEvent): Promise<string | undefined> {
	const label = m.object || "このコントロール";
	const ev = m.eventName ?? "クリック";
	const open = async (file: string, name: string) => {
		const code = await documentOf(file);
		await reveal(code, bodyLineOf(code, name) ?? 0);
	};
	const assign = async (full: string) => (full.toLowerCase() === m.macro.replace(/^.*!/, "").toLowerCase() ? undefined : edit(what, [{ op: "macro", sheet: m.sheet, object: m.object, shape: m.shape, macro: full }]));
	// Its `@on` handler (an object with a name: the handler names it).
	const handler = async () => {
		const code = await documentOf(sheetFile);
		let found = findHandler(code.getText(), m.object, "Click");
		if (found === undefined) {
			await append(code, (text, indent, newline) => shapeHandlerText(text, m.object, indent, newline));
			await code.save();
			found = findHandler(code.getText(), m.object, "Click");
		}
		const html = await assign(shapeMacro(m.codeName, m.object));
		await reveal(code, found?.bodyLine ?? 0);
		return html;
	};
	const generated = m.object !== "" && m.macro !== "" && isShapeMacro(m.macro, m.codeName, m.object);
	if (!m.choose && m.macro && !generated) {
		const found = await findMacro(root, m.macro);
		if (found === undefined) {
			throw new Error(`${label} の${ev}イベントに登録された ${m.macro} は、プロジェクトの RAS のファイルにありません（ブックの VBA だけにある関数かもしれません）。`);
		}
		await open(found.file, found.name);
		return undefined;
	}
	if (!m.choose && m.object !== "") {
		return handler();
	}
	type Target = { file: string; module: string; name: string };
	const fresh = newMacroName(await textOf(sheetFile), m.object, m.shape);
	let chosen: Target | undefined = { file: sheetFile, module: m.codeName, name: fresh };
	let create = true;
	if (m.choose) {
		type Item = vscode.QuickPickItem & { target?: Target; create?: boolean; handler?: boolean };
		const items: Item[] =
			m.object !== ""
				? [
						{
							label: `$(symbol-event) @on("${m.object}.Click")`,
							description: path.basename(sheetFile),
							detail: findHandler(await textOf(sheetFile), m.object, "Click") ? `ハンドラーがあります${generated ? "（登録されています）" : ""}` : "ハンドラーを追加します",
							handler: true,
						},
					]
				: [{ label: `$(add) 新しい関数: ${m.codeName}.${fresh}`, description: path.basename(sheetFile), target: chosen, create: true }];
		for (const f of await macroFiles(root)) {
			const macros = macrosOf(await textOf(f.file));
			if (macros.length > 0) {
				items.push({ label: f.module, kind: vscode.QuickPickItemKind.Separator });
			}
			for (const x of macros) {
				const full = `${f.module}.${x.name}`;
				items.push({ label: full, description: vscode.workspace.asRelativePath(f.file), detail: m.macro.replace(/^.*!/, "").toLowerCase() === full.toLowerCase() ? "登録されています" : undefined, target: { file: f.file, module: f.module, name: x.name } });
			}
		}
		const picked = await vscode.window.showQuickPick(items, { placeHolder: `${label} の${ev}イベントに登録するハンドラー・関数を選んでください` });
		if (picked?.handler) {
			return handler();
		}
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
	const html = await assign(`${target.module}.${target.name}`);
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
		show(await page(root, file, before, undefined, before, undefined, editable), `シートプレビュー: ${path.basename(file)}`, true, { root, file, editable });
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
		show(await page(root, book, undefined, undefined, undefined, sheet, true), `シートプレビュー: ${sheet}（${path.basename(book)}）`, true, { root, file: book, sheet, editable: true });
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
			show(html, `シートプレビュー: ${book.name}（デバッグ中）`, false, undefined);
			this.shown = key;
		} catch {
			// The program went on, or the session ended, before the answer: nothing to show.
		} finally {
			this.busy = false;
		}
	}
}

export function registerSheetView(context: vscode.ExtensionContext): vscode.Disposable {
	extensionContext = context;
	return vscode.Disposable.from(
		{ dispose: () => void shown?.run?.stop() },
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
