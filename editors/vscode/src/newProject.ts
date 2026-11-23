// "New RAS project": runs `ras init` in a folder, with a workbook (a new empty one, or one the
// user already has), writes the sample code into the workbook (`ras pack`) and opens the result:
// the language server starts, src/main.ras and the workbook's sheet preview open.

import * as cp from "node:child_process";
import * as fs from "node:fs";
import * as path from "node:path";
import * as vscode from "vscode";
import { errorMessage } from "./server";
import { findServerExecutable } from "./serverPath";
import { isNewProjectFolder, projectBookName, setWorkbook, workbookFileNameError, WORKBOOK_EXTENSIONS } from "./newProjectFiles";

/** globalState: a project created for a folder that the window is reopened on, to open its files after the reload. */
const PENDING_OPEN = "ras.newProject.pendingOpen";

interface PendingOpen {
	root: string;
	book: string;
}

let output: vscode.OutputChannel | undefined;

interface RunResult {
	ok: boolean;
	stdout: string;
	stderr: string;
}

function runRas(exe: string, args: string[], cwd: string): Promise<RunResult> {
	const channel = (output ??= vscode.window.createOutputChannel("RAS"));
	channel.appendLine(`> ras ${args.join(" ")}`);
	return new Promise((resolve) =>
		cp.execFile(exe, args, { cwd, windowsHide: true }, (error, stdout, stderr) => {
			for (const text of [stdout, stderr]) {
				if (text.trim() !== "") channel.appendLine(text.trimEnd());
			}
			resolve({ ok: !error, stdout, stderr });
		}),
	);
}

/** The `error: …` line of a failed run. */
function failure(run: RunResult, fallback: string): string {
	return run.stderr.split(/\r?\n/).find((l) => l.startsWith("error:"))?.slice("error:".length).trim() || fallback;
}

/** The folder the project goes to: the clicked folder, a workspace folder without a project, else one the user picks. */
async function targetFolder(resource: vscode.Uri | undefined): Promise<string | undefined> {
	if (resource?.scheme === "file") {
		try {
			return fs.statSync(resource.fsPath).isDirectory() ? resource.fsPath : path.dirname(resource.fsPath);
		} catch {
			// Fall through.
		}
	}
	const folders = (vscode.workspace.workspaceFolders ?? []).filter((f) => f.uri.scheme === "file" && !fs.existsSync(path.join(f.uri.fsPath, "ras.toml")));
	if (folders.length === 1) return folders[0].uri.fsPath;
	if (folders.length > 1) {
		const picked = await vscode.window.showQuickPick(
			folders.map((f) => ({ label: f.name, description: f.uri.fsPath, root: f.uri.fsPath })),
			{ title: "新しい RAS プロジェクト", placeHolder: "プロジェクトを作成するフォルダーを選んでください" },
		);
		return picked?.root;
	}
	const picked = await vscode.window.showOpenDialog({
		title: "新しい RAS プロジェクトを作成するフォルダー",
		canSelectFiles: false,
		canSelectFolders: true,
		canSelectMany: false,
		openLabel: "ここに作成",
	});
	return picked?.[0]?.fsPath;
}

function workbooksIn(dir: string): string[] {
	try {
		return fs
			.readdirSync(dir, { withFileTypes: true })
			.filter((e) => e.isFile() && !e.name.startsWith("~$") && WORKBOOK_EXTENSIONS.includes(path.extname(e.name).toLowerCase()))
			.map((e) => e.name)
			.sort();
	} catch {
		return [];
	}
}

type BookChoice = { kind: "new" } | { kind: "existing"; file: string };

/** Which workbook the project uses: a new one, one in the folder, or one copied in from elsewhere. */
async function chooseBook(root: string): Promise<BookChoice | undefined> {
	type Item = vscode.QuickPickItem & { choice?: BookChoice; browse?: boolean };
	const items: Item[] = [
		{ label: "$(new-file) 新しい空のブックを作成", description: ".xlsm", choice: { kind: "new" } },
		...workbooksIn(root).map((name): Item => ({ label: `$(file) ${name}`, description: "このフォルダーのブックを使う", choice: { kind: "existing", file: path.join(root, name) } })),
		{ label: "$(folder-opened) ほかの場所のブックを取り込む…", description: "フォルダーへコピーして使う", browse: true },
	];
	const picked = await vscode.window.showQuickPick(items, { title: "新しい RAS プロジェクト: ブック", placeHolder: "プロジェクトで使うブックを選んでください" });
	if (!picked) return undefined;
	if (!picked.browse) return picked.choice;
	const files = await vscode.window.showOpenDialog({
		title: "取り込むブック",
		canSelectMany: false,
		openLabel: "取り込む",
		filters: { "Excel ブック": WORKBOOK_EXTENSIONS.map((e) => e.slice(1)) },
	});
	const file = files?.[0]?.fsPath;
	return file === undefined ? undefined : { kind: "existing", file };
}

async function askNewBookName(root: string): Promise<string | undefined> {
	const name = await vscode.window.showInputBox({
		title: "新しい RAS プロジェクト: ブックの名前",
		prompt: `${root} に作成するマクロ有効ブックの名前`,
		value: projectBookName(path.basename(root)),
		valueSelection: [0, projectBookName(path.basename(root)).length - ".xlsm".length],
		validateInput: (value) => {
			const file = withBookExtension(value.trim());
			return workbookFileNameError(file) ?? (fs.existsSync(path.join(root, file)) ? "同じ名前のファイルがすでにあります。" : undefined);
		},
	});
	return name === undefined ? undefined : withBookExtension(name.trim());
}

function withBookExtension(name: string): string {
	return path.extname(name) === "" ? `${name}.xlsm` : name;
}

async function newProject(resource: vscode.Uri | undefined): Promise<void> {
	if (!vscode.workspace.isTrusted) {
		void vscode.window.showErrorMessage("RAS: ワークスペースを信頼してからプロジェクトを作成してください。");
		return;
	}
	const root = await targetFolder(resource);
	if (root === undefined) return;
	if (fs.existsSync(path.join(root, "ras.toml"))) {
		void vscode.window.showErrorMessage(`RAS: ${root} にはすでに ras.toml があります。`);
		return;
	}
	const choice = await chooseBook(root);
	if (!choice) return;
	const newBook = choice.kind === "new" ? await askNewBookName(root) : undefined;
	if (choice.kind === "new" && newBook === undefined) return;

	const exe = findServerExecutable((vscode.workspace.getConfiguration("ras").get<string>("path") ?? "").trim(), root);
	if (exe === undefined) {
		const openSettings = "設定を開く";
		if ((await vscode.window.showErrorMessage("RAS: ras 実行ファイルが見つかりません。ras をインストールするか、設定 ras.path に実行ファイルのパスを指定してください。", openSettings)) === openSettings) {
			await vscode.commands.executeCommand("workbench.action.openSettings", "ras.path");
		}
		return;
	}

	const book = await vscode.window.withProgress({ location: vscode.ProgressLocation.Notification, title: "RAS: プロジェクトを作成しています…" }, () =>
		createProject(exe, root, choice, newBook),
	);
	if (book === undefined) return;
	await openProject(root, book);
}

/**
 * Creates the project and returns the workbook's file name (relative to `root`), or `undefined`
 * after reporting a failure. The workbook gets the sample code (`ras pack`); a failed pack is a
 * warning only: the project is there.
 */
async function createProject(exe: string, root: string, choice: BookChoice, newBook: string | undefined): Promise<string | undefined> {
	const showOutput = "出力を表示";
	const fail = async (message: string) => {
		if ((await vscode.window.showErrorMessage(`RAS: ${message}`, showOutput)) === showOutput) output?.show(true);
		return undefined;
	};

	// The workbook the project uses, in the folder.
	let source: string | undefined; // The workbook as found or copied (an .xlsx is converted from it).
	let copied = false;
	if (choice.kind === "existing") {
		const name = path.basename(choice.file);
		const inRoot = path.join(root, name);
		if (path.resolve(path.dirname(choice.file)).toLowerCase() !== path.resolve(root).toLowerCase()) {
			if (fs.existsSync(inRoot)) return fail(`${root} には同じ名前のファイル ${name} がすでにあります。`);
			await fs.promises.copyFile(choice.file, inRoot);
			copied = true;
		}
		source = name;
	}
	const ext = source === undefined ? ".xlsm" : path.extname(source).toLowerCase();
	const convert = ext === ".xlsx";
	const book = newBook ?? (convert ? `${path.basename(source!, path.extname(source!))}.xlsm` : source!);
	if (convert && fs.existsSync(path.join(root, book))) {
		if (copied) await fs.promises.rm(path.join(root, source!), { force: true });
		return fail(`${book} がすでにあるため、${source} をマクロ有効ブックに変換できません。`);
	}

	// ras.toml, src/main.ras, AGENTS.md, .gitignore (and the new empty workbook).
	const init = await runRas(exe, ["init", root, ...(convert ? [] : ["--workbook", book])], root);
	if (!init.ok) {
		if (copied) await fs.promises.rm(path.join(root, source!), { force: true });
		return fail(failure(init, "プロジェクトを作成できませんでした。"));
	}
	if (convert) {
		const toml = path.join(root, "ras.toml");
		await fs.promises.writeFile(toml, setWorkbook(await fs.promises.readFile(toml, "utf8"), book), "utf8");
	}

	// The sample code into the workbook; an existing workbook's own VBA is taken over into the project.
	const pack = await runRas(exe, ["--project", root, "pack", ...(convert ? [source!, "--convert"] : [])], root);
	if (convert && copied) await fs.promises.rm(path.join(root, source!), { force: true });
	if (!pack.ok) {
		if (convert) return fail(failure(pack, `${source} をマクロ有効ブックに変換できませんでした。`));
		void vscode.window.showWarningMessage(`RAS: プロジェクトは作成しましたが、ブックへの書き込みに失敗しました: ${failure(pack, "ビルド エラーがあります。")}`, showOutput).then((c) => {
			if (c === showOutput) output?.show(true);
		});
	}
	return book;
}

/** Opens what was created: in this window when the folder is in the workspace, else reopens the window on the folder. */
async function openProject(root: string, book: string): Promise<void> {
	const inWorkspace = (vscode.workspace.workspaceFolders ?? []).some((f) => f.uri.scheme === "file" && isSameOrInside(root, f.uri.fsPath));
	if (inWorkspace) {
		await showCreated(root, book);
		return;
	}
	const pending: PendingOpen = { root, book };
	await context?.globalState.update(PENDING_OPEN, pending);
	// An empty window becomes the project's; a window with other folders stays as it is.
	await vscode.commands.executeCommand("vscode.openFolder", vscode.Uri.file(root), { forceNewWindow: (vscode.workspace.workspaceFolders ?? []).length > 0 });
}

function isSameOrInside(dir: string, folder: string): boolean {
	const relative = path.relative(folder, dir);
	return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative));
}

async function showCreated(root: string, book: string): Promise<void> {
	const main = path.join(root, "src", "main.ras");
	if (fs.existsSync(main)) await vscode.window.showTextDocument(vscode.Uri.file(main), { preview: false });
	const bookPath = path.join(root, book);
	if (fs.existsSync(bookPath)) await vscode.commands.executeCommand("ras.book.show", vscode.Uri.file(bookPath));
	const openInExcel = "Excel で開く";
	const actions = process.platform === "win32" ? [openInExcel] : [];
	void vscode.window.showInformationMessage(`RAS: プロジェクトを作成しました（ブック: ${book}）。`, ...actions).then((choice) => {
		if (choice === openInExcel) void vscode.commands.executeCommand("ras.client.openInExcel");
	});
}

/** After the window was reopened on a project created by {@link newProject}: opens its files. */
async function openPending(): Promise<void> {
	const pending = context?.globalState.get<PendingOpen>(PENDING_OPEN);
	if (!pending) return;
	const here = (vscode.workspace.workspaceFolders ?? []).some((f) => f.uri.scheme === "file" && path.resolve(f.uri.fsPath).toLowerCase() === path.resolve(pending.root).toLowerCase());
	if (!here) return;
	await context?.globalState.update(PENDING_OPEN, undefined);
	await showCreated(pending.root, pending.book);
}

/**
 * Keeps `ras.canCreateProject` current: a workspace folder that is empty (or holds only workbooks
 * and hidden entries) and no ras.toml anywhere. The Explorer's title bar and the status bar show
 * the button then.
 */
function trackNewProjectContext(onChange: (can: boolean) => void): vscode.Disposable {
	const update = async () => {
		const roots = (vscode.workspace.workspaceFolders ?? []).filter((f) => f.uri.scheme === "file").map((f) => f.uri.fsPath);
		let can = roots.some((root) => {
			try {
				return isNewProjectFolder(fs.readdirSync(root));
			} catch {
				return false;
			}
		});
		if (can) can = (await vscode.workspace.findFiles("**/ras.toml", "**/node_modules/**", 1)).length === 0;
		await vscode.commands.executeCommand("setContext", "ras.canCreateProject", can);
		onChange(can);
	};
	const refresh = () => void update();
	// The top level of each workspace folder (and ras.toml anywhere).
	let watchers: vscode.Disposable[] = [];
	const watch = () => {
		for (const d of watchers) d.dispose();
		watchers = (vscode.workspace.workspaceFolders ?? []).flatMap((f) => {
			const watcher = vscode.workspace.createFileSystemWatcher(new vscode.RelativePattern(f, "*"), false, true, false);
			return [watcher, watcher.onDidCreate(refresh), watcher.onDidDelete(refresh)];
		});
	};
	const toml = vscode.workspace.createFileSystemWatcher("**/ras.toml", false, true, false);
	watch();
	void update();
	return vscode.Disposable.from(
		toml,
		toml.onDidCreate(refresh),
		toml.onDidDelete(refresh),
		vscode.workspace.onDidChangeWorkspaceFolders(() => {
			watch();
			refresh();
		}),
		{ dispose: () => watchers.forEach((d) => d.dispose()) },
	);
}

let context: vscode.ExtensionContext | undefined;

/** @param onCanCreate Called with whether the workspace is one to start a project in, whenever that is worked out. */
export function registerProjectCommands(extensionContext: vscode.ExtensionContext, onCanCreate: (can: boolean) => void): vscode.Disposable {
	context = extensionContext;
	void openPending().catch(() => undefined);
	return vscode.Disposable.from(
		trackNewProjectContext(onCanCreate),
		vscode.commands.registerCommand("ras.project.new", (resource?: unknown) =>
			newProject(resource instanceof vscode.Uri ? resource : undefined).catch((error) =>
				vscode.window.showErrorMessage(`RAS: プロジェクトを作成できませんでした: ${errorMessage(error)}`),
			),
		),
	);
}
