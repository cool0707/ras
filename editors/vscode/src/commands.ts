import * as cp from "node:child_process";
import * as vscode from "vscode";
import { isGeneratedVba } from "./generatedVba";
import type { GeneratedVbaView } from "./generatedVba";
import { errorMessage, ServerNotRunningError } from "./server";
import type { RasServer } from "./server";
import { findServerExecutable } from "./serverPath";
import { projectRoot } from "./userForm/commands";

// Every command only forwards to the server through `workspace/executeCommand`
// (spec §7.4). vscode-languageclient registers the commands a server declares
// under the server's own ids, so the palette commands use distinct `ras.client.*`
// ids and call these:
const SERVER_COMMANDS = {
	showGeneratedVba: "ras.showGeneratedVba",
	syncNow: "ras.syncNow",
	openInExcel: "ras.openInExcel",
	listMacros: "ras.listMacros",
	runMacro: "ras.runMacro",
} as const;

export function registerCommands(server: RasServer, vbaView: GeneratedVbaView): vscode.Disposable {
	const command = (id: string, title: string, run: (...args: unknown[]) => Promise<void>) =>
		vscode.commands.registerCommand(id, async (...args: unknown[]) => {
			try {
				await run(...args);
			} catch (error) {
				await reportFailure(title, error);
			}
		});

	return vscode.Disposable.from(
		command("ras.client.showGeneratedVba", "生成 VBA を表示", (resource) => showGeneratedVba(server, vbaView, resource)),
		command("ras.client.syncNow", "今すぐ同期", () => syncNow(server)),
		command("ras.client.openInExcel", "Excel で開く", async () => {
			const result = await withProgress("RAS: Excel で開いています…", () => server.execute(SERVER_COMMANDS.openInExcel));
			showResult(result, "Excel で開きました。", "Excel で開けませんでした。");
		}),
		// The run lens above a function passes its macro name (`Module.Proc`).
		command("ras.client.runMacro", "マクロを実行", (name) => runMacro(server, typeof name === "string" ? name : undefined)),
		command("ras.restartServer", "言語サーバーを再起動", async () => {
			if (!vscode.workspace.isTrusted) {
				void vscode.window.showInformationMessage("RAS 言語サーバーは、ワークスペースを信頼すると起動します。");
				return;
			}
			await server.restart();
		}),
	);
}

async function showGeneratedVba(server: RasServer, vbaView: GeneratedVbaView, resource: unknown): Promise<void> {
	const editor = findRasEditor(resource);
	if (!editor) {
		void vscode.window.showInformationMessage("RAS ファイル（.ras）のエディターで実行してください。");
		return;
	}
	const result = await server.execute(SERVER_COMMANDS.showGeneratedVba, [
		server.toProtocol(editor.document.uri, editor.selection.active),
	]);
	if (result === null || result === undefined) {
		void vscode.window.showInformationMessage("このファイルに対応する生成 VBA はありません。");
		return;
	}
	if (!isGeneratedVba(result)) {
		throw new Error("言語サーバーから想定外の応答が返されました。");
	}
	await vbaView.show(editor.document.uri, result);
}

async function syncNow(server: RasServer): Promise<void> {
	await vscode.workspace.saveAll(false);
	const result = await withProgress("RAS: 同期しています…", () => server.execute(SERVER_COMMANDS.syncNow));
	if (isOffline(result)) {
		// No Excel has the workbook open: write to the workbook file instead (spec §14.9, `ras pack`).
		await packOffline();
		return;
	}
	const drifted = driftedModules(result);
	if (drifted.length === 0) {
		showResult(result, "同期しました。", "同期に失敗しました。");
		return;
	}
	// Modules edited in the VBE are left alone until the user decides (spec §14.7).
	const overwrite = "RAS で上書き";
	const choice = await vscode.window.showWarningMessage(
		`RAS: 次のモジュールは Excel（VBE）側で編集されているため、上書きしていません: ${drifted.join(", ")}`,
		overwrite,
	);
	if (choice !== overwrite) {
		return;
	}
	const forced = await withProgress("RAS: 同期しています…", () => server.execute(SERVER_COMMANDS.syncNow, [{ overwrite: true }]));
	showResult(forced, "同期しました。", "同期に失敗しました。");
}

/** The sync failed only because no Excel has the workbook open. */
function isOffline(result: unknown): boolean {
	return typeof result === "object" && result !== null && (result as { offline?: unknown }).offline === true;
}

let packOutput: vscode.OutputChannel | undefined;

/** Runs `ras pack`: writes the generated VBA and the user forms straight into the workbook file. */
async function packOffline(overwrite = false): Promise<void> {
	const root = await projectRoot(undefined, "書き込むプロジェクトを選んでください");
	if (!root) {
		void vscode.window.showErrorMessage("ras.toml のあるフォルダーを開いてください。");
		return;
	}
	const exe = findServerExecutable((vscode.workspace.getConfiguration("ras").get<string>("path") ?? "").trim(), root);
	if (exe === undefined) {
		throw new Error("ras 実行ファイルが見つかりません。設定 ras.path を確認してください。");
	}
	const args = ["--project", root, "pack", ...(overwrite ? ["--overwrite"] : [])];
	const output = (packOutput ??= vscode.window.createOutputChannel("RAS"));
	output.appendLine(`> ras ${args.join(" ")}`);
	const run = await withProgress("RAS: ブックファイルへ書き込んでいます…", () =>
		new Promise<{ ok: boolean; stdout: string; stderr: string }>((resolve) =>
			cp.execFile(exe, args, { cwd: root, windowsHide: true }, (error, stdout, stderr) => resolve({ ok: !error, stdout, stderr })),
		),
	);
	for (const text of [run.stdout, run.stderr]) {
		if (text.trim() !== "") output.appendLine(text.trimEnd());
	}
	if (run.ok) {
		const line = run.stdout.trim().split(/\r?\n/).pop();
		void vscode.window.showInformationMessage(`RAS: ${line || "ブックファイルへ書き込みました。"}（Excel は未接続のため、オフラインで書き込みました）`);
		return;
	}
	// Errors are on stderr as `error: …`; build errors come with their diagnostics.
	const errorLine = run.stderr.split(/\r?\n/).find((l) => l.startsWith("error:"))?.slice("error:".length).trim();
	const message = errorLine || "ブックファイルへの書き込みに失敗しました（ビルド エラーがあります）。";
	const overwriteAction = "RAS で上書き";
	const showOutput = "出力を表示";
	const actions = !overwrite && run.stderr.includes("--overwrite") ? [overwriteAction, showOutput] : [showOutput];
	const choice = await vscode.window.showErrorMessage(`RAS: ${message}`, ...actions);
	if (choice === overwriteAction) {
		await packOffline(true);
	} else if (choice === showOutput) {
		output.show(true);
	}
}

/** Names in the `drifted` list of a sync result. */
function driftedModules(result: unknown): string[] {
	const drifted = (typeof result === "object" && result !== null ? result : {}) as { drifted?: unknown };
	return Array.isArray(drifted.drifted) ? drifted.drifted.filter((name): name is string => typeof name === "string") : [];
}

async function runMacro(server: RasServer, name: string | undefined): Promise<void> {
	// The macro runs from the saved sources, which the server syncs to the workbook first.
	await vscode.workspace.saveAll(false);
	const macro = name ?? (await pickMacro(server));
	if (macro === undefined) {
		return;
	}
	const result = await withProgress(`RAS: ${macro} を実行しています…`, () => server.execute(SERVER_COMMANDS.runMacro, [macro]));
	showResult(result, `${macro} を実行しました。`, `${macro} の実行に失敗しました。`);
}

async function pickMacro(server: RasServer): Promise<string | undefined> {
	const listed = await server.execute(SERVER_COMMANDS.listMacros);
	const macros = Array.isArray(listed) ? listed.filter((name): name is string => typeof name === "string") : [];
	if (macros.length === 0) {
		void vscode.window.showInformationMessage("実行できるマクロがありません。");
		return undefined;
	}
	return vscode.window.showQuickPick(macros, { title: "RAS: マクロを実行", placeHolder: "実行するマクロを選択してください" });
}

/** The RAS editor a command applies to: the one a menu was invoked on, else the active one. */
function findRasEditor(resource: unknown): vscode.TextEditor | undefined {
	const active = vscode.window.activeTextEditor;
	const candidates = resource instanceof vscode.Uri
		? [active, ...vscode.window.visibleTextEditors].filter((editor) => editor?.document.uri.toString() === resource.toString())
		: [active];
	return candidates.find((editor) => editor?.document.languageId === "ras");
}

/** Shows the `{ ok, message }` result of a server command. */
function showResult(result: unknown, succeeded: string, failed: string): void {
	const { ok, message } = (typeof result === "object" && result !== null ? result : {}) as { ok?: unknown; message?: unknown };
	const text = typeof message === "string" && message !== "" ? message : undefined;
	if (ok === false) {
		void vscode.window.showErrorMessage(text ?? failed);
	} else {
		void vscode.window.showInformationMessage(text ?? succeeded);
	}
}

function withProgress<T>(title: string, task: () => Promise<T>): Thenable<T> {
	return vscode.window.withProgress({ location: vscode.ProgressLocation.Window, title }, task);
}

async function reportFailure(title: string, error: unknown): Promise<void> {
	if (error instanceof ServerNotRunningError) {
		const restart = "再起動";
		const choice = await vscode.window.showWarningMessage(error.message, restart);
		if (choice === restart) {
			await vscode.commands.executeCommand("ras.restartServer");
		}
		return;
	}
	void vscode.window.showErrorMessage(`RAS「${title}」に失敗しました: ${errorMessage(error)}`);
}
