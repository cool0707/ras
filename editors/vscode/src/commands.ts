import * as vscode from "vscode";
import { isGeneratedVba } from "./generatedVba";
import type { GeneratedVbaView } from "./generatedVba";
import { errorMessage, ServerNotRunningError } from "./server";
import type { RasServer } from "./server";

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
		command("ras.client.syncNow", "今すぐ同期", async () => {
			const result = await withProgress("RAS: 同期しています…", () => server.execute(SERVER_COMMANDS.syncNow));
			showResult(result, "同期しました。", "同期に失敗しました。");
		}),
		command("ras.client.openInExcel", "Excel で開く", async () => {
			const result = await withProgress("RAS: Excel で開いています…", () => server.execute(SERVER_COMMANDS.openInExcel));
			showResult(result, "Excel で開きました。", "Excel で開けませんでした。");
		}),
		command("ras.client.runMacro", "マクロを実行", () => runMacro(server)),
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

async function runMacro(server: RasServer): Promise<void> {
	const listed = await server.execute(SERVER_COMMANDS.listMacros);
	const macros = Array.isArray(listed) ? listed.filter((name): name is string => typeof name === "string") : [];
	if (macros.length === 0) {
		void vscode.window.showInformationMessage("実行できるマクロがありません。");
		return;
	}
	const macro = await vscode.window.showQuickPick(macros, { title: "RAS: マクロを実行", placeHolder: "実行するマクロを選択してください" });
	if (macro === undefined) {
		return;
	}
	const result = await withProgress(`RAS: ${macro} を実行しています…`, () => server.execute(SERVER_COMMANDS.runMacro, [macro]));
	showResult(result, `${macro} を実行しました。`, `${macro} の実行に失敗しました。`);
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
