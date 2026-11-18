import * as vscode from "vscode";
import { registerCommands } from "./commands";
import { GeneratedVbaView } from "./generatedVba";
import { RasServer } from "./server";
import { SyncStatusBar } from "./statusBar";

// The extension is deliberately thin (spec §7.4): it contributes the grammar,
// starts `ras lsp --stdio`, forwards four commands to it and mirrors the sync
// status it reports. All language intelligence lives in the server.

let server: RasServer | undefined;

export function activate(context: vscode.ExtensionContext): void {
	const rasServer = new RasServer();
	const statusBar = new SyncStatusBar();
	const vbaView = new GeneratedVbaView();
	server = rasServer;

	// rasServer is not a subscription: VS Code disposes subscriptions without
	// waiting for deactivate(), and the server needs its output channel until
	// the shutdown handshake has finished.
	context.subscriptions.push(
		statusBar,
		vbaView,
		rasServer.onDidChangeState((state) => statusBar.setServerState(state)),
		rasServer.onDidReceiveSyncStatus((params) => statusBar.setSyncStatus(params)),
		registerCommands(rasServer, vbaView),
		vscode.workspace.onDidChangeConfiguration((event) => {
			if (event.affectsConfiguration("ras.path") || event.affectsConfiguration("ras.server.extraArgs")) {
				void rasServer.restart();
			}
		}),
		vscode.workspace.onDidGrantWorkspaceTrust(() => void rasServer.start()),
	);

	void rasServer.start();
}

export async function deactivate(): Promise<void> {
	const rasServer = server;
	server = undefined;
	await rasServer?.stop();
	rasServer?.dispose();
}
