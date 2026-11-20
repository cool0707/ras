import * as vscode from "vscode";
import { registerCommands } from "./commands";
import { registerDebugging } from "./debug";
import { GeneratedVbaView } from "./generatedVba";
import { registerModuleCommands } from "./newModule";
import { RasServer } from "./server";
import { SyncStatusBar } from "./statusBar";
import { registerUserFormCommands } from "./userForm/commands";
import { UserFormDesignerProvider } from "./userForm/provider";

// The extension is deliberately thin (spec §7.4): it contributes the grammar,
// starts `ras lsp --stdio`, forwards four commands to it and mirrors the sync
// status it reports. All language intelligence lives in the server.
// Debugging (spec §14.10) builds with `ras build --target js` and hands the
// resulting Node process to VS Code's built-in JavaScript debugger.

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
		registerDebugging(context),
		UserFormDesignerProvider.register(context),
		registerUserFormCommands(),
		registerModuleCommands(),
		trackProjectContext(),
		vscode.workspace.onDidChangeConfiguration((event) => {
			if (event.affectsConfiguration("ras.path") || event.affectsConfiguration("ras.server.extraArgs")) {
				void rasServer.restart();
			}
		}),
		vscode.workspace.onDidGrantWorkspaceTrust(() => void rasServer.start()),
	);

	void rasServer.start();
}

/** Keeps the `ras.projectOpen` context key (a ras.toml in the workspace) current for the Explorer's buttons. */
function trackProjectContext(): vscode.Disposable {
	const update = async () => {
		const found = await vscode.workspace.findFiles("**/ras.toml", "**/node_modules/**", 1);
		await vscode.commands.executeCommand("setContext", "ras.projectOpen", found.length > 0);
	};
	const watcher = vscode.workspace.createFileSystemWatcher("**/ras.toml", false, true, false);
	const refresh = () => void update();
	void update();
	return vscode.Disposable.from(
		watcher,
		watcher.onDidCreate(refresh),
		watcher.onDidDelete(refresh),
		vscode.workspace.onDidChangeWorkspaceFolders(refresh),
	);
}

export async function deactivate(): Promise<void> {
	const rasServer = server;
	server = undefined;
	await rasServer?.stop();
	rasServer?.dispose();
}
