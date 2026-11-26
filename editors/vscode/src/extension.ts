import * as vscode from "vscode";
import { registerCommands } from "./commands";
import { registerDebugging } from "./debug";
import { GeneratedVbaView } from "./generatedVba";
import { registerModuleCommands } from "./newModule";
import { registerProjectCommands } from "./newProject";
import { RasServer } from "./server";
import { registerSheetView } from "./sheetView";
import { registerSheetCommands } from "./sheetCommands";
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
	onServerWanted = () => statusBar.setServerWanted(true);

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
		registerSheetView(context),
		registerSheetCommands(rasServer),
		UserFormDesignerProvider.register(context),
		registerUserFormCommands(),
		registerModuleCommands(),
		registerProjectCommands(context, (can) => statusBar.setCanCreateProject(can)),
		trackProjectContext(startServer),
		vscode.workspace.onDidOpenTextDocument((document) => {
			if (document.languageId === "ras") startServer();
		}),
		vscode.workspace.onDidChangeConfiguration((event) => {
			if (serverWanted && (event.affectsConfiguration("ras.path") || event.affectsConfiguration("ras.server.extraArgs"))) {
				void rasServer.restart();
			}
		}),
		vscode.workspace.onDidGrantWorkspaceTrust(() => {
			if (serverWanted) void rasServer.start();
		}),
	);

	// The extension also activates in folders without a project (for "New RAS project" in the
	// Explorer): the server starts once there is a ras.toml or a .ras file is opened.
	if (vscode.workspace.textDocuments.some((document) => document.languageId === "ras")) startServer();
}

/** Whether the language server is to run: a project or a .ras file was seen. */
let serverWanted = false;

let onServerWanted: (() => void) | undefined;

function startServer(): void {
	if (serverWanted) return;
	serverWanted = true;
	onServerWanted?.();
	void server?.start();
}

/**
 * Keeps the `ras.projectOpen` / `ras.noProject` context keys (a ras.toml in the workspace) current
 * for the Explorer's buttons, and calls `onProject` whenever there is one.
 */
function trackProjectContext(onProject: () => void): vscode.Disposable {
	const update = async () => {
		const found = await vscode.workspace.findFiles("**/ras.toml", "**/node_modules/**", 1);
		await vscode.commands.executeCommand("setContext", "ras.projectOpen", found.length > 0);
		await vscode.commands.executeCommand("setContext", "ras.noProject", found.length === 0);
		if (found.length > 0) onProject();
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
	serverWanted = false;
	await rasServer?.stop();
	rasServer?.dispose();
}
