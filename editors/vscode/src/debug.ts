import * as cp from "node:child_process";
import * as fs from "node:fs";
import * as path from "node:path";
import * as vscode from "vscode";
import { findServerExecutable } from "./serverPath";

// Debugging (spec §14.10). `ras build --target js` turns the project into JavaScript with
// source maps; that JavaScript runs in a Node process and reaches Excel through the
// `ras_com` addon. The debugger itself is VS Code's built-in JavaScript debugger: this
// file only builds, starts the process and hands its inspector port over.
//
// The process is VS Code's own runtime in Node mode (`ELECTRON_RUN_AS_NODE`), so no
// separate Node installation is needed. The built-in debugger's `launch` request cannot
// be used for that — it removes `ELECTRON_RUN_AS_NODE` from the environment — hence the
// process is started here and the debugger attaches to it.

const DEBUG_TYPE = "ras";

interface RasDebugConfiguration extends vscode.DebugConfiguration {
	/** Function to run: `Module.proc` or just `proc`. Asked for when omitted. */
	entry?: string;
	/** Project folder (the one with ras.toml). Defaults to the workspace folder. */
	project?: string;
	/** Node executable to run the program with. Defaults to VS Code's own runtime. */
	runtimeExecutable?: string;
	/** Path of the `ras_com` addon. Defaults to the one shipped with the extension. */
	comAddon?: string;
	/** Module that stands in for Excel (for tests on machines without Excel). */
	mockHost?: string;
	/** What to do when Excel shows a dialog while the program waits for it. */
	dialogs?: "notify" | "dismiss" | "off";
	stopOnEntry?: boolean;
}

interface DebugManifest {
	entries: string[];
}

/** Processes started for debug sessions, by the id stored in the attach configuration. */
const running = new Map<string, cp.ChildProcess>();
let nextId = 1;
let output: vscode.OutputChannel | undefined;

function channel(): vscode.OutputChannel {
	output ??= vscode.window.createOutputChannel("RAS Debug");
	return output;
}

export function registerDebugging(context: vscode.ExtensionContext): vscode.Disposable {
	return vscode.Disposable.from(
		vscode.debug.registerDebugConfigurationProvider(DEBUG_TYPE, new RasDebugConfigurationProvider(context)),
		vscode.debug.onDidTerminateDebugSession((session) => {
			const id = session.configuration.__rasProcess as string | undefined;
			const child = id === undefined ? undefined : running.get(id);
			if (id !== undefined && child !== undefined) {
				running.delete(id);
				child.kill();
			}
		}),
		vscode.commands.registerCommand("ras.debug.start", async () => {
			const folder = projectFolder(undefined);
			await vscode.debug.startDebugging(folder, { type: DEBUG_TYPE, request: "launch", name: "RAS: デバッグ" });
		}),
		new vscode.Disposable(() => {
			for (const child of running.values()) {
				child.kill();
			}
			running.clear();
			output?.dispose();
		}),
	);
}

function projectFolder(folder: vscode.WorkspaceFolder | undefined): vscode.WorkspaceFolder | undefined {
	if (folder !== undefined) {
		return folder;
	}
	const active = vscode.window.activeTextEditor?.document.uri;
	return (active && vscode.workspace.getWorkspaceFolder(active)) ?? vscode.workspace.workspaceFolders?.[0];
}

/** Walks up from `start` to the directory that holds ras.toml. */
function findProjectRoot(start: string): string | undefined {
	let dir = start;
	for (;;) {
		if (fs.existsSync(path.join(dir, "ras.toml"))) {
			return dir;
		}
		const parent = path.dirname(dir);
		if (parent === dir) {
			return undefined;
		}
		dir = parent;
	}
}

class RasDebugConfigurationProvider implements vscode.DebugConfigurationProvider {
	constructor(private readonly context: vscode.ExtensionContext) {}

	provideDebugConfigurations(): vscode.DebugConfiguration[] {
		return [{ type: DEBUG_TYPE, request: "launch", name: "RAS: デバッグ" }];
	}

	// The configuration is replaced by one of type `node` here, in the first resolution
	// phase: only there does VS Code activate and consult the debugger of the new type.
	async resolveDebugConfiguration(folder: vscode.WorkspaceFolder | undefined, config: RasDebugConfiguration): Promise<vscode.DebugConfiguration | undefined> {
		try {
			// Variables are substituted after this phase, so the common one is handled here.
			const base = projectFolder(folder)?.uri.fsPath;
			for (const key of ["project", "runtimeExecutable", "comAddon", "mockHost"] as const) {
				const value = config[key];
				if (typeof value === "string" && base !== undefined) {
					config[key] = value.replace(/\$\{workspaceFolder\}/g, base);
				}
			}
			return await this.resolve(folder, config);
		} catch (error) {
			void vscode.window.showErrorMessage(`RAS: デバッグを開始できません。${error instanceof Error ? error.message : String(error)}`);
			return undefined;
		}
	}

	private async resolve(folder: vscode.WorkspaceFolder | undefined, config: RasDebugConfiguration): Promise<vscode.DebugConfiguration | undefined> {
		if (!vscode.workspace.isTrusted) {
			throw new Error("ワークスペースを信頼するまで、プログラムは実行できません。");
		}
		// F5 without a launch.json arrives as an empty configuration.
		if (!config.type) {
			config.type = DEBUG_TYPE;
			config.request = "launch";
			config.name = "RAS: デバッグ";
		}
		const base = projectFolder(folder)?.uri.fsPath;
		const activeFile = vscode.window.activeTextEditor?.document.uri;
		const start = config.project ? path.resolve(base ?? process.cwd(), config.project) : activeFile?.scheme === "file" ? path.dirname(activeFile.fsPath) : base;
		const root = (start && findProjectRoot(start)) ?? (base && findProjectRoot(base));
		if (root === undefined) {
			throw new Error("ras.toml が見つかりません。RAS プロジェクトのフォルダーを開いてください。");
		}

		await vscode.workspace.saveAll(false);
		await build(root);
		const debugDir = path.join(root, "build", "debug");
		const manifest = JSON.parse(fs.readFileSync(path.join(debugDir, "ras-debug.json"), "utf8")) as DebugManifest;
		const entry = config.entry ?? (await pickEntry(manifest.entries));
		if (entry === undefined) {
			return undefined;
		}

		const addon = comAddon(this.context, config);
		if (process.platform === "win32" && addon === undefined && !config.mockHost) {
			void vscode.window.showWarningMessage(
				"RAS: Excel に接続するためのアドオン（ras_com.node）が見つかりません。Excel を操作する箇所でエラーになります。設定 ras.debug.comAddon で場所を指定できます。",
			);
		}
		const env: NodeJS.ProcessEnv = { ...process.env, ELECTRON_RUN_AS_NODE: "1" };
		// Left over from VS Code's own processes; they must not reach the debuggee.
		delete env.NODE_OPTIONS;
		delete env.VSCODE_INSPECTOR_OPTIONS;
		if (addon !== undefined) {
			env.RAS_COM_ADDON = addon;
		}
		if (config.mockHost) {
			env.RAS_MOCK_HOST = path.resolve(root, config.mockHost);
		}
		const dialogs = config.dialogs || setting("debug.dialogs");
		if (dialogs) {
			env.RAS_DIALOGS = dialogs;
		}
		const runtime = config.runtimeExecutable || setting("debug.runtimeExecutable") || process.execPath;
		const main = path.join(debugDir, "ras-main.js");
		const out = channel();
		out.appendLine(`> ${entry}`);

		if (config.noDebug) {
			const child = cp.spawn(runtime, ["--enable-source-maps", main, entry], { cwd: root, env });
			child.stdout.on("data", (data: Buffer) => out.append(data.toString()));
			child.stderr.on("data", (data: Buffer) => {
				const text = data.toString();
				out.append(text);
				text.split(/\r?\n/).forEach(notifyDialog);
			});
			child.on("exit", (code) => out.appendLine(`（終了コード ${code ?? "?"}）`));
			out.show(true);
			return undefined;
		}

		const child = cp.spawn(runtime, ["--enable-source-maps", "--inspect-brk=127.0.0.1:0", main, entry], { cwd: root, env });
		const port = await inspectorPort(child, out);
		const id = String(nextId++);
		running.set(id, child);
		child.on("exit", () => running.delete(id));

		const dir = debugDir.replace(/\\/g, "/");
		return {
			type: "node",
			request: "attach",
			name: config.name,
			address: "127.0.0.1",
			port,
			// The process waits at its first line until the breakpoints are in place.
			continueOnAttach: !config.stopOnEntry,
			sourceMaps: true,
			outFiles: [`${dir}/**/*.js`],
			resolveSourceMapLocations: [`${root.replace(/\\/g, "/")}/**`],
			skipFiles: ["<node_internals>/**", `${dir}/ras-runtime.js`, `${dir}/ras-main.js`],
			smartStep: true,
			cwd: root,
			__rasProcess: id,
		};
	}
}

function setting(name: string): string {
	return (vscode.workspace.getConfiguration("ras").get<string>(name) ?? "").trim();
}

function rasExecutable(root: string): string {
	const exe = findServerExecutable(setting("path"), root);
	if (exe === undefined) {
		throw new Error("ras 実行ファイルが見つかりません。設定 ras.path を確認してください。");
	}
	return exe;
}

/** Runs `ras build --target js`; rejects with the compiler's messages when it fails. */
function build(root: string): Promise<void> {
	const exe = rasExecutable(root);
	return new Promise((resolve, reject) => {
		cp.execFile(exe, ["--project", root, "build", "--target", "js"], { cwd: root, windowsHide: true }, (error, stdout, stderr) => {
			const out = channel();
			if (stderr.trim() !== "") {
				out.appendLine(stderr.trimEnd());
			}
			if (error) {
				out.show(true);
				reject(new Error("ビルドに失敗しました（「出力」パネルの RAS Debug を確認してください）。"));
			} else {
				out.appendLine(stdout.trimEnd());
				resolve();
			}
		});
	});
}

async function pickEntry(entries: string[]): Promise<string | undefined> {
	if (entries.length === 0) {
		throw new Error("実行できる関数がありません。引数なしの関数か、トップレベルの文が必要です。");
	}
	if (entries.length === 1) {
		return entries[0];
	}
	return vscode.window.showQuickPick(entries, { title: "RAS: デバッグする関数", placeHolder: "実行する関数を選んでください" });
}

/** The `ras_com` addon: explicit path, the one shipped with the extension, or one next to `ras`. */
function comAddon(context: vscode.ExtensionContext, config: RasDebugConfiguration): string | undefined {
	const explicit = config.comAddon || setting("debug.comAddon");
	if (explicit) {
		return path.resolve(explicit);
	}
	if (process.platform !== "win32") {
		return undefined;
	}
	const candidates = [path.join(context.extensionPath, "bin", `win32-${process.arch}`, "ras_com.node")];
	const ras = findServerExecutable(setting("path"), undefined);
	if (ras !== undefined) {
		candidates.push(path.join(path.dirname(ras), "ras_com.node"));
	}
	return candidates.find((file) => fs.existsSync(file));
}

/**
 * The run time reports a dialog of Excel that the program is waiting for (spec §14.10.6) on
 * a line of its own. Excel is behind the editor at that moment, so the line becomes a notification.
 */
function notifyDialog(line: string): void {
	const match = /^\[ras\] (Excel .*)$/.exec(line);
	if (match === null) {
		return;
	}
	const message = `RAS: ${match[1]}`;
	if (/閉じました/.test(match[1])) {
		void vscode.window.showInformationMessage(message);
	} else {
		void vscode.window.showWarningMessage(message);
	}
}

const INSPECTOR_NOISE = /^(Debugger listening on |For help, see: |Debugger attached\.|Waiting for the debugger to disconnect)/;

/** Waits for the inspector to announce its port; afterwards forwards what the process writes. */
function inspectorPort(child: cp.ChildProcess, out: vscode.OutputChannel): Promise<number> {
	return new Promise((resolve, reject) => {
		let pending = "";
		let found = false;
		const timer = setTimeout(() => fail(new Error("デバッグ対象のプロセスが応答しません。")), 15000);
		const fail = (error: Error) => {
			clearTimeout(timer);
			if (!found) {
				child.kill();
				reject(error);
			}
		};
		child.on("error", (error) => fail(new Error(`プロセスを起動できません: ${error.message}`)));
		child.on("exit", (code) => fail(new Error(`プロセスが終了しました（終了コード ${code ?? "?"}）。${pending.trim()}`)));
		child.stdout?.on("data", (data: Buffer) => out.append(data.toString()));
		child.stderr?.on("data", (data: Buffer) => {
			pending += data.toString();
			const lines = pending.split(/\r?\n/);
			pending = lines.pop() ?? "";
			for (const line of lines) {
				const match = /^Debugger listening on ws:\/\/[^:]+:(\d+)\//.exec(line);
				if (match && !found) {
					found = true;
					clearTimeout(timer);
					resolve(Number(match[1]));
				} else if (!INSPECTOR_NOISE.test(line)) {
					// Uncaught errors are printed here when the program ends with one.
					out.appendLine(line);
					notifyDialog(line);
					if (found) {
						vscode.debug.activeDebugConsole.appendLine(line);
					}
				}
			}
		});
	});
}
