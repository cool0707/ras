import * as vscode from "vscode";
import { CloseAction, ExecuteCommandRequest, LanguageClient, State } from "vscode-languageclient/node";
import type { ErrorHandler, LanguageClientOptions, ServerOptions } from "vscode-languageclient/node";
import { findServerExecutable } from "./serverPath";

export type ServerState = "stopped" | "starting" | "running";

/** Server → client notification that drives the status bar. */
const SYNC_STATUS_NOTIFICATION = "ras/syncStatus";

/** Commands of the run / debug lenses the server puts above functions that can be started. */
const RUN_LENS_COMMANDS = new Set(["ras.client.runMacro", "ras.debug.start"]);

/** Thrown by {@link RasServer.execute} when there is no server to talk to. */
export class ServerNotRunningError extends Error {
	constructor() {
		super("RAS 言語サーバーが起動していません。");
		this.name = "ServerNotRunningError";
	}
}

/**
 * Owns the language client: starts `ras lsp --stdio`, restarts it on request,
 * and forwards commands to it. Everything language-related (completion, hover,
 * diagnostics, semantic tokens, formatting ...) is provided by the server
 * through the stock client; nothing is implemented here (spec §7.4).
 */
export class RasServer implements vscode.Disposable {
	private client: LanguageClient | undefined;
	private currentState: ServerState = "stopped";
	/** Serializes start / stop / restart so they never overlap. */
	private queue: Promise<void> = Promise.resolve();

	private readonly output = vscode.window.createOutputChannel("RAS Language Server");
	private readonly watchers = [
		vscode.workspace.createFileSystemWatcher("**/ras.toml"),
		vscode.workspace.createFileSystemWatcher("**/*.ras"),
	];
	private readonly stateEmitter = new vscode.EventEmitter<ServerState>();
	private readonly syncStatusEmitter = new vscode.EventEmitter<unknown>();

	readonly onDidChangeState = this.stateEmitter.event;
	/** Raw params of each `ras/syncStatus` notification. */
	readonly onDidReceiveSyncStatus = this.syncStatusEmitter.event;

	start(): Promise<void> {
		return this.enqueue(() => this.doStart());
	}

	stop(): Promise<void> {
		return this.enqueue(() => this.doStop());
	}

	restart(): Promise<void> {
		return this.enqueue(async () => {
			await this.doStop();
			await this.doStart();
		});
	}

	/** Sends `workspace/executeCommand` to the server and returns its result. */
	async execute(command: string, args: unknown[] = []): Promise<unknown> {
		const client = this.client;
		if (!client || client.state === State.Stopped) {
			throw new ServerNotRunningError();
		}
		// While the client is still starting, sendRequest waits for it.
		return client.sendRequest(ExecuteCommandRequest.type, { command, arguments: args });
	}

	/** `{ uri, position }` in the form the server knows the document by. */
	toProtocol(uri: vscode.Uri, position: vscode.Position): { uri: string; position: { line: number; character: number } } {
		const client = this.client;
		if (!client) {
			throw new ServerNotRunningError();
		}
		return {
			uri: client.code2ProtocolConverter.asUri(uri),
			position: client.code2ProtocolConverter.asPosition(position),
		};
	}

	dispose(): void {
		for (const watcher of this.watchers) {
			watcher.dispose();
		}
		this.stateEmitter.dispose();
		this.syncStatusEmitter.dispose();
		this.output.dispose();
	}

	private enqueue(task: () => Promise<void>): Promise<void> {
		this.queue = this.queue.then(task).catch((error: unknown) => {
			this.output.appendLine(`[ras-vscode] ${errorMessage(error)}`);
		});
		return this.queue;
	}

	private async doStart(): Promise<void> {
		if (this.client) {
			return;
		}
		if (!vscode.workspace.isTrusted) {
			// Restricted Mode: syntax highlighting only. activate() starts the
			// server once the workspace is trusted.
			return;
		}

		const config = vscode.workspace.getConfiguration("ras");
		const configuredPath = (config.get<string>("path") ?? "").trim();
		const extraArgs = config.get<string[]>("server.extraArgs") ?? [];
		const command = findServerExecutable(configuredPath, vscode.workspace.workspaceFolders?.[0]?.uri.fsPath);
		if (!command) {
			void this.reportStartFailure(
				configuredPath === ""
					? "ras 実行ファイルが PATH に見つかりません。ras をインストールするか、設定 ras.path に実行ファイルのパスを指定してください。"
					: `設定 ras.path の実行ファイルが見つかりません: ${configuredPath}`,
				false,
			);
			return;
		}

		// `transport` is left unset on purpose: the client then talks over the
		// process's stdin/stdout without appending a second `--stdio` argument.
		const serverOptions: ServerOptions = {
			command,
			args: ["lsp", "--stdio", ...extraArgs],
		};
		let started = false;
		let defaultErrorHandler: ErrorHandler | undefined;
		const clientOptions: LanguageClientOptions = {
			documentSelector: [{ language: "ras" }],
			synchronize: { fileEvents: this.watchers },
			outputChannel: this.output,
			middleware: {
				// The run / debug lenses above functions can be turned off with `ras.codeLens.run`.
				provideCodeLenses: async (document, token, next) => {
					const lenses = await next(document, token);
					if (lenses && !vscode.workspace.getConfiguration("ras", document).get<boolean>("codeLens.run", true)) {
						return lenses.filter((lens) => !RUN_LENS_COMMANDS.has(lens.command?.command ?? ""));
					}
					return lenses;
				},
			},
			// A failed start is reported once, below, with a way to fix `ras.path`.
			initializationFailedHandler: () => false,
			errorHandler: {
				error: (error, message, count) => (defaultErrorHandler ??= client.createDefaultErrorHandler()).error(error, message, count),
				// The stock policy restarts a crashed server a few times. Keep it for a
				// server that was running, but never respawn one that died while starting
				// (wrong binary, no `lsp` subcommand ...) or that has been replaced.
				closed: () =>
					started && this.client === client
						? (defaultErrorHandler ??= client.createDefaultErrorHandler()).closed()
						: { action: CloseAction.DoNotRestart, handled: true },
			},
		};
		// The id "ras" makes the client read the `ras.trace.server` setting.
		const client: LanguageClient = new LanguageClient("ras", "RAS Language Server", serverOptions, clientOptions);
		client.onNotification(SYNC_STATUS_NOTIFICATION, (params: unknown) => this.syncStatusEmitter.fire(params));
		client.onDidChangeState((event) => {
			if (this.client === client) {
				this.setState(event.newState === State.Running ? "running" : event.newState === State.Starting ? "starting" : "stopped");
			}
		});

		this.client = client;
		let failure = "サーバープロセスが起動中に終了しました。";
		try {
			await client.start();
		} catch (error) {
			failure = errorMessage(error);
		}
		// start() can resolve even though the server died during the handshake,
		// so the client's state decides, not the promise.
		if (client.isRunning()) {
			started = true;
			return;
		}
		this.client = undefined;
		this.setState("stopped");
		void this.reportStartFailure(`RAS 言語サーバーを起動できませんでした（${command}）: ${failure}`, true);
	}

	private async doStop(): Promise<void> {
		const client = this.client;
		this.client = undefined;
		try {
			if (client?.needsStop()) {
				await client.stop();
			}
		} catch (error) {
			// A server that does not shut down cleanly must not block a restart.
			this.output.appendLine(`[ras-vscode] ${errorMessage(error)}`);
		} finally {
			this.setState("stopped");
		}
	}

	private setState(state: ServerState): void {
		if (state !== this.currentState) {
			this.currentState = state;
			this.stateEmitter.fire(state);
		}
	}

	private async reportStartFailure(message: string, offerOutput: boolean): Promise<void> {
		const openSettings = "設定を開く";
		const showOutput = "出力を表示";
		const choice = await vscode.window.showErrorMessage(message, ...(offerOutput ? [openSettings, showOutput] : [openSettings]));
		if (choice === openSettings) {
			await vscode.commands.executeCommand("workbench.action.openSettings", "ras.path");
		} else if (choice === showOutput) {
			this.output.show(true);
		}
	}
}

export function errorMessage(error: unknown): string {
	return error instanceof Error ? error.message : String(error);
}
