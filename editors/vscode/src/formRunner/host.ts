import { randomBytes } from "node:crypto";
import * as net from "node:net";
import * as vscode from "vscode";
import { applyUpdate, initialState, LineReader } from "./protocol";
import type { FormState, PanelMessage, RuntimeMessage, UserAction } from "./protocol";

// Debug run of a user form (spec §14.10.7). The form's code runs in the debug process; the
// form is drawn here, each in a webview panel of its own. The debug process connects to the
// TCP server of a FormHost (127.0.0.1, the port is passed in RAS_FORM_HOST), reports what to
// draw and receives what the user does.

/** One per debug run: a server that accepts the one debug process, and the panels of its forms. */
export class FormHost implements vscode.Disposable {
	private readonly server = net.createServer((socket) => this.accept(socket));
	private socket: net.Socket | undefined;
	private readonly forms = new Map<number, FormPanel>();
	private disposed = false;

	private constructor(private readonly resources: vscode.Uri) {}

	static async start(context: vscode.ExtensionContext): Promise<FormHost> {
		const host = new FormHost(vscode.Uri.joinPath(context.extensionUri, "dist", "webview"));
		await new Promise<void>((resolve, reject) => {
			host.server.once("error", reject);
			host.server.listen(0, "127.0.0.1", () => resolve());
		});
		return host;
	}

	get port(): number {
		return (this.server.address() as net.AddressInfo).port;
	}

	private accept(socket: net.Socket): void {
		if (this.socket !== undefined || this.disposed) {
			socket.destroy();
			return;
		}
		this.socket = socket;
		socket.setEncoding("utf8");
		socket.setNoDelay(true);
		const reader = new LineReader();
		socket.on("data", (chunk: string) => {
			for (const line of reader.push(chunk)) {
				let msg: RuntimeMessage;
				try {
					msg = JSON.parse(line) as RuntimeMessage;
				} catch {
					continue;
				}
				this.receive(msg);
			}
		});
		socket.on("error", () => undefined);
		socket.on("close", () => this.ended());
	}

	/** Passes what the user did on to the debug process. */
	send(id: number, action: UserAction): void {
		if (this.socket?.writable) {
			this.socket.write(`${JSON.stringify({ ...action, id })}\n`);
		}
	}

	private receive(msg: RuntimeMessage): void {
		switch (msg.t) {
			case "show": {
				let form = this.forms.get(msg.id);
				if (form === undefined) {
					form = new FormPanel(this, msg.id, this.resources);
					this.forms.set(msg.id, form);
				}
				form.show(initialState(msg));
				break;
			}
			case "update":
				this.forms.get(msg.id)?.update(msg);
				break;
			case "hide":
			case "unload":
				this.forms.get(msg.id)?.close();
				this.forms.delete(msg.id);
				break;
			case "error":
				this.forms.get(msg.id)?.post({ type: "error", message: msg.message });
				break;
			case "blocked":
				this.forms.get(msg.id)?.post({ type: "blocked" });
				break;
		}
	}

	/** The debug process has ended: its forms are gone. */
	private ended(): void {
		for (const form of this.forms.values()) {
			form.close();
		}
		this.forms.clear();
	}

	dispose(): void {
		this.disposed = true;
		this.ended();
		this.socket?.destroy();
		this.server.close();
	}
}

/** The panel of one shown form. It is closed when the form is hidden or unloaded. */
class FormPanel {
	private panel: vscode.WebviewPanel | undefined;
	private state: FormState | undefined;
	private ready = false;
	private pending: PanelMessage[] = [];
	/** Closed by the debug process (hidden or unloaded), not by the user. */
	private closing = false;

	constructor(
		private readonly host: FormHost,
		private readonly id: number,
		private readonly resources: vscode.Uri,
	) {}

	show(state: FormState): void {
		this.closing = false;
		this.state = state;
		this.post({ type: "state", state });
	}

	update(msg: Extract<RuntimeMessage, { t: "update" }>): void {
		if (this.state === undefined) {
			return;
		}
		this.state = applyUpdate(this.state, msg);
		this.post({ type: "state", state: this.state, focus: msg.focus });
		if (this.panel !== undefined) {
			this.panel.title = title(this.state);
		}
	}

	post(message: PanelMessage): void {
		const panel = this.open();
		if (panel === undefined) {
			return;
		}
		if (this.ready) {
			void panel.webview.postMessage(message);
		} else if (message.type === "state") {
			// Only the latest state matters.
			this.pending = [...this.pending.filter((m) => m.type !== "state"), message];
		} else {
			this.pending.push(message);
		}
	}

	close(): void {
		this.closing = true;
		this.panel?.dispose();
	}

	/** The panel, opened again if the user closed it while the form stayed (QueryClose). */
	private open(): vscode.WebviewPanel | undefined {
		if (this.panel !== undefined) {
			return this.panel;
		}
		if (this.closing || this.state === undefined) {
			return undefined;
		}
		const panel = vscode.window.createWebviewPanel("ras.formRunner", title(this.state), { viewColumn: vscode.ViewColumn.Beside, preserveFocus: false }, {
			enableScripts: true,
			localResourceRoots: [this.resources],
			retainContextWhenHidden: true,
		});
		panel.webview.html = runnerHtml(panel.webview, this.resources);
		this.panel = panel;
		this.ready = false;
		this.pending = [{ type: "state", state: this.state }];
		panel.webview.onDidReceiveMessage((message: { type?: string; action?: UserAction }) => {
			if (message.type === "ready") {
				this.ready = true;
				for (const m of this.pending) {
					void panel.webview.postMessage(m);
				}
				this.pending = [];
			} else if (message.type === "action" && message.action) {
				this.host.send(this.id, message.action);
			}
		});
		panel.onDidDispose(() => {
			this.panel = undefined;
			this.ready = false;
			// Closing the tab is the form's close button: the form decides (QueryClose).
			if (!this.closing) {
				this.host.send(this.id, { t: "close" });
			}
		});
		return panel;
	}
}

function title(state: FormState): string {
	const caption = typeof state.form.Caption === "string" && state.form.Caption !== "" ? state.form.Caption : state.name;
	return `${caption}（${state.name}）`;
}

function runnerHtml(webview: vscode.Webview, resources: vscode.Uri): string {
	const nonce = randomBytes(24).toString("hex");
	const script = webview.asWebviewUri(vscode.Uri.joinPath(resources, "formRunner.js"));
	const style = webview.asWebviewUri(vscode.Uri.joinPath(resources, "formRunner.css"));
	return `<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'nonce-${nonce}'; style-src ${webview.cspSource}; img-src ${webview.cspSource} data:;">
<link rel="stylesheet" href="${style}"><title>RAS フォーム</title></head>
<body><div id="app"></div><script nonce="${nonce}" src="${script}"></script></body></html>`;
}
