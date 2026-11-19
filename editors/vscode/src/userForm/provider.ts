// UserForm Designer: a custom editor over forms/<Name>.form.json.
//
// Ported from xlflow's editors/vscode/src/userFormEditor/provider.ts
// (https://github.com/harumiWeb/xlflow, MIT License, Copyright (c) 2026 harumiWeb).
// Unlike xlflow, the Designer is a CustomTextEditorProvider: every Designer
// action becomes an ordinary edit of the JSON TextDocument, so VS Code's own
// undo/redo, dirty state, save and hot exit apply, and the edit logic runs in
// the extension (engine.ts) instead of a language server.

import { randomBytes } from "node:crypto";
import * as path from "node:path";
import * as vscode from "vscode";
import { applyToSpec, copyControls, EditError } from "./engine";
import { bodyIndent, eventsForType, eventTarget, findHandler, handlerText } from "./events";
import type { EventDescriptor } from "./events";
import { minimalEdit } from "./jsonSource";
import { designerDocument } from "./model";
import { designerStrings, isWebviewMessage } from "./protocol";
import type { ClipboardSnapshot, DesignerDocument, DesignerError, HostMessage, SemanticOperation, WebviewMessage } from "./protocol";
import type { ParsedSpec } from "./spec";
import { flattenControls, parseSpec, previewDocument, problems, propertyGrid } from "./spec";
import { subtreeIds } from "./toolbox";
import { pasteAllocations } from "./clipboard";

export const designerViewType = "ras.userFormDesigner";
export const FORM_SUFFIX = ".form.json";

interface ClipboardContent {
	snapshot: ClipboardSnapshot;
	/** Copied controls (flat), used to allocate fresh names and IDs on paste. */
	controls: { id: string; type: string }[];
	projectKey: string;
	hasPicturePaths: boolean;
}

/** Shared by every Designer panel of the window. */
class DesignerClipboard {
	private content: ClipboardContent | undefined;
	private readonly repetitions = new Map<string, number>();
	private readonly changed = new vscode.EventEmitter<void>();
	readonly onDidChange = this.changed.event;
	get available() {
		return this.content !== undefined;
	}
	read() {
		return this.content;
	}
	write(content: ClipboardContent) {
		this.content = content;
		this.repetitions.clear();
		this.changed.fire();
	}
	nextOffset(key: string) {
		return 8 * ((this.repetitions.get(key) ?? 0) + 1);
	}
	pasted(key: string) {
		this.repetitions.set(key, (this.repetitions.get(key) ?? 0) + 1);
	}
	dispose() {
		this.changed.dispose();
	}
}

/** Name of the code-behind (`forms/<Name>.bas`) that belongs to a FormSpec. */
export function codeBehindUri(spec: vscode.Uri): vscode.Uri {
	const base = path.posix.basename(spec.path);
	const stem = base.toLowerCase().endsWith(FORM_SUFFIX) ? base.slice(0, -FORM_SUFFIX.length) : base.replace(/\.json$/i, "");
	return vscode.Uri.joinPath(spec, "..", `${stem}.bas`);
}

const projectKey = (uri: vscode.Uri) => {
	const folder = vscode.workspace.getWorkspaceFolder(uri)?.uri.toString() ?? path.posix.dirname(uri.path);
	return process.platform === "win32" ? folder.toLowerCase() : folder;
};

export class UserFormDesignerProvider implements vscode.CustomTextEditorProvider, vscode.Disposable {
	private readonly clipboard = new DesignerClipboard();

	constructor(private readonly extensionUri: vscode.Uri) {}

	static register(context: vscode.ExtensionContext): vscode.Disposable {
		const provider = new UserFormDesignerProvider(context.extensionUri);
		return vscode.Disposable.from(
			provider,
			vscode.window.registerCustomEditorProvider(designerViewType, provider, {
				webviewOptions: { retainContextWhenHidden: true },
				supportsMultipleEditorsPerDocument: true,
			}),
		);
	}

	dispose() {
		this.clipboard.dispose();
	}

	resolveCustomTextEditor(document: vscode.TextDocument, panel: vscode.WebviewPanel): void {
		const resources = vscode.Uri.joinPath(this.extensionUri, "dist", "webview");
		panel.webview.options = { enableScripts: true, localResourceRoots: [resources] };
		panel.webview.html = designerHtml(panel.webview, resources);

		let ready = false;
		let latest: { version: number; spec: ParsedSpec; document: DesignerDocument } | undefined;
		let busy = false;
		const send = (message: HostMessage) => {
			if (ready) void panel.webview.postMessage(message);
		};
		const sendState = () => {
			send({ type: "draftState", dirty: document.isDirty, conflicted: false, saving: false });
			send({ type: "historyState", canUndo: true, canRedo: true });
			send({ type: "clipboardState", available: this.clipboard.available });
		};
		const update = () => {
			const version = document.version;
			const parsed = parseSpec(document.getText());
			if (!parsed.ok) {
				latest = undefined;
				send({ type: "invalidDocument", version, error: { code: "invalidSpec", message: parsed.message } });
				return;
			}
			const spec = parsed.spec;
			const designer = designerDocument(previewDocument(spec), problems(spec), true);
			latest = { version, spec, document: designer };
			send({
				type: "document",
				version,
				document: designer,
				editable: true,
				propertyEditable: true,
				propertyEditorVersion: 1,
				structuralEditable: true,
				parentEditable: true,
				pageEditable: true,
				tabEditable: true,
				productivityEditable: true,
				propertyGrid: propertyGrid(spec),
			});
		};
		const stale = (): DesignerError => ({ code: "staleDocument", message: "ドキュメントが変更されました。更新後のフォームで操作し直してください。" });
		const current = (version: number) => (latest && latest.version === version && document.version === version ? latest : undefined);

		const writeText = async (text: string): Promise<boolean> => {
			const before = document.getText();
			const change = minimalEdit(before, text);
			if (!change) return true;
			const edit = new vscode.WorkspaceEdit();
			edit.replace(document.uri, new vscode.Range(document.positionAt(change.start), document.positionAt(change.end)), change.text);
			return vscode.workspace.applyEdit(edit);
		};
		const run = async (version: number, operations: SemanticOperation[]): Promise<{ error?: DesignerError; selectionIds?: string[] }> => {
			const snapshot = current(version);
			if (!snapshot) return { error: stale() };
			try {
				const cancelled = await this.confirmDeletion(snapshot.document, operations);
				if (cancelled) return { error: cancelled };
				if (document.version !== version) return { error: stale() };
				const result = applyToSpec(snapshot.spec, operations);
				if (!(await writeText(result.text))) return { error: { code: "editFailed", message: "フォームを書き換えられませんでした。" } };
				return result.selectionIds ? { selectionIds: result.selectionIds } : {};
			} catch (error) {
				if (error instanceof EditError) return { error: error.toDesignerError() };
				return { error: { code: "editFailed", message: String(error) } };
			}
		};

		const clipboardAction = async (message: Extract<WebviewMessage, { type: "clipboard" }>) => {
			const snapshot = current(message.version);
			if (!snapshot) return { error: stale() };
			const key = projectKey(document.uri);
			try {
				const capture = (ids: string[]): ClipboardContent => {
					const copied = copyControls(snapshot.spec.text, ids);
					const controls = flattenControls((JSON.parse(copied.snapshot.text) as { controls: [] }).controls).map(({ control }) => ({
						id: control.id,
						type: control.type,
					}));
					return { snapshot: copied.snapshot, controls, projectKey: key, hasPicturePaths: copied.hasPicturePaths };
				};
				const pasteInto = async (content: ClipboardContent, parentId: string, offset: number) => {
					if (content.hasPicturePaths && content.projectKey !== key)
						return { error: { code: "crossProjectPicture", message: "画像のパスを持つコントロールは同じプロジェクト内にだけ貼り付けられます。" } };
					const op: SemanticOperation = {
						type: "pasteControls",
						snapshot: content.snapshot,
						allocations: pasteAllocations(content, snapshot.document),
						parentId,
						offset: { left: offset, top: offset },
					};
					return run(message.version, [op]);
				};
				switch (message.action) {
					case "copy":
						this.clipboard.write(capture(message.controlIds));
						return {};
					case "cut": {
						const content = capture(message.controlIds);
						const result = await run(message.version, [{ type: "removeControls", controlIds: message.controlIds, cascade: true }]);
						if (!result.error) this.clipboard.write(content);
						return result;
					}
					case "paste": {
						const content = this.clipboard.read();
						if (!content) return { error: { code: "clipboardEmpty", message: "クリップボードが空です。" } };
						const pasteKey = JSON.stringify([document.uri.toString(), message.parentId]);
						const result = await pasteInto(content, message.parentId, this.clipboard.nextOffset(pasteKey));
						if (!result.error) this.clipboard.pasted(pasteKey);
						return result;
					}
					case "duplicate": {
						const parentId = snapshot.document.controls.find((c) => c.id === message.controlIds[0])?.parentId ?? "";
						return pasteInto(capture(message.controlIds), parentId, 8);
					}
				}
			} catch (error) {
				if (error instanceof EditError) return { error: error.toDesignerError() };
				return { error: { code: "copyUnavailable", message: String(error) } };
			}
		};

		const subscriptions: vscode.Disposable[] = [
			vscode.workspace.onDidChangeTextDocument((event) => {
				if (event.document === document) {
					update();
					sendState();
				}
			}),
			vscode.workspace.onDidSaveTextDocument((saved) => {
				if (saved === document) sendState();
			}),
			this.clipboard.onDidChange(() => send({ type: "clipboardState", available: this.clipboard.available })),
			vscode.window.onDidChangeActiveColorTheme(() => send({ type: "themeChanged" })),
			panel.webview.onDidReceiveMessage(async (message: unknown) => {
				if (!isWebviewMessage(message)) return;
				const reply = (requestId: number, result: { error?: DesignerError; selectionIds?: string[] }) =>
					send({ type: "editResult", requestId, ...result });
				switch (message.type) {
					case "ready":
						ready = true;
						send({ type: "localization", strings: designerStrings });
						sendState();
						update();
						return;
					case "openText":
						await vscode.window.showTextDocument(document, { viewColumn: vscode.ViewColumn.Beside });
						return;
					case "save":
						await document.save();
						return;
					case "openProblem": {
						const problem = current(message.version)?.document.warnings[message.index];
						const range = problem?.range;
						const selection = range ? new vscode.Range(range.start.line, range.start.character, range.end.line, range.end.character) : undefined;
						await vscode.window.showTextDocument(document, { viewColumn: vscode.ViewColumn.Beside, selection });
						return;
					}
					case "history":
						await vscode.commands.executeCommand(message.action);
						reply(message.requestId, {});
						return;
					case "picturePick":
					case "pictureCommit":
						send({
							type: message.type === "picturePick" ? "pictureSelection" : "editResult",
							requestId: message.requestId,
							error: { code: "unsupported", message: "画像の取り込みには対応していません。JSON の picture.path を直接編集してください。" },
						});
						return;
					case "pictureCancel":
					case "resolveConflict":
						return;
					case "openEvent":
						try {
							await this.openEvent(document, latest?.document, message.targetId, message.choose);
							send({ type: "eventResult", requestId: message.requestId });
						} catch (error) {
							send({ type: "eventResult", requestId: message.requestId, error: error instanceof Error ? error.message : String(error) });
						}
						return;
					case "clipboard":
					case "edit": {
						if (busy) {
							reply(message.requestId, { error: { code: "busy", message: "前の操作を処理しています。" } });
							return;
						}
						busy = true;
						try {
							reply(message.requestId, message.type === "clipboard" ? ((await clipboardAction(message)) ?? {}) : await run(message.version, message.operations));
						} finally {
							busy = false;
						}
						return;
					}
				}
			}),
		];
		panel.onDidDispose(() => subscriptions.forEach((s) => s.dispose()));
		update();
	}

	/** Asks before a deletion that takes descendant controls with it. */
	private async confirmDeletion(document: DesignerDocument, operations: SemanticOperation[]): Promise<DesignerError | undefined> {
		const op = operations[0];
		let ids: string[] = [];
		if (op.type === "removeControl" || op.type === "removePage") ids = [op.controlId];
		else if (op.type === "removeControls") ids = op.controlIds;
		else return;
		const deleted = new Set<string>();
		for (const id of ids) for (const d of subtreeIds(document.controls, id)) deleted.add(d);
		const descendants = deleted.size - ids.length;
		if (!descendants) return;
		const names = ids.map((id) => document.controls.find((c) => c.id === id)?.name ?? id).join(", ");
		const action = "削除";
		const answer = await vscode.window.showWarningMessage(`「${names}」とその中の ${descendants} 個のコントロールを削除しますか?`, { modal: true }, action);
		return answer === action ? undefined : { code: "editCancelled", message: "削除を取り消しました。" };
	}

	/** Opens (or appends) the handler of an event in forms/<Name>.bas. */
	private async openEvent(spec: vscode.TextDocument, document: DesignerDocument | undefined, targetId: string | undefined, choose: boolean) {
		if (!document) throw new Error("フォームを読み込めていません。");
		const target = eventTarget(document.controls, targetId);
		if (!target) throw new Error("このコントロールのイベントは扱えません（カスタム コントロールのハンドラーは手で書いてください）。");
		const events = eventsForType(target.type, target.objectName);
		let event: EventDescriptor | undefined = events.find((e) => e.isDefault) ?? events[0];
		if (choose) {
			const items = [...events]
				.sort((a, b) => Number(b.isDefault) - Number(a.isDefault))
				.map((e) => ({ label: e.name, description: e.signature, event: e }));
			const picked = await vscode.window.showQuickPick(items, { placeHolder: "開くイベントを選んでください" });
			event = picked?.event;
		}
		if (!event) return;
		const uri = codeBehindUri(spec.uri);
		let code: vscode.TextDocument;
		try {
			code = await vscode.workspace.openTextDocument(uri);
		} catch {
			const create = new vscode.WorkspaceEdit();
			create.createFile(uri, { ignoreIfExists: true });
			await vscode.workspace.applyEdit(create);
			code = await vscode.workspace.openTextDocument(uri);
		}
		let handler = findHandler(code.getText(), target.objectName, event.name);
		if (!handler) {
			const text = code.getText();
			const editorConfig = vscode.workspace.getConfiguration("editor", uri);
			const tabSize = editorConfig.get<number>("tabSize", 4);
			const fallback = editorConfig.get<boolean>("insertSpaces", true) ? " ".repeat(tabSize > 0 ? tabSize : 4) : "\t";
			const newline = code.eol === vscode.EndOfLine.CRLF ? "\r\n" : "\n";
			const edit = new vscode.WorkspaceEdit();
			edit.insert(uri, code.positionAt(text.length), handlerText(text, event.signature, bodyIndent(text, fallback), newline));
			await vscode.workspace.applyEdit(edit);
			handler = findHandler(code.getText(), target.objectName, event.name);
		}
		const line = handler?.bodyLine ?? 0;
		const position = new vscode.Position(line, code.lineAt(line).text.length);
		await vscode.window.showTextDocument(code, { viewColumn: vscode.ViewColumn.Beside, selection: new vscode.Range(position, position) });
	}
}

export function designerHtml(webview: vscode.Webview, resources: vscode.Uri): string {
	const nonce = randomBytes(24).toString("hex");
	const script = webview.asWebviewUri(vscode.Uri.joinPath(resources, "userFormDesigner.js"));
	const style = webview.asWebviewUri(vscode.Uri.joinPath(resources, "userFormDesigner.css"));
	return `<!DOCTYPE html><html lang="ja"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'nonce-${nonce}'; style-src ${webview.cspSource}; font-src ${webview.cspSource}; img-src ${webview.cspSource} data:;">
<link rel="stylesheet" href="${style}"><title>${designerStrings.header}</title></head>
<body><div id="app"></div><script nonce="${nonce}" src="${script}"></script></body></html>`;
}
