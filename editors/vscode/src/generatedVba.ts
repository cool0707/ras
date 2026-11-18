import * as vscode from "vscode";

const SCHEME = "ras-vba";
const PLACEHOLDER = "' 生成 VBA を表示するには、RAS ファイルで「RAS: 生成 VBA を表示」を実行してください。\n";

/** Result of the server command `ras.showGeneratedVba`. */
export interface GeneratedVba {
	moduleName: string;
	text: string;
	/** 0-based line of `text` that corresponds to the cursor, or -1. */
	line: number;
}

export function isGeneratedVba(value: unknown): value is GeneratedVba {
	const candidate = value as Partial<GeneratedVba> | null | undefined;
	return typeof candidate === "object" && candidate !== null
		&& typeof candidate.moduleName === "string"
		&& typeof candidate.text === "string"
		&& typeof candidate.line === "number";
}

/**
 * Shows the VBA the server generated for a RAS file in a read-only virtual
 * document (`ras-vba:` scheme) and highlights the line that corresponds to the
 * cursor (spec §14.8). The text and the line both come from the server.
 */
export class GeneratedVbaView implements vscode.TextDocumentContentProvider, vscode.Disposable {
	private readonly contents = new Map<string, string>();
	private readonly changeEmitter = new vscode.EventEmitter<vscode.Uri>();
	private readonly highlight = vscode.window.createTextEditorDecorationType({
		isWholeLine: true,
		backgroundColor: new vscode.ThemeColor("editor.findMatchHighlightBackground"),
		overviewRulerColor: new vscode.ThemeColor("editorOverviewRuler.findMatchForeground"),
		overviewRulerLane: vscode.OverviewRulerLane.Full,
	});
	private readonly subscriptions: vscode.Disposable[];

	readonly onDidChange = this.changeEmitter.event;

	constructor() {
		this.subscriptions = [
			vscode.workspace.registerTextDocumentContentProvider(SCHEME, this),
			vscode.workspace.onDidCloseTextDocument((document) => {
				if (document.uri.scheme === SCHEME) {
					this.contents.delete(document.uri.toString());
				}
			}),
		];
	}

	provideTextDocumentContent(uri: vscode.Uri): string {
		return this.contents.get(uri.toString()) ?? PLACEHOLDER;
	}

	/** Opens (or refreshes) the generated VBA of `source` beside the editor and highlights `generated.line`. */
	async show(source: vscode.Uri, generated: GeneratedVba): Promise<void> {
		// One virtual document per RAS file. The `.vba` extension selects the
		// built-in `vb` language; the query keeps same-named modules apart.
		const moduleName = generated.moduleName.replace(/[\\/]/g, "_") || "module";
		const uri = vscode.Uri.from({ scheme: SCHEME, path: `/${moduleName}.vba`, query: source.toString() });
		const key = uri.toString();

		const alreadyOpen = vscode.workspace.textDocuments.find((document) => document.uri.toString() === key);
		this.contents.set(key, generated.text);
		if (alreadyOpen && alreadyOpen.getText() !== generated.text) {
			await this.refresh(uri);
		}

		let document = await vscode.workspace.openTextDocument(uri);
		if (document.languageId !== "vb") {
			try {
				document = await vscode.languages.setTextDocumentLanguage(document, "vb");
			} catch {
				// The built-in Visual Basic extension is disabled; show plain text.
			}
		}

		const visible = vscode.window.visibleTextEditors.find((editor) => editor.document.uri.toString() === key);
		const editor = await vscode.window.showTextDocument(document, {
			viewColumn: visible?.viewColumn ?? vscode.ViewColumn.Beside,
			preserveFocus: true,
			preview: true,
		});

		if (Number.isInteger(generated.line) && generated.line >= 0 && generated.line < editor.document.lineCount) {
			const range = editor.document.lineAt(generated.line).range;
			editor.selection = new vscode.Selection(range.start, range.start);
			editor.revealRange(range, vscode.TextEditorRevealType.InCenterIfOutsideViewport);
			editor.setDecorations(this.highlight, [range]);
		} else {
			editor.setDecorations(this.highlight, []);
		}
	}

	dispose(): void {
		for (const subscription of this.subscriptions) {
			subscription.dispose();
		}
		this.changeEmitter.dispose();
		this.highlight.dispose();
		this.contents.clear();
	}

	/** Asks VS Code to reload an open virtual document and waits until it has. */
	private refresh(uri: vscode.Uri): Promise<void> {
		return new Promise<void>((resolve) => {
			const key = uri.toString();
			const finish = () => {
				clearTimeout(timeout);
				listener.dispose();
				resolve();
			};
			const listener = vscode.workspace.onDidChangeTextDocument((event) => {
				if (event.document.uri.toString() === key) {
					finish();
				}
			});
			// Fallback for the case where VS Code sees no difference and fires no event.
			const timeout = setTimeout(finish, 1000);
			this.changeEmitter.fire(uri);
		});
	}
}
