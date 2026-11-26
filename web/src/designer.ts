// The UserForm Designer of the playground: the VS Code extension's Designer (its webview,
// editors/vscode/webview/userFormDesigner) in a frame of the page, over a form's
// forms/<Name>.form.json held in memory. What the extension's provider (src/userForm/provider.ts)
// does with VS Code's documents is done here with the playground's files: every edit of the
// Designer becomes new text of the file, and its events become handlers of forms/<Name>.ras.

import { pasteAllocations } from "../../editors/vscode/src/userForm/clipboard";
import { applyToSpec, copyControls, EditError } from "../../editors/vscode/src/userForm/engine";
import { bodyIndent, eventsForType, eventTarget, findHandler, handlerText, rasParams } from "../../editors/vscode/src/userForm/events";
import type { EventDescriptor } from "../../editors/vscode/src/userForm/events";
import { designerDocument } from "../../editors/vscode/src/userForm/model";
import { designerStrings, isWebviewMessage } from "../../editors/vscode/src/userForm/protocol";
import type { ClipboardSnapshot, DesignerDocument, DesignerError, HostMessage, SemanticOperation, WebviewMessage } from "../../editors/vscode/src/userForm/protocol";
import type { ParsedSpec } from "../../editors/vscode/src/userForm/spec";
import { flattenControls, parseSpec, previewDocument, problems, propertyGrid } from "../../editors/vscode/src/userForm/spec";
import { subtreeIds } from "../../editors/vscode/src/userForm/toolbox";
import { versioned } from "./versioned";

/** What the Designer needs from the page. */
export interface DesignerHooks {
    /** The text of a file of the project (undefined: there is none). */
    text(path: string): string | undefined;
    /** Writes a file of the project (made when there is none). */
    write(path: string, text: string): void;
    /** Shows the form's JSON in the text editor. */
    openText(path: string, at?: { line: number; col: number }): void;
    /** Opens a file of the project at a line (0-based), the cursor at its end. */
    openCode(path: string, line: number): void;
    /** Asks for one of the items (its index), or undefined when cancelled. */
    pick(title: string, items: { label: string; description?: string; detail?: string }[]): Promise<number | undefined>;
}

/** The code-behind of a form: `forms/<Name>.ras` for `forms/<Name>.form.json`. */
export const codeBehindPath = (spec: string) => spec.replace(/\.form\.json$/i, ".ras");

interface ClipboardContent {
    snapshot: ClipboardSnapshot;
    controls: { id: string; type: string }[];
}

const stale = (): DesignerError => ({ code: "staleDocument", message: "ドキュメントが変更されました。更新後のフォームで操作し直してください。" });

export class FormDesigner {
    /** The form shown (its `.form.json`), or "" when the Designer is hidden. */
    private path = "";
    private ready = false;
    private version = 0;
    private latest: { version: number; spec: ParsedSpec; document: DesignerDocument } | undefined;
    private busy = false;
    /** The texts before (undo) and after (redo) the Designer's edits, per form. */
    private readonly history = new Map<string, { undo: string[]; redo: string[] }>();
    private clipboard: ClipboardContent | undefined;
    private readonly repetitions = new Map<string, number>();

    constructor(
        private readonly frame: HTMLIFrameElement,
        private readonly hooks: DesignerHooks,
    ) {
        window.addEventListener("message", (e) => {
            if (e.source !== this.frame.contentWindow) return;
            const m = e.data && (e.data as { rasFrame?: unknown }).rasFrame;
            if (isWebviewMessage(m)) void this.receive(m);
            else this.unknown(m);
        });
    }

    get shown(): string {
        return this.path;
    }

    /** Shows the form of `path` (`forms/<Name>.form.json`). */
    show(path: string): void {
        this.frame.hidden = false;
        if (this.path === path && this.ready) {
            this.update();
            return;
        }
        this.path = path;
        this.version++;
        if (!this.frame.src) {
            this.frame.src = versioned("designer.html");
        } else if (this.ready) {
            // A Designer that showed another form starts over (its selection and history are that form's).
            this.ready = false;
            this.frame.contentWindow?.location.reload();
        }
    }

    hide(): void {
        this.frame.hidden = true;
        this.path = "";
    }

    /** The text of a form changed (in the text editor, or the file was renamed or removed). */
    changed(path: string): void {
        if (path !== this.path) return;
        this.version++;
        this.update();
    }

    /** A form renamed: its history goes with it. */
    renamed(from: string, to: string): void {
        const h = this.history.get(from);
        if (h) {
            this.history.delete(from);
            this.history.set(to, h);
        }
        if (this.path === from) this.path = to;
    }

    /**
     * An edit this page does not know (a Designer of another build, e.g. one kept in the
     * browser's cache): it is answered, so that the Designer does not wait for it.
     */
    private unknown(m: unknown): void {
        const request = m as { type?: unknown; requestId?: unknown } | null;
        if (!request || typeof request !== "object" || typeof request.requestId !== "number") return;
        if (request.type !== "edit" && request.type !== "clipboard" && request.type !== "history") return;
        this.send({
            type: "editResult",
            requestId: request.requestId,
            error: { code: "editFailed", message: "この操作を処理できませんでした。ページを再読み込みしてください。" },
        });
    }

    private send(message: HostMessage): void {
        if (this.ready) this.frame.contentWindow?.postMessage(message, "*");
    }

    private sendState(): void {
        const h = this.history.get(this.path);
        this.send({ type: "draftState", dirty: false, conflicted: false, saving: false });
        this.send({ type: "historyState", canUndo: !!h?.undo.length, canRedo: !!h?.redo.length });
        this.send({ type: "clipboardState", available: this.clipboard !== undefined });
    }

    private update(): void {
        if (!this.path) return;
        const text = this.hooks.text(this.path);
        if (text === undefined) return;
        const version = this.version;
        const parsed = parseSpec(text);
        if (!parsed.ok) {
            this.latest = undefined;
            this.send({ type: "invalidDocument", version, error: { code: "invalidSpec", message: parsed.message } });
            return;
        }
        const preview = previewDocument(parsed.spec);
        const designer = designerDocument(preview, problems(parsed.spec), true);
        this.latest = { version, spec: parsed.spec, document: designer };
        this.send({
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
            propertyGrid: propertyGrid(parsed.spec),
        });
        this.sendState();
    }

    private current(version: number) {
        return this.latest && this.latest.version === version && this.version === version ? this.latest : undefined;
    }

    /** Writes the form's new text as an edit of the Designer (one step of its undo). */
    private write(text: string): void {
        const before = this.hooks.text(this.path);
        if (before === undefined || before === text) return;
        const h = this.history.get(this.path) ?? { undo: [], redo: [] };
        h.undo.push(before);
        if (h.undo.length > 100) h.undo.shift();
        h.redo = [];
        this.history.set(this.path, h);
        this.version++;
        this.hooks.write(this.path, text);
        this.update();
    }

    private run(version: number, operations: SemanticOperation[]): { error?: DesignerError; selectionIds?: string[] } {
        const snapshot = this.current(version);
        if (!snapshot) return { error: stale() };
        try {
            const cancelled = confirmDeletion(snapshot.document, operations);
            if (cancelled) return { error: cancelled };
            const result = applyToSpec(snapshot.spec, operations);
            this.write(result.text);
            return result.selectionIds ? { selectionIds: result.selectionIds } : {};
        } catch (error) {
            if (error instanceof EditError) return { error: error.toDesignerError() };
            return { error: { code: "editFailed", message: String(error) } };
        }
    }

    private travel(redo: boolean): void {
        const h = this.history.get(this.path);
        const text = this.hooks.text(this.path);
        if (!h || text === undefined) return;
        const [from, to] = redo ? [h.redo, h.undo] : [h.undo, h.redo];
        const next = from.pop();
        if (next === undefined) return;
        to.push(text);
        this.version++;
        this.hooks.write(this.path, next);
        this.update();
    }

    private clipboardAction(message: Extract<WebviewMessage, { type: "clipboard" }>): { error?: DesignerError; selectionIds?: string[] } {
        const snapshot = this.current(message.version);
        if (!snapshot) return { error: stale() };
        try {
            const capture = (ids: string[]): ClipboardContent => {
                const copied = copyControls(snapshot.spec.text, ids);
                const controls = flattenControls((JSON.parse(copied.snapshot.text) as { controls: [] }).controls).map(({ control }) => ({ id: control.id, type: control.type }));
                return { snapshot: copied.snapshot, controls };
            };
            const pasteInto = (content: ClipboardContent, parentId: string, offset: number) =>
                this.run(message.version, [
                    {
                        type: "pasteControls",
                        snapshot: content.snapshot,
                        allocations: pasteAllocations(content, snapshot.document, () => crypto.randomUUID()),
                        parentId,
                        offset: { left: offset, top: offset },
                    },
                ]);
            switch (message.action) {
                case "copy":
                    this.clipboard = capture(message.controlIds);
                    this.repetitions.clear();
                    this.sendState();
                    return {};
                case "cut": {
                    const content = capture(message.controlIds);
                    const result = this.run(message.version, [{ type: "removeControls", controlIds: message.controlIds, cascade: true }]);
                    if (!result.error) {
                        this.clipboard = content;
                        this.repetitions.clear();
                        this.sendState();
                    }
                    return result;
                }
                case "paste": {
                    const content = this.clipboard;
                    if (!content) return { error: { code: "clipboardEmpty", message: "クリップボードが空です。" } };
                    const key = JSON.stringify([this.path, message.parentId]);
                    const times = this.repetitions.get(key) ?? 0;
                    const result = pasteInto(content, message.parentId, 8 * (times + 1));
                    if (!result.error) this.repetitions.set(key, times + 1);
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
        return {};
    }

    /** Opens (or appends) the handler of an event in the code-behind (forms/<Name>.ras). */
    private async openEvent(targetId: string | undefined, choose: boolean): Promise<void> {
        const document = this.latest?.document;
        if (!document) throw new Error("フォームを読み込めていません。");
        const target = eventTarget(document.controls, targetId);
        if (!target) throw new Error("このコントロールのイベントは扱えません（カスタム コントロールのハンドラーは手で書いてください）。");
        const events = eventsForType(target.type, target.objectName);
        const code = codeBehindPath(this.path);
        let event: EventDescriptor | undefined = events.find((e) => e.isDefault) ?? events[0];
        if (choose) {
            const sorted = [...events].sort((a, b) => Number(b.isDefault) - Number(a.isDefault));
            const text = this.hooks.text(code) ?? "";
            const i = await this.hooks.pick(
                `${target.objectName} の開くイベント`,
                sorted.map((e) => ({ label: e.name, description: `(${rasParams(e.signature)})`, detail: findHandler(text, target.objectName, e.name) ? "ハンドラーがあります" : undefined })),
            );
            event = i === undefined ? undefined : sorted[i];
        }
        if (!event) return;
        openHandler(this.hooks, code, target.objectName, event);
    }

    private async receive(message: WebviewMessage): Promise<void> {
        const reply = (requestId: number, result: { error?: DesignerError; selectionIds?: string[] }) => this.send({ type: "editResult", requestId, ...result });
        switch (message.type) {
            case "ready":
                this.ready = true;
                this.send({ type: "localization", strings: designerStrings });
                this.update();
                return;
            case "openText":
                this.hooks.openText(this.path);
                return;
            case "save":
                return;
            case "openProblem": {
                const range = this.current(message.version)?.document.warnings[message.index]?.range;
                this.hooks.openText(this.path, range ? { line: range.start.line, col: range.start.character } : undefined);
                return;
            }
            case "history":
                this.travel(message.action === "redo");
                reply(message.requestId, {});
                return;
            case "picturePick":
                this.send({
                    type: "pictureSelection",
                    requestId: message.requestId,
                    error: { code: "pictureUnavailable", message: "プレイグラウンドでは画像を設定できません（VS Code の拡張機能で設定してください）。" },
                });
                return;
            case "pictureCancel":
            case "resolveConflict":
                return;
            case "openEvent":
                try {
                    await this.openEvent(message.targetId, message.choose);
                    this.send({ type: "eventResult", requestId: message.requestId });
                } catch (error) {
                    this.send({ type: "eventResult", requestId: message.requestId, error: error instanceof Error ? error.message : String(error) });
                }
                return;
            case "pictureCommit":
                reply(message.requestId, { error: { code: "pictureUnavailable", message: "プレイグラウンドでは画像を設定できません。" } });
                return;
            case "clipboard":
            case "edit": {
                if (this.busy) {
                    reply(message.requestId, { error: { code: "busy", message: "前の操作を処理しています。" } });
                    return;
                }
                this.busy = true;
                try {
                    reply(message.requestId, message.type === "clipboard" ? this.clipboardAction(message) : this.run(message.version, message.operations));
                } finally {
                    this.busy = false;
                }
                return;
            }
        }
    }
}

/** Asks before a deletion that takes descendant controls with it. */
function confirmDeletion(document: DesignerDocument, operations: SemanticOperation[]): DesignerError | undefined {
    const op = operations[0];
    let ids: string[] = [];
    if (op.type === "removeControl" || op.type === "removePage") ids = [op.controlId];
    else if (op.type === "removeControls") ids = op.controlIds;
    else return undefined;
    const deleted = new Set<string>();
    for (const id of ids) for (const d of subtreeIds(document.controls, id)) deleted.add(d);
    const descendants = deleted.size - ids.length;
    if (!descendants) return undefined;
    const names = ids.map((id) => document.controls.find((c) => c.id === id)?.name ?? id).join(", ");
    return confirm(`「${names}」とその中の ${descendants} 個のコントロールを削除しますか?`) ? undefined : { code: "editCancelled", message: "削除を取り消しました。" };
}

/** Opens the `@on("<object>.<event>")` handler of a file, appended when there is none. */
export function openHandler(hooks: DesignerHooks, file: string, objectName: string, event: EventDescriptor): void {
    let text = hooks.text(file) ?? "";
    let handler = findHandler(text, objectName, event.name);
    if (!handler) {
        text += handlerText(text, objectName, event, bodyIndent(text, "    "), "\n");
        hooks.write(file, text);
        handler = findHandler(text, objectName, event.name);
    }
    hooks.openCode(file, handler?.bodyLine ?? 0);
}
