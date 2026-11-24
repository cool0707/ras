// The explorer: the project's files as a tree of folders, as the Explorer of VS Code shows them.
// Files are opened, created, renamed and deleted from it (in place, in the tree), with the mouse,
// a context menu or the keyboard (↑↓ ←→ Enter, F2, Delete).

export interface ExplorerHooks {
    /** The project's files (their paths). */
    files(): string[];
    /** The file open in the editor ("" for none). */
    active(): string;
    /** The errors of each file. */
    errors(): Map<string, number>;
    open(path: string): void;
    /** A new file at `path` (the name typed in a folder): false keeps the name being typed. */
    create(path: string): boolean;
    /** The file renamed: false keeps the name being typed. */
    rename(from: string, to: string): boolean;
    remove(path: string): void;
    /** Why a file cannot be renamed or deleted (undefined: it can). */
    fixed(path: string): { rename?: string; remove?: string };
    /** A message for the user (the page's toast). */
    message(text: string): void;
}

interface Node {
    name: string;
    path: string;
    folder: boolean;
    children: Node[];
}

type Editing = { kind: "rename"; path: string } | { kind: "new"; folder: string };

const STORE = "ras-playground:explorer";

function tree(paths: string[]): Node[] {
    const root: Node = { name: "", path: "", folder: true, children: [] };
    for (const p of paths) {
        const parts = p.split("/");
        let cur = root;
        for (let i = 0; i < parts.length - 1; i++) {
            const path = parts.slice(0, i + 1).join("/");
            let next = cur.children.find((c) => c.folder && c.path === path);
            if (!next) {
                next = { name: parts[i], path, folder: true, children: [] };
                cur.children.push(next);
            }
            cur = next;
        }
        cur.children.push({ name: parts[parts.length - 1], path: p, folder: false, children: [] });
    }
    const sort = (nodes: Node[]) => {
        nodes.sort((a, b) => (a.folder !== b.folder ? (a.folder ? -1 : 1) : a.name.localeCompare(b.name, undefined, { sensitivity: "base", numeric: true })));
        for (const n of nodes) sort(n.children);
    };
    sort(root.children);
    return root.children;
}

/** The codicon of a file. */
function icon(path: string): string {
    if (path === "ras.toml") return "settings-gear";
    if (path.endsWith(".form.json")) return "window";
    if (path.endsWith(".test.ras")) return "beaker";
    if (path.endsWith(".d.ras")) return "symbol-interface";
    if (path.endsWith(".ras")) return "file-code";
    return "file";
}

function parentOf(path: string): string {
    const i = path.lastIndexOf("/");
    return i < 0 ? "" : path.slice(0, i);
}

function codicon(name: string): HTMLElement {
    const i = document.createElement("i");
    i.className = `codicon codicon-${name}`;
    i.setAttribute("aria-hidden", "true");
    return i;
}

export class Explorer {
    private collapsed = new Set<string>();
    /** The row with the keyboard focus (a file or folder path). */
    private focused = "";
    private editing: Editing | null = null;
    private menu: HTMLElement | null = null;
    /** While the tree is drawn again (the input it removes is not left for that). */
    private rendering = false;

    constructor(
        private readonly el: HTMLElement,
        private readonly hooks: ExplorerHooks,
    ) {
        try {
            const saved = JSON.parse(localStorage.getItem(STORE) || "null");
            if (saved && Array.isArray(saved.collapsed)) for (const p of saved.collapsed) if (typeof p === "string") this.collapsed.add(p);
        } catch {
            // All open.
        }
        el.setAttribute("role", "tree");
        el.addEventListener("keydown", (e) => this.key(e));
        // Right-click on the empty part of the tree: the root's menu.
        el.addEventListener("contextmenu", (e) => {
            if (e.target !== el) return;
            e.preventDefault();
            this.showMenu(e.clientX, e.clientY, [
                { label: "新しいファイル…", icon: "new-file", run: () => this.newFile("src") },
                { label: "新しいフォーム…", icon: "window", run: () => this.newFile("forms") },
            ]);
        });
        document.addEventListener("pointerdown", (e) => {
            if (this.menu && !this.menu.contains(e.target as globalThis.Node)) this.closeMenu();
        });
        window.addEventListener("blur", () => this.closeMenu());
    }

    /** A new file typed in `folder` (the folder of the focused row when not given). */
    newFile(folder?: string): void {
        if (folder === undefined) {
            const f = this.focused || this.hooks.active();
            folder = this.hooks.files().includes(f) ? parentOf(f) : f;
            if (!folder) folder = "src";
        }
        // The folder and those above it are opened.
        for (let p = folder; p; p = parentOf(p)) this.collapsed.delete(p);
        this.editing = { kind: "new", folder };
        this.render();
    }

    rename(path: string): void {
        const why = this.hooks.fixed(path).rename;
        if (why) {
            this.hooks.message(why);
            return;
        }
        this.editing = { kind: "rename", path };
        this.render();
    }

    remove(path: string): void {
        const why = this.hooks.fixed(path).remove;
        if (why) {
            this.hooks.message(why);
            return;
        }
        if (!confirm(`${path} を削除しますか？`)) return;
        this.hooks.remove(path);
    }

    collapseAll(): void {
        const all = new Set<string>();
        for (const p of this.hooks.files()) for (let d = parentOf(p); d; d = parentOf(d)) all.add(d);
        this.collapsed = all;
        // The focus goes to the top folder of the row it was on.
        if (this.focused.includes("/")) this.focused = this.focused.split("/")[0];
        this.keep();
        this.render();
    }

    /** The active file shown (its folders opened). */
    reveal(path: string): void {
        let changed = false;
        for (let p = parentOf(path); p; p = parentOf(p)) changed = this.collapsed.delete(p) || changed;
        if (changed) this.keep();
    }

    private keep() {
        try {
            localStorage.setItem(STORE, JSON.stringify({ collapsed: [...this.collapsed] }));
        } catch {
            // Not kept without storage.
        }
    }

    render(): void {
        // The name being typed is kept when the tree is drawn again (after a compile).
        const typing = this.el.querySelector<HTMLInputElement>(".tree-input");
        const kept = typing && this.editing ? { value: typing.value, start: typing.selectionStart, end: typing.selectionEnd, focus: document.activeElement === typing } : null;
        const hadFocus = this.el.contains(document.activeElement) && !typing;
        this.rendering = true;
        try {
            this.el.textContent = "";
        } finally {
            this.rendering = false;
        }
        const files = this.hooks.files();
        const active = this.hooks.active();
        const errors = this.hooks.errors();
        if (this.focused && !files.includes(this.focused) && !files.some((f) => f.startsWith(`${this.focused}/`))) this.focused = "";
        if (!this.focused) this.focused = active;
        const errorsIn = (n: Node): number => (n.folder ? n.children.reduce((s, c) => s + errorsIn(c), 0) : errors.get(n.path) || 0);
        let input: HTMLInputElement | null = null;
        const walk = (nodes: Node[], depth: number, folder: string) => {
            if (this.editing?.kind === "new" && this.editing.folder === folder) input = this.inputRow(depth, folder, "");
            for (const n of nodes) {
                if (this.editing?.kind === "rename" && this.editing.path === n.path) {
                    input = this.inputRow(depth, folder, n.name);
                    continue;
                }
                const row = document.createElement("div");
                row.className = "tree-row";
                row.setAttribute("role", "treeitem");
                row.setAttribute("aria-level", String(depth + 1));
                row.dataset.path = n.path;
                row.style.paddingLeft = `${6 + depth * 12}px`;
                row.title = n.path;
                row.tabIndex = n.path === this.focused ? 0 : -1;
                const errs = errorsIn(n);
                if (errs) row.classList.add("has-error");
                const open = n.folder && !this.collapsed.has(n.path);
                if (n.folder) {
                    row.setAttribute("aria-expanded", String(open));
                    row.append(codicon(open ? "chevron-down" : "chevron-right"));
                } else {
                    row.setAttribute("aria-selected", String(n.path === active));
                    if (n.path === active) row.classList.add("active");
                    row.append(el("span", "twistie"), codicon(icon(n.path)));
                }
                if (n.path === this.focused) row.classList.add("focused");
                row.append(el("span", "tree-name", n.name));
                if (errs) {
                    const badge = el("span", "tree-badge", n.folder ? "●" : String(errs));
                    badge.title = `エラー ${errs} 件`;
                    row.append(badge);
                }
                const actions = el("span", "tree-actions");
                if (n.folder) {
                    actions.append(this.action("new-file", n.path.split("/")[0] === "forms" ? "新しいフォーム" : "新しいファイル", () => this.newFile(n.path)));
                } else {
                    const fixed = this.hooks.fixed(n.path);
                    if (!fixed.rename) actions.append(this.action("edit", "名前の変更（F2）", () => this.rename(n.path)));
                    if (!fixed.remove) actions.append(this.action("trash", "削除（Delete）", () => this.remove(n.path)));
                }
                row.append(actions);
                row.onclick = () => {
                    this.focused = n.path;
                    if (n.folder) this.toggle(n.path);
                    else this.hooks.open(n.path);
                };
                row.oncontextmenu = (e) => {
                    e.preventDefault();
                    this.focused = n.path;
                    this.render();
                    this.showMenu(e.clientX, e.clientY, this.menuOf(n));
                };
                this.el.append(row);
                if (open) walk(n.children, depth + 1, n.path);
            }
        };
        const nodes = tree(files);
        // A new file in a folder that has none yet (forms/ of a project without forms).
        if (this.editing?.kind === "new" && this.editing.folder && !files.some((f) => f.startsWith(`${(this.editing as { folder: string }).folder}/`))) {
            const parts = this.editing.folder.split("/");
            let list = nodes;
            for (let i = 0; i < parts.length; i++) {
                const path = parts.slice(0, i + 1).join("/");
                let n = list.find((x) => x.folder && x.path === path);
                if (!n) {
                    n = { name: parts[i], path, folder: true, children: [] };
                    list.unshift(n);
                }
                list = n.children;
            }
        }
        walk(nodes, 0, "");
        if (input) {
            const i = input as HTMLInputElement;
            if (kept) {
                i.value = kept.value;
                if (kept.focus) i.focus();
                i.setSelectionRange(kept.start, kept.end);
            } else {
                i.focus();
                const dot = i.value.indexOf(".");
                i.setSelectionRange(0, dot > 0 ? dot : i.value.length);
                i.scrollIntoView({ block: "nearest" });
            }
        } else if (hadFocus) {
            this.row(this.focused)?.focus();
        }
        this.el.querySelector<HTMLElement>(".tree-row.active")?.scrollIntoView({ block: "nearest" });
    }

    private action(name: string, title: string, run: () => void): HTMLElement {
        const b = document.createElement("button");
        b.className = "tree-action";
        b.title = title;
        b.setAttribute("aria-label", title);
        b.tabIndex = -1;
        b.append(codicon(name));
        b.onclick = (e) => {
            e.stopPropagation();
            run();
        };
        return b;
    }

    private menuOf(n: Node): MenuItem[] {
        if (n.folder) {
            const forms = n.path.split("/")[0] === "forms";
            return [
                { label: forms ? "新しいフォーム…" : "新しいファイル…", icon: "new-file", run: () => this.newFile(n.path) },
                { label: this.collapsed.has(n.path) ? "展開" : "折りたたむ", icon: this.collapsed.has(n.path) ? "expand-all" : "collapse-all", run: () => this.toggle(n.path) },
                { label: "パスのコピー", icon: "copy", run: () => void navigator.clipboard?.writeText(n.path) },
            ];
        }
        const fixed = this.hooks.fixed(n.path);
        return [
            { label: "開く", icon: icon(n.path), run: () => this.hooks.open(n.path) },
            { label: "新しいファイル…", icon: "new-file", run: () => this.newFile(parentOf(n.path) || "src") },
            { label: "名前の変更…", icon: "edit", key: "F2", disabled: !!fixed.rename, run: () => this.rename(n.path) },
            { label: "削除", icon: "trash", key: "Delete", disabled: !!fixed.remove, run: () => this.remove(n.path) },
            { label: "パスのコピー", icon: "copy", run: () => void navigator.clipboard?.writeText(n.path) },
        ];
    }

    private toggle(path: string) {
        if (this.collapsed.has(path)) this.collapsed.delete(path);
        else this.collapsed.add(path);
        this.keep();
        this.render();
        this.row(path)?.focus();
    }

    private row(path: string): HTMLElement | null {
        for (const r of this.el.querySelectorAll<HTMLElement>(".tree-row")) if (r.dataset.path === path) return r;
        return null;
    }

    /** The row a name is typed in: a new file in `folder`, or the file renamed. */
    private inputRow(depth: number, folder: string, name: string): HTMLInputElement {
        const editing = this.editing!;
        const row = el("div", "tree-row editing");
        row.style.paddingLeft = `${6 + depth * 12}px`;
        const path = editing.kind === "rename" ? editing.path : `${folder ? `${folder}/` : ""}new.ras`;
        row.append(el("span", "twistie"), codicon(icon(path)));
        const input = document.createElement("input");
        input.className = "tree-input";
        input.value = name;
        input.spellcheck = false;
        input.setAttribute("aria-label", editing.kind === "rename" ? `${editing.path} の新しい名前` : `${folder || "ルート"} に追加するファイルの名前`);
        input.placeholder = folder.split("/")[0] === "forms" ? "フォームの名前" : folder.split("/")[0] === "test" ? "〜.test.ras" : "〜.ras";
        let done = false;
        const finish = (commit: boolean) => {
            if (done) return;
            const typed = input.value.trim();
            if (commit && typed && !(editing.kind === "rename" && typed === name)) {
                const to = `${folder ? `${folder}/` : ""}${typed}`;
                // The tree the hook draws again is without the input.
                done = true;
                this.editing = null;
                const ok = editing.kind === "rename" ? this.hooks.rename(editing.path, to) : this.hooks.create(to);
                if (!ok) {
                    done = false;
                    this.editing = editing;
                    if (input.isConnected) input.focus();
                    else this.render();
                    return;
                }
                this.focused = editing.kind === "rename" ? to : this.hooks.active();
            }
            done = true;
            this.editing = null;
            this.render();
            this.row(this.focused)?.focus();
        };
        input.onkeydown = (e) => {
            e.stopPropagation();
            if (e.key === "Enter") {
                e.preventDefault();
                finish(true);
            } else if (e.key === "Escape") {
                e.preventDefault();
                finish(false);
            }
        };
        // A click elsewhere: the name taken as typed (an empty one, nothing), as VS Code does.
        input.onblur = () => {
            if (done || this.rendering || !input.isConnected) return;
            const typed = input.value.trim();
            if (!typed || (editing.kind === "rename" && typed === name)) finish(false);
            else finish(true);
            // A name refused on blur is given up rather than keep the focus.
            if (!done) finish(false);
        };
        row.append(input);
        this.el.append(row);
        return input;
    }

    private key(e: KeyboardEvent) {
        if (e.target instanceof HTMLInputElement) return;
        const rows = [...this.el.querySelectorAll<HTMLElement>(".tree-row[data-path]")];
        const i = rows.findIndex((r) => r.dataset.path === this.focused);
        const cur = rows[i];
        const path = cur?.dataset.path ?? "";
        const isFolder = cur?.hasAttribute("aria-expanded") ?? false;
        const go = (r: HTMLElement | undefined) => {
            if (!r) return;
            this.focused = r.dataset.path!;
            for (const x of rows) {
                x.tabIndex = x === r ? 0 : -1;
                x.classList.toggle("focused", x === r);
            }
            r.focus();
            r.scrollIntoView({ block: "nearest" });
        };
        switch (e.key) {
            case "ArrowDown":
                go(rows[Math.min(i + 1, rows.length - 1)]);
                break;
            case "ArrowUp":
                go(rows[Math.max(i - 1, 0)]);
                break;
            case "Home":
                go(rows[0]);
                break;
            case "End":
                go(rows[rows.length - 1]);
                break;
            case "ArrowRight":
                if (isFolder && this.collapsed.has(path)) this.toggle(path);
                else if (isFolder) go(rows[i + 1]);
                break;
            case "ArrowLeft":
                if (isFolder && !this.collapsed.has(path)) this.toggle(path);
                else go(this.row(parentOf(path)) ?? undefined);
                break;
            case "Enter":
            case " ":
                if (isFolder) this.toggle(path);
                else if (path) this.hooks.open(path);
                break;
            case "F2":
                if (path && !isFolder) this.rename(path);
                break;
            case "Delete":
                if (path && !isFolder) this.remove(path);
                break;
            default:
                return;
        }
        e.preventDefault();
    }

    private showMenu(x: number, y: number, items: MenuItem[]) {
        this.closeMenu();
        const menu = el("div", "menu");
        menu.setAttribute("role", "menu");
        for (const it of items) {
            const b = document.createElement("button");
            b.className = "menu-item";
            b.setAttribute("role", "menuitem");
            b.disabled = !!it.disabled;
            b.append(codicon(it.icon), el("span", "menu-label", it.label));
            if (it.key) b.append(el("span", "menu-key", it.key));
            b.onclick = () => {
                this.closeMenu();
                it.run();
            };
            menu.append(b);
        }
        menu.addEventListener("keydown", (e) => {
            const all = [...menu.querySelectorAll<HTMLButtonElement>(".menu-item:not(:disabled)")];
            const at = all.indexOf(document.activeElement as HTMLButtonElement);
            if (e.key === "Escape") {
                this.closeMenu();
                this.row(this.focused)?.focus();
            } else if (e.key === "ArrowDown") all[(at + 1) % all.length]?.focus();
            else if (e.key === "ArrowUp") all[(at - 1 + all.length) % all.length]?.focus();
            else return;
            e.preventDefault();
        });
        document.body.append(menu);
        const r = menu.getBoundingClientRect();
        menu.style.left = `${Math.max(4, Math.min(x, innerWidth - r.width - 4))}px`;
        menu.style.top = `${Math.max(4, Math.min(y, innerHeight - r.height - 4))}px`;
        this.menu = menu;
        menu.querySelector<HTMLButtonElement>(".menu-item:not(:disabled)")?.focus();
    }

    private closeMenu() {
        this.menu?.remove();
        this.menu = null;
    }
}

interface MenuItem {
    label: string;
    icon: string;
    key?: string;
    disabled?: boolean;
    run(): void;
}

function el<K extends keyof HTMLElementTagNameMap>(tag: K, cls?: string, text?: string): HTMLElementTagNameMap[K] {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined) e.textContent = text;
    return e;
}
