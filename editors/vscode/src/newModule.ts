// "New standard module": creates src/<name>.ras in a RAS project and opens it.

import * as fs from "node:fs";
import * as path from "node:path";
import * as vscode from "vscode";
import { projectRoot } from "./userForm/commands";

/**
 * The VBA name of the module a source file becomes (mirrors `vba_module_name`
 * in ras-compiler: `str_util` → `StrUtil`, `1st` → `Mod1st`; reserved words
 * are not checked here).
 */
export function vbaModuleName(stem: string): string {
	let out = "";
	let upper = true;
	for (const c of stem) {
		if (c === "_" || c === "-" || c === "." || c === " ") {
			upper = true;
			continue;
		}
		out += upper ? c.toUpperCase() : c;
		upper = false;
	}
	if (out === "" || /^[0-9]/.test(out)) out = `Mod${out}`;
	return [...out].slice(0, 31).join("");
}

/** A file name usable for a module (without `.ras`). */
export function moduleFileNameError(name: string): string | undefined {
	if (name === "") return "名前を入力してください。";
	if (!/^[A-Za-z0-9_-]+$/.test(name)) return "英数字と _ - だけからなる名前を入力してください。";
	if (vbaModuleName(name).toLowerCase() === "vba") return "この名前は使えません。";
	return undefined;
}

/** VBA names of the modules a project already has (src/**\/*.ras and forms/*.ras), lower-cased. */
function existingModuleNames(root: string): Set<string> {
	const names = new Set<string>();
	const walk = (dir: string) => {
		let entries: fs.Dirent[];
		try {
			entries = fs.readdirSync(dir, { withFileTypes: true });
		} catch {
			return;
		}
		for (const entry of entries) {
			if (entry.isDirectory()) walk(path.join(dir, entry.name));
			else if (entry.name.endsWith(".ras") && !entry.name.endsWith(".d.ras")) names.add(vbaModuleName(entry.name.slice(0, -4)).toLowerCase());
		}
	};
	walk(path.join(root, "src"));
	try {
		for (const file of fs.readdirSync(path.join(root, "forms"))) {
			if (file.endsWith(".form.json")) names.add(file.slice(0, -".form.json".length).toLowerCase());
		}
	} catch {
		// No forms folder.
	}
	return names;
}

/** The folder a new module goes to: the clicked folder when it is inside src/, else src/. */
function targetFolder(root: string, resource: vscode.Uri | undefined): string {
	const src = path.join(root, "src");
	if (resource?.scheme !== "file") return src;
	let dir = resource.fsPath;
	try {
		if (!fs.statSync(dir).isDirectory()) dir = path.dirname(dir);
	} catch {
		return src;
	}
	const relative = path.relative(src, dir);
	return relative === "" || (!relative.startsWith("..") && !path.isAbsolute(relative)) ? dir : src;
}

function nextName(dir: string, taken: Set<string>): string {
	for (let i = 1; ; i++) {
		const name = `module${i}`;
		if (!fs.existsSync(path.join(dir, `${name}.ras`)) && !taken.has(vbaModuleName(name).toLowerCase())) return name;
	}
}

async function newModule(resource?: vscode.Uri): Promise<void> {
	const root = await projectRoot(resource, "モジュールを作成するプロジェクトを選んでください");
	if (!root) {
		void vscode.window.showErrorMessage("フォルダーを開いてから標準モジュールを作成してください。");
		return;
	}
	const dir = targetFolder(root, resource);
	const taken = existingModuleNames(root);
	const name = await vscode.window.showInputBox({
		title: "新しい標準モジュール",
		prompt: `${dir} に作成する .ras ファイルの名前（VBA のモジュール名はファイル名から決まります）`,
		value: nextName(dir, taken),
		validateInput: (value) => {
			const stem = value.trim().replace(/\.ras$/i, "");
			const error = moduleFileNameError(stem);
			if (error) return error;
			if (fs.existsSync(path.join(dir, `${stem}.ras`))) return "同じ名前のファイルがすでにあります。";
			const vbaName = vbaModuleName(stem);
			if (taken.has(vbaName.toLowerCase())) return `VBA のモジュール名 ${vbaName} は、ほかのモジュールかフォームが使っています。`;
			return { message: `VBA のモジュール名: ${vbaName}`, severity: vscode.InputBoxValidationSeverity.Info };
		},
	});
	if (name === undefined) return;
	const stem = name.trim().replace(/\.ras$/i, "");
	const file = vscode.Uri.file(path.join(dir, `${stem}.ras`));
	await vscode.workspace.fs.createDirectory(vscode.Uri.file(dir));
	await vscode.workspace.fs.writeFile(file, Buffer.from(newModuleSource(vbaModuleName(stem)), "utf8"));
	const editor = await vscode.window.showTextDocument(file);
	const end = editor.document.lineAt(editor.document.lineCount - 1).range.end;
	editor.selection = new vscode.Selection(end, end);
}

/** The text of a new standard module. */
export function newModuleSource(vbaName: string): string {
	return [
		`// 標準モジュール ${vbaName}。\`export\` した関数はほかのモジュールから呼び出せます。`,
		"// 引数なし・戻り値なしで `export` した関数は、Excel のマクロとして実行できます。",
		"",
		"",
	].join("\n");
}

export function registerModuleCommands(): vscode.Disposable {
	return vscode.commands.registerCommand("ras.module.new", (resource?: unknown) =>
		newModule(resource instanceof vscode.Uri ? resource : undefined).catch((error) =>
			vscode.window.showErrorMessage(`標準モジュールを作成できませんでした: ${error instanceof Error ? error.message : String(error)}`),
		),
	);
}
