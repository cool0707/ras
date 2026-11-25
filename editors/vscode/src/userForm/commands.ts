// Commands around the UserForm Designer: create a form, open a FormSpec in the Designer.

import * as fs from "node:fs";
import * as path from "node:path";
import * as vscode from "vscode";
import { formatJson } from "./jsonSource";
import { designerViewType, FORM_SUFFIX } from "./provider";
import { newFormSpec } from "./spec";
import { systemUiFont } from "./systemFont";

/** A VBA identifier usable as a component name (at most 31 characters). */
export function formNameError(name: string): string | undefined {
	if (!/^[A-Za-z][A-Za-z0-9_]*$/.test(name)) return "英字で始まり、英数字と _ だけからなる名前を入力してください。";
	if (name.length > 31) return "31 文字以内で入力してください。";
	if (name.toLowerCase() === "vba") return "この名前は使えません。";
	return undefined;
}

export function findProjectRoot(start: string | undefined): string | undefined {
	for (let dir = start; dir; ) {
		if (fs.existsSync(path.join(dir, "ras.toml"))) return dir;
		const parent = path.dirname(dir);
		if (parent === dir) break;
		dir = parent;
	}
	return undefined;
}

/** The RAS project (folder with ras.toml) a command applies to: the clicked resource's, the active editor's, else a workspace folder. */
export async function projectRoot(resource?: vscode.Uri, placeHolder = "フォームを作成するプロジェクトを選んでください"): Promise<string | undefined> {
	const fromResource = resource?.scheme === "file" ? findProjectRoot(resource.fsPath) : undefined;
	if (fromResource) return fromResource;
	const active = vscode.window.activeTextEditor?.document.uri;
	const fromActive = active?.scheme === "file" ? findProjectRoot(path.dirname(active.fsPath)) : undefined;
	if (fromActive) return fromActive;
	const roots = (vscode.workspace.workspaceFolders ?? []).filter((f) => f.uri.scheme === "file").map((f) => f.uri.fsPath);
	const projects = roots.filter((root) => fs.existsSync(path.join(root, "ras.toml")));
	const candidates = projects.length ? projects : roots;
	if (candidates.length <= 1) return candidates[0];
	return vscode.window.showQuickPick(candidates, { placeHolder });
}

async function newUserForm(resource?: vscode.Uri): Promise<void> {
	const root = await projectRoot(resource);
	if (!root) {
		void vscode.window.showErrorMessage("フォルダーを開いてからユーザーフォームを作成してください。");
		return;
	}
	const formsDir = path.join(root, "forms");
	const name = await vscode.window.showInputBox({
		title: "新しいユーザーフォーム",
		prompt: `${formsDir} に作成するフォームの名前`,
		value: nextName(formsDir),
		validateInput: (value) => formNameError(value.trim()) ?? (fs.existsSync(path.join(formsDir, `${value.trim()}${FORM_SUFFIX}`)) ? "同じ名前のフォームがすでにあります。" : undefined),
	});
	if (!name) return;
	const formName = name.trim();
	const spec = vscode.Uri.file(path.join(formsDir, `${formName}${FORM_SUFFIX}`));
	const code = vscode.Uri.file(path.join(formsDir, `${formName}.ras`));
	await vscode.workspace.fs.createDirectory(vscode.Uri.file(formsDir));
	await vscode.workspace.fs.writeFile(spec, Buffer.from(formatJson(newFormSpec(formName, await systemUiFont())), "utf8"));
	if (!fs.existsSync(code.fsPath)) await vscode.workspace.fs.writeFile(code, Buffer.from(newCodeBehind(formName), "utf8"));
	await vscode.commands.executeCommand("vscode.openWith", spec, designerViewType);
}

/** The code-behind of a new form (the same text as `ras form new`). */
export function newCodeBehind(name: string): string {
	return [
		`// フォーム ${name} のコードビハインド。\`this\` がフォームで、コントロールは \`this.<名前>\` で参照します。`,
		`// イベントは @on("<コントロール名>.<イベント>") を付けた関数で受け取ります（フォーム自身は UserForm）。`,
		"//",
		`// @on("UserForm.Initialize")`,
		"// function initialize() {",
		"// }",
		"",
	].join("\n");
}

function nextName(formsDir: string): string {
	for (let i = 1; ; i++) if (!fs.existsSync(path.join(formsDir, `UserForm${i}${FORM_SUFFIX}`))) return `UserForm${i}`;
}

export function registerUserFormCommands(): vscode.Disposable {
	return vscode.Disposable.from(
		vscode.commands.registerCommand("ras.userForm.new", (resource?: unknown) =>
			newUserForm(resource instanceof vscode.Uri ? resource : undefined).catch((error) => vscode.window.showErrorMessage(`ユーザーフォームを作成できませんでした: ${error instanceof Error ? error.message : String(error)}`)),
		),
		vscode.commands.registerCommand("ras.userForm.openDesigner", async (resource?: vscode.Uri) => {
			const uri = resource instanceof vscode.Uri ? resource : vscode.window.activeTextEditor?.document.uri;
			if (!uri) return;
			await vscode.commands.executeCommand("vscode.openWith", uri, designerViewType);
		}),
		vscode.commands.registerCommand("ras.userForm.openSource", async (resource?: vscode.Uri) => {
			const uri = resource instanceof vscode.Uri ? resource : activeCustomEditorUri();
			if (!uri) return;
			await vscode.commands.executeCommand("vscode.openWith", uri, "default", vscode.ViewColumn.Beside);
		}),
	);
}

function activeCustomEditorUri(): vscode.Uri | undefined {
	const input = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
	return input instanceof vscode.TabInputCustom ? input.uri : undefined;
}
