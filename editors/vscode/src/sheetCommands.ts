// "RAS: シートを追加 / 複製 / 削除 / 名前を変更": the sheets of the project's workbook and their
// files in `objects/` change together (spec §14.16). The work is the server's (`ras.sheet`
// through `workspace/executeCommand`): it changes the workbook open in Excel, or the workbook
// file when it is closed.

import * as path from "node:path";
import * as vscode from "vscode";
import { errorMessage } from "./server";
import type { RasServer } from "./server";
import { sheetOfRasFile } from "./sheetFile";
import { projectRoot } from "./userForm/commands";

const SHEET_COMMAND = "ras.sheet";

interface SheetRow {
	name: string;
	codeName: string;
	worksheet: boolean;
	visible: boolean;
	file: string | null;
}

export interface SheetResult {
	ok: boolean;
	message?: string;
	refused?: string[];
	action?: string;
	name?: string;
	codeName?: string;
	live?: boolean;
	created?: string[];
	removed?: string[];
	renamed?: { from: string; to: string }[];
	references?: number;
	formulas?: number;
	warnings?: string[];
	sheets?: SheetRow[];
}

function asResult(value: unknown): SheetResult {
	return typeof value === "object" && value !== null && typeof (value as { ok?: unknown }).ok === "boolean" ? (value as SheetResult) : { ok: false, message: "言語サーバーから想定外の応答が返されました。" };
}

/**
 * What the tabs of the sheet grid give (its menu, its dialog of a new sheet, a double click on a
 * tab): the sheet, and the names already chosen there — no more is asked.
 */
interface FromGrid {
	sheet: string;
	fromGrid: true;
	name?: string;
	codeName?: string;
}

function fromGrid(arg: unknown): FromGrid | undefined {
	return typeof arg === "object" && arg !== null && (arg as { fromGrid?: unknown }).fromGrid === true ? (arg as FromGrid) : undefined;
}

/** What a menu was invoked on: a file of `objects/`, or a tab of the sheet grid (`{ sheet }`). */
function sheetOfArgument(arg: unknown): string | undefined {
	if (arg instanceof vscode.Uri) {
		return sheetOfRasFile(arg.fsPath);
	}
	if (typeof arg === "object" && arg !== null && typeof (arg as { sheet?: unknown }).sheet === "string") {
		return (arg as { sheet: string }).sheet;
	}
	if (arg === undefined) {
		const active = vscode.window.activeTextEditor?.document.uri;
		return active === undefined ? undefined : sheetOfRasFile(active.fsPath);
	}
	return undefined;
}

async function execute(server: RasServer, title: string, args: Record<string, unknown>): Promise<SheetResult> {
	return vscode.window.withProgress({ location: vscode.ProgressLocation.Window, title }, async () => asResult(await server.execute(SHEET_COMMAND, [args])));
}

async function listSheets(server: RasServer): Promise<SheetRow[] | undefined> {
	const result = await execute(server, "RAS: シートを読み込んでいます…", { action: "list" });
	if (!result.ok) {
		void vscode.window.showErrorMessage(`RAS: ${result.message ?? "シートの一覧を取得できませんでした。"}`);
		return undefined;
	}
	return (result.sheets ?? []).filter((s) => s.worksheet);
}

/** The sheet a command is about: the one the menu was invoked on, else one the user picks. */
async function pickSheet(server: RasServer, arg: unknown, placeHolder: string): Promise<{ sheet: string; row?: SheetRow } | undefined> {
	const sheets = await listSheets(server);
	if (sheets === undefined) {
		return undefined;
	}
	const wanted = sheetOfArgument(arg);
	if (wanted !== undefined) {
		const lower = wanted.toLowerCase();
		return { sheet: wanted, row: sheets.find((s) => s.codeName.toLowerCase() === lower) ?? sheets.find((s) => s.name.toLowerCase() === lower) };
	}
	const picked = await vscode.window.showQuickPick(
		sheets.map((s) => ({ label: s.name, description: `コード名 ${s.codeName}`, detail: s.file ?? undefined, row: s })),
		{ title: "RAS: シート", placeHolder },
	);
	return picked === undefined ? undefined : { sheet: picked.row.codeName, row: picked.row };
}

const NAME_RULE = /[:\\/?*[\]]/;

function validateName(value: string, required: boolean): string | undefined {
	if (value === "") {
		return required ? "見出しの名前を入力してください。" : undefined;
	}
	if ([...value].length > 31 || NAME_RULE.test(value) || value.startsWith("'") || value.endsWith("'")) {
		return "31 文字までで、: \\ / ? * [ ] は使えません（前後を ' にもできません）。";
	}
	return undefined;
}

function validateCodeName(value: string): string | undefined {
	if (value === "") {
		return undefined;
	}
	return /^\p{L}[\p{L}\p{N}_]*$/u.test(value) && [...value].length <= 31 ? undefined : "文字で始まり、文字・数字・_ だけの 31 文字までの名前にしてください。";
}

/** Tells what was done, opens the file of a new sheet and shows the warnings. */
async function report(root: string | undefined, result: SheetResult, done: string): Promise<void> {
	if (!result.ok) {
		void vscode.window.showErrorMessage(`RAS: ${result.message ?? "シートを変更できませんでした。"}`);
		return;
	}
	const details: string[] = [];
	if ((result.references ?? 0) > 0) {
		details.push(`ソースの参照を ${result.references} 箇所書き換えました`);
	}
	if ((result.formulas ?? 0) > 0) {
		details.push(`ブックの数式・名前の定義を ${result.formulas} 箇所書き換えました`);
	}
	void vscode.window.showInformationMessage(`RAS: ${done}${details.length > 0 ? `（${details.join("、")}）` : ""}`);
	for (const warning of result.warnings ?? []) {
		void vscode.window.showWarningMessage(`RAS: ${warning}`);
	}
	const created = result.created?.[0] ?? result.renamed?.[0]?.to;
	if (root !== undefined && created !== undefined && (result.action === "add" || result.action === "copy")) {
		await vscode.window.showTextDocument(vscode.Uri.file(path.join(root, created)), { preview: false });
	}
	// The sheet grid, when it is open, shows the workbook file again.
	if (result.live === false && result.action !== "remove" && root !== undefined && result.codeName !== undefined) {
		await vscode.commands.executeCommand("ras.sheet.refresh", vscode.Uri.file(path.join(root, "objects", `${result.codeName}.ras`)));
	}
}

/** The project a command is about: that of the file it was invoked on, else of the active editor. */
function rootOf(arg: unknown): Promise<string | undefined> {
	return projectRoot(arg instanceof vscode.Uri ? arg : undefined, "プロジェクトを選んでください");
}

async function addSheet(server: RasServer, arg: unknown): Promise<SheetResult | undefined> {
	const grid = fromGrid(arg);
	const name = grid?.name ?? (await vscode.window.showInputBox({ title: "RAS: シートを追加", prompt: "見出しの名前", validateInput: (v) => validateName(v.trim(), true) }));
	if (name === undefined) {
		return undefined;
	}
	const codeName = grid !== undefined ? (grid.codeName ?? "") : await vscode.window.showInputBox({ title: "RAS: シートを追加", prompt: "コード名（空欄なら Sheet<番号> を自動で付けます）", validateInput: (v) => validateCodeName(v.trim()) });
	if (codeName === undefined) {
		return undefined;
	}
	// On a sheet's file (or tab): right after that sheet. Else at the end.
	const after = sheetOfArgument(arg);
	const result = await execute(server, "RAS: シートを追加しています…", { action: "add", name: name.trim(), codeName: codeName.trim(), after });
	await report(await rootOf(arg), result, `シート「${result.name}」（コード名 ${result.codeName}）を追加しました。`);
	return result;
}

async function copySheet(server: RasServer, arg: unknown): Promise<SheetResult | undefined> {
	const target = await pickSheet(server, arg, "複製するシートを選んでください");
	if (target === undefined) {
		return undefined;
	}
	// From the grid's tab ("コピーを作成"): named as Excel names a copy.
	const grid = fromGrid(arg);
	const name = grid !== undefined ? "" : await vscode.window.showInputBox({ title: "RAS: シートを複製", prompt: "複製先の見出しの名前（空欄なら「<元の名前> (2)」）", validateInput: (v) => validateName(v.trim(), false) });
	if (name === undefined) {
		return undefined;
	}
	const codeName = grid !== undefined ? "" : await vscode.window.showInputBox({ title: "RAS: シートを複製", prompt: "複製先のコード名（空欄なら Sheet<番号> を自動で付けます）", validateInput: (v) => validateCodeName(v.trim()) });
	if (codeName === undefined) {
		return undefined;
	}
	const result = await execute(server, "RAS: シートを複製しています…", { action: "copy", sheet: target.sheet, name: name.trim(), codeName: codeName.trim() });
	await report(await rootOf(arg), result, `シート「${result.name}」（コード名 ${result.codeName}）を複製して作りました。`);
	return result;
}

async function removeSheet(server: RasServer, arg: unknown, sheetName?: string): Promise<SheetResult | undefined> {
	const target = sheetName !== undefined ? { sheet: sheetName, row: undefined } : await pickSheet(server, arg, "削除するシートを選んでください");
	if (target === undefined) {
		return undefined;
	}
	const shown = target.row !== undefined ? `「${target.row.name}」（コード名 ${target.row.codeName}）` : `「${target.sheet}」`;
	const remove = "削除する";
	if (sheetName === undefined) {
		const choice = await vscode.window.showWarningMessage(`シート${shown}を、ブックから削除します。objects/ のファイルも削除します。`, { modal: true }, remove);
		if (choice !== remove) {
			return undefined;
		}
	}
	let result = await execute(server, "RAS: シートを削除しています…", { action: "remove", sheet: target.sheet });
	if (!result.ok && result.refused !== undefined) {
		// Something refers to the sheet, or its file has code: the user decides (spec §14.16).
		const choice = await vscode.window.showWarningMessage(`シート${shown}を削除すると、次のものが失われます。`, { modal: true, detail: result.refused.map((r) => `・${r}`).join("\n") }, remove);
		if (choice !== remove) {
			return undefined;
		}
		result = await execute(server, "RAS: シートを削除しています…", { action: "remove", sheet: target.sheet, force: true });
	}
	await report(await rootOf(arg), result, `シート${shown}を削除しました。`);
	return result;
}

async function renameSheet(server: RasServer, arg: unknown): Promise<SheetResult | undefined> {
	const target = await pickSheet(server, arg, "名前を変えるシートを選んでください");
	if (target === undefined) {
		return undefined;
	}
	// From the grid's tab (its name edited there): the tab's name only.
	const grid = fromGrid(arg);
	const name = grid?.name ?? (await vscode.window.showInputBox({ title: "RAS: シートの名前を変更", prompt: "見出しの名前", value: target.row?.name ?? "", validateInput: (v) => validateName(v.trim(), true) }));
	if (name === undefined) {
		return undefined;
	}
	const codeName = grid !== undefined ? (target.row?.codeName ?? "") : await vscode.window.showInputBox({
		title: "RAS: シートの名前を変更",
		prompt: "コード名（変えると、objects/ のファイル名とソースの参照も変えます）",
		value: target.row?.codeName ?? "",
		validateInput: (v) => (v.trim() === "" ? "コード名を入力してください。" : validateCodeName(v.trim())),
	});
	if (codeName === undefined) {
		return undefined;
	}
	const newName = name.trim() !== (target.row?.name ?? "") ? name.trim() : undefined;
	const newCodeName = codeName.trim() !== (target.row?.codeName ?? "") ? codeName.trim() : undefined;
	if (newName === undefined && newCodeName === undefined) {
		return undefined;
	}
	// The files are saved first: the references in the sources are rewritten on disk.
	await vscode.workspace.saveAll(false);
	const result = await execute(server, "RAS: シートの名前を変えています…", { action: "rename", sheet: target.sheet, name: newName, codeName: newCodeName });
	await report(await rootOf(arg), result, `シートの名前を「${result.name}」（コード名 ${result.codeName}）に変えました。`);
	return result;
}

/** A file of `objects/` was renamed in the Explorer: the sheet's code name can follow. */
async function onRenamed(server: RasServer, event: vscode.FileRenameEvent): Promise<void> {
	for (const { oldUri, newUri } of event.files) {
		const [from, to] = [sheetOfRasFile(oldUri.fsPath), sheetOfRasFile(newUri.fsPath)];
		if (from === undefined || to === undefined || from.toLowerCase() === to.toLowerCase() || validateCodeName(to) !== undefined) {
			continue;
		}
		const rename = "コード名も変える";
		const choice = await vscode.window.showInformationMessage(`objects/${from}.ras の名前を ${to}.ras に変えました。ブックのシートのコード名も ${from} から ${to} に変えますか？（ソースの参照も書き換えます）`, rename, "変えない");
		if (choice !== rename) {
			continue;
		}
		try {
			// The command renames the file itself, with the references: back to the old name first.
			await vscode.workspace.saveAll(false);
			await vscode.workspace.fs.rename(newUri, oldUri, { overwrite: false });
			const result = await execute(server, "RAS: シートのコード名を変えています…", { action: "rename", sheet: from, codeName: to });
			await report(await rootOf(oldUri), result, `シートのコード名を ${from} から ${to} に変えました。`);
		} catch (error) {
			void vscode.window.showErrorMessage(`RAS: シートのコード名を変えられませんでした: ${errorMessage(error)}`);
		}
	}
}

/** A file of `objects/` was deleted in the Explorer: the sheet can go with it. */
async function onDeleted(server: RasServer, event: vscode.FileDeleteEvent): Promise<void> {
	for (const uri of event.files) {
		const sheet = sheetOfRasFile(uri.fsPath);
		if (sheet === undefined) {
			continue;
		}
		const remove = "シートも削除する";
		const choice = await vscode.window.showInformationMessage(`objects/${sheet}.ras を削除しました。ブックのシート ${sheet} も削除しますか？（削除しなければ、次の同期でファイルを作り直します）`, remove, "削除しない");
		if (choice === remove) {
			try {
				await removeSheet(server, uri, sheet);
			} catch (error) {
				void vscode.window.showErrorMessage(`RAS: シートを削除できませんでした: ${errorMessage(error)}`);
			}
		}
	}
}

export function registerSheetCommands(server: RasServer): vscode.Disposable {
	// (What was done is the command's result: the sheet grid draws the page after it.)
	const command = (id: string, title: string, run: (arg: unknown) => Promise<SheetResult | undefined>) =>
		vscode.commands.registerCommand(id, async (arg?: unknown): Promise<SheetResult | undefined> => {
			try {
				return await run(arg);
			} catch (error) {
				void vscode.window.showErrorMessage(`RAS「${title}」に失敗しました: ${errorMessage(error)}`);
				return undefined;
			}
		});
	return vscode.Disposable.from(
		command("ras.sheet.add", "シートを追加", (arg) => addSheet(server, arg)),
		command("ras.sheet.copy", "シートを複製", (arg) => copySheet(server, arg)),
		command("ras.sheet.remove", "シートを削除", (arg) => removeSheet(server, arg)),
		command("ras.sheet.rename", "シートの名前を変更", (arg) => renameSheet(server, arg)),
		vscode.workspace.onDidRenameFiles((event) => void onRenamed(server, event)),
		vscode.workspace.onDidDeleteFiles((event) => void onDeleted(server, event)),
	);
}
