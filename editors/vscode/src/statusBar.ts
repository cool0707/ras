import * as vscode from "vscode";
import type { ServerState } from "./server";

/** Params of the `ras/syncStatus` notification (spec §7.4: 接続先ブック・未同期件数・ドリフト警告). */
interface SyncStatus {
	connected: boolean;
	workbook: string | null;
	pending: number;
	drift: number;
	message?: string;
}

/** Status bar item that mirrors what the server reports; it computes nothing itself. */
export class SyncStatusBar implements vscode.Disposable {
	private readonly item = vscode.window.createStatusBarItem("ras.syncStatus", vscode.StatusBarAlignment.Left, 0);
	/** "Excel で開く", right of the status while the workbook is not open (live sync needs Excel, so Windows only). */
	private readonly openItem = vscode.window.createStatusBarItem("ras.openInExcel", vscode.StatusBarAlignment.Left, -0.001);
	private serverState: ServerState = "stopped";
	private status: SyncStatus | undefined;
	/** Whether the workspace has a project (or a .ras file): otherwise the server is not started. */
	private serverWanted = false;
	/** No project yet, in a folder to start one in: the item offers "New RAS project". */
	private canCreateProject = false;

	constructor() {
		this.item.name = "RAS 同期状態";
		this.openItem.name = "RAS: Excel で開く";
		this.openItem.text = "$(link-external) Excel で開く";
		this.openItem.command = "ras.client.openInExcel";
		this.render();
	}

	setServerWanted(wanted: boolean): void {
		this.serverWanted = wanted;
		this.render();
	}

	setCanCreateProject(can: boolean): void {
		this.canCreateProject = can;
		this.render();
	}

	setServerState(state: ServerState): void {
		this.serverState = state;
		if (state !== "running") {
			// A status reported by a previous server process is no longer true.
			this.status = undefined;
		}
		this.render();
	}

	/** @param params Raw params of a `ras/syncStatus` notification. */
	setSyncStatus(params: unknown): void {
		this.status = parseSyncStatus(params);
		this.render();
	}

	dispose(): void {
		this.item.dispose();
		this.openItem.dispose();
	}

	private render(): void {
		const item = this.item;
		item.backgroundColor = undefined;
		if (!this.serverWanted && this.serverState === "stopped") {
			this.openItem.hide();
			if (this.canCreateProject) {
				item.text = "$(new-folder) RAS プロジェクトを作成";
				item.tooltip = "このフォルダーに、Excel ブックと RAS のソースを入れたプロジェクトを作成します。";
				item.command = "ras.project.new";
				item.show();
			} else {
				item.hide();
			}
			return;
		}
		item.show();
		const showOpen = process.platform === "win32" && this.serverState === "running" && !this.status?.connected;
		if (showOpen) {
			this.openItem.tooltip = this.status?.workbook ? `${this.status.workbook} を Excel で開いて同期します` : "対象のブックを Excel で開いて同期します";
			this.openItem.show();
		} else {
			this.openItem.hide();
		}

		if (this.serverState !== "running") {
			const starting = this.serverState === "starting";
			item.text = starting ? "$(loading~spin) RAS: 起動中" : "$(circle-slash) RAS: 停止中";
			item.tooltip = starting ? "RAS 言語サーバーを起動しています…" : "RAS 言語サーバーは起動していません。クリックすると再起動します。";
			item.command = starting ? undefined : "ras.restartServer";
			return;
		}

		const status = this.status;
		const tooltip = new vscode.MarkdownString();
		let text: string;
		if (status?.connected) {
			text = `$(sync) RAS: ${status.workbook ? baseName(status.workbook) : "接続済み"}`;
			tooltip.appendMarkdown("**RAS 同期: 接続済み**");
		} else {
			text = "$(debug-disconnect) RAS: 未接続";
			tooltip.appendMarkdown("**RAS 同期: 未接続**");
		}
		if (status?.workbook) {
			tooltip.appendMarkdown("\n\nブック: ").appendText(status.workbook);
		}
		if (status && status.pending > 0) {
			text += ` $(arrow-up)${status.pending}`;
			tooltip.appendMarkdown(`\n\n未同期: ${status.pending} 件`);
		}
		if (status && status.drift > 0) {
			text += ` $(warning)${status.drift}`;
			item.backgroundColor = new vscode.ThemeColor("statusBarItem.warningBackground");
			tooltip.appendMarkdown(`\n\n**ドリフト警告: ${status.drift} 件** — Excel（VBE）側で編集された RAS 管理モジュールがあります。`);
		}
		if (status?.message) {
			tooltip.appendMarkdown("\n\n").appendText(status.message);
		}
		tooltip.appendMarkdown(status?.connected ? "\n\nクリックすると今すぐ同期します。" : "\n\nクリックすると、ブックファイルへ直接書き込みます（オフライン。ブックを閉じている必要があります）。");

		item.text = text;
		item.tooltip = tooltip;
		item.command = "ras.client.syncNow";
	}
}

function parseSyncStatus(params: unknown): SyncStatus {
	const raw = (typeof params === "object" && params !== null ? params : {}) as Record<string, unknown>;
	const count = (value: unknown) => (typeof value === "number" && Number.isFinite(value) && value > 0 ? Math.floor(value) : 0);
	return {
		connected: raw.connected === true,
		workbook: typeof raw.workbook === "string" && raw.workbook !== "" ? raw.workbook : null,
		pending: count(raw.pending),
		drift: count(raw.drift),
		message: typeof raw.message === "string" && raw.message !== "" ? raw.message : undefined,
	};
}

/** File name of a workbook path; `ras.toml` uses `/`, Excel reports `\`. */
function baseName(workbook: string): string {
	return workbook.split(/[\\/]/).pop() || workbook;
}
