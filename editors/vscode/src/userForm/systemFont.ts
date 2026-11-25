// The font new forms and controls are made with: the system's UI font
// (System.Drawing.SystemFonts.MenuFont) at 10pt, instead of the MS UI Gothic 9pt
// that MSForms uses when a control has none.

import { execFile } from "node:child_process";
import os from "node:os";
import type { FontValue } from "./protocol";
import { fallbackNewFont, newFontSize } from "./toolbox";

const script =
	"Add-Type -AssemblyName System.Drawing; $f = [System.Drawing.SystemFonts]::MenuFont; " +
	"[Console]::OutputEncoding = [Text.Encoding]::UTF8; Write-Output ($f.Name + \"`t\" + $f.GdiCharSet)";

/** Windows, or WSL (whose extension host can still run the Windows PowerShell). */
const canAskWindows = () => process.platform === "win32" || /microsoft/i.test(os.release());

function ask(): Promise<FontValue> {
	if (!canAskWindows()) return Promise.resolve({ ...fallbackNewFont });
	return new Promise((resolve) => {
		execFile(
			"powershell.exe",
			["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script],
			{ timeout: 15000, windowsHide: true, encoding: "utf8" },
			(error, stdout) => {
				const [name, charset] = (error ? "" : stdout).trim().split("\t");
				if (!name) return resolve({ ...fallbackNewFont });
				const font: FontValue = { name, size: newFontSize };
				const set = Number(charset);
				if (Number.isInteger(set) && set >= 0 && set <= 255) font.charset = set;
				resolve(font);
			},
		);
	});
}

let pending: Promise<FontValue> | undefined;
let known: FontValue | undefined;

/** The system UI font at 10pt, looked up once (PowerShell takes a moment to start). */
export function systemUiFont(): Promise<FontValue> {
	pending ??= ask().then((font) => (known = font));
	return pending;
}

/** The font if it has been looked up already. */
export function knownSystemUiFont(): FontValue | undefined {
	return known;
}
