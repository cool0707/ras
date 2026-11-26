// Names and text of a new standard module, shared by the extension (newModule.ts) and the
// playground (web/src/main.ts). No VS Code API here.

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

/** The text of a new standard module. */
export function newModuleSource(vbaName: string): string {
	return [
		`// 標準モジュール ${vbaName}。\`export\` した関数はほかのモジュールから呼び出せます。`,
		"// 引数なし・戻り値なしで `export` した関数は、Excel のマクロとして実行できます。",
		"",
		"",
	].join("\n");
}
