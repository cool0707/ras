// Which worksheet a RAS file is the module of, for showing that sheet's grid (`ras book view --sheet`).

/**
 * The code name of the worksheet the RAS file at `filePath` is the module of, or `undefined`:
 * a file directly in `objects/` is the module of the sheet its name is the code name of
 * (`objects/Sheet1.ras`, spec §12.2). `objects/ThisWorkbook.ras` is the workbook's, not a sheet's.
 */
export function sheetOfRasFile(filePath: string): string | undefined {
	const parts = filePath.replace(/\\/g, "/").split("/");
	const fileName = parts[parts.length - 1] ?? "";
	if (parts.length < 2 || parts[parts.length - 2].toLowerCase() !== "objects") {
		return undefined;
	}
	const named = /^(.+)\.ras$/i.exec(fileName);
	if (named === null || /\.d$/i.test(named[1]) || /^ThisWorkbook$/i.test(named[1])) {
		return undefined;
	}
	return named[1];
}
