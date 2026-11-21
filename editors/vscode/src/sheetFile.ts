// Which worksheet a RAS file is about, for showing that sheet's grid (`ras book view --sheet`).

/**
 * The code name of the worksheet the RAS file `fileName` (base name) with the source `text` is
 * about, or `undefined`:
 *
 * - `Sheet1Events.ras` (`Sheet1Events2.ras`): the file sync creates or takes over for the sheet
 *   `Sheet1` (spec §14.7.1, §14.7.2). `ThisWorkbookEvents.ras` is the workbook's, not a sheet's.
 * - otherwise the sheet of the first event handler of the file (`@on("Sheet1.Change")`).
 */
export function sheetOfRasFile(fileName: string, text: string): string | undefined {
	const named = /^([A-Za-z_\u0080-￿][\w\u0080-￿]*?)Events\d*\.ras$/i.exec(fileName);
	if (named !== null && !/^(ThisWorkbook|Workbook)$/i.test(named[1])) {
		return named[1];
	}
	const handler = /@on\(\s*["']([A-Za-z_\u0080-￿][\w\u0080-￿]*)\./g;
	for (let m = handler.exec(text); m !== null; m = handler.exec(text)) {
		if (!/^(Workbook|Application|App)$/i.test(m[1])) {
			return m[1];
		}
	}
	return undefined;
}
