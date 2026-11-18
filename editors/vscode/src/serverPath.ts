import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";

/**
 * Locates the `ras` executable (spec §7.4: `ras.path`, or PATH when it is unset).
 *
 * @param configured Value of the `ras.path` setting (already trimmed). Empty means "search PATH".
 * @param baseDir Directory that a relative `ras.path` is resolved against.
 * @returns Absolute path of the executable, or `undefined` when none was found.
 */
export function findServerExecutable(configured: string, baseDir: string | undefined): string | undefined {
	const isWindows = process.platform === "win32";
	const withExe = (name: string) => (isWindows && path.extname(name) === "" ? `${name}.exe` : name);

	if (configured === "") {
		return findOnPath(withExe("ras"));
	}

	let candidate = configured;
	if (candidate === "~" || candidate.startsWith("~/") || candidate.startsWith("~\\")) {
		candidate = path.join(os.homedir(), candidate.slice(1));
	}
	if (!path.isAbsolute(candidate) && !/[\\/]/.test(candidate)) {
		// A bare command name such as "ras-nightly" is looked up on PATH.
		return findOnPath(withExe(candidate));
	}
	if (!path.isAbsolute(candidate)) {
		candidate = path.resolve(baseDir ?? process.cwd(), candidate);
	}
	return [candidate, withExe(candidate)].find(isExecutableFile);
}

function findOnPath(fileName: string): string | undefined {
	for (const directory of (process.env.PATH ?? "").split(path.delimiter)) {
		if (directory === "") {
			continue;
		}
		const candidate = path.join(directory, fileName);
		if (isExecutableFile(candidate)) {
			return candidate;
		}
	}
	return undefined;
}

function isExecutableFile(file: string): boolean {
	try {
		if (!fs.statSync(file).isFile()) {
			return false;
		}
		if (process.platform !== "win32") {
			fs.accessSync(file, fs.constants.X_OK);
		}
		return true;
	} catch {
		return false;
	}
}
