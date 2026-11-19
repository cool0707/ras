// Runs the UserForm Designer unit tests (test/userForm/*.test.ts).
// They import the extension's TypeScript modules, so each test file is bundled
// with esbuild first and the bundle is run with Node's test runner.

import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import * as esbuild from "esbuild";

const dir = "test/userForm";
const tests = fs.readdirSync(dir).filter((name) => name.endsWith(".test.ts"));
const out = fs.mkdtempSync(path.join(os.tmpdir(), "ras-userform-test-"));
try {
	await esbuild.build({
		entryPoints: tests.map((name) => path.join(dir, name)),
		outdir: out,
		outExtension: { ".js": ".mjs" },
		bundle: true,
		format: "esm",
		platform: "node",
		target: "node20",
		logLevel: "warning",
	});
	const files = tests.map((name) => path.join(out, name.replace(/\.ts$/, ".mjs")));
	const result = spawnSync(process.execPath, ["--test", "--test-reporter=spec", ...files], { stdio: "inherit" });
	process.exitCode = result.status ?? 1;
} finally {
	fs.rmSync(out, { recursive: true, force: true });
}
