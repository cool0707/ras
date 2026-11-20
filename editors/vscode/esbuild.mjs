// Bundles src/extension.ts and its npm dependencies into dist/extension.js, and the
// UserForm Designer webview (webview/userFormDesigner) into dist/webview/.
//
//   node esbuild.mjs               development build (source map, not minified)
//   node esbuild.mjs --production  minified build used by `vsce package`
//   node esbuild.mjs --watch       rebuild on change

import fs from "node:fs";
import * as esbuild from "esbuild";

const production = process.argv.includes("--production");
const watch = process.argv.includes("--watch");

/**
 * The bundle redistributes the npm packages it contains, so each of them must be
 * listed in THIRD_PARTY_NOTICES.md. Fails the build when one is missing.
 */
const checkNotices = {
	name: "check-third-party-notices",
	setup(build) {
		build.onEnd((result) => {
			if (!result.metafile) {
				return undefined;
			}
			const notices = fs.readFileSync("THIRD_PARTY_NOTICES.md", "utf8");
			const marker = "node_modules/";
			const bundled = new Set();
			for (const input of Object.keys(result.metafile.inputs)) {
				const index = input.lastIndexOf(marker);
				if (index >= 0) {
					const [scope, name] = input.slice(index + marker.length).split("/");
					bundled.add(scope.startsWith("@") ? `${scope}/${name}` : scope);
				}
			}
			const missing = [...bundled].filter((name) => !notices.includes(`### ${name}\n`)).sort();
			if (missing.length === 0) {
				return undefined;
			}
			return {
				errors: [{ text: `THIRD_PARTY_NOTICES.md has no "### <package>" section for bundled package(s): ${missing.join(", ")}` }],
			};
		});
	},
};

const context = await esbuild.context({
	entryPoints: ["src/extension.ts"],
	outfile: "dist/extension.js",
	bundle: true,
	format: "cjs",
	platform: "node",
	// VS Code 1.90 (the `engines.vscode` baseline) runs extensions on Node.js 20.
	target: "node20",
	external: ["vscode"],
	minify: production,
	sourcemap: !production,
	sourcesContent: false,
	metafile: true,
	logLevel: "info",
	plugins: [checkNotices],
});

// The UserForm Designer and the panel of a form being debugged (formRunner) run in
// webviews: browser bundles with their own CSS (and the codicon font), loaded under a
// strict content security policy.
const webview = await esbuild.context({
	entryPoints: { userFormDesigner: "webview/userFormDesigner/App.tsx", formRunner: "webview/formRunner/main.ts" },
	outdir: "dist/webview",
	bundle: true,
	format: "iife",
	platform: "browser",
	target: "es2022",
	jsx: "automatic",
	jsxImportSource: "preact",
	loader: { ".ttf": "file" },
	assetNames: "[name]",
	minify: production,
	sourcemap: !production,
	sourcesContent: false,
	metafile: true,
	logLevel: "info",
	plugins: [checkNotices],
});

if (watch) {
	await Promise.all([context.watch(), webview.watch()]);
} else {
	try {
		await Promise.all([context.rebuild(), webview.rebuild()]);
	} catch {
		// esbuild has already printed the errors.
		process.exitCode = 1;
	} finally {
		await Promise.all([context.dispose(), webview.dispose()]);
	}
}
