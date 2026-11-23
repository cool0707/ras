// Builds the playground into web/dist/ (what GitHub Pages serves).
//
//   node scripts/build.mjs            build (web/pkg/ must exist: node scripts/build-wasm.mjs)
//   node scripts/build.mjs --serve    build, then serve web/dist/ on http://localhost:8000
//
// The samples are the projects of web/samples/ and examples/: their files are put in the bundle.
import { build } from "esbuild";
import { copyFileSync, existsSync, mkdirSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from "node:fs";
import { createServer } from "node:http";
import { dirname, extname, join, relative } from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);

const web = dirname(dirname(fileURLToPath(import.meta.url)));
const root = dirname(web);
const dist = join(web, "dist");
const production = !process.argv.includes("--serve");

if (!existsSync(join(web, "pkg", "ras_wasm_bg.wasm"))) {
    console.error("web/pkg/ がありません。先に node scripts/build-wasm.mjs を実行してください");
    process.exit(1);
}

/** A project's files: ras.toml, the .ras files of src/, types/, objects/, test/ and the forms of forms/. */
function project(dir) {
    const files = [];
    const walk = (d) => {
        for (const n of readdirSync(d).sort()) {
            const p = join(d, n);
            if (statSync(p).isDirectory()) walk(p);
            else if (n.endsWith(".ras") || n.endsWith(".form.json")) files.push({ path: relative(dir, p).split("\\").join("/"), text: readFileSync(p, "utf8") });
        }
    };
    for (const sub of ["src", "types", "objects", "forms", "test"]) if (existsSync(join(dir, sub))) walk(join(dir, sub));
    let toml = readFileSync(join(dir, "ras.toml"), "utf8");
    // The playground's workbook is Book.xlsx (a new one, or the one read into the page).
    if (!/^\s*workbook\s*=/m.test(toml)) toml = toml.replace(/^\[project\][^\n]*\n/m, (h) => `${h}workbook = "Book.xlsx"\n`);
    files.unshift({ path: "ras.toml", text: toml });
    return files;
}

const samples = [
    { id: "hello", title: "はじめての RAS（表と数式・テスト）", files: project(join(web, "samples", "hello")) },
    { id: "dialogs", title: "ダイアログと実行時エラー", files: project(join(web, "samples", "dialogs")) },
    { id: "forms", title: "ユーザーフォームとシートのコントロール", files: project(join(web, "samples", "forms")) },
    { id: "sales-report", title: "examples/sales-report", files: project(join(root, "examples", "sales-report")) },
];

const samplesPlugin = {
    name: "samples",
    setup(b) {
        b.onResolve({ filter: /^virtual:samples$/ }, () => ({ path: "samples", namespace: "samples" }));
        b.onLoad({ filter: /.*/, namespace: "samples" }, () => ({ contents: `export default ${JSON.stringify(samples)};`, loader: "js" }));
    },
};

rmSync(dist, { recursive: true, force: true });
mkdirSync(dist, { recursive: true });
const common = {
    bundle: true,
    format: "esm",
    target: "es2022",
    minify: production,
    sourcemap: production ? false : "inline",
    logLevel: "warning",
    define: { "process.env.NODE_ENV": JSON.stringify(production ? "production" : "development") },
};
// The page uses modules of the VS Code extension (the Designer's engine, the sheets' events):
// the few Node modules they import are the browser's.
const browserModules = {
    name: "browser-modules",
    setup(b) {
        b.onResolve({ filter: /^(node:)?crypto$/ }, () => ({ path: "crypto", namespace: "browser-modules" }));
        b.onLoad({ filter: /^crypto$/, namespace: "browser-modules" }, () => ({ contents: "export const randomUUID = () => crypto.randomUUID();", loader: "js" }));
        b.onResolve({ filter: /^(node:)?path$/ }, () => ({ path: require.resolve("path-browserify") }));
    },
};
await build({ ...common, entryPoints: [join(web, "src", "main.ts")], outfile: join(dist, "app.js"), plugins: [samplesPlugin, browserModules], nodePaths: [join(web, "node_modules")] });
// The VS Code extension's UserForm Designer and form panel (its webviews), for the frames of
// designer.html and runner.html. Their packages are the playground's (package.json).
const vscode = join(root, "editors", "vscode");
await build({
    ...common,
    format: "iife",
    entryPoints: { userFormDesigner: join(vscode, "webview", "userFormDesigner", "App.tsx"), formRunner: join(vscode, "webview", "formRunner", "main.ts") },
    outdir: join(dist, "designer"),
    jsx: "automatic",
    jsxImportSource: "preact",
    loader: { ".ttf": "file" },
    assetNames: "[name]",
    nodePaths: [join(web, "node_modules")],
    tsconfig: join(vscode, "tsconfig.webview.json"),
});
await build({ ...common, entryPoints: [join(web, "src", "worker.ts")], outfile: join(dist, "worker.js") });
copyFileSync(join(web, "pkg", "ras_wasm_bg.wasm"), join(dist, "ras_wasm_bg.wasm"));
for (const f of ["index.html", "style.css", "sw.js", "designer.html", "runner.html", "frame.css"]) copyFileSync(join(web, "static", f), join(dist, f));
// The icon is the VS Code extension's (also embedded in ras.exe: crates/ras/assets/ras.ico).
copyFileSync(join(root, "editors", "vscode", "images", "icon.png"), join(dist, "icon.png"));
writeFileSync(join(dist, ".nojekyll"), "");
console.log(`web/dist/ を作成しました（${readdirSync(dist).join(", ")}）`);

if (!production) {
    const types = { ".png": "image/png", ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".wasm": "application/wasm", ".json": "application/json" };
    createServer((req, res) => {
        const url = new URL(req.url, "http://localhost");
        let file = join(dist, decodeURIComponent(url.pathname));
        if (existsSync(file) && statSync(file).isDirectory()) file = join(file, "index.html");
        if (!file.startsWith(dist) || !existsSync(file)) {
            res.writeHead(404).end("not found");
            return;
        }
        res.writeHead(200, { "Content-Type": types[extname(file)] || "application/octet-stream" });
        res.end(readFileSync(file));
    }).listen(8000, () => console.log("http://localhost:8000/"));
}
