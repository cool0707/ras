// Builds crates/ras-wasm for the browser into web/pkg/ (wasm-bindgen --target web).
//
//   node scripts/build-wasm.mjs [--dev]
//
// Needs the wasm32-unknown-unknown target (`rustup target add wasm32-unknown-unknown`) and
// wasm-bindgen-cli of the version crates/ras-wasm uses (`cargo install wasm-bindgen-cli
// --version <that version> --locked`). wasm-opt (binaryen) makes the file smaller when it is on PATH.
import { execFileSync, spawnSync } from "node:child_process";
import { existsSync, mkdirSync, statSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const web = dirname(dirname(fileURLToPath(import.meta.url)));
const root = dirname(web);
const dev = process.argv.includes("--dev");
const profile = dev ? "debug" : "release";
const run = (cmd, args, opts = {}) => execFileSync(cmd, args, { stdio: "inherit", cwd: root, ...opts });

run("cargo", ["build", "-p", "ras-wasm", "--target", "wasm32-unknown-unknown", "--locked", ...(dev ? [] : ["--release"])]);
const wasm = join(root, "target", "wasm32-unknown-unknown", profile, "ras_wasm.wasm");
const out = join(web, "pkg");
mkdirSync(out, { recursive: true });
run("wasm-bindgen", ["--target", "web", "--out-dir", out, wasm]);
const built = join(out, "ras_wasm_bg.wasm");
if (!dev && spawnSync("wasm-opt", ["--version"]).status === 0) {
    run("wasm-opt", ["-Oz", "--enable-bulk-memory", "--enable-nontrapping-float-to-int", "--enable-sign-ext", "--enable-mutable-globals", "-o", built, built]);
}
if (existsSync(built)) console.log(`${built}: ${(statSync(built).size / 1048576).toFixed(1)} MB`);
