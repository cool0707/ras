// npm test: bundles test/*.test.ts for Node and runs them with node --test (web/pkg must exist).
import { build } from "esbuild";
import { spawnSync } from "node:child_process";
import { mkdtempSync, readdirSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const web = dirname(dirname(fileURLToPath(import.meta.url)));
const out = mkdtempSync(join(tmpdir(), "ras-playground-test-"));
const tests = readdirSync(join(web, "test")).filter((f) => f.endsWith(".test.ts"));
try {
    const files = [];
    for (const t of tests) {
        const outfile = join(web, "test", `.${t.replace(/\.ts$/, ".mjs")}`);
        await build({ entryPoints: [join(web, "test", t)], bundle: true, platform: "node", format: "esm", outfile, logLevel: "warning", packages: "bundle" });
        files.push(outfile);
    }
    const r = spawnSync(process.execPath, ["--test", ...files], { stdio: "inherit" });
    for (const f of files) rmSync(f, { force: true });
    process.exitCode = r.status ?? 1;
} finally {
    rmSync(out, { recursive: true, force: true });
}
