declare module "virtual:samples" {
    const samples: { id: string; title: string; files: { path: string; text: string }[] }[];
    export default samples;
}

// Node modules the VS Code extension's modules import, which the page bundles with the
// browser's (scripts/build.mjs).
declare module "node:crypto" {
    export function randomUUID(): string;
}
declare module "path" {
    import path = require("path-browserify");
    export = path;
}

/** An id of the build (scripts/build.mjs), put on the files the page loads. */
declare const RAS_BUILD: string;
