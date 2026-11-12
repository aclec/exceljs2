// Builds the self-contained browser bundles in dist/ with esbuild.
//
//   dist/exceljs.js       dist/exceljs.min.js       (+ .map)  from lib/exceljs.browser.js
//   dist/exceljs.bare.js  dist/exceljs.bare.min.js  (+ .map)  from lib/exceljs.bare.js
//   dist/LICENSE
//
// The bundles target ES2020 and expose the global `ExcelJS`; when loaded through a CommonJS or AMD
// loader they export it there instead (UMD-compatible, like the previous browserify standalone build).
// Node builtins are replaced by small browser polyfills (see ALIASES and scripts/browser-shims/).
import { createRequire } from "node:module";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import * as esbuild from "esbuild";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(path.join(root, "package.json"));
const pkg = require("./package.json");
const dist = path.join(root, "dist");
const shim = (name) => path.join(root, "scripts", "browser-shims", name);
// trailing slash: resolve the npm polyfill package, not the node builtin of the same name
const polyfill = (name) => require.resolve(`${name}/`);

const ALIASES = {
    buffer: polyfill("buffer"),
    crypto: shim("crypto.js"),
    events: polyfill("events"),
    fs: shim("empty.js"),
    stream: require.resolve("stream-browserify"),
    string_decoder: polyfill("string_decoder"),
    util: shim("util.js"),
    "util-deprecate": shim("util-deprecate.js"),
    // lib/utils/deflate.js falls back to fflate when node:zlib is not there
    zlib: shim("empty.js"),
};

// lib/ and its dependencies may use either the bare or the node: prefixed specifier
for (const name of Object.keys(ALIASES)) ALIASES[`node:${name}`] = ALIASES[name];

const BUNDLES = [
    { entry: "lib/exceljs.browser.js", name: "exceljs" },
    { entry: "lib/exceljs.bare.js", name: "exceljs.bare" },
];

const banner = `/*! ExcelJS v${pkg.version} | ${pkg.license} License | https://github.com/aclec/exceljs2 (fork of exceljs/exceljs) */`;
const footer = [
    'if (typeof module === "object" && module.exports) { module.exports = ExcelJS; }',
    'else if (typeof define === "function" && define.amd) { define([], function () { return ExcelJS; }); }',
].join("\n");

function options(entry, outfile, minify) {
    return {
        absWorkingDir: root,
        entryPoints: [entry],
        outfile,
        bundle: true,
        format: "iife",
        globalName: "ExcelJS",
        platform: "browser",
        target: ["es2020"],
        minify,
        sourcemap: "linked",
        charset: "ascii",
        legalComments: "eof",
        inject: [shim("globals.mjs")],
        define: { "process.browser": "true", global: "globalThis" },
        alias: ALIASES,
        banner: { js: banner },
        footer: { js: footer },
        logLevel: "warning",
    };
}

const started = Date.now();
fs.rmSync(dist, { recursive: true, force: true });
fs.mkdirSync(dist);

await Promise.all(
    BUNDLES.flatMap(({ entry, name }) => [
        esbuild.build(options(entry, path.join(dist, `${name}.js`), false)),
        esbuild.build(options(entry, path.join(dist, `${name}.min.js`), true)),
    ])
);
fs.copyFileSync(path.join(root, "LICENSE"), path.join(dist, "LICENSE"));

for (const file of fs.readdirSync(dist).sort()) {
    const { size } = fs.statSync(path.join(dist, file));
    process.stdout.write(`  dist/${file.padEnd(28)} ${(size / 1024).toFixed(1).padStart(8)} KiB\n`);
}
process.stdout.write(`built in ${Date.now() - started} ms\n`);
