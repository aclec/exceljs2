// Compiles src/**/*.ts to lib/ (CommonJS + .d.ts + source maps) with tsc -p tsconfig.build.json, then copies the
// non-TypeScript assets of src/ (xml, json) next to the compiled files. Fails on any type error.
import { spawnSync } from "node:child_process";
import { createRequire } from "node:module";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const require = createRequire(path.join(root, "package.json"));
const src = path.join(root, "src");
const lib = path.join(root, "lib");

fs.rmSync(lib, { recursive: true, force: true });

const tsc = require.resolve("typescript/bin/tsc");
const { status } = spawnSync(process.execPath, [tsc, "-p", path.join(root, "tsconfig.build.json")], { stdio: "inherit" });
if (status !== 0) {
    process.exitCode = status ?? 1;
} else {
    for (const entry of fs.readdirSync(src, { recursive: true, withFileTypes: true })) {
        if (!entry.isFile() || entry.name.endsWith(".ts")) continue;
        const from = path.join(entry.parentPath, entry.name);
        const to = path.join(lib, path.relative(src, from));
        fs.mkdirSync(path.dirname(to), { recursive: true });
        fs.copyFileSync(from, to);
    }
}
