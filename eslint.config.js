"use strict";

// ESLint flat config. Only correctness / quality / security rules live here: all formatting is
// owned by Prettier (.prettierrc), and eslint-config-prettier (last) turns off any rule that
// could disagree with it.
const js = require("@eslint/js");
const n = require("eslint-plugin-n");
const globals = require("globals");
const prettier = require("eslint-config-prettier");
const tseslint = require("typescript-eslint");

// crypto is listed because its browser shim (scripts/browser-shims/crypto.js) only implements what
// lib/utils/encryptor.js uses (getHashes, createHash, randomBytes): other lib files go through Encryptor.
const NODE_ONLY_MODULES = ["child_process", "crypto", "fs", "http", "https", "net", "os", "path", "vm", "worker_threads"];

module.exports = [
    {
        ignores: [
            "build/**",
            "coverage/**",
            "dist/**",
            "graphify-out/**",
            // compiled from src/ by `bun run build:lib`
            "lib/**",
            "out/**",
            "spec/out/**",
            // ad-hoc manual scripts: not run by any suite or CI, some need packages that were removed (sax)
            "test/**",
        ],
    },
    js.configs.recommended,
    n.configs["flat/recommended-script"],
    {
        plugins: {
            // eslint-plugin-n is the maintained fork of eslint-plugin-node with the same rule names; registering it
            // under the old prefix keeps existing `eslint-disable node/...` comments in lib/ valid.
            node: n,
        },
        languageOptions: {
            ecmaVersion: 2022,
            sourceType: "commonjs",
            globals: { ...globals.node },
        },
        rules: {
            // kept from the previous .eslintrc (non-formatting rules)
            "no-console": ["error", { allow: ["warn"] }],
            "no-restricted-syntax": ["error", "ForInStatement", "LabeledStatement", "WithStatement"],
            // caughtErrors: 'none' was the default of ESLint < 9, which the previous config relied on
            "no-unused-vars": ["error", { vars: "all", args: "none", ignoreRestSiblings: true, caughtErrors: "none" }],
            "no-use-before-define": ["error", { variables: false, classes: false, functions: false }],
            "prefer-destructuring": ["warn", { array: false, object: true }],
            // switched off in the previous .eslintrc and flagged by the recommended set on existing code
            "no-prototype-builtins": "off",

            // security
            "no-eval": "error",
            "no-implied-eval": "error",
            "no-new-func": "error",
            "no-proto": "error",
            "no-script-url": "error",

            // quality
            eqeqeq: ["error", "smart"],
            "no-var": "error",
            "prefer-const": "error",
            // new in eslint:recommended 10; they flag existing lib/ code (1 + 3 places), to be fixed there
            "no-unassigned-vars": "warn",
            "no-useless-assignment": "warn",

            // Node: the supported versions come from package.json "engines"
            "n/no-deprecated-api": "error",
            // globalThis.crypto and Readable.fromWeb are stable from Node 23 but available (experimental) in 22.12
            "n/no-unsupported-features/node-builtins": ["error", { allowExperimental: true }],
        },
    },
    {
        // Code reachable from the browser bundle (src/exceljs.browser.ts, src/exceljs.bare.ts) must not use
        // Node-only modules. The exceptions are the streaming API and the file helpers, which are not
        // available in the browser (fs is replaced by an empty module there).
        files: ["src/**/*.ts"],
        ignores: [
            "src/exceljs.nodejs.ts",
            "src/stream/**",
            "src/csv/csv.ts", // CSV#readFile / CSV#writeFile
            "src/utils/encryptor.ts", // crypto, polyfilled in the bundle by scripts/browser-shims/crypto.js
            "src/utils/utils.ts", // utils.fs.exists, used by readFile
            "src/xlsx/xlsx.ts", // XLSX#readFile / XLSX#writeFile
        ],
        rules: {
            "no-restricted-imports": [
                "error",
                {
                    paths: NODE_ONLY_MODULES.flatMap((m) => [m, `node:${m}`]).map((name) => ({
                        name,
                        message: "browser bundle: Node-only module",
                    })),
                    patterns: [
                        {
                            group: NODE_ONLY_MODULES.flatMap((m) => [`${m}/*`, `node:${m}/*`]),
                            message: "browser bundle: Node-only module",
                        },
                    ],
                },
            ],
        },
    },
    // TypeScript sources: type-aware rules on top of the shared ones above. `any` is banned: use precise types,
    // generics or `unknown` + narrowing.
    ...tseslint.configs.recommendedTypeChecked.map((config) => ({ ...config, files: ["src/**/*.ts"] })),
    {
        files: ["src/**/*.ts"],
        languageOptions: {
            sourceType: "module",
            parserOptions: { projectService: true, tsconfigRootDir: __dirname },
        },
        rules: {
            // replaced by the TypeScript-aware versions
            "no-unused-vars": "off",
            "@typescript-eslint/no-unused-vars": ["error", { vars: "all", args: "none", ignoreRestSiblings: true, caughtErrors: "none" }],
            "no-use-before-define": "off",
            "@typescript-eslint/no-use-before-define": ["error", { variables: false, classes: false, functions: false, typedefs: false }],
            // eslint-plugin-n resolves require() of .js files; tsc checks the imports of the .ts sources
            "n/no-missing-import": "off",
            "n/no-missing-require": "off",
            "n/no-unpublished-import": "off",

            "@typescript-eslint/no-explicit-any": "error",
            "@typescript-eslint/consistent-type-imports": ["error", { fixStyle: "inline-type-imports" }],
            "@typescript-eslint/consistent-type-exports": "error",
            "@typescript-eslint/explicit-module-boundary-types": "error",
            "@typescript-eslint/no-non-null-assertion": "off",
            "@typescript-eslint/ban-ts-comment": ["error", { "ts-expect-error": "allow-with-description", minimumDescriptionLength: 10 }],
            "@typescript-eslint/switch-exhaustiveness-check": "error",
            "@typescript-eslint/prefer-nullish-coalescing": "off",
            "@typescript-eslint/no-unnecessary-type-assertion": "error",
        },
    },
    {
        files: ["**/*.mjs"],
        languageOptions: { sourceType: "module" },
    },
    {
        // development-only files: they are not published, may use devDependencies and exit the process
        files: ["benchmark.js", "eslint.config.js", "scripts/**", "spec/**"],
        rules: {
            "n/no-process-exit": "off",
            "n/no-unpublished-import": "off",
            "n/no-unpublished-require": "off",
        },
    },
    {
        files: ["spec/**"],
        languageOptions: {
            globals: { ...globals.mocha, expect: "readonly", verquire: "readonly" },
        },
        rules: {
            // row values are written as sparse arrays on purpose ([, 'A1', 'B1'] = columns start at 1)
            "no-sparse-arrays": "off",
            // chai / dirty-chai assertions are expressions (expect(x).to.be.true())
            "no-unused-expressions": "off",
        },
    },
    {
        // functions passed to page.evaluate() run inside the browser page
        files: ["spec/browser/**"],
        languageOptions: { globals: { ExcelJS: "readonly" } },
    },
    prettier,
];
