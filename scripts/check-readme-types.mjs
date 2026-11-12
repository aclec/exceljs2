/* eslint-disable no-console -- command-line report */
// Type-checks the JavaScript examples of README.md against the public declarations (index.d.ts), as strict TypeScript.
//
//   node scripts/check-readme-types.mjs [file.md]       (bun run check:readme; default README.md)
//
// Each fenced ```javascript / ```js block becomes its own in-memory module: the block body is wrapped in an async
// function, its `import` declarations are hoisted, and the free variables that the examples assume (`workbook`,
// `worksheet`, `cell`, `filename`, `stream`...) are declared with their public types (PRELUDE, STREAMING_WRITER). The
// README lists alternatives in one block (`const sheet = ...` several times): a declaration that redeclares a name of
// the same block opens a nested scope instead of failing. The package is imported by its own name (`@aclec/exceljs2`,
// resolved to index.d.ts through the `exports` of package.json) or as `exceljs` (mapped to index.d.ts).
//
// Not reported: implicit `any` (a JavaScript example cannot carry type annotations), and the KNOWN errors below: valid
// JavaScript that strict TypeScript correctly rejects (a value that must be narrowed first). A KNOWN entry that no
// longer matches an error fails the check, so that the list cannot go stale.
import { createRequire } from "node:module";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const ts = createRequire(path.join(root, "package.json"))("typescript");
const readmePath = path.resolve(process.argv[2] ?? path.join(root, "README.md"));
const readmeName = path.relative(process.cwd(), readmePath).startsWith("..") ? readmePath : path.relative(process.cwd(), readmePath);

/** Implicit-any diagnostics: a JavaScript example has no annotations. */
const IGNORED_CODES = new Set([7005, 7006, 7008, 7019, 7031, 7034]);

/**
 * Errors that are expected: the example is valid JavaScript, strict TypeScript requires narrowing first (a JavaScript
 * example cannot carry the `!` or the cast). Matched by error code and by a fragment of the README line that the error
 * points to (`line`) or of the error message (`message`). spec/typescript/align-readme.ts checks the narrowed flows.
 */
const KNOWN = [
    // cell.note is `string | Comment | undefined` (a plain-text note is a string; a cell without a note has none)
    ...["note.margins = {", "note.protection = {", "note.editAs = 'twoCells';"].flatMap((line) => [
        { code: 2339, line },
        { code: 2532, line },
    ]),
    // worksheet.getTable(name) is undefined when the sheet has no table of that name
    { code: 18048, message: "'table' is possibly 'undefined'." },
    // `const options = { worksheets: 'emit' }` widens 'emit' to string (TypeScript needs `as const`)
    { code: 2345, line: "new ExcelJS.stream.xlsx.WorkbookReader('./file.xlsx', options)" },
];

/** Free variables of the examples, declared once per module (a block that declares the same name shadows it). */
const PRELUDE = [
    'import * as ExcelJS from "@aclec/exceljs2";',
    'import * as fs from "fs";',
    "declare const workbook: ExcelJS.Workbook;",
    "declare const workbookWriter: ExcelJS.stream.xlsx.WorkbookWriter;",
    "declare const worksheet: ExcelJS.Worksheet;",
    "declare const ws: ExcelJS.Worksheet;",
    "declare const sheet: ExcelJS.Worksheet;",
    "declare const row: ExcelJS.Row;",
    "declare const cell: ExcelJS.Cell;",
    "declare const filename: string;",
    'declare const stream: import("stream").PassThrough;',
    "declare const data: ExcelJS.Buffer;",
    "declare const options: Partial<ExcelJS.WorksheetProtection>;",
    "declare const imageId: number;",
    "declare const imageId1: number;",
    "declare const imageId2: number;",
    "declare const myFonts: { arial: Partial<ExcelJS.Font> };",
    "declare const myInput: { value: string };",
    "declare function createAndFillWorkbook(): ExcelJS.Workbook;",
    "declare function dayjs(value: unknown): { format(template: string): string };",
    // the README asserts with chai (`expect(x).to.equal(y)`) and jest (`expect(x).toEqual(y)`) styles
    "interface ReadmeExpectation {",
    "    to: { equal(v: unknown): void; be: { true: void; false: void }; have: { members(v: unknown[]): void } };",
    "    not: ReadmeExpectation;",
    "    toBe(v: unknown): void;",
    "    toEqual(v: unknown): void;",
    "    toBeNull(): void;",
    "}",
    "declare function expect(actual: unknown): ReadmeExpectation;",
];

/** Under the "Streaming XLSX Writer" heading, `workbook` and `worksheet` are the streaming writers. */
const STREAMING_WRITER = [
    "declare const workbook: ExcelJS.stream.xlsx.WorkbookWriter;",
    "declare const worksheet: ExcelJS.stream.xlsx.WorksheetWriter;",
    "declare const records: AsyncIterable<Record<string, ExcelJS.CellValue>>;",
    "declare const i: number;",
    "declare const theName: string;",
    "declare const someOtherDetail: string;",
];

/** The name declared by a prelude line (`declare const x`, `declare function x`, `import * as x`), or undefined. */
function declaredName(line) {
    return /^(?:declare (?:const|function) |import \* as )(\w+)/.exec(line)?.[1];
}

/** The fenced js blocks of the README: first code line (1-based), code, nearest heading. */
function extractBlocks(text) {
    const lines = text.split("\n");
    const blocks = [];
    let heading = "";
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (/^#{1,6} /.test(line)) heading = line;
        const fence = /^```(javascript|js)\s*$/.exec(line);
        if (!fence) continue;
        const start = i + 1;
        let end = start;
        while (end < lines.length && !lines[end].startsWith("```")) end++;
        blocks.push({ line: start + 1, code: lines.slice(start, end).join("\n"), heading });
        i = end;
    }
    return blocks;
}

/** Names bound by a top-level statement, and whether they are block-scoped (`let` / `const` / class / import). */
function declaredNames(statement) {
    const names = [];
    const collect = (binding) => {
        if (ts.isIdentifier(binding)) names.push(binding.text);
        else for (const element of binding.elements) if (!ts.isOmittedExpression(element)) collect(element.name);
    };
    if (ts.isVariableStatement(statement)) {
        for (const declaration of statement.declarationList.declarations) collect(declaration.name);
    } else if ((ts.isFunctionDeclaration(statement) || ts.isClassDeclaration(statement)) && statement.name) {
        names.push(statement.name.text);
    } else if (ts.isImportDeclaration(statement) && statement.importClause) {
        const clause = statement.importClause;
        if (clause.name) names.push(clause.name.text);
        const bindings = clause.namedBindings;
        if (bindings && ts.isNamespaceImport(bindings)) names.push(bindings.name.text);
        if (bindings && ts.isNamedImports(bindings)) for (const element of bindings.elements) names.push(element.name.text);
    }
    return names;
}

/**
 * Turns one README block into in-memory modules. Returns `{ text, lineMap }` per module; `lineMap[i]` is the README
 * line of the module's line `i` (null for the prelude).
 */
function toModules(block) {
    const sf = ts.createSourceFile("block.ts", block.code, ts.ScriptTarget.ES2022, true, ts.ScriptKind.TS);
    const extra = /Streaming XLSX Writer/.test(block.heading) ? STREAMING_WRITER : [];
    const lineOf = (pos) => block.line + sf.getLineAndCharacterOfPosition(pos).line;

    // hoist the imports (a second import of the same name starts another module), nest the redeclarations
    const importGroups = [[]];
    let imported = new Set();
    const insertions = []; // positions where a nested scope opens
    const blanked = []; // [start, end) of the hoisted imports
    let scope = new Set();
    for (const statement of sf.statements) {
        const names = declaredNames(statement);
        if (ts.isImportDeclaration(statement)) {
            if (names.some((name) => imported.has(name))) {
                importGroups.push([]);
                imported = new Set();
            }
            names.forEach((name) => imported.add(name));
            importGroups[importGroups.length - 1].push({ text: statement.getText(sf), line: lineOf(statement.getStart(sf)) });
            blanked.push([statement.getStart(sf), statement.getEnd()]);
            continue;
        }
        if (names.some((name) => scope.has(name))) {
            insertions.push(statement.getStart(sf));
            scope = new Set();
        }
        names.forEach((name) => scope.add(name));
    }

    let body = "";
    let last = 0;
    const cuts = [...insertions.map((pos) => ({ pos, kind: "open" })), ...blanked.map(([pos, end]) => ({ pos, end, kind: "blank" }))];
    cuts.sort((a, b) => a.pos - b.pos);
    for (const cut of cuts) {
        body += block.code.slice(last, cut.pos);
        if (cut.kind === "open") {
            body += "{ ";
            last = cut.pos;
        } else {
            body += block.code.slice(cut.pos, cut.end).replace(/[^\n]/g, " ");
            last = cut.end;
        }
    }
    body += block.code.slice(last) + "\n" + "}".repeat(insertions.length);

    return importGroups.map((imports, index) => {
        const ownNames = new Set(
            imports.flatMap((i) => ts.createSourceFile("i.ts", i.text, ts.ScriptTarget.ES2022).statements.flatMap(declaredNames))
        );
        const overridden = new Set(extra.map(declaredName));
        const prelude = [...PRELUDE.filter((line) => !overridden.has(declaredName(line))), ...extra].filter(
            (line) => !ownNames.has(declaredName(line))
        );
        const lines = [];
        const lineMap = [];
        for (const i of imports) {
            lines.push(i.text);
            lineMap.push(i.line);
        }
        for (const line of prelude) {
            lines.push(line);
            lineMap.push(null);
        }
        lines.push("export async function readmeExample(): Promise<void> {");
        lineMap.push(null);
        if (index === 0) {
            body.split("\n").forEach((line, offset) => {
                lines.push(line);
                lineMap.push(block.line + offset);
            });
        }
        lines.push("}");
        lineMap.push(null);
        return { text: lines.join("\n"), lineMap };
    });
}

const readme = fs.readFileSync(readmePath, "utf8");
const readmeLines = readme.split("\n");
const modules = new Map();
for (const block of extractBlocks(readme)) {
    toModules(block).forEach((module, index) => {
        modules.set(path.join(root, "__readme__", `line-${block.line}-${index}.ts`), module);
    });
}

const options = {
    strict: true,
    noEmit: true,
    target: ts.ScriptTarget.ES2022,
    lib: ["lib.es2022.d.ts"],
    module: ts.ModuleKind.NodeNext,
    moduleResolution: ts.ModuleResolutionKind.NodeNext,
    esModuleInterop: true,
    skipLibCheck: true,
    types: ["node"],
    // the README imports `exceljs` (the fork installed under that name through a package alias, see "Installation")
    paths: { exceljs: ["./index.d.ts"] },
    pathsBasePath: root,
};
const host = ts.createCompilerHost(options);
const { getSourceFile, fileExists, readFile } = host;
host.fileExists = (file) => modules.has(path.resolve(file)) || fileExists.call(host, file);
host.readFile = (file) => modules.get(path.resolve(file))?.text ?? readFile.call(host, file);
host.getSourceFile = (file, languageVersion, ...rest) => {
    const module = modules.get(path.resolve(file));
    return module
        ? ts.createSourceFile(file, module.text, languageVersion, true)
        : getSourceFile.call(host, file, languageVersion, ...rest);
};
const program = ts.createProgram([...modules.keys()], options, host);

const errors = [];
const used = new Set();
let knownErrors = 0;
for (const diagnostic of ts.getPreEmitDiagnostics(program)) {
    if (IGNORED_CODES.has(diagnostic.code)) continue;
    const message = ts.flattenDiagnosticMessageText(diagnostic.messageText, "\n");
    const module = diagnostic.file && modules.get(path.resolve(diagnostic.file.fileName));
    if (!module) {
        errors.push(`${diagnostic.file ? diagnostic.file.fileName : "(global)"}: TS${diagnostic.code} ${message}`);
        continue;
    }
    const { line, character } = diagnostic.file.getLineAndCharacterOfPosition(diagnostic.start);
    const readmeLine = module.lineMap[line];
    if (readmeLine === null || readmeLine === undefined) {
        errors.push(`${path.basename(diagnostic.file.fileName)} (prelude line ${line + 1}): TS${diagnostic.code} ${message}`);
        continue;
    }
    const known = KNOWN.findIndex(
        (k) =>
            k.code === diagnostic.code &&
            (k.line === undefined || readmeLines[readmeLine - 1].includes(k.line)) &&
            (k.message === undefined || message.includes(k.message))
    );
    if (known >= 0) {
        used.add(known);
        knownErrors++;
        continue;
    }
    errors.push(`${readmeName}:${readmeLine}:${character + 1} TS${diagnostic.code} ${message.split("\n").join("\n    ")}`);
}
KNOWN.forEach((k, index) => {
    if (!used.has(index)) errors.push(`stale KNOWN entry (no matching error): TS${k.code} ${JSON.stringify(k.line ?? k.message)}`);
});

if (errors.length) {
    console.error(errors.join("\n"));
    console.error(`\n${errors.length} error(s) in the examples of ${readmeName}`);
    process.exit(1);
}
console.log(`${readmeName}: ${modules.size} examples type-check (${knownErrors} known errors, see KNOWN)`);
