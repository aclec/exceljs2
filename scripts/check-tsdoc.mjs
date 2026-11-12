/* eslint-disable no-console -- command-line report */
// Fails when an exported symbol of src/ or a public member of an exported class has no TSDoc comment (/** ... */).
//
//   node scripts/check-tsdoc.mjs        (bun run check:tsdoc)
//
// Not reported: private / protected / `_`-prefixed members; `override` members and the members of an implemented
// interface of the same file (their TSDoc is inherited); a setter whose getter is documented; the overload signatures
// after a documented first one; in xforms, `map` / `model` / `parser` (documented on BaseXform) and `tag` (the element
// name, part of the xform contract documented on BaseXform).
import { createRequire } from "node:module";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const ts = createRequire(path.join(root, "package.json"))("typescript");
const src = path.join(root, "src");

const XFORM_CONTRACT = new Set(["map", "model", "parser", "tag"]);

const hasDoc = (node, sf) => (ts.getLeadingCommentRanges(sf.text, node.getFullStart()) || []).some((r) => sf.text.startsWith("/**", r.pos));
const modifiers = (node) => (ts.canHaveModifiers(node) ? ts.getModifiers(node) || [] : []).map((m) => m.kind);
const memberName = (member, sf) => (member.name ? member.name.getText(sf) : "");

const missing = [];
for (const file of fs.readdirSync(src, { recursive: true })) {
    if (!file.endsWith(".ts")) continue;
    const full = path.join(src, file);
    const sf = ts.createSourceFile(full, fs.readFileSync(full, "utf8"), ts.ScriptTarget.ES2022, true);
    const report = (node, name) => missing.push(`src/${file}:${sf.getLineAndCharacterOfPosition(node.getStart(sf)).line + 1} ${name}`);

    // names exported by `export default X` / `export { X }`, and the documented members of the interfaces of the file
    const exportedNames = new Set();
    const interfaceDocs = new Map();
    for (const st of sf.statements) {
        if (ts.isExportAssignment(st) && ts.isIdentifier(st.expression)) exportedNames.add(st.expression.text);
        if (ts.isExportDeclaration(st) && st.exportClause && ts.isNamedExports(st.exportClause)) {
            for (const e of st.exportClause.elements) exportedNames.add((e.propertyName || e.name).text);
        }
        if (ts.isInterfaceDeclaration(st)) {
            interfaceDocs.set(st.name.text, new Set(st.members.filter((m) => hasDoc(m, sf)).map((m) => memberName(m, sf))));
        }
    }

    const seenFunctions = new Set();
    for (const st of sf.statements) {
        const name = st.name && ts.isIdentifier(st.name) ? st.name.text : undefined;
        const exported = modifiers(st).includes(ts.SyntaxKind.ExportKeyword) || (name !== undefined && exportedNames.has(name));
        if (ts.isFunctionDeclaration(st) && name) {
            // the TSDoc of overloaded functions is on their first signature
            if (seenFunctions.has(name)) continue;
            seenFunctions.add(name);
        }
        if (!exported) continue;
        if (ts.isClassDeclaration(st) && name) {
            if (!hasDoc(st, sf)) report(st, `class ${name}`);
            const heritage = st.heritageClauses || [];
            const isXform = name.endsWith("Xform") && heritage.some((h) => h.token === ts.SyntaxKind.ExtendsKeyword);
            const inherited = new Set();
            for (const h of heritage) {
                for (const type of h.types) for (const n of interfaceDocs.get(type.expression.getText(sf)) || []) inherited.add(n);
            }
            const documentedGetters = new Set(st.members.filter((m) => ts.isGetAccessor(m) && hasDoc(m, sf)).map((m) => memberName(m, sf)));
            const seenMethods = new Set();
            for (const m of st.members) {
                const kinds = modifiers(m);
                if (kinds.includes(ts.SyntaxKind.PrivateKeyword) || kinds.includes(ts.SyntaxKind.ProtectedKeyword)) continue;
                if (kinds.includes(ts.SyntaxKind.OverrideKeyword)) continue;
                if (ts.isConstructorDeclaration(m) || ts.isClassStaticBlockDeclaration(m) || ts.isIndexSignatureDeclaration(m)) continue;
                const n = memberName(m, sf);
                if (!n || n.startsWith("_") || n.startsWith("#") || n.startsWith("[")) continue;
                if (inherited.has(n) || (isXform && XFORM_CONTRACT.has(n))) continue;
                if (ts.isSetAccessor(m) && documentedGetters.has(n)) continue;
                if (ts.isMethodDeclaration(m)) {
                    if (seenMethods.has(n)) continue;
                    seenMethods.add(n);
                }
                if (!hasDoc(m, sf)) report(m, `${name}.${n}`);
            }
        } else if (
            ts.isFunctionDeclaration(st) ||
            ts.isInterfaceDeclaration(st) ||
            ts.isTypeAliasDeclaration(st) ||
            ts.isVariableStatement(st) ||
            ts.isEnumDeclaration(st)
        ) {
            if (!hasDoc(st, sf)) report(st, name ?? st.declarationList.declarations[0].name.getText(sf));
        }
    }
}

if (missing.length) {
    console.error(`TSDoc missing on ${missing.length} exported symbol(s) / public member(s):\n${missing.join("\n")}`);
    process.exitCode = 1;
}
