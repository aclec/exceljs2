// Validates XML files against the schemas of ECMA-376 / ISO/IEC 29500-4 (transitional) with xmllint, when both are
// available: set OOXML_SCHEMAS to the directory of the schemas (dml-chart.xsd, dml-spreadsheetDrawing.xsd and the
// schemas they import), e.g. `OOXML_SCHEMAS=/path/to/ISO-IEC29500-4_2016 bun run test`. Without them, the specs that
// validate are skipped.
const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");

/**
 * @returns a function `(xsd, files) => { ok, output }` validating files (paths) against a schema of the directory (its
 *   file name), or null when OOXML_SCHEMAS or xmllint is missing
 */
function schemaValidator() {
    const dir = process.env.OOXML_SCHEMAS;
    if (!dir || !fs.existsSync(path.join(dir, "dml-chart.xsd"))) {
        return null;
    }
    const probe = spawnSync("xmllint", ["--version"], { encoding: "utf8" });
    if (probe.error || probe.status !== 0) {
        return null;
    }
    return (xsd, files) => {
        const result = spawnSync("xmllint", ["--noout", "--schema", path.join(dir, xsd), ...files], { encoding: "utf8" });
        return { ok: result.status === 0, output: `${result.stdout}${result.stderr}` };
    };
}

// the schema of each kind of part validated by validateDrawingParts
const PART_SCHEMAS = [
    { pattern: /^xl\/charts\/chart\d+\.xml$/, xsd: "dml-chart.xsd" },
    { pattern: /^xl\/drawings\/drawing\d+\.xml$/, xsd: "dml-spreadsheetDrawing.xsd" },
    { pattern: /^xl\/chartsheets\/sheet\d+\.xml$/, xsd: "sml.xsd" },
];

/**
 * Validates the chart, drawing and chartsheet parts of a file against their schemas (dml-chart.xsd,
 * dml-spreadsheetDrawing.xsd, sml.xsd), when the schemas and xmllint are available.
 *
 * @param parts - the parts of the file (name → bytes, spec/utils/xlsx-tamper.js unzipParts)
 * @param dir - where the parts are written for xmllint (emptied first)
 * @returns null when OOXML_SCHEMAS or xmllint is missing, else `{ ok, output, count }` (count: the parts validated)
 */
function validateDrawingParts(parts, dir) {
    const validate = schemaValidator();
    if (!validate) {
        return null;
    }
    fs.rmSync(dir, { recursive: true, force: true });
    fs.mkdirSync(dir, { recursive: true });
    const files = new Map(PART_SCHEMAS.map(({ xsd }) => [xsd, []]));
    for (const name of Object.keys(parts)) {
        const schema = PART_SCHEMAS.find(({ pattern }) => pattern.test(name));
        if (schema) {
            const file = path.join(dir, name.replace(/\//g, "-"));
            fs.writeFileSync(file, parts[name]);
            files.get(schema.xsd).push(file);
        }
    }
    let ok = true;
    let output = "";
    let count = 0;
    for (const [xsd, list] of files) {
        if (list.length) {
            const result = validate(xsd, list);
            ok = ok && result.ok;
            output += result.output;
            count += list.length;
        }
    }
    return { ok, output, count };
}

module.exports = { schemaValidator, validateDrawingParts };
