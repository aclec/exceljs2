// A pivot chart (a chart part with a `<c:pivotSource>`) is written as a static chart when the workbook written does not
// carry the pivot table it names (a pivot table dropped on read, the worksheet of the pivot table removed or renamed...),
// as Excel does when the pivot table of a pivot chart is deleted: without its `<c:pivotSource>` (CT_ChartSpace), its
// `<c:pivotFmts>` (CT_Chart) and the pivot options of its extensions (c14:pivotOptions, c16:pivotOptions16). The series
// keep their formulas and cached values. A chart that names a pivot table written is written as read.
//
// The chart part comes from an untrusted file: it is parsed once with saxes (linear, nesting bounded as in parse-sax), into
// a tree that is edited (whole elements removed) and written back through XmlStream (names as parsed, attribute values
// and text encoded). Comments and processing instructions are not kept. A chart that is not well-formed is written as read,
// as every chart is carried (chart-parts.ts does not parse the charts).

import { SaxesParser } from "saxes";
import XmlStream from "../utils/xml-stream";

/** Whether a pivot table named by the `<c:pivotSource><c:name>` of a pivot chart is written (pivot-parts.ts PivotPartWriter). */
export interface PivotSourceLookup {
    /** @returns `true` when the pivot table that `source` names (`[Book1.xlsx]Sheet1!PivotTable1`) is written */
    hasPivotTable(source: string): boolean;
}

// the nesting bound of parse-sax: a chart is nested less than 30 deep
const MAX_DEPTH = 1000;

// the extensions of CT_ChartSpace that hold the pivot options of a pivot chart, by uri (MS-ODRAWXML), and their elements
const PIVOT_EXTENSION_URIS: ReadonlySet<string> = new Set([
    "{781A3756-C4B2-4CAC-9D66-4F8BD8637D16}",
    "{E28EC0CA-F0BB-4C9C-879D-F8772B89E7AC}",
]);
const PIVOT_OPTIONS: ReadonlySet<string> = new Set(["pivotOptions", "pivotOptions16"]);

/** An element of the chart part, as parsed. */
interface ChartNode {
    readonly name: string;
    readonly attributes: Readonly<Record<string, string>>;
    children: (ChartNode | string)[];
}

// the local name of a qualified name
function localName(name: string): string {
    return name.slice(name.indexOf(":") + 1);
}

function elements(node: ChartNode): ChartNode[] {
    return node.children.filter((child): child is ChartNode => typeof child !== "string");
}

function textOf(node: ChartNode): string {
    return node.children.filter((child): child is string => typeof child === "string").join("");
}

/** The root element of a chart part, or undefined when the part is not well-formed (or nested too deep). */
function parseChart(text: string): ChartNode | undefined {
    const parser = new SaxesParser();
    const stack: ChartNode[] = [];
    let root: ChartNode | undefined;
    let failed = false;
    parser.on("error", () => {
        failed = true;
    });
    parser.on("opentag", (tag) => {
        if (stack.length >= MAX_DEPTH) {
            throw new Error("too deep");
        }
        const node: ChartNode = { name: tag.name, attributes: tag.attributes, children: [] };
        const parent = stack[stack.length - 1];
        if (parent) {
            parent.children.push(node);
        } else {
            root ??= node;
        }
        stack.push(node);
    });
    const addText = (value: string): void => {
        stack[stack.length - 1]?.children.push(value);
    };
    parser.on("text", addText);
    parser.on("cdata", addText);
    parser.on("closetag", () => {
        stack.pop();
    });
    try {
        parser.write(text).close();
    } catch {
        return undefined;
    }
    return failed ? undefined : root;
}

/** Removes the pivot elements of a chart space (mutation). */
function removePivotElements(chartSpace: ChartNode): void {
    chartSpace.children = chartSpace.children.filter((child) => typeof child === "string" || localName(child.name) !== "pivotSource");
    for (const child of elements(chartSpace)) {
        const local = localName(child.name);
        if (local === "chart") {
            child.children = child.children.filter((node) => typeof node === "string" || localName(node.name) !== "pivotFmts");
        } else if (local === "extLst") {
            child.children = child.children.filter(
                (ext) =>
                    typeof ext === "string" ||
                    localName(ext.name) !== "ext" ||
                    !(
                        PIVOT_EXTENSION_URIS.has(ext.attributes.uri ?? "") ||
                        elements(ext).some((option) => PIVOT_OPTIONS.has(localName(option.name)))
                    )
            );
        }
    }
}

/** Writes a chart part (with the XML declaration of the parts), without recursion. */
function renderChart(root: ChartNode): string {
    const xml = new XmlStream();
    xml.openXml(XmlStream.StdDocAttributes);
    // the nodes to write, last first; null: the end of the element opened before
    const pending: (ChartNode | string | null)[] = [root];
    for (let node = pending.pop(); node !== undefined; node = pending.pop()) {
        if (node === null) {
            xml.closeNode();
        } else if (typeof node === "string") {
            xml.writeText(node);
        } else {
            xml.openNode(node.name, node.attributes);
            pending.push(null);
            for (let i = node.children.length - 1; i >= 0; i--) {
                const child = node.children[i];
                if (child !== undefined) {
                    pending.push(child);
                }
            }
        }
    }
    return xml.xml;
}

/**
 * The bytes of a chart part to write: as read, or, for a pivot chart whose pivot table is not written, the static chart.
 *
 * @param pivotTables - the pivot tables written (none: every pivot chart is written static)
 */
export function chartPartBytes(data: Uint8Array, pivotTables: PivotSourceLookup | undefined): Uint8Array {
    const text = new TextDecoder().decode(data);
    if (!text.includes("pivotSource")) {
        return data;
    }
    const root = parseChart(text);
    if (!root || localName(root.name) !== "chartSpace") {
        return data;
    }
    const pivotSource = elements(root).find((child) => localName(child.name) === "pivotSource");
    if (!pivotSource) {
        return data;
    }
    const nameNode = elements(pivotSource).find((child) => localName(child.name) === "name");
    if (nameNode && pivotTables && pivotTables.hasPivotTable(textOf(nameNode))) {
        return data;
    }
    removePivotElements(root);
    return new TextEncoder().encode(renderChart(root));
}
