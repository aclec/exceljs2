import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { ChartXmlAttribute, ChartXmlElement } from "../../../types";

// The namespaces that a chart anchor is read with (resolved from the declarations in scope, whatever their prefixes)
const SPREADSHEET_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MARKUP_COMPATIBILITY_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const CHART_NS = new Set([
    "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "http://schemas.microsoft.com/office/drawing/2014/chartex",
]);

/**
 * The children of `<xdr:wsDr>` whose chart may be carried: the anchors, and `<mc:AlternateContent>` (Excel writes the
 * anchor of a chartEx chart in one, its `<mc:Choice Requires="cx1">` and a `<mc:Fallback>` text box).
 */
export const ANCHOR_ELEMENTS: ReadonlySet<string> = new Set([
    "xdr:twoCellAnchor",
    "xdr:oneCellAnchor",
    "xdr:absoluteAnchor",
    "mc:AlternateContent",
]);

// A character that cannot be in an XML name: saxes already rejects such names, checked again as the names are written
// back verbatim (no backtracking: one character class, unanchored test)
const NOT_NAME_CHAR = /[\s<>&"'/=]/;

// the prefix of a qualified name ("" for none) and its local name
function splitName(name: string): [prefix: string, local: string] {
    const colon = name.indexOf(":");
    return colon === -1 ? ["", name] : [name.slice(0, colon), name.slice(colon + 1)];
}

// the namespace prefix declared by an attribute ("" for the default namespace), or undefined for another attribute
function declaredPrefix(name: string): string | undefined {
    if (name === "xmlns") {
        return "";
    }
    return name.startsWith("xmlns:") ? name.slice(6) : undefined;
}

/** An element being captured, and the prefixes it declares (undone when it ends). */
interface OpenElement {
    readonly name: string;
    readonly attributes: ChartXmlAttribute[];
    readonly children: (ChartXmlElement | string)[];
    readonly declared: readonly string[];
}

/**
 * Captures one child of a drawing (an anchor, `<xdr:twoCellAnchor>`, `<xdr:oneCellAnchor>` or `<xdr:absoluteAnchor>`, or a
 * `<mc:AlternateContent>` holding anchors), every event from its start tag to its end tag, as a tree of elements: the
 * anchor of a chart is carried opaque and written back as parsed (through XmlStream, which encodes the attribute values
 * and the text). The relationship ids (attributes of the relationships namespace) and the ids of the shapes
 * (`<xdr:cNvPr id>`) are marked, to be renumbered on write.
 *
 * Linear in the size of the anchor: the namespaces it declares are kept as a stack per prefix (pushed and popped, never
 * copied), those of `<xdr:wsDr>` are looked up in its map, and only the ones the anchor uses are declared on it on write.
 * The depth is bounded by the SAX parser (parse-sax MAX_DEPTH).
 */
export class ChartAnchorCapture {
    private readonly stack: OpenElement[] = [];
    /** the namespaces declared inside the anchor: prefix → URIs, innermost last */
    private readonly bindings = new Map<string, string[]>();
    /** the prefixes the anchor uses that only <xdr:wsDr> declares */
    private readonly rootPrefixes = new Set<string>();
    private readonly rootScope: ReadonlyMap<string, string>;
    private readonly writtenScope: ReadonlyMap<string, string>;
    private hasChart = false;
    private valid = true;
    private captured: ChartXmlElement | null = null;

    /**
     * @param rootScope - the namespaces declared on the `<xdr:wsDr>` read (prefix → URI); not copied
     * @param writtenScope - those declared on the `<xdr:wsDr>` written (not declared again on the anchor)
     */
    constructor(rootScope: ReadonlyMap<string, string>, writtenScope: ReadonlyMap<string, string>) {
        this.rootScope = rootScope;
        this.writtenScope = writtenScope;
    }

    /** Marks the capture as not to be carried (e.g. an anchor of the element is a picture, read as an image). */
    discard(): void {
        this.valid = false;
    }

    // the URI of a prefix in scope ("" for the default namespace); a prefix declared by <xdr:wsDr> only is noted as used
    private lookup(prefix: string): string | undefined {
        const uris = this.bindings.get(prefix);
        if (uris && uris.length) {
            return uris[uris.length - 1];
        }
        this.rootPrefixes.add(prefix);
        return this.rootScope.get(prefix);
    }

    /** An element starts (the anchor itself first). */
    open(node: SaxTag): void {
        const parent = this.stack[this.stack.length - 1];
        const entries = Object.entries(node.attributes);
        const declared: string[] = [];
        for (const [name, value] of entries) {
            const prefix = declaredPrefix(name);
            if (prefix !== undefined) {
                const uris = this.bindings.get(prefix);
                if (uris) {
                    uris.push(value);
                } else {
                    this.bindings.set(prefix, [value]);
                }
                declared.push(prefix);
            }
        }
        const [elementPrefix, elementLocal] = splitName(node.name);
        const elementNs = this.lookup(elementPrefix);
        if (NOT_NAME_CHAR.test(node.name) || (!parent && elementLocal === "AlternateContent" && elementNs !== MARKUP_COMPATIBILITY_NS)) {
            this.valid = false;
        }
        const attributes: ChartXmlAttribute[] = [];
        let hasRel = false;
        for (const [name, value] of entries) {
            if (NOT_NAME_CHAR.test(name)) {
                this.valid = false;
            }
            const [prefix, local] = splitName(name);
            const ns = prefix && prefix !== "xmlns" ? this.lookup(prefix) : undefined;
            if (ns === RELATIONSHIPS_NS) {
                attributes.push({ name, value, role: "rel" });
                hasRel = true;
                continue;
            }
            if (name === "id" && elementLocal === "cNvPr" && elementNs === SPREADSHEET_DRAWING_NS) {
                attributes.push({ name, value, role: "shapeId" });
                continue;
            }
            attributes.push({ name, value });
            // markup compatibility names prefixes in values: <mc:Choice Requires="cx1">, mc:Ignorable="a14"
            if (
                (name === "Requires" && elementNs === MARKUP_COMPATIBILITY_NS) ||
                (local === "Ignorable" && ns === MARKUP_COMPATIBILITY_NS)
            ) {
                for (const required of value.split(" ")) {
                    if (required) {
                        this.lookup(required);
                    }
                }
            }
        }
        if (hasRel && elementLocal === "chart" && elementNs !== undefined && CHART_NS.has(elementNs)) {
            this.hasChart = true;
        }
        this.stack.push({ name: node.name, attributes, children: [], declared });
    }

    /** Text inside the anchor. */
    text(text: string): void {
        const element = this.stack[this.stack.length - 1];
        if (element) {
            element.children.push(text);
        }
    }

    /**
     * An element ends.
     * @returns `true` when it is the anchor (the capture is complete: see {@link result})
     */
    close(): boolean {
        const element = this.stack.pop();
        if (!element) {
            return true;
        }
        for (const prefix of element.declared) {
            this.bindings.get(prefix)?.pop();
        }
        const parent = this.stack[this.stack.length - 1];
        if (parent) {
            parent.children.push({ name: element.name, attributes: element.attributes, children: element.children });
            return false;
        }
        // the anchor is written in another <xdr:wsDr>: it declares the namespaces of the file's that it uses (unless the
        // <xdr:wsDr> written declares them the same)
        const { attributes } = element;
        for (const prefix of this.rootPrefixes) {
            const uri = this.rootScope.get(prefix);
            if (uri !== undefined && this.writtenScope.get(prefix) !== uri) {
                attributes.push({ name: prefix ? `xmlns:${prefix}` : "xmlns", value: uri });
            }
        }
        this.captured = { name: element.name, attributes, children: element.children };
        return true;
    }

    /** The anchor captured, when it is complete, holds a chart (`<c:chart r:id>` or `<cx:chart r:id>`) and has valid names. */
    get result(): ChartXmlElement | null {
        return this.hasChart && this.valid ? this.captured : null;
    }
}

/** The relationship ids an anchor references (attributes marked `rel`), each once, in document order. */
export function anchorRelIds(anchor: ChartXmlElement): string[] {
    const ids = new Set<string>();
    const stack: ChartXmlElement[] = [anchor];
    for (let element = stack.pop(); element; element = stack.pop()) {
        for (const attribute of element.attributes) {
            if (attribute.role === "rel") {
                ids.add(attribute.value);
            }
        }
        for (let i = element.children.length - 1; i >= 0; i--) {
            const child = element.children[i];
            if (typeof child !== "string" && child) {
                stack.push(child);
            }
        }
    }
    return [...ids];
}

/**
 * Writes a chart anchor: its elements, attributes and text as parsed (encoded by XmlStream), each relationship id replaced
 * by the new id of `ids`, each shape id by the next one of `nextShapeId`.
 *
 * @throws Error for a relationship id without its new id (the anchor was not prepared)
 */
export function renderChartAnchor(
    xmlStream: XmlStream,
    element: ChartXmlElement,
    ids: ReadonlyMap<string, string>,
    nextShapeId: () => number
): void {
    xmlStream.openNode(element.name);
    for (const { name, value, role } of element.attributes) {
        if (role === "rel") {
            const id = ids.get(value);
            if (id === undefined) {
                throw new Error(`A chart anchor references the relationship ${JSON.stringify(value)}, which was not prepared`);
            }
            xmlStream.addAttribute(name, id);
        } else if (role === "shapeId") {
            xmlStream.addAttribute(name, nextShapeId());
        } else {
            xmlStream.addAttribute(name, value);
        }
    }
    for (const child of element.children) {
        if (typeof child === "string") {
            xmlStream.writeText(child);
        } else {
            renderChartAnchor(xmlStream, child, ids, nextShapeId);
        }
    }
    xmlStream.closeNode();
}
