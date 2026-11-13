import type { SaxTag } from "../../../utils/parse-sax";
import XmlStream from "../../../utils/xml-stream";
import type { ChartXmlAttribute, ChartXmlElement, ChartsheetPartModel } from "../../../types";
import BaseXform from "../base-xform";
import { renderChartAnchor } from "../drawing/chart-anchor";

const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
// the relationships namespace, transitional and strict: an attribute of either is a relationship id
const RELATIONSHIPS_NS: ReadonlySet<string> = new Set([
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
]);
// the namespace of the `xml:` prefix, declared implicitly
const XML_NS = "http://www.w3.org/XML/1998/namespace";

// A character that cannot be in an XML name: saxes already rejects such names, checked again as the names are written
// back verbatim (one character class, unanchored test: no backtracking)
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
 * A chartsheet part, `xl/chartsheets/sheetN.xml` (CT_Chartsheet: sheetPr, sheetViews, sheetProtection, customSheetViews,
 * pageMargins, pageSetup, headerFooter, drawing, legacyDrawing, legacyDrawingHF, drawingHF, picture, webPublishItems,
 * extLst), carried opaque: parsing captures the `<chartsheet>` element as a tree of elements, attributes and text, which
 * {@link render} writes back as parsed (through XmlStream, which encodes the values and the text). Its namespace
 * declarations are attributes like the others: written back where they were.
 *
 * Only the drawing (and through it the chart) is carried with a chartsheet: the elements that reference another part by
 * a relationship id are left out (`<legacyDrawing>`, `<legacyDrawingHF>`, `<drawingHF>`, `<picture>`, or any other
 * element with an attribute of the relationships namespace), except `<pageSetup>`, which keeps its attributes but the
 * relationship id of its printer settings (optional in CT_CsPageSetup). The `r:id` of the `<drawing>` (a child of
 * `<chartsheet>`, exactly one) is marked `rel`, to be replaced on write.
 *
 * The part is refused (no model) when its root is not a `<chartsheet>` of SpreadsheetML (transitional), when a prefix is
 * not declared, or without one `<drawing r:id>`. Linear in the size of the part: the namespaces are a stack per prefix
 * (pushed and popped, never copied), the subtree of an element left out is skipped; the depth is bounded by the SAX
 * parser (parse-sax MAX_DEPTH).
 */
class ChartsheetXform extends BaseXform<ChartsheetPartModel> {
    private stack: OpenElement[] = [];
    private bindings = new Map<string, string[]>();
    // the depth inside an element left out (0: not in one)
    private skipDepth = 0;
    private valid = true;
    private drawingRelId: string | undefined;
    private drawings = 0;
    private done = false;

    /**
     * Writes a chartsheet: its `<chartsheet>` element as parsed, the relationship id of its `<drawing>` replaced by
     * `drawingRId`.
     */
    override render(xmlStream: XmlStream, model: ChartsheetPartModel & { readonly drawingRId: string }): void {
        xmlStream.openXml(XmlStream.StdDocAttributes);
        renderChartAnchor(xmlStream, model.sheet, new Map([[model.drawingRelId, model.drawingRId]]), () => {
            // never called: no attribute of a chartsheet is marked shapeId
            throw new Error("A chartsheet has no shape id");
        });
    }

    override reset(): void {
        super.reset();
        this.stack = [];
        this.bindings = new Map();
        this.skipDepth = 0;
        this.valid = true;
        this.drawingRelId = undefined;
        this.drawings = 0;
        this.done = false;
    }

    // the URI of a prefix in scope ("" for the default namespace), undefined when it is not declared
    private lookup(prefix: string): string | undefined {
        const uris = this.bindings.get(prefix);
        if (uris && uris.length) {
            return uris[uris.length - 1];
        }
        return prefix === "xml" ? XML_NS : undefined;
    }

    /** @returns `true`: every element of the part is accepted */
    override parseOpen(node: SaxTag): boolean {
        if (this.done) {
            return true;
        }
        if (this.skipDepth) {
            this.skipDepth++;
            return true;
        }
        const depth = this.stack.length;
        const entries = Object.entries(node.attributes);
        // the declarations first: they apply to the element and its attributes
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
        if (NOT_NAME_CHAR.test(node.name) || (elementPrefix && elementNs === undefined)) {
            this.valid = false;
        }
        if (depth === 0 && (elementLocal !== "chartsheet" || elementNs !== MAIN_NS)) {
            this.valid = false;
        }
        const isDrawing = depth === 1 && elementLocal === "drawing" && elementNs === MAIN_NS;
        const isPageSetup = elementLocal === "pageSetup" && elementNs === MAIN_NS;

        const attributes: ChartXmlAttribute[] = [];
        let skip = false;
        for (const [name, value] of entries) {
            if (NOT_NAME_CHAR.test(name)) {
                this.valid = false;
            }
            const [prefix, local] = splitName(name);
            if (!prefix || prefix === "xmlns") {
                attributes.push({ name, value });
                continue;
            }
            const ns = this.lookup(prefix);
            if (ns === undefined) {
                this.valid = false;
            } else if (RELATIONSHIPS_NS.has(ns)) {
                if (isDrawing && local === "id" && this.drawingRelId === undefined) {
                    this.drawingRelId = value;
                    attributes.push({ name, value, role: "rel" });
                } else if (!isDrawing && !isPageSetup) {
                    // an element referencing a part that is not carried
                    skip = true;
                }
                continue;
            }
            attributes.push({ name, value });
        }
        if (isDrawing) {
            this.drawings++;
        }
        if (skip && depth > 0) {
            // undone as the element is not pushed
            for (const prefix of declared) {
                this.bindings.get(prefix)?.pop();
            }
            this.skipDepth = 1;
            return true;
        }
        this.stack.push({ name: node.name, attributes, children: [], declared });
        return true;
    }

    override parseText(text: string): void {
        if (this.done || this.skipDepth) {
            return;
        }
        const element = this.stack[this.stack.length - 1];
        if (element) {
            element.children.push(text);
        }
    }

    /** @returns `false` once `<chartsheet>` (the root) ends */
    override parseClose(): boolean {
        if (this.done) {
            return false;
        }
        if (this.skipDepth) {
            this.skipDepth--;
            return true;
        }
        const element = this.stack.pop();
        if (!element) {
            return true;
        }
        for (const prefix of element.declared) {
            this.bindings.get(prefix)?.pop();
        }
        const captured: ChartXmlElement = { name: element.name, attributes: element.attributes, children: element.children };
        const parent = this.stack[this.stack.length - 1];
        if (parent) {
            parent.children.push(captured);
            return true;
        }
        this.done = true;
        if (this.valid && this.drawings === 1 && this.drawingRelId !== undefined) {
            this.model = { sheet: captured, drawingRelId: this.drawingRelId };
        }
        return false;
    }
}

export default ChartsheetXform;
