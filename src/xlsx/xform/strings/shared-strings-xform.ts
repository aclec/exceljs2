import type { SaxTag } from "../../../utils/parse-sax";
import XmlStream from "../../../utils/xml-stream";
import type { RichTextValue, SharedStringLookup, SharedStringRegistry, SharedStringValue, SharedStringsModel } from "../../../types";
import BaseXform from "../base-xform";
import SharedStringXform from "./shared-string-xform";

/**
 * The shared strings table, `xl/sharedStrings.xml` (`<sst>`), and its manager.
 *
 * Writing: the cells add their strings with {@link add} during prepare (each distinct string is stored once, each
 * reference counted), then the table is rendered from {@link model}. Reading: parsing fills {@link model}, and the cells
 * look their strings up with {@link getString} during reconcile.
 *
 * The model is `{ values, count }`, `count` being the number of references to the shared strings in the workbook (cells),
 * `values` its unique strings.
 */
class SharedStringsXform extends BaseXform<SharedStringsModel> implements SharedStringRegistry, SharedStringLookup {
    // set by the constructor, never reset
    declare model: SharedStringsModel;

    /** Index of each plain string already added (keys are cell values: a Map). */
    hash: Map<string, number>;
    /** Index of each rich text already added, by its XML. */
    rich: Map<string, number>;
    /**
     * Index of each rich text object already added: the cells of a loaded workbook that reference the same shared string
     * share its object, and serialising it again for each of them was O(cells x runs).
     */
    richIndexes: WeakMap<RichTextValue, number>;
    private _sharedStringXform?: SharedStringXform;

    /** @param model - the initial table (default: empty) */
    constructor(model?: SharedStringsModel) {
        super();

        this.model = model || {
            values: [],
            count: 0,
        };
        this.hash = new Map();
        this.rich = new Map();
        this.richIndexes = new WeakMap();
    }

    /** The xform of one shared string (created on first use). */
    get sharedStringXform(): SharedStringXform {
        return this._sharedStringXform || (this._sharedStringXform = new SharedStringXform());
    }

    /** The XML of the shared strings part ({@link render} of {@link model}). */
    get xml(): string {
        return this.toXml(this.model);
    }

    /** The unique strings. */
    get values(): SharedStringValue[] {
        return this.model.values;
    }

    /** The number of unique strings. */
    get uniqueCount(): number {
        return this.model.values.length;
    }

    /** The number of references to the strings. */
    get count(): number {
        return this.model.count;
    }

    /** @returns the string at an index (`undefined` when out of range: on read, the index comes from the file) */
    getString(index: number): SharedStringValue | undefined {
        return this.model.values[index];
    }

    /**
     * Adds a reference to a string or rich text.
     *
     * @returns its index in the table
     */
    add(value: SharedStringValue): number {
        return typeof value === "string" ? this.addText(value) : this.addRichText(value);
    }

    /**
     * Adds a reference to a plain string.
     *
     * @returns its index in the table
     */
    addText(value: string): number {
        let index = this.hash.get(value);
        if (index === undefined) {
            index = this.model.values.length;
            this.hash.set(value, index);
            this.model.values.push(value);
        }
        this.model.count++;
        return index;
    }

    /**
     * Adds a reference to a rich text: equal rich texts (same XML) share an index.
     *
     * @returns its index in the table
     */
    addRichText(value: RichTextValue): number {
        let index = this.richIndexes.get(value);
        if (index === undefined) {
            const xml = this.sharedStringXform.toXml(value);
            index = this.rich.get(xml);
            if (index === undefined) {
                index = this.model.values.length;
                this.rich.set(xml, index);
                this.model.values.push(value);
            }
            this.richIndexes.set(value, index);
        }
        this.model.count++;
        return index;
    }

    // <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    // <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="<%=totalRefs%>" uniqueCount="<%=count%>">
    //   <si><t><%=text%></t></si>
    //   <si><r><rPr></rPr><t></t></r></si>
    // </sst>

    override render(xmlStream: XmlStream, model: SharedStringsModel): void {
        xmlStream.openXml(XmlStream.StdDocAttributes);

        xmlStream.openNode("sst", {
            xmlns: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            count: model.count,
            uniqueCount: model.values.length,
        });

        const sx = this.sharedStringXform;
        for (const value of model.values) {
            sx.render(xmlStream, value);
        }
        xmlStream.closeNode();
    }

    /**
     * @returns `true` for `<sst>` and its strings
     * @throws Error on any other element
     */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case "sst":
                return true;
            case "si":
                this.parser = this.sharedStringXform;
                this.parser.parseOpen(node);
                return true;
            default:
                throw new Error(`Unexpected xml node in parseOpen: ${JSON.stringify(node)}`);
        }
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    /** @throws Error on an unexpected end tag */
    override parseClose(name: string): boolean {
        if (this.parser) {
            if (!this.parser.parseClose(name)) {
                // the shared string closed: its model is complete (at least "")
                const value = this.sharedStringXform.model;
                if (value !== null && value !== undefined) {
                    this.model.values.push(value);
                    this.model.count++;
                }
                this.parser = undefined;
            }
            return true;
        }
        switch (name) {
            case "sst":
                return false;
            default:
                throw new Error(`Unexpected xml node in parseClose: ${name}`);
        }
    }
}

export default SharedStringsXform;
