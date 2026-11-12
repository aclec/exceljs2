import utils from "./utils";

// utils.ts is not typed yet: its xmlEncode is inferred as (any) => any. Call it directly once it is typed (eslint
// no-unnecessary-type-assertion will flag this assertion then).
const xmlEncode = (text: string): string => utils.xmlEncode(text) as string;

/**
 * A value written as an XML attribute or as text: converted with `toString()`, then XML-encoded (special characters escaped,
 * characters invalid in XML 1.0 removed).
 */
export type XmlValue = string | number | boolean;

/**
 * Attributes of an XML element, in rendering order. An `undefined` value is skipped (the attribute is not written), which
 * lets callers build the object with optional attributes.
 */
export type XmlAttributes = Readonly<Record<string, XmlValue | undefined>>;

/** State saved by {@link XmlStream.addRollback}, restored by {@link XmlStream.rollback}. */
interface Rollback {
    readonly xml: number;
    readonly stack: number;
    readonly leaf: boolean;
    readonly open: boolean;
}

// Every attribute of an element as one string (` name="value"`...): one array entry per element instead of one per
// attribute keeps the final join() cheap. Own enumerable keys only, in insertion order (as Object.keys).
function renderAttributes(attributes: XmlAttributes): string {
    let xml = "";
    const names = Object.keys(attributes);
    for (let i = 0; i < names.length; i++) {
        const name = names[i]!;
        const value = attributes[name];
        if (value !== undefined) {
            xml += ` ${name}="${xmlEncode(value.toString())}"`;
        }
    }
    return xml;
}

/**
 * Builds an XML document in memory, element by element: the writer used by every xform `render()`.
 *
 * Elements are opened with {@link openNode} and closed with {@link closeNode}; an element closed without content is written
 * as an empty-element tag (`<a/>`). Attribute values and text are XML-encoded here, so callers pass raw values. Names
 * (element and attribute) are written as given: they must be valid XML names chosen by the code, never data.
 *
 * {@link addRollback} / {@link commit} / {@link rollback} let a renderer discard what it wrote since a point (e.g. an
 * element that ends up without any attribute).
 *
 * @example
 * const xmlStream = new XmlStream();
 * xmlStream.openXml(XmlStream.StdDocAttributes);
 * xmlStream.openNode("sst", { count: 1 });
 * xmlStream.leafNode("t", undefined, "a < b");
 * xmlStream.closeNode();
 * xmlStream.xml; // '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst count="1"><t>a &lt; b</t></sst>'
 */
class XmlStream {
    /** The attributes of the standard XML declaration of the parts of an xlsx file (`version`, `encoding`, `standalone`). */
    static readonly StdDocAttributes: XmlAttributes = {
        version: "1.0",
        encoding: "UTF-8",
        standalone: "yes",
    };

    /** The document so far, as fragments joined by {@link xml}. */
    private _xml: string[] = [];
    /** The names of the open elements, innermost last. */
    private _stack: string[] = [];
    private _rollbacks: Rollback[] = [];
    /** True while the innermost element has no content yet (it is closed as `<name/>`). */
    leaf = false;
    /** True while the start tag of the innermost element is not terminated yet (`>` not written): attributes can be added. */
    open = false;

    /** The name of the innermost open element, or `undefined` when none is open. */
    get tos(): string | undefined {
        return this._stack.length ? this._stack[this._stack.length - 1] : undefined;
    }

    /** A position in the output, which grows whenever something is written: compare two values to know whether anything was. */
    get cursor(): number {
        return this._xml.length;
    }

    /**
     * Writes the XML declaration (`<?xml ...?>` and a line feed).
     *
     * @param docAttributes - usually {@link XmlStream.StdDocAttributes}
     */
    openXml(docAttributes?: XmlAttributes): void {
        const xml = this._xml;
        // <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        xml.push("<?xml");
        if (docAttributes) {
            xml.push(renderAttributes(docAttributes));
        }
        xml.push("?>\n");
    }

    /**
     * Opens an element: its start tag stays open, so attributes can still be added with {@link addAttribute} until content
     * (a child element or text) is written.
     *
     * @param name - the element name, written as is
     * @param attributes - its attributes; `undefined` values are skipped
     */
    openNode(name: string, attributes?: XmlAttributes): void {
        // the start tag as one fragment, with the ">" of the parent's start tag if still open
        const start = attributes ? `<${name}${renderAttributes(attributes)}` : `<${name}`;
        this._xml.push(this.open && this._stack.length ? `>${start}` : start);
        this._stack.push(name);
        this.leaf = true;
        this.open = true;
    }

    /**
     * Adds an attribute to the element just opened.
     *
     * @param value - `undefined` writes nothing
     * @throws Error if the start tag of the current element is already terminated (content was written)
     */
    addAttribute(name: string, value: XmlValue | undefined): void {
        if (!this.open) {
            throw new Error("Cannot write attributes to node if it is not open");
        }
        if (value !== undefined) {
            this._xml.push(` ${name}="${xmlEncode(value.toString())}"`);
        }
    }

    /**
     * Adds attributes to the element just opened; `undefined` values are skipped.
     *
     * @throws Error if the start tag of the current element is already terminated (content was written)
     */
    addAttributes(attrs: XmlAttributes | undefined): void {
        if (!this.open) {
            throw new Error("Cannot write attributes to node if it is not open");
        }
        if (attrs) {
            this._xml.push(renderAttributes(attrs));
        }
    }

    /** Writes text content in the current element (XML-encoded). */
    writeText(text: XmlValue): void {
        const xml = this._xml;
        if (this.open) {
            xml.push(">");
            this.open = false;
        }
        this.leaf = false;
        xml.push(xmlEncode(text.toString()));
    }

    /**
     * Writes a fragment of XML as is (not encoded) in the current element.
     *
     * @param xml - well-formed XML, built by an xform (never data)
     */
    writeXml(xml: string): void {
        if (this.open) {
            this._xml.push(">");
            this.open = false;
        }
        this.leaf = false;
        this._xml.push(xml);
    }

    /**
     * Closes the innermost open element: `/>` if it has no content, else its end tag. Nothing is checked: closing more
     * elements than were opened writes a broken document.
     */
    closeNode(): void {
        const node = this._stack.pop();
        if (this.leaf) {
            this._xml.push("/>");
        } else {
            this._xml.push(`</${node ?? ""}>`);
        }
        this.open = false;
        this.leaf = false;
    }

    /**
     * Writes a whole element: {@link openNode}, then the text if any, then {@link closeNode}.
     *
     * @param text - its text content; `undefined` writes an empty element (`0` and `""` are written)
     */
    leafNode(name: string, attributes?: XmlAttributes, text?: XmlValue): void {
        // the whole element as one fragment: the same output and state as openNode, writeText, closeNode
        const start = attributes ? `<${name}${renderAttributes(attributes)}` : `<${name}`;
        // zeros need to be written
        const element = text === undefined ? `${start}/>` : `${start}>${xmlEncode(text.toString())}</${name}>`;
        this._xml.push(this.open && this._stack.length ? `>${element}` : element);
        this.open = false;
        this.leaf = false;
    }

    /** Closes every open element. */
    closeAll(): void {
        while (this._stack.length) {
            this.closeNode();
        }
    }

    /**
     * Saves the current state, to be restored by {@link rollback} or dropped by {@link commit}. Rollback points nest.
     *
     * @returns the {@link cursor} at this point
     */
    addRollback(): number {
        this._rollbacks.push({
            xml: this._xml.length,
            stack: this._stack.length,
            leaf: this.leaf,
            open: this.open,
        });
        return this.cursor;
    }

    /** Keeps what was written since the last {@link addRollback} (drops that rollback point). */
    commit(): void {
        this._rollbacks.pop();
    }

    /**
     * Discards what was written since the last {@link addRollback}, including the elements opened since then.
     *
     * @throws TypeError if there is no rollback point
     */
    rollback(): void {
        const r = this._rollbacks.pop()!;
        if (this._xml.length > r.xml) {
            this._xml.splice(r.xml, this._xml.length - r.xml);
        }
        if (this._stack.length > r.stack) {
            this._stack.splice(r.stack, this._stack.length - r.stack);
        }
        this.leaf = r.leaf;
        this.open = r.open;
    }

    /** The document: closes every open element first (so it is final), then returns the XML. */
    get xml(): string {
        this.closeAll();
        return this._xml.join("");
    }
}

export default XmlStream;
