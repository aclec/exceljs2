import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { FontModel } from "../../../types";
import ColorXform from "./color-xform";
import BooleanXform from "../simple/boolean-xform";
import FloatXform from "../simple/float-xform";
import IntegerXform from "../simple/integer-xform";
import StringXform from "../simple/string-xform";
import UnderlineXform from "./underline-xform";

import BaseXform, { ownEntry, type AnyXform } from "../base-xform";

/** The options of a {@link FontXform}: the names of its element and of its font name element. */
export interface FontXformOptions {
    /** The font element: `font` (styles part), `rPr` (rich text run). */
    readonly tagName: string;
    /** The font name element: `name` (styles part), `rFont` (rich text run). */
    readonly fontNameTag: string;
}

/** A child element of a font: the model property it holds, and its xform (whose model is that property). */
interface FontEntry {
    readonly prop: keyof FontModel;
    readonly xform: AnyXform;
}

/**
 * A font (CT_Font, or CT_RPrElt for the run of a rich text): one child element per property of {@link FontModel}, in
 * the order of the map (`b`, `i`, `u`, `charset`, `color`, `condense`, `extend`, `family`, `outline`, `vertAlign`,
 * `scheme`, `shadow`, `strike`, `sz`, then the name). Parsing keeps the properties whose value is truthy.
 */
class FontXform extends BaseXform<FontModel> {
    /** The options of the font of the styles part. */
    static readonly OPTIONS: FontXformOptions = {
        tagName: "font",
        fontNameTag: "name",
    };

    /** The element names. */
    readonly options: FontXformOptions;

    declare map: Readonly<Record<string, FontEntry>>;

    /** The entries of {@link map}, in render order. */
    private readonly entries: readonly FontEntry[];

    /** While parsing: the entry of the child being parsed. */
    private entry?: FontEntry;

    constructor(options?: FontXformOptions) {
        super();

        this.options = options || FontXform.OPTIONS;

        const map: Record<string, FontEntry> = {
            b: { prop: "bold", xform: new BooleanXform({ tag: "b", attr: "val" }) },
            i: { prop: "italic", xform: new BooleanXform({ tag: "i", attr: "val" }) },
            u: { prop: "underline", xform: new UnderlineXform() },
            charset: { prop: "charset", xform: new IntegerXform({ tag: "charset", attr: "val" }) },
            color: { prop: "color", xform: new ColorXform() },
            condense: { prop: "condense", xform: new BooleanXform({ tag: "condense", attr: "val" }) },
            extend: { prop: "extend", xform: new BooleanXform({ tag: "extend", attr: "val" }) },
            family: { prop: "family", xform: new IntegerXform({ tag: "family", attr: "val" }) },
            outline: { prop: "outline", xform: new BooleanXform({ tag: "outline", attr: "val" }) },
            vertAlign: { prop: "vertAlign", xform: new StringXform({ tag: "vertAlign", attr: "val" }) },
            scheme: { prop: "scheme", xform: new StringXform({ tag: "scheme", attr: "val" }) },
            shadow: { prop: "shadow", xform: new BooleanXform({ tag: "shadow", attr: "val" }) },
            strike: { prop: "strike", xform: new BooleanXform({ tag: "strike", attr: "val" }) },
            // CT_FontSize: xsd:double (e.g. 10.5)
            sz: { prop: "size", xform: new FloatXform({ tag: "sz", attr: "val" }) },
        };
        map[this.options.fontNameTag] = {
            prop: "name",
            xform: new StringXform({ tag: this.options.fontNameTag, attr: "val" }),
        };
        this.map = map;
        this.entries = Object.values(map);
    }

    get tag(): string {
        return this.options.tagName;
    }

    override render(xmlStream: XmlStream, model: FontModel): void {
        const { entries } = this;

        xmlStream.openNode(this.options.tagName);
        for (let i = 0; i < entries.length; i++) {
            const entry = entries[i]!;
            entry.xform.render(xmlStream, model[entry.prop]);
        }
        xmlStream.closeNode();
    }

    /**
     * @returns `true` for the font element (and its descendants while a child is parsed), `false` for an unknown
     *   element; for a property element, what its xform returns (nothing for most)
     */
    override parseOpen(node: SaxTag): boolean | void {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        const entry = ownEntry(this.map, node.name);
        if (entry) {
            this.entry = entry;
            this.parser = entry.xform;
            return this.parser.parseOpen(node);
        }
        switch (node.name) {
            case this.options.tagName:
                this.model = {};
                return true;
            default:
                return false;
        }
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    override parseClose(name: string): boolean {
        if (this.parser && !this.parser.parseClose(name)) {
            const value = this.parser.model;
            if (value && this.entry && this.model) {
                // each xform of the map parses the type of the property it is mapped to
                (this.model as Record<keyof FontModel, unknown>)[this.entry.prop] = value;
            }
            this.parser = undefined;
            this.entry = undefined;
            return true;
        }
        switch (name) {
            case this.options.tagName:
                return false;
            default:
                return true;
        }
    }
}

export default FontXform;
