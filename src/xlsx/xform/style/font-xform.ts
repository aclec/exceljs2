import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { Font } from "../../../../index";
import type { FontModel } from "../../../types";
import ColorXform from "./color-xform";
import BooleanXform from "../simple/boolean-xform";
import FloatXform from "../simple/float-xform";
import IntegerXform from "../simple/integer-xform";
import StringXform from "../simple/string-xform";
import UnderlineXform from "./underline-xform";

import { isOneOf } from "../../../utils/utils";
import BaseXform, { ownEntry, type AnyXform } from "../base-xform";

/** The options of a {@link FontXform}: the names of its element and of its font name element. */
export interface FontXformOptions {
    /** The font element: `font` (styles part), `rPr` (rich text run). */
    readonly tagName: string;
    /** The font name element: `name` (styles part), `rFont` (rich text run). */
    readonly fontNameTag: string;
}

/** A child element of a font: its xform, which writes and reads one property of the font. */
interface FontEntry {
    /** The xform of the element (the parser of its events). */
    readonly xform: AnyXform;
    /** Writes the element of the property (nothing when the font does not have it). */
    render(xmlStream: XmlStream, model: FontModel): void;
    /** Sets the property from the model of {@link xform}, once its element is parsed (only a truthy value). */
    read(model: FontModel): void;
}

/** The xform of the property `K` of a font: it renders the property, and parses a value of type `V`. */
type PropertyXform<K extends keyof FontModel, V> = BaseXform<V> & {
    render(xmlStream: XmlStream, model: FontModel[K]): void;
};

/**
 * The entry of the property `prop`.
 *
 * @param toProperty - the property for a parsed value (`undefined`: an invalid value, ignored)
 */
function entry<K extends keyof FontModel, V>(prop: K, xform: PropertyXform<K, V>, toProperty: (value: V) => FontModel[K]): FontEntry {
    return {
        xform,
        render(xmlStream: XmlStream, model: FontModel): void {
            xform.render(xmlStream, model[prop]);
        },
        read(model: FontModel): void {
            const value = xform.model;
            if (value) {
                const property = toProperty(value);
                if (property !== undefined) {
                    model[prop] = property;
                }
            }
        },
    };
}

// ECMA-376 ST_VerticalAlignRun
const VERTICAL_ALIGNS: Readonly<Record<NonNullable<Font["vertAlign"]>, true>> = { baseline: true, superscript: true, subscript: true };
// ECMA-376 ST_FontScheme
const SCHEMES: Readonly<Record<Font["scheme"], true>> = { none: true, major: true, minor: true };

const same = <T>(value: T): T => value;

/**
 * A font (CT_Font, or CT_RPrElt for the run of a rich text): one child element per property of {@link FontModel}, in
 * the order of the map (`b`, `i`, `u`, `charset`, `color`, `condense`, `extend`, `family`, `outline`, `vertAlign`,
 * `scheme`, `shadow`, `strike`, `sz`, then the name). Parsing keeps the properties whose value is truthy; a `vertAlign`
 * or a `scheme` that is not one of its ECMA-376 values is ignored.
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
            b: entry("bold", new BooleanXform({ tag: "b", attr: "val" }), same),
            i: entry("italic", new BooleanXform({ tag: "i", attr: "val" }), same),
            u: entry("underline", new UnderlineXform(), same),
            charset: entry("charset", new IntegerXform({ tag: "charset", attr: "val" }), same),
            color: entry("color", new ColorXform(), same),
            condense: entry("condense", new BooleanXform({ tag: "condense", attr: "val" }), same),
            extend: entry("extend", new BooleanXform({ tag: "extend", attr: "val" }), same),
            family: entry("family", new IntegerXform({ tag: "family", attr: "val" }), same),
            outline: entry("outline", new BooleanXform({ tag: "outline", attr: "val" }), same),
            vertAlign: entry("vertAlign", new StringXform({ tag: "vertAlign", attr: "val" }), (value) =>
                isOneOf(VERTICAL_ALIGNS, value) ? value : undefined
            ),
            scheme: entry("scheme", new StringXform({ tag: "scheme", attr: "val" }), (value) =>
                isOneOf(SCHEMES, value) ? value : undefined
            ),
            shadow: entry("shadow", new BooleanXform({ tag: "shadow", attr: "val" }), same),
            strike: entry("strike", new BooleanXform({ tag: "strike", attr: "val" }), same),
            // CT_FontSize: xsd:double (e.g. 10.5)
            sz: entry("size", new FloatXform({ tag: "sz", attr: "val" }), same),
        };
        map[this.options.fontNameTag] = entry("name", new StringXform({ tag: this.options.fontNameTag, attr: "val" }), same);
        this.map = map;
        this.entries = Object.values(map);
    }

    get tag(): string {
        return this.options.tagName;
    }

    override render(xmlStream: XmlStream, model: FontModel): void {
        xmlStream.openNode(this.options.tagName);
        for (const child of this.entries) {
            child.render(xmlStream, model);
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
        const child = ownEntry(this.map, node.name);
        if (child) {
            this.entry = child;
            this.parser = child.xform;
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
            if (this.entry && this.model) {
                this.entry.read(this.model);
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
