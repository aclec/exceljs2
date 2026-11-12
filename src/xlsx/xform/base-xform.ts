import parseSax, { type SaxEvent, type SaxInput, type SaxTag } from "../../utils/parse-sax";
import { parseBoolean } from "../../utils/utils";
import XmlStream, { type XmlAttributes, type XmlValue } from "../../utils/xml-stream";

/** Any xform, whatever its model: the type of child xforms held without their precise type (`parser`, maps). */
export type AnyXform = BaseXform<unknown>;

/**
 * The child xforms of a composite xform, by element name. An entry is the xform itself, or an object holding it under
 * `xform` (with other fields of the owner's choosing, e.g. the model property it fills). {@link BaseXform.reset} resets
 * every entry.
 *
 * Subclasses redeclare `map` with its precise type (a type literal, not an interface, so that it stays assignable to this
 * index signature): `declare map: { readonly t: TextXform; readonly r: RichTextXform };`.
 */
export type XformMap = Readonly<Record<string, AnyXform | { readonly xform: AnyXform }>>;

/**
 * The options of the xforms of an element holding one value (simple/: string, integer, float, date): in an attribute
 * (`<sz val="11"/>` with `attr: "val"`) or as text (`<dc:title>Report</dc:title>` without `attr`).
 */
export interface ValueXformOptions {
    /** The element. */
    readonly tag: string;
    /** The attribute holding the value; absent: the value is the text of the element. */
    readonly attr?: string;
    /** Fixed attributes written on the element. */
    readonly attrs?: XmlAttributes;
}

/**
 * The own entry of a map for a key read from a file (an element name): never an inherited member (`constructor`,
 * `__proto__`, ...), which a plain `map[name]` lookup would return.
 *
 * @example
 * this.parser = ownEntry(this.map, node.name);
 */
export function ownEntry<T>(map: Readonly<Record<string, T>>, key: string): T | undefined {
    return Object.hasOwn(map, key) ? map[key] : undefined;
}

/**
 * Whether a string read from a file is one of the members of a union of string literals, given as the keys of a record
 * (an object literal typed `Record<T, true>` lists every member of `T`, and only them: the compiler checks both).
 *
 * @example
 * const EDIT_AS: Record<ImageEditAs, true> = { twoCell: true, oneCell: true, absolute: true };
 * const editAs = isOneOf(EDIT_AS, node.attributes.editAs) ? node.attributes.editAs : "oneCell";
 */
export function isOneOf<T extends string>(members: Readonly<Record<T, true>>, value: string | undefined): value is T {
    return value !== undefined && Object.hasOwn(members, value);
}

/** Batches of SAX events, as yielded by `parseSax()`. */
export type SaxEventBatches = AsyncIterable<readonly SaxEvent[]>;

/**
 * The base class of the xforms: the transforms between an XML element of an xlsx part and its JSON model.
 *
 * ## The contract
 *
 * Writing a part, for a model `M`:
 * 1. {@link prepare}`(model, options)`: optional, mutates the model so that it is ready to render (e.g. replaces a style
 *    object by its `styleId`, collects shared strings). `options` carries the workbook-wide state of the write.
 * 2. {@link render}`(xmlStream, model)`: writes the element (or nothing, when the model is empty) to the {@link XmlStream}.
 *
 * Reading a part, with a SAX parser:
 * 1. {@link parseOpen}`(node)` for each start tag, {@link parseText}`(text)` for each chunk of text, {@link parseClose}`(name)`
 *    for each end tag. An xform that accepts the element returns `true` from `parseOpen` (some return nothing for their own
 *    tag: the parent decides when to delegate); `parseClose` returns `false` when its own element ends (the model is
 *    complete) and `true` while it still expects more. A composite xform keeps the child it delegates to in `parser` and
 *    forwards the events to it until the child's `parseClose` returns `false`.
 * 2. The result is {@link model}. {@link reconcile}`(model, options)`: optional, resolves what needs the other parts (style
 *    ids, shared string indexes, relationships), mutating the model.
 *
 * The parsed model is untrusted: attribute values are strings from the file (see `SaxTag`). Bounds and validation belong
 * in the xform (see the invariants of CLAUDE.md: linear parsing, row/column/merge/formula limits).
 *
 * ## Typing a subclass
 *
 * - `TModel` is the model rendered and parsed. Narrow the parameters of the overrides to the precise types: `render(xmlStream,
 *   model: M | undefined)` when an absent model renders nothing, `prepare(model: M, options: CellPrepareOptions)` with the
 *   shared option types of `src/types` (method parameters are bivariant, so overrides may narrow them, add optional ones,
 *   or require `options`).
 * - Declare the fields set in the constructor or lazily with `declare` or an initializer; never as a plain field without
 *   initializer that a base constructor sets (ES2022 [[Define]] semantics reset it to `undefined`). Redeclare `map` (and
 *   `model`, when the subclass guarantees more) with `declare`.
 * - `tag` is not declared here: declare it as a getter (`get tag(): string`) or a readonly field, as the subclass needs.
 * - Look child xforms up by element name with {@link ownEntry} (element names come from the file).
 * - Read a child's result through its typed reference (`this.map.t.model`, after checking `this.parser === this.map.t`),
 *   not through `this.parser` (an {@link AnyXform}, whose model is `unknown`).
 *
 * @example
 * // <c:val v="42"/> ⇄ 42
 * class ValXform extends BaseXform<number> {
 *     get tag(): string {
 *         return "c:val";
 *     }
 *     override render(xmlStream: XmlStream, model: number | undefined): void {
 *         if (model !== undefined) xmlStream.leafNode(this.tag, { v: model });
 *     }
 *     override parseOpen(node: SaxTag): boolean {
 *         if (node.name !== this.tag) return false;
 *         this.model = BaseXform.toIntValue(node.attributes.v, 0);
 *         return true;
 *     }
 *     override parseClose(): boolean {
 *         return false;
 *     }
 * }
 *
 * @typeParam TModel - the JSON model of the element
 */
class BaseXform<TModel = unknown> {
    /**
     * The model: the one parsed, once the element is closed (or being built while parsing), or the one an xform is
     * constructed with. `undefined` before anything set it, `null` after {@link reset}.
     */
    declare model: TModel | null | undefined;

    /** The child xforms by element name, for xforms composed of others (see {@link XformMap}). */
    declare map?: XformMap;

    /** While parsing: the child xform that the current events are delegated to. */
    declare parser?: AnyXform;

    // ============================================================
    // Virtual Interface

    /**
     * Prepares (mutates) the model before {@link render}. Default: nothing.
     *
     * @param options - the state of the write shared by the xforms (shared strings, styles, ...)
     */
    prepare(model: NoInfer<TModel>, options?: unknown): void {
        // optional preparation (mutation) of model so it is ready for write
    }

    /**
     * Writes the model as XML. Default: nothing.
     *
     * @param rest - extra arguments some xforms take (`ListXform` passes the index of the item)
     */
    render(xmlStream: XmlStream, model: NoInfer<TModel>, ...rest: unknown[]): void {
        // convert model to xml
    }

    /**
     * Called for each start tag (and empty-element tag) while this xform parses.
     *
     * @returns `true` if the element is accepted (it is this xform's own or one of its descendants)
     */
    parseOpen(node: SaxTag): boolean | void {
        // XML node opened
    }

    /** Called for each chunk of text while this xform parses (a text node may come in several chunks). */
    parseText(text: string): void {
        // chunk of text encountered for current node
    }

    /**
     * Called for each end tag (also of empty-element tags) while this xform parses.
     *
     * @returns `false` when this xform's element is closed (its model is complete), `true` while it expects more. Default:
     *   `false` (an xform without children is done at its first end tag).
     */
    parseClose(name: string): boolean {
        // XML node closed
        return false;
    }

    /**
     * Completes (mutates) the parsed model with what the other parts provide. Default: nothing.
     *
     * @param options - the state of the read shared by the xforms (styles, shared strings, relationships, ...)
     */
    reconcile(model: NoInfer<TModel>, options?: unknown): void {
        // optional post-parse step (opposite to prepare)
    }

    // ============================================================

    /** Forgets the last parse: sets {@link model} to `null`, and resets the children of {@link map}. */
    reset(): void {
        // to make sure parses don't bleed to next iteration
        this.model = null;

        // if we have a map - reset them too
        const { map } = this;
        if (map) {
            for (const entry of Object.values(map)) {
                if (entry instanceof BaseXform) {
                    entry.reset();
                } else if (entry.xform) {
                    entry.xform.reset();
                }
            }
        }
    }

    /**
     * Feeds SAX events to this xform until its element is closed (`parseClose` returns `false`) or the events end.
     *
     * @param saxParser - batches of events, from `parseSax()`
     * @returns {@link model}
     */
    async parse(saxParser: SaxEventBatches): Promise<TModel | null | undefined> {
        for await (const events of saxParser) {
            for (const event of events) {
                if (event.eventType === "opentag") {
                    this.parseOpen(event.value);
                } else if (event.eventType === "text") {
                    this.parseText(event.value);
                } else if (event.eventType === "closetag") {
                    if (!this.parseClose(event.value.name)) {
                        return this.model;
                    }
                }
            }
        }
        return this.model;
    }

    /**
     * Parses an XML document with this xform.
     *
     * @param stream - the document (e.g. a zip entry stream), see `parseSax()`
     * @returns {@link model}
     * @throws Error on malformed or truncated XML, or on whatever the xform rejects
     */
    async parseStream(stream: SaxInput): Promise<TModel | null | undefined> {
        return this.parse(parseSax(stream));
    }

    /**
     * Renders a model to an XML string (without XML declaration, unless the xform writes one).
     *
     * @param model - what the {@link render} of the xform accepts (many accept a missing model and write nothing)
     */
    toXml(model: Parameters<this["render"]>[1]): string {
        const xmlStream = new XmlStream();
        this.render(xmlStream, model);
        return xmlStream.xml;
    }

    // ============================================================
    // Useful Utilities: model value ⇄ attribute string. A `...Attribute` helper returns `undefined` (attribute not written)
    // when the value is undefined or equals the default `dflt`, unless `always`; a `...Value` helper returns `dflt` when
    // the attribute is absent.

    /**
     * @returns the value as a string, or `undefined` when it is undefined or equals `dflt` (then `dflt` if `always`)
     */
    static toAttribute<T extends XmlValue>(value: T | undefined, dflt?: T, always = false): string | T | undefined {
        if (value === undefined) {
            if (always) {
                return dflt;
            }
        } else if (always || value !== dflt) {
            return value.toString();
        }
        return undefined;
    }

    /** {@link toAttribute} for a string value. */
    static toStringAttribute(value: string | undefined, dflt?: string, always = false): string | undefined {
        return BaseXform.toAttribute(value, dflt, always);
    }

    /** @returns the attribute (`undefined` when it is absent) */
    static toStringValue(attr: string | undefined): string | undefined;
    /** @returns the attribute, or `dflt` when it is absent */
    static toStringValue<D>(attr: string | undefined, dflt: D): string | D;
    static toStringValue<D>(attr: string | undefined, dflt?: D): string | D | undefined {
        return attr === undefined ? dflt : attr;
    }

    /** @returns `"1"` / `"0"`, or `undefined` when the value is undefined or equals `dflt` (then `dflt` if `always`) */
    static toBoolAttribute(value: boolean | undefined, dflt?: boolean, always = false): "1" | "0" | boolean | undefined {
        if (value === undefined) {
            if (always) {
                return dflt;
            }
        } else if (always || value !== dflt) {
            return value ? "1" : "0";
        }
        return undefined;
    }

    /**
     * @returns `true` for the xsd:boolean `"1"` or `"true"` (any other value is `false`), or `undefined` when the attribute is
     *   absent
     */
    static toBoolValue(attr: string | undefined): boolean | undefined;
    /** @returns `true` for the xsd:boolean `"1"` or `"true"` (any other value is `false`), or `dflt` when the attribute is absent */
    static toBoolValue<D>(attr: string | undefined, dflt: D): boolean | D;
    static toBoolValue<D>(attr: string | undefined, dflt?: D): boolean | D | undefined {
        return attr === undefined ? dflt : parseBoolean(attr);
    }

    /** {@link toAttribute} for an integer value. */
    static toIntAttribute(value: number | undefined, dflt?: number, always = false): string | number | undefined {
        return BaseXform.toAttribute(value, dflt, always);
    }

    /** @returns `parseInt(attr, 10)` (`NaN` when it is not a number), or `undefined` when the attribute is absent */
    static toIntValue(attr: string | undefined): number | undefined;
    /** @returns `parseInt(attr, 10)` (`NaN` when it is not a number), or `dflt` when the attribute is absent */
    static toIntValue<D>(attr: string | undefined, dflt: D): number | D;
    static toIntValue<D>(attr: string | undefined, dflt?: D): number | D | undefined {
        return attr === undefined ? dflt : parseInt(attr, 10);
    }

    /** {@link toAttribute} for a number value. */
    static toFloatAttribute(value: number | undefined, dflt?: number, always = false): string | number | undefined {
        return BaseXform.toAttribute(value, dflt, always);
    }

    /** @returns `parseFloat(attr)` (`NaN` when it is not a number), or `undefined` when the attribute is absent */
    static toFloatValue(attr: string | undefined): number | undefined;
    /** @returns `parseFloat(attr)` (`NaN` when it is not a number), or `dflt` when the attribute is absent */
    static toFloatValue<D>(attr: string | undefined, dflt: D): number | D;
    static toFloatValue<D>(attr: string | undefined, dflt?: D): number | D | undefined {
        return attr === undefined ? dflt : parseFloat(attr);
    }
}

export default BaseXform;
