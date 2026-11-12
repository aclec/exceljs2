import BaseXform from "./base-xform";
import XmlStream, { type XmlAttributes } from "../../utils/xml-stream";

/**
 * A static XML element: its tag, attributes (`$`), children (`c`) and text (`t`).
 *
 * @example
 * const model: StaticModel = { tag: "name", $: { attr: "value" }, c: [{ tag: "child" }], t: "some text" };
 */
export interface StaticModel {
    readonly tag: string;
    readonly $?: XmlAttributes;
    readonly c?: readonly StaticModel[];
    readonly t?: string;
}

function build(xmlStream: XmlStream, model: StaticModel): void {
    xmlStream.openNode(model.tag, model.$);
    if (model.c) {
        for (const child of model.c) {
            build(xmlStream, child);
        }
    }
    if (model.t) {
        xmlStream.writeText(model.t);
    }
    xmlStream.closeNode();
}

/**
 * An optimisation for static (unimportant and unchanging) XML: renders a fixed element, built once, and skips it when
 * parsing.
 *
 * It is stateless, apart from its static model, and so can be used as a singleton. Being stateless, it only tracks the
 * end of its root element while parsing. Known issues: `parseOpen` always returns `true` (the parent must know when to
 * start using this xform), and if the root tag is recursive, the parsing will behave unpredictably.
 */
class StaticXform extends BaseXform<never> {
    /** The element rendered. */
    private readonly _model: StaticModel;
    /** Its XML, built on first render. */
    private _xml?: string;

    /** @param model - the element, rendered as is */
    constructor(model: StaticModel) {
        super();
        this._model = model;
    }

    /** Writes the static element (the `model` argument is ignored). */
    override render(xmlStream: XmlStream): void {
        if (!this._xml) {
            const stream = new XmlStream();
            build(stream, this._model);
            this._xml = stream.xml;
        }
        xmlStream.writeXml(this._xml);
    }

    /** @returns `true`: every element is accepted (see the class documentation) */
    override parseOpen(): boolean {
        return true;
    }

    override parseText(): void {}

    /** @returns `false` at the end tag of the static element */
    override parseClose(name: string): boolean {
        return name !== this._model.tag;
    }
}

export default StaticXform;
