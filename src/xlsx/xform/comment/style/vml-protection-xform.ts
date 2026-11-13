import type { SaxTag } from "../../../../utils/parse-sax";
import type XmlStream from "../../../../utils/xml-stream";
import BaseXform from "../../base-xform";

/** The options of a {@link VmlProtectionXform}. */
export interface VmlTagOptions {
    /** The element, e.g. `x:Locked`. */
    readonly tag: string;
}

/**
 * A protection flag of a VML note, `<x:Locked>` / `<x:LockText>`, whose text is `True` or `False`. Parsing keeps the
 * text in {@link text} (the last chunk of text), not in the model.
 */
class VmlProtectionXform extends BaseXform<string> {
    private readonly _model: VmlTagOptions;

    /** While parsing: the text of the element (`""` until a text is read). */
    declare text?: string;

    constructor(model: VmlTagOptions) {
        super();
        this._model = model;
    }

    get tag(): string {
        return this._model && this._model.tag;
    }

    /** Writes the element with the text `model` (an empty element when it is undefined). */
    override render(xmlStream: XmlStream, model: string | undefined): void {
        xmlStream.leafNode(this.tag, undefined, model);
    }

    /** @returns `true` for its element */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case this.tag:
                this.text = "";
                return true;
            default:
                return false;
        }
    }

    override parseText(text: string): void {
        this.text = text;
    }

    override parseClose(): boolean {
        return false;
    }
}

export default VmlProtectionXform;
