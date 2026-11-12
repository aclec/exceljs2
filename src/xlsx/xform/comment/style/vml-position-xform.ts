import type { SaxTag } from "../../../../utils/parse-sax";
import type XmlStream from "../../../../utils/xml-stream";
import BaseXform from "../../base-xform";
import type { VmlTagOptions } from "./vml-protection-xform";

/** The parsed model of a {@link VmlPositionXform}: `{ [tag]: true }` when its element is present. */
export type VmlPositionModel = Record<string, true>;

/**
 * A positioning flag of a VML note, `<x:MoveWithCells/>` / `<x:SizeWithCells/>`: written according to the `editAs` of
 * the note (see `render`).
 */
class VmlPositionXform extends BaseXform<VmlPositionModel> {
    private readonly _model: VmlTagOptions;

    constructor(model: VmlTagOptions) {
        super();
        this._model = model;
    }

    get tag(): string {
        return this._model && this._model.tag;
    }

    /**
     * Writes the empty element when `model` is `type[2]` (`absolute`), or for `x:SizeWithCells` when it is `type[1]`
     * (`oneCells`).
     *
     * @param model - the `editAs` of the note (any other value writes nothing)
     * @param type - the editAs values: `["twoCells", "oneCells", "absolute"]`
     */
    override render(xmlStream: XmlStream, model: string | VmlPositionModel | undefined, type: readonly string[]): void {
        if (model === type[2]) {
            xmlStream.leafNode(this.tag);
        } else if (this.tag === "x:SizeWithCells" && model === type[1]) {
            xmlStream.leafNode(this.tag);
        }
    }

    /** @returns `true` for its element */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case this.tag:
                this.model = { [this.tag]: true };
                return true;
            default:
                return false;
        }
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default VmlPositionXform;
