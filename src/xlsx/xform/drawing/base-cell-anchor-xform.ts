import type { SaxTag } from "../../../utils/parse-sax";
import type { ImageEditAs } from "../../../../index";
import type { DrawingAnchorModel, DrawingPictureModel, DrawingReconcileOptions, MediaModel } from "../../../types";
import { isOneOf } from "../../../utils/utils";
import BaseXform, { ownEntry, type AnyXform } from "../base-xform";

// ECMA-376 ST_EditAs
const EDIT_AS: Readonly<Record<ImageEditAs, true>> = { twoCell: true, oneCell: true, absolute: true };

/**
 * The base of the anchors of a drawing (two-cell and one-cell): parsing of the element and its children, and the
 * resolution of the medium of the picture. A subclass provides the `tag`, the `map` of the children, `render` and
 * `parseClose` (which builds the model from {@link parsedEditAs} and the children).
 */
abstract class BaseCellAnchorXform extends BaseXform<DrawingAnchorModel> {
    /** The element. */
    abstract get tag(): string;

    declare map: Readonly<Record<string, AnyXform>>;

    /** While parsing: the `editAs` attribute of the anchor (`oneCell` when it is absent or invalid). */
    protected parsedEditAs: ImageEditAs = "oneCell";

    /** @returns `true`: every element inside the anchor is accepted (unknown ones are skipped) */
    override parseOpen(node: SaxTag): boolean {
        if (this.parser) {
            this.parser.parseOpen(node);
            return true;
        }
        switch (node.name) {
            case this.tag: {
                this.reset();
                // the model is built when the element closes; a missing or invalid editAs is the default of the
                // anchor: twoCell for a <xdr:twoCellAnchor> (CT_TwoCellAnchor), oneCell for the others (no attribute)
                const { editAs } = node.attributes;
                this.parsedEditAs = isOneOf(EDIT_AS, editAs) ? editAs : this.tag === "xdr:twoCellAnchor" ? "twoCell" : "oneCell";
                break;
            }
            default: {
                const parser = ownEntry(this.map, node.name);
                this.parser = parser;
                if (parser) {
                    parser.parseOpen(node);
                }
                break;
            }
        }
        return true;
    }

    override parseText(text: string): void {
        if (this.parser) {
            this.parser.parseText(text);
        }
    }

    /**
     * @returns the workbook medium of a parsed picture: the image its relationship targets (`.../media/<name>.<ext>`),
     *   found by name in `options.mediaIndex`
     */
    reconcilePicture(model: DrawingPictureModel | null | undefined, options: DrawingReconcileOptions): MediaModel | undefined {
        if (model && model.rId) {
            const rel = options.rels && ownEntry(options.rels, model.rId);
            const target = rel && rel.Target;
            // take the part after the last '/media/' without a backtracking regex
            // (/.*\/media\/(.+)/ was cubic on Targets like '/media//media//media/...')
            const mediaPos = target ? target.lastIndexOf("/media/") : -1;
            const match = target && mediaPos !== -1 && target.slice(mediaPos + 7).match(/^.+[.][a-zA-Z]{3,4}/);
            if (match) {
                const [name] = match;
                const mediaId = ownEntry(options.mediaIndex, name);
                return mediaId === undefined ? undefined : options.media[mediaId];
            }
        }
        return undefined;
    }
}

export default BaseCellAnchorXform;
