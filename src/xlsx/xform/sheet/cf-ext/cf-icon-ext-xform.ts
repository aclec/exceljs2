import type { SaxTag } from "../../../../utils/parse-sax";
import type XmlStream from "../../../../utils/xml-stream";
import type { CfIconModel } from "../../../../types";
import BaseXform from "../../base-xform";

/** `<x14:cfIcon iconSet="3Arrows" iconId="0"/>`: a custom icon of an x14 icon set. */
class CfIconExtXform extends BaseXform<CfIconModel> {
    get tag(): string {
        return "x14:cfIcon";
    }

    override render(xmlStream: XmlStream, model: CfIconModel): void {
        xmlStream.leafNode(this.tag, {
            iconSet: model.iconSet,
            iconId: model.iconId,
        });
    }

    /** Reads the attributes of the element (the parent delegates `<x14:cfIcon>` only). */
    override parseOpen({ attributes }: SaxTag): void {
        this.model = {
            // unvalidated ST_IconSetType
            iconSet: attributes.iconSet as CfIconModel["iconSet"],
            iconId: BaseXform.toIntValue(attributes.iconId),
        };
    }

    /** @returns `false` at the end of `<x14:cfIcon>` */
    override parseClose(name: string): boolean {
        return name !== this.tag;
    }
}

export default CfIconExtXform;
