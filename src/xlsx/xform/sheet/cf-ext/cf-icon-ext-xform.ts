import type { SaxTag } from "../../../../utils/parse-sax";
import type XmlStream from "../../../../utils/xml-stream";
import type { CfIconModel } from "../../../../types";
import BaseXform from "../../base-xform";
import { iconIdAttribute, isIconSet, validIconSet } from "../cf/icon-sets";

/** `<x14:cfIcon iconSet="3Arrows" iconId="0"/>`: a custom icon of an x14 icon set. */
class CfIconExtXform extends BaseXform<CfIconModel> {
    get tag(): string {
        return "x14:cfIcon";
    }

    /**
     * Writes the icon: its icon set and its `iconId` as given (`0` when absent for `NoIcons`, its only "icon").
     *
     * @throws Error for an `iconSet` that is not an icon set, or an `iconId` that is absent (but for `NoIcons`) or not an
     *   integer from 0 to the number of icons of the set - 1 ([MS-XLSX] CT_CfIcon requires both)
     */
    override render(xmlStream: XmlStream, model: CfIconModel): void {
        // both attributes are required (the default icon set is written too)
        const iconSet = validIconSet(model.iconSet);
        xmlStream.leafNode(this.tag, {
            iconSet,
            iconId: iconIdAttribute(iconSet, model.iconId),
        });
    }

    /**
     * Reads the attributes of the element (the parent delegates `<x14:cfIcon>` only). An icon of an unknown (or absent)
     * icon set is no icon (`NoIcons`, its only icon 0).
     */
    override parseOpen({ attributes }: SaxTag): void {
        const { iconSet } = attributes;
        this.model =
            iconSet !== undefined && isIconSet(iconSet)
                ? { iconSet, iconId: BaseXform.toIntValue(attributes.iconId) }
                : { iconSet: "NoIcons", iconId: 0 };
    }

    /** @returns `false` at the end of `<x14:cfIcon>` */
    override parseClose(name: string): boolean {
        return name !== this.tag;
    }
}

export default CfIconExtXform;
