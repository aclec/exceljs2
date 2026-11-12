import type { SaxTag } from "../../../../utils/parse-sax";
import type XmlStream from "../../../../utils/xml-stream";
import type { CfvoModel, IconSetRuleModel } from "../../../../types";
import BaseXform from "../../base-xform";
import { parseOptionalBoolean } from "../../../../utils/utils";
import CompositeXform from "../../composite-xform";

import CfvoXform from "./cfvo-xform";
import { DEFAULT_ICON_SET, iconSetAttribute, isIconSet } from "./icon-sets";

/**
 * The model of `<iconSet>` (merged into the rule). The parsed `iconSet` is `"3TrafficLights1"` (the default) when absent
 * or unknown.
 */
export interface IconSetModel {
    iconSet: NonNullable<IconSetRuleModel["iconSet"]>;
    reverse: boolean | undefined;
    showValue: boolean | undefined;
    cfvo: CfvoModel[];
}

/** `<iconSet>` of an `iconSet` rule (Excel 2007 icon sets): the icon set and its thresholds. */
class IconSetXform extends CompositeXform<IconSetModel> {
    /** The xform of the thresholds. */
    readonly cfvoXform = new CfvoXform();

    declare map: { readonly cfvo: CfvoXform };

    constructor() {
        super();

        this.map = {
            cfvo: this.cfvoXform,
        };
    }

    get tag(): string {
        return "iconSet";
    }

    /**
     * Writes the icon set (`iconSet` is not written when it is absent or the default, `"3TrafficLights1"`).
     *
     * @param model - the rule
     * @throws Error when the rule has no `cfvo` (CT_IconSet requires them), or an `iconSet` that is not an icon set of
     *   ECMA-376 (the x14 ones are written by the x14 rule)
     */
    override render(xmlStream: XmlStream, model: Pick<IconSetRuleModel, "iconSet" | "reverse" | "showValue" | "cfvo">): void {
        const { cfvo } = model;
        if (!cfvo) {
            throw new Error("An iconSet conditional formatting rule needs cfvo");
        }
        xmlStream.openNode(this.tag, {
            iconSet: iconSetAttribute(model.iconSet, false),
            reverse: BaseXform.toBoolAttribute(model.reverse, false),
            showValue: BaseXform.toBoolAttribute(model.showValue, true),
        });

        for (const item of cfvo) {
            this.cfvoXform.render(xmlStream, item);
        }

        xmlStream.closeNode();
    }

    override createNewModel({ attributes }: SaxTag): IconSetModel {
        const { iconSet } = attributes;
        return {
            iconSet: iconSet !== undefined && isIconSet(iconSet) ? iconSet : DEFAULT_ICON_SET,
            reverse: parseOptionalBoolean(attributes.reverse),
            showValue: parseOptionalBoolean(attributes.showValue),
            cfvo: [],
        };
    }

    override onParserClose(name: string, model: IconSetModel): void {
        // the cfvo model is set by its start tag
        const cfvo = this.cfvoXform.model;
        if (name === "cfvo" && cfvo) {
            model.cfvo.push(cfvo);
        }
    }
}

export default IconSetXform;
