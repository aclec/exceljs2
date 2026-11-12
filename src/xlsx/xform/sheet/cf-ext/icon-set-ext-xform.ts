import type { SaxTag } from "../../../../utils/parse-sax";
import type XmlStream from "../../../../utils/xml-stream";
import type { CfIconModel, CfvoModel, IconSetRuleModel } from "../../../../types";
import { parseBoolean } from "../../../../utils/utils";
import BaseXform from "../../base-xform";
import { parseOptionalBoolean } from "../../../../utils/utils";
import CompositeXform from "../../composite-xform";

import CfvoExtXform from "./cfvo-ext-xform";
import CfIconExtXform from "./cf-icon-ext-xform";
import { DEFAULT_ICON_SET, iconSetAttribute, isIconSet } from "../cf/icon-sets";

/**
 * The model of `<x14:iconSet>` (merged into the rule). The parsed `iconSet` is `"3TrafficLights1"` (the default) when
 * absent or unknown.
 */
export interface IconSetExtModel {
    cfvo: CfvoModel[];
    iconSet: NonNullable<IconSetRuleModel["iconSet"]>;
    reverse: boolean;
    showValue: boolean;
    /** The `custom` attribute (an xsd:boolean), `undefined` when absent. */
    custom: boolean | undefined;
    /** The custom icons (`<x14:cfIcon>`), when there are some. */
    icons?: CfIconModel[];
}

/** `<x14:iconSet>` of an x14 `iconSet` rule: the icon set, its thresholds and its custom icons. */
class IconSetExtXform extends CompositeXform<IconSetExtModel> {
    /** The xform of the thresholds. */
    readonly cfvoXform = new CfvoExtXform();
    /** The xform of the custom icons. */
    readonly cfIconXform = new CfIconExtXform();

    declare map: { readonly "x14:cfvo": CfvoExtXform; readonly "x14:cfIcon": CfIconExtXform };

    constructor() {
        super();

        this.map = {
            "x14:cfvo": this.cfvoXform,
            "x14:cfIcon": this.cfIconXform,
        };
    }

    get tag(): string {
        return "x14:iconSet";
    }

    /**
     * Writes the icon set (`iconSet` is not written when it is absent or the default, `"3TrafficLights1"`), with
     * `custom="1"` and its custom icons when it has `icons`.
     *
     * @param model - the rule
     * @throws Error when the rule has no `cfvo` (CT_IconSet requires them), an `iconSet` that is not an icon set, or
     *   `icons` that are not one per threshold (see also {@link CfIconExtXform.render})
     */
    override render(xmlStream: XmlStream, model: Pick<IconSetRuleModel, "iconSet" | "reverse" | "showValue" | "cfvo" | "icons">): void {
        const { cfvo, icons } = model;
        if (!cfvo) {
            throw new Error("An iconSet conditional formatting rule needs cfvo");
        }
        if (icons && icons.length !== cfvo.length) {
            throw new Error(
                `An icon set with custom icons needs one icon per threshold (cfvo): ${icons.length} icons for ${cfvo.length} thresholds`
            );
        }
        xmlStream.openNode(this.tag, {
            iconSet: iconSetAttribute(model.iconSet, true),
            reverse: BaseXform.toBoolAttribute(model.reverse, false),
            showValue: BaseXform.toBoolAttribute(model.showValue, true),
            custom: icons !== undefined ? "1" : undefined,
        });

        for (const item of cfvo) {
            this.cfvoXform.render(xmlStream, item);
        }

        if (icons) {
            for (const icon of icons) {
                this.cfIconXform.render(xmlStream, icon);
            }
        }

        xmlStream.closeNode();
    }

    override createNewModel({ attributes }: SaxTag): IconSetExtModel {
        const { iconSet } = attributes;
        return {
            cfvo: [],
            iconSet: iconSet !== undefined && isIconSet(iconSet) ? iconSet : DEFAULT_ICON_SET,
            reverse: parseBoolean(attributes.reverse, false),
            showValue: parseBoolean(attributes.showValue, true),
            custom: parseOptionalBoolean(attributes.custom),
        };
    }

    override onParserClose(name: string, model: IconSetExtModel): void {
        // the element was opened before its children (createNewModel); a child model is set by its start tag
        switch (name) {
            case "x14:cfvo": {
                const cfvo = this.cfvoXform.model;
                if (cfvo) {
                    model.cfvo.push(cfvo);
                }
                break;
            }

            case "x14:cfIcon": {
                const icon = this.cfIconXform.model;
                if (icon) {
                    model.icons = model.icons || [];
                    model.icons.push(icon);
                }
                break;
            }
        }
    }
}

export default IconSetExtXform;
