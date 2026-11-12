import type { SaxTag } from "../../../../utils/parse-sax";
import type XmlStream from "../../../../utils/xml-stream";
import type { CfvoModel, IconSetRuleModel } from "../../../../types";
import BaseXform from "../../base-xform";
import CompositeXform from "../../composite-xform";

import CfvoXform from "./cfvo-xform";

/** The model of `<iconSet>` (merged into the rule). The parsed `iconSet` is the attribute (unvalidated). */
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
     * @param model - the rule
     * @throws Error when the rule has no `cfvo` (CT_IconSet requires them)
     */
    override render(xmlStream: XmlStream, model: Pick<IconSetRuleModel, "iconSet" | "reverse" | "showValue" | "cfvo">): void {
        const { cfvo } = model;
        if (!cfvo) {
            throw new Error("An iconSet conditional formatting rule needs cfvo");
        }
        xmlStream.openNode(this.tag, {
            iconSet: BaseXform.toStringAttribute(model.iconSet, "3TrafficLights"),
            reverse: BaseXform.toBoolAttribute(model.reverse, false),
            showValue: BaseXform.toBoolAttribute(model.showValue, true),
        });

        for (const item of cfvo) {
            this.cfvoXform.render(xmlStream, item);
        }

        xmlStream.closeNode();
    }

    override createNewModel({ attributes }: SaxTag): IconSetModel {
        return {
            // unvalidated ST_IconSetType
            iconSet: BaseXform.toStringValue(attributes.iconSet, "3TrafficLights") as IconSetModel["iconSet"],
            reverse: BaseXform.toBoolValue(attributes.reverse),
            showValue: BaseXform.toBoolValue(attributes.showValue),
            cfvo: [],
        };
    }

    override onParserClose(name: string): void {
        if (name === "cfvo") {
            // the element was opened before its children; the cfvo model is set by its start tag
            this.model!.cfvo.push(this.cfvoXform.model!);
        }
    }
}

export default IconSetXform;
