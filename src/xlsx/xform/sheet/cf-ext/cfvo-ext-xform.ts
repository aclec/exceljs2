import type { SaxTag } from "../../../../utils/parse-sax";
import type XmlStream from "../../../../utils/xml-stream";
import type { CfvoModel } from "../../../../types";
import CompositeXform from "../../composite-xform";

import CfvoXform, { cfvoType } from "../cf/cfvo-xform";
import FExtXform from "./f-ext-xform";

/**
 * `<x14:cfvo type="num"><xm:f>50</xm:f></x14:cfvo>`: a threshold of an x14 data bar or icon set, its value in `<xm:f>`.
 * A parsed `type` that is absent or unknown is `num`; an empty `<xm:f>` is the value 0.
 */
class CfvoExtXform extends CompositeXform<CfvoModel> {
    /** The xform of the value. */
    readonly fExtXform = new FExtXform();

    declare map: { readonly "xm:f": FExtXform };

    constructor() {
        super();

        this.map = {
            "xm:f": this.fExtXform,
        };
    }

    get tag(): string {
        return "x14:cfvo";
    }

    override render(xmlStream: XmlStream, model: CfvoModel): void {
        xmlStream.openNode(this.tag, {
            type: model.type,
        });
        if (model.value !== undefined) {
            this.fExtXform.render(xmlStream, model.value);
        }
        xmlStream.closeNode();
    }

    override createNewModel(node: SaxTag): CfvoModel {
        return {
            type: cfvoType(node.attributes.type),
        };
    }

    override onParserClose(name: string, model: CfvoModel): void {
        switch (name) {
            case "xm:f": {
                // <xm:f> holds text
                const text = this.fExtXform.model;
                model.value = text ? CfvoXform.toValue(String(text)) : 0;
                break;
            }
        }
    }
}

export default CfvoExtXform;
