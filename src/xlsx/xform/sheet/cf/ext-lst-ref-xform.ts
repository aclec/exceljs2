import type XmlStream from "../../../../utils/xml-stream";
import type { ConditionalFormattingRuleModel } from "../../../../types";
import BaseXform from "../../base-xform";
import CompositeXform from "../../composite-xform";

/** The model of the reference of a main rule to its x14 version (merged into the rule). */
export interface ExtLstRefModel {
    x14Id?: string;
}

/** `<x14:id>{…}</x14:id>`: the id of the x14 version of a rule. */
class X14IdXform extends BaseXform<string> {
    get tag(): string {
        return "x14:id";
    }

    override render(xmlStream: XmlStream, model: string | undefined): void {
        xmlStream.leafNode(this.tag, undefined, model);
    }

    /** Starts the text (the parent delegates `<x14:id>` only). */
    override parseOpen(): void {
        this.model = "";
    }

    override parseText(text: string): void {
        this.model = `${String(this.model)}${text}`;
    }

    /** @returns `false` at the end of `<x14:id>` */
    override parseClose(name: string): boolean {
        return name !== this.tag;
    }
}

/** The `<ext>` holding the `<x14:id>` of a rule. */
class ExtXform extends CompositeXform<ExtLstRefModel> {
    /** The xform of the id. */
    readonly idXform = new X14IdXform();

    declare map: { readonly "x14:id": X14IdXform };

    constructor() {
        super();

        this.map = {
            "x14:id": this.idXform,
        };
    }

    get tag(): string {
        return "ext";
    }

    override render(xmlStream: XmlStream, model: Pick<ConditionalFormattingRuleModel, "x14Id">): void {
        xmlStream.openNode(this.tag, {
            uri: "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}",
            "xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
        });

        this.idXform.render(xmlStream, model.x14Id);

        xmlStream.closeNode();
    }

    override createNewModel(): ExtLstRefModel {
        return {};
    }

    override onParserClose(_name: string, model: ExtLstRefModel): void {
        model.x14Id = this.idXform.model ?? undefined;
    }
}

/**
 * `<extLst><ext uri="{B025F937-…}"><x14:id>{…}</x14:id></ext></extLst>` of a main `<cfRule>`: the reference to the x14
 * version of the rule (its `x14Id`).
 */
class ExtLstRefXform extends CompositeXform<ExtLstRefModel> {
    declare map: { readonly ext: ExtXform };

    constructor() {
        super();
        this.map = {
            ext: new ExtXform(),
        };
    }

    get tag(): string {
        return "extLst";
    }

    override render(xmlStream: XmlStream, model: Pick<ConditionalFormattingRuleModel, "x14Id">): void {
        xmlStream.openNode(this.tag);
        this.map.ext.render(xmlStream, model);
        xmlStream.closeNode();
    }

    override createNewModel(): ExtLstRefModel {
        return {};
    }

    override onParserClose(_name: string, model: ExtLstRefModel): void {
        const ext = this.map.ext.model;
        if (ext && ext.x14Id !== undefined) {
            model.x14Id = ext.x14Id;
        }
    }
}

export default ExtLstRefXform;
