import type XmlStream from "../../../utils/xml-stream";
import type { ConditionalFormattingModel } from "../../../types";
import CompositeXform from "../composite-xform";

import ConditionalFormattingsExt from "./cf-ext/conditional-formattings-ext-xform";
import type { ConditionalFormattingExtModel } from "./cf-ext/conditional-formatting-ext-xform";

/** What the worksheet extension renders: the conditional formattings of the worksheet (those with x14 rules). */
export interface ExtLstRenderModel {
    conditionalFormattings: ConditionalFormattingModel[];
}

/**
 * The parsed worksheet extension, by element name: `x14:conditionalFormattings`. Also accepts the render model (see
 * {@link ExtLstRenderModel}).
 */
export interface ExtLstModel {
    "x14:conditionalFormattings"?: ConditionalFormattingExtModel[];
    conditionalFormattings?: ConditionalFormattingModel[];
}

/** The `<ext>` of the x14 conditional formattings. */
class ExtXform extends CompositeXform<ExtLstModel> {
    /** The xform of the conditional formattings. */
    readonly conditionalFormattings = new ConditionalFormattingsExt();

    declare map: { readonly "x14:conditionalFormattings": ConditionalFormattingsExt };

    constructor() {
        super();
        this.map = {
            "x14:conditionalFormattings": this.conditionalFormattings,
        };
    }

    get tag(): string {
        return "ext";
    }

    /** @returns whether a conditional formatting has an x14 rule */
    hasContent(model: ExtLstRenderModel): boolean {
        return this.conditionalFormattings.hasContent(model.conditionalFormattings);
    }

    override prepare(model: ExtLstRenderModel): void {
        this.conditionalFormattings.prepare(model.conditionalFormattings);
    }

    override render(xmlStream: XmlStream, model: ExtLstRenderModel): void {
        xmlStream.openNode("ext", {
            uri: "{78C0D931-6437-407d-A8EE-F0AAD7539E65}",
            "xmlns:x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
        });

        this.conditionalFormattings.render(xmlStream, model.conditionalFormattings);

        xmlStream.closeNode();
    }

    override createNewModel(): ExtLstModel {
        return {};
    }

    override onParserClose(name: string): void {
        if (name === "x14:conditionalFormattings") {
            // the element was opened before its child (createNewModel)
            this.model![name] = this.conditionalFormattings.model ?? undefined;
        }
    }
}

/**
 * `<extLst>` of a worksheet: its extensions. Only the x14 conditional formattings are written (when a rule needs them)
 * and read.
 */
class ExtLstXform extends CompositeXform<ExtLstModel> {
    /** The xform of the `<ext>`. */
    readonly ext = new ExtXform();

    declare map: { readonly ext: ExtXform };

    constructor() {
        super();

        this.map = {
            ext: this.ext,
        };
    }

    get tag(): string {
        return "extLst";
    }

    /** Gives the x14 rules their `x14Id` (mutates the rules). */
    override prepare(model: ExtLstRenderModel): void {
        this.ext.prepare(model);
    }

    /** @returns whether a conditional formatting has an x14 rule */
    hasContent(model: ExtLstRenderModel): boolean {
        return this.ext.hasContent(model);
    }

    override render(xmlStream: XmlStream, model: ExtLstRenderModel): void {
        if (!this.hasContent(model)) {
            return;
        }

        xmlStream.openNode("extLst");
        this.ext.render(xmlStream, model);
        xmlStream.closeNode();
    }

    override createNewModel(): ExtLstModel {
        return {};
    }

    override onParserClose(): void {
        // the element was opened before its child (createNewModel)
        Object.assign(this.model!, this.ext.model);
    }
}

export default ExtLstXform;
