import type XmlStream from "../../../../utils/xml-stream";
import type { CfFormula } from "../../../../types";
import BaseXform from "../../base-xform";

/** `<formula>…</formula>` of a `<cfRule>`: the model is the formula text (parsing concatenates its chunks). */
class FormulaXform extends BaseXform<CfFormula> {
    get tag(): string {
        return "formula";
    }

    override render(xmlStream: XmlStream, model: CfFormula): void {
        xmlStream.leafNode(this.tag, undefined, model);
    }

    /** Starts the text (the parent delegates `<formula>` only). */
    override parseOpen(): void {
        this.model = "";
    }

    override parseText(text: string): void {
        this.model = `${String(this.model)}${text}`;
    }

    /** @returns `false` at the end of `<formula>` */
    override parseClose(name: string): boolean {
        return name !== this.tag;
    }
}

export default FormulaXform;
