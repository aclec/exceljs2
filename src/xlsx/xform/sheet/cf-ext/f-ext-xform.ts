import type XmlStream from "../../../../utils/xml-stream";
import type { CfFormula } from "../../../../types";
import BaseXform from "../../base-xform";

/**
 * `<xm:f>…</xm:f>`: a formula of an x14 rule, or the value of an x14 threshold. The model is the text (parsing
 * concatenates its chunks).
 */
class FExtXform extends BaseXform<CfFormula> {
    get tag(): string {
        return "xm:f";
    }

    override render(xmlStream: XmlStream, model: CfFormula): void {
        xmlStream.leafNode(this.tag, undefined, model);
    }

    /** Starts the text (the parent delegates `<xm:f>` only). */
    override parseOpen(): void {
        this.model = "";
    }

    override parseText(text: string): void {
        this.model = `${String(this.model)}${text}`;
    }

    /** @returns `false` at the end of `<xm:f>` */
    override parseClose(name: string): boolean {
        return name !== this.tag;
    }
}

export default FExtXform;
