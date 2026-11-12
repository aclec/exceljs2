import type XmlStream from "../../../../utils/xml-stream";
import BaseXform from "../../base-xform";

/** `<xm:sqref>A1:B10</xm:sqref>`: the range of an x14 conditional formatting. */
class SqrefExtXform extends BaseXform<string> {
    get tag(): string {
        return "xm:sqref";
    }

    override render(xmlStream: XmlStream, model: string | undefined): void {
        xmlStream.leafNode(this.tag, undefined, model);
    }

    /** Starts the text (the parent delegates `<xm:sqref>` only). */
    override parseOpen(): void {
        this.model = "";
    }

    override parseText(text: string): void {
        this.model = `${String(this.model)}${text}`;
    }

    /** @returns `false` at the end of `<xm:sqref>` */
    override parseClose(name: string): boolean {
        return name !== this.tag;
    }
}

export default SqrefExtXform;
