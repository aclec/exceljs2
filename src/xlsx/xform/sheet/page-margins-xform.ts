import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { Margins } from "../../../../index";
import BaseXform from "../base-xform";

/**
 * `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>` (inches). Rendered when a
 * margin is set; parsing gives every margin, with Excel's defaults for the absent ones.
 */
class PageMarginsXform extends BaseXform<Margins> {
    get tag(): string {
        return "pageMargins";
    }

    override render(xmlStream: XmlStream, model: Partial<Margins> | null | undefined): void {
        if (model) {
            const { left, right, top, bottom, header, footer } = model;
            if (
                left !== undefined ||
                right !== undefined ||
                top !== undefined ||
                bottom !== undefined ||
                header !== undefined ||
                footer !== undefined
            ) {
                xmlStream.leafNode(this.tag, { left, right, top, bottom, header, footer });
            }
        }
    }

    /** @returns `true` for `<pageMargins>` */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case this.tag: {
                const { attributes } = node;
                this.model = {
                    left: parseFloat(attributes.left || "0.7"),
                    right: parseFloat(attributes.right || "0.7"),
                    top: parseFloat(attributes.top || "0.75"),
                    bottom: parseFloat(attributes.bottom || "0.75"),
                    header: parseFloat(attributes.header || "0.3"),
                    footer: parseFloat(attributes.footer || "0.3"),
                };
                return true;
            }
            default:
                return false;
        }
    }

    override parseText(): void {}

    override parseClose(): boolean {
        return false;
    }
}

export default PageMarginsXform;
