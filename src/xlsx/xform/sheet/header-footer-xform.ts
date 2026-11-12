import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { HeaderFooterModel } from "../../../types";
import BaseXform from "../base-xform";

/** The header / footer texts: the child elements of `<headerFooter>`, in CT_HeaderFooter order. */
type HeaderFooterText = "oddHeader" | "oddFooter" | "evenHeader" | "evenFooter" | "firstHeader" | "firstFooter";

/**
 * `<headerFooter differentFirst="1" differentOddEven="1"><oddHeader>…</oddHeader>…</headerFooter>`: the headers and
 * footers of the printed pages. Rendered only when a flag or a (non-empty string) text is set. Parsing keeps the flags
 * present and the texts.
 */
class HeaderFooterXform extends BaseXform<HeaderFooterModel> {
    /** While parsing: the text element being read. */
    declare currentNode: HeaderFooterText | undefined;

    get tag(): string {
        return "headerFooter";
    }

    override render(xmlStream: XmlStream, model: HeaderFooterModel | null | undefined): void {
        if (model) {
            xmlStream.addRollback();

            let createTag = false;

            xmlStream.openNode("headerFooter");
            if (model.differentFirst) {
                xmlStream.addAttribute("differentFirst", "1");
                createTag = true;
            }
            if (model.differentOddEven) {
                xmlStream.addAttribute("differentOddEven", "1");
                createTag = true;
            }
            if (model.oddHeader && typeof model.oddHeader === "string") {
                xmlStream.leafNode("oddHeader", undefined, model.oddHeader);
                createTag = true;
            }
            if (model.oddFooter && typeof model.oddFooter === "string") {
                xmlStream.leafNode("oddFooter", undefined, model.oddFooter);
                createTag = true;
            }
            if (model.evenHeader && typeof model.evenHeader === "string") {
                xmlStream.leafNode("evenHeader", undefined, model.evenHeader);
                createTag = true;
            }
            if (model.evenFooter && typeof model.evenFooter === "string") {
                xmlStream.leafNode("evenFooter", undefined, model.evenFooter);
                createTag = true;
            }
            if (model.firstHeader && typeof model.firstHeader === "string") {
                xmlStream.leafNode("firstHeader", undefined, model.firstHeader);
                createTag = true;
            }
            if (model.firstFooter && typeof model.firstFooter === "string") {
                xmlStream.leafNode("firstFooter", undefined, model.firstFooter);
                createTag = true;
            }

            if (createTag) {
                xmlStream.closeNode();
                xmlStream.commit();
            } else {
                xmlStream.rollback();
            }
        }
    }

    /** @returns `true` for `<headerFooter>` and its text elements */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case "headerFooter": {
                const model: HeaderFooterModel = (this.model = {});
                if (node.attributes.differentFirst) {
                    model.differentFirst = parseInt(node.attributes.differentFirst, 0) === 1;
                }
                if (node.attributes.differentOddEven) {
                    model.differentOddEven = parseInt(node.attributes.differentOddEven, 0) === 1;
                }
                return true;
            }

            case "oddHeader":
            case "oddFooter":
            case "evenHeader":
            case "evenFooter":
            case "firstHeader":
            case "firstFooter":
                this.currentNode = node.name;
                return true;

            default:
                return false;
        }
    }

    /** Sets the text of the text element being read (the last chunk wins). */
    override parseText(text: string): void {
        const { currentNode } = this;
        if (currentNode !== undefined) {
            // a text element is only read inside <headerFooter>, which created the model
            this.model![currentNode] = text;
        }
    }

    /** @returns `true` at the end of a text element, `false` otherwise (the end of `<headerFooter>`) */
    override parseClose(): boolean {
        if (this.currentNode !== undefined) {
            this.currentNode = undefined;
            return true;
        }
        return false;
    }
}

export default HeaderFooterXform;
