import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import BaseXform from "../base-xform";

/**
 * `<HeadingPairs>` of `docProps/app.xml`: one "Worksheets" heading with the number of worksheets, and a "Charts" heading
 * with the number of chartsheets when there are any. Rendered from the worksheets and the chartsheets (only their counts
 * are used); not parsed (no model).
 */
class AppHeadingPairsXform extends BaseXform<readonly unknown[]> {
    /** @param chartsheets - the chartsheets (none by default) */
    override render(xmlStream: XmlStream, model: readonly unknown[], chartsheets: readonly unknown[] = []): void {
        const headings: [string, number][] = [["Worksheets", model.length]];
        if (chartsheets.length) {
            headings.push(["Charts", chartsheets.length]);
        }
        xmlStream.openNode("HeadingPairs");
        xmlStream.openNode("vt:vector", { size: 2 * headings.length, baseType: "variant" });

        for (const [heading, count] of headings) {
            xmlStream.openNode("vt:variant");
            xmlStream.leafNode("vt:lpstr", undefined, heading);
            xmlStream.closeNode();

            xmlStream.openNode("vt:variant");
            xmlStream.leafNode("vt:i4", undefined, count);
            xmlStream.closeNode();
        }

        xmlStream.closeNode();
        xmlStream.closeNode();
    }

    /** @returns `true` for `<HeadingPairs>` (its content is skipped) */
    override parseOpen(node: SaxTag): boolean {
        // no parsing
        return node.name === "HeadingPairs";
    }

    override parseText(): void {}

    override parseClose(name: string): boolean {
        return name !== "HeadingPairs";
    }
}

export default AppHeadingPairsXform;
