import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import BaseXform from "../base-xform";

/**
 * `<HeadingPairs>` of `docProps/app.xml`: one "Worksheets" heading with the number of worksheets. Rendered from the
 * worksheets (only their count is used); not parsed (no model).
 */
class AppHeadingPairsXform extends BaseXform<readonly unknown[]> {
    override render(xmlStream: XmlStream, model: readonly unknown[]): void {
        xmlStream.openNode("HeadingPairs");
        xmlStream.openNode("vt:vector", { size: 2, baseType: "variant" });

        xmlStream.openNode("vt:variant");
        xmlStream.leafNode("vt:lpstr", undefined, "Worksheets");
        xmlStream.closeNode();

        xmlStream.openNode("vt:variant");
        xmlStream.leafNode("vt:i4", undefined, model.length);
        xmlStream.closeNode();

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
