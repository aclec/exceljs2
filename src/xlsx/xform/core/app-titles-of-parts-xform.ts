import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import BaseXform from "../base-xform";

/** `<TitlesOfParts>` of `docProps/app.xml`: the names of the worksheets. Rendered only (parsing skips it, no model). */
class AppTitlesOfPartsXform extends BaseXform<readonly { readonly name: string }[]> {
    override render(xmlStream: XmlStream, model: readonly { readonly name: string }[]): void {
        xmlStream.openNode("TitlesOfParts");
        xmlStream.openNode("vt:vector", { size: model.length, baseType: "lpstr" });

        for (const sheet of model) {
            xmlStream.leafNode("vt:lpstr", undefined, sheet.name);
        }

        xmlStream.closeNode();
        xmlStream.closeNode();
    }

    /** @returns `true` for `<TitlesOfParts>` (its content is skipped) */
    override parseOpen(node: SaxTag): boolean {
        // no parsing
        return node.name === "TitlesOfParts";
    }

    override parseText(): void {}

    override parseClose(name: string): boolean {
        return name !== "TitlesOfParts";
    }
}

export default AppTitlesOfPartsXform;
