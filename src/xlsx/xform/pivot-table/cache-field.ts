import type { PivotCacheFieldModel } from "../../../types";
import { xmlEncode } from "../../../utils/utils";

/**
 * A `<cacheField>` of a pivot cache definition, rendered as an XML string: a string field lists its shared items, an
 * integer field (`sharedItems: null`) has none. Every value is written with `xmlEncode(String(value))`.
 */
class CacheField {
    /** The header value of the column. */
    readonly name: unknown;
    /** The distinct values of a row or column field, `null` for an integer (value) field. */
    readonly sharedItems: readonly unknown[] | null;

    constructor({ name, sharedItems }: PivotCacheFieldModel) {
        // string type
        //
        // {
        //   'name': 'A',
        //   'sharedItems': ['a1', 'a2', 'a3']
        // }
        //
        // or
        //
        // integer type
        //
        // {
        //   'name': 'D',
        //   'sharedItems': null
        // }
        this.name = name;
        this.sharedItems = sharedItems;
    }

    /** @returns the XML of the `<cacheField>` */
    render(): string {
        // PivotCache Field: http://www.datypic.com/sc/ooxml/e-ssml_cacheField-1.html
        // Shared Items: http://www.datypic.com/sc/ooxml/e-ssml_sharedItems-1.html

        // integer types
        if (this.sharedItems === null) {
            // TK(2023-07-18): left out attributes... minValue="5" maxValue="45"
            return `<cacheField name="${xmlEncode(String(this.name))}" numFmtId="0">
      <sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" />
    </cacheField>`;
        }

        // string types
        return `<cacheField name="${xmlEncode(String(this.name))}" numFmtId="0">
      <sharedItems count="${this.sharedItems.length}">
        ${this.sharedItems.map((item) => `<s v="${xmlEncode(String(item))}" />`).join("")}
      </sharedItems>
    </cacheField>`;
    }
}

export default CacheField;
