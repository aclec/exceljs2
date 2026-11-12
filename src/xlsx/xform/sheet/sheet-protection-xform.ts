import type { SaxTag } from "../../../utils/parse-sax";
import type XmlStream from "../../../utils/xml-stream";
import type { XmlValue } from "../../../utils/xml-stream";
import type { SheetProtectionModel } from "../../../types";
import { parseBoolean } from "../../../utils/utils";
import BaseXform from "../base-xform";

function booleanToXml<V extends string>(model: boolean | undefined, value: V): V | undefined {
    return model ? value : undefined;
}

/**
 * An xsd:boolean attribute that differs from its default (`true` for the protected sheet, `false` for an allowed
 * operation): `true` when it says so, `undefined` otherwise (absent, or the default).
 */
function xmlToBoolean(value: string | undefined, equals: boolean): true | undefined {
    return value !== undefined && parseBoolean(value) === equals ? true : undefined;
}

/** An xsd:boolean attribute that forbids (`"1"` / `"true"`): `false`, `undefined` otherwise. */
function xmlToForbidden(value: string | undefined): false | undefined {
    return parseBoolean(value) ? false : undefined;
}

/**
 * `<sheetProtection>`: what a protected worksheet still allows, and the password hash. In the model, `false` forbids
 * (`selectLockedCells`, `selectUnlockedCells`, `objects`, `scenarios`: written as `"1"` when `false`) and `true` allows
 * (the other operations: written as `"0"` when `true`); parsing gives `true` / `false` for the attributes present and
 * `undefined` keys for the others. The hash and `objects` / `scenarios` are written for a protected sheet (`sheet`) only.
 */
class SheetProtectionXform extends BaseXform<SheetProtectionModel> {
    get tag(): string {
        return "sheetProtection";
    }

    override render(xmlStream: XmlStream, model: SheetProtectionModel | null | undefined): void {
        if (model) {
            const attributes: Record<string, XmlValue | undefined> = {
                sheet: booleanToXml(model.sheet, "1"),
                selectLockedCells: model.selectLockedCells === false ? "1" : undefined,
                selectUnlockedCells: model.selectUnlockedCells === false ? "1" : undefined,
                formatCells: booleanToXml(model.formatCells, "0"),
                formatColumns: booleanToXml(model.formatColumns, "0"),
                formatRows: booleanToXml(model.formatRows, "0"),
                insertColumns: booleanToXml(model.insertColumns, "0"),
                insertRows: booleanToXml(model.insertRows, "0"),
                insertHyperlinks: booleanToXml(model.insertHyperlinks, "0"),
                deleteColumns: booleanToXml(model.deleteColumns, "0"),
                deleteRows: booleanToXml(model.deleteRows, "0"),
                sort: booleanToXml(model.sort, "0"),
                autoFilter: booleanToXml(model.autoFilter, "0"),
                pivotTables: booleanToXml(model.pivotTables, "0"),
            };
            if (model.sheet) {
                attributes.algorithmName = model.algorithmName;
                attributes.hashValue = model.hashValue;
                attributes.saltValue = model.saltValue;
                attributes.spinCount = model.spinCount;
                attributes.objects = booleanToXml(model.objects === false, "1");
                attributes.scenarios = booleanToXml(model.scenarios === false, "1");
            }
            if (Object.values(attributes).some((value) => value !== undefined)) {
                xmlStream.leafNode(this.tag, attributes);
            }
        }
    }

    /** @returns `true` for `<sheetProtection>` */
    override parseOpen(node: SaxTag): boolean {
        switch (node.name) {
            case this.tag: {
                const { attributes } = node;
                const model: SheetProtectionModel = (this.model = {
                    sheet: xmlToBoolean(attributes.sheet, true),
                    objects: xmlToForbidden(attributes.objects),
                    scenarios: xmlToForbidden(attributes.scenarios),
                    selectLockedCells: xmlToForbidden(attributes.selectLockedCells),
                    selectUnlockedCells: xmlToForbidden(attributes.selectUnlockedCells),
                    formatCells: xmlToBoolean(attributes.formatCells, false),
                    formatColumns: xmlToBoolean(attributes.formatColumns, false),
                    formatRows: xmlToBoolean(attributes.formatRows, false),
                    insertColumns: xmlToBoolean(attributes.insertColumns, false),
                    insertRows: xmlToBoolean(attributes.insertRows, false),
                    insertHyperlinks: xmlToBoolean(attributes.insertHyperlinks, false),
                    deleteColumns: xmlToBoolean(attributes.deleteColumns, false),
                    deleteRows: xmlToBoolean(attributes.deleteRows, false),
                    sort: xmlToBoolean(attributes.sort, false),
                    autoFilter: xmlToBoolean(attributes.autoFilter, false),
                    pivotTables: xmlToBoolean(attributes.pivotTables, false),
                });
                if (attributes.algorithmName) {
                    model.algorithmName = attributes.algorithmName;
                    model.hashValue = attributes.hashValue;
                    model.saltValue = attributes.saltValue;
                    // NaN when absent
                    model.spinCount = BaseXform.toIntValue(attributes.spinCount, NaN);
                }
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

export default SheetProtectionXform;
