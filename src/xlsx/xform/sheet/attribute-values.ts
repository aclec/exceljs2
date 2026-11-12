// Readers of the attribute values of the worksheet elements (strings of the file, untrusted).
import { parseBoolean } from "../../../utils/utils";

/**
 * A type guard for the string literals of a union (the values of an ECMA-376 enumeration such as ST_Orientation), from
 * a record that has every one of them as a key: the record type makes the list exhaustive, and the guard checks the own
 * keys of the record (never an inherited member: the value comes from the file).
 *
 * @example
 * const isOrientation = isLiteralOf<"portrait" | "landscape">({ portrait: true, landscape: true });
 * const orientation = isOrientation(attributes.orientation) ? attributes.orientation : "portrait";
 */
export function isLiteralOf<T extends string>(literals: Readonly<Record<T, true>>): (value: string | undefined) => value is T {
    return (value: string | undefined): value is T => value !== undefined && Object.hasOwn(literals, value);
}

/** An xsd:boolean attribute (see utils.parseBoolean: `"1"` / `"true"` are true), `undefined` when it is absent. */
export function parseOptionalBoolean(value: string | undefined): boolean | undefined {
    return value === undefined ? undefined : parseBoolean(value);
}
