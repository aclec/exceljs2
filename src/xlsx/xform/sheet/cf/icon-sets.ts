import type { IconSetTypes } from "../../../../../index";

// The icon sets (ST_IconSetType) and their number of icons: the ECMA-376 type of the main <iconSet>, and the x14 type of
// [MS-XLSX] (x14:iconSet, x14:cfIcon), which adds 3Stars, 3Triangles, 5Boxes and NoIcons. NoIcons has a single "icon"
// (none): Excel writes its custom icon as iconId="0". [MS-XLSX] CT_CfIcon: iconId is required, at most the number of
// icons of its set - 1.
const ICON_COUNTS: Readonly<Record<IconSetTypes, number>> = {
    "3Arrows": 3,
    "3ArrowsGray": 3,
    "3Flags": 3,
    "3TrafficLights1": 3,
    "3TrafficLights2": 3,
    "3Signs": 3,
    "3Symbols": 3,
    "3Symbols2": 3,
    "4Arrows": 4,
    "4ArrowsGray": 4,
    "4RedToBlack": 4,
    "4Rating": 4,
    "4TrafficLights": 4,
    "5Arrows": 5,
    "5ArrowsGray": 5,
    "5Rating": 5,
    "5Quarters": 5,
    "3Stars": 3,
    "3Triangles": 3,
    "5Boxes": 5,
    NoIcons: 1,
};

/** The icon sets that only the x14 version of ST_IconSetType has (Excel 2010): a rule using one is written in the extension. */
const X14_ONLY_ICON_SETS: ReadonlySet<string> = new Set<IconSetTypes>(["3Stars", "3Triangles", "5Boxes", "NoIcons"]);

/** The icon set of a rule or an icon read from a file without one (the default of ST_IconSetType). */
export const DEFAULT_ICON_SET = "3TrafficLights1" satisfies IconSetTypes;

/**
 * The name that versions before 5.0 documented (and read) as the default icon set: not an ST_IconSetType value, it is
 * written as the default one ({@link DEFAULT_ICON_SET}) so that code following the old documentation keeps working.
 */
const LEGACY_DEFAULT_ICON_SET = "3TrafficLights";

/** @returns whether `value` is an icon set of the x14 ST_IconSetType (a superset of the ECMA-376 one) */
export function isIconSet(value: string): value is IconSetTypes {
    return Object.hasOwn(ICON_COUNTS, value);
}

/** @returns whether `iconSet` is one of the icon sets that only the x14 extension can hold */
export function isX14OnlyIconSet(iconSet: string | undefined): boolean {
    return iconSet !== undefined && X14_ONLY_ICON_SETS.has(iconSet);
}

/**
 * The `iconSet` attribute of an `<iconSet>` (`x14` false) or `<x14:iconSet>` (`x14` true): the icon set, `undefined` (not
 * written) for the default one or when absent.
 *
 * @throws Error for a value that is not an icon set, or one that only the x14 extension has in the main `<iconSet>`
 */
export function iconSetAttribute(iconSet: string | undefined, x14: boolean): IconSetTypes | undefined {
    if (iconSet === undefined || iconSet === DEFAULT_ICON_SET || iconSet === LEGACY_DEFAULT_ICON_SET) {
        return undefined;
    }
    const valid = validIconSet(iconSet);
    if (!x14 && X14_ONLY_ICON_SETS.has(valid)) {
        throw new Error(`The icon set "${valid}" can only be written in the x14 extension of the worksheet`);
    }
    return valid;
}

/**
 * @returns the icon set given, checked (the legacy `"3TrafficLights"` gives the default one)
 * @throws Error when it is absent or not an icon set of the x14 ST_IconSetType
 */
export function validIconSet(iconSet: string | undefined): IconSetTypes {
    if (iconSet === LEGACY_DEFAULT_ICON_SET) {
        return DEFAULT_ICON_SET;
    }
    if (iconSet === undefined || !isIconSet(iconSet)) {
        throw new Error(
            `Unknown icon set ${iconSet === undefined ? "(none)" : `"${iconSet}"`} (ST_IconSetType), e.g. "${DEFAULT_ICON_SET}"`
        );
    }
    return iconSet;
}

/**
 * The `iconId` attribute of an `<x14:cfIcon>` (required by [MS-XLSX]): the index of the icon in its icon set, `0` when
 * absent for `NoIcons` (its only "icon").
 *
 * @param iconSet - the icon set of the icon (validated by {@link iconSetAttribute})
 * @param iconId - the index given
 * @throws Error when the index is absent (but for `NoIcons`), or not an integer from 0 to the number of icons - 1
 */
export function iconIdAttribute(iconSet: IconSetTypes, iconId: number | undefined): number {
    if (iconId === undefined && iconSet === "NoIcons") {
        return 0;
    }
    const count = ICON_COUNTS[iconSet];
    if (iconId === undefined || !Number.isInteger(iconId) || iconId < 0 || iconId >= count) {
        throw new Error(`The iconId of an icon of "${iconSet}" must be an integer from 0 to ${count - 1}, not ${String(iconId)}`);
    }
    return iconId;
}
