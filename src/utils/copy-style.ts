import type { GradientStop, Style } from "../../index";

/** A style object seen as a map of its properties (styles are plain objects). */
type StyleObject = Record<string, unknown>;

// a shallow copy of obj, and of its nestKeys properties that are set
const oneDepthCopy = <T extends object>(obj: T, nestKeys: readonly string[]): T => {
    const source = obj as StyleObject;
    const memo: StyleObject = {};
    for (const key of nestKeys) {
        if (source[key]) memo[key] = { ...source[key] };
    }
    return { ...obj, ...memo };
};

// dst[key] = a one-depth copy of src[key], when it is set
const setIfExists = (src: object, dst: object, key: string, nestKeys: readonly string[] = []): void => {
    const value = (src as StyleObject)[key];
    if (value) (dst as StyleObject)[key] = oneDepthCopy(value, nestKeys);
};

const isEmptyObj = (obj: object): boolean => Object.keys(obj).length === 0;

/**
 * A copy of a style deep enough that changing the copy's font, alignment, protection, border (and its edges) or fill
 * (and its colours, centre and gradient stops) does not change the original.
 * @returns The copy; an empty style for an empty one; style itself when it is null or undefined.
 */
function copyStyle(style: Partial<Style>): Partial<Style>;
function copyStyle(style: Partial<Style> | null | undefined): Partial<Style> | null | undefined;
function copyStyle(style: Partial<Style> | null | undefined): Partial<Style> | null | undefined {
    if (!style) return style;
    if (isEmptyObj(style)) return {};

    const copied = { ...style };

    setIfExists(style, copied, "font", ["color"]);
    setIfExists(style, copied, "alignment");
    setIfExists(style, copied, "protection");
    if (style.border) {
        setIfExists(style, copied, "border");
        // set just above, from style.border
        const border = copied.border!;
        setIfExists(style.border, border, "top", ["color"]);
        setIfExists(style.border, border, "left", ["color"]);
        setIfExists(style.border, border, "bottom", ["color"]);
        setIfExists(style.border, border, "right", ["color"]);
        setIfExists(style.border, border, "diagonal", ["color"]);
    }

    if (style.fill) {
        setIfExists(style, copied, "fill", ["fgColor", "bgColor", "center"]);
        // gradient fills only
        const { stops } = style.fill as { stops?: readonly GradientStop[] };
        if (stops) {
            // set just above, from style.fill
            (copied.fill as { stops?: GradientStop[] }).stops = stops.map((s) => oneDepthCopy(s, ["color"]));
        }
    }

    return copied;
}

export { copyStyle };
