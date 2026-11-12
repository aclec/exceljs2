import type { Color, Style } from "../../index";

/** A shallow copy of value, with a shallow copy of its color when it is set. */
function withColorCopy<V extends { color?: Partial<Color> | undefined }>(value: V): V {
    const copy = { ...value };
    if (value.color) copy.color = { ...value.color };
    return copy;
}

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

    if (style.font) copied.font = withColorCopy(style.font);
    if (style.alignment) copied.alignment = { ...style.alignment };
    if (style.protection) copied.protection = { ...style.protection };
    const { border } = style;
    if (border) {
        const borderCopy = { ...border };
        if (border.top) borderCopy.top = withColorCopy(border.top);
        if (border.left) borderCopy.left = withColorCopy(border.left);
        if (border.bottom) borderCopy.bottom = withColorCopy(border.bottom);
        if (border.right) borderCopy.right = withColorCopy(border.right);
        if (border.diagonal) borderCopy.diagonal = withColorCopy(border.diagonal);
        copied.border = borderCopy;
    }

    if (style.fill) {
        const fill = { ...style.fill };
        // pattern fills
        if ("fgColor" in fill && fill.fgColor) fill.fgColor = { ...fill.fgColor };
        if ("bgColor" in fill && fill.bgColor) fill.bgColor = { ...fill.bgColor };
        // gradient fills
        if ("center" in fill && fill.center) fill.center = { ...fill.center };
        if ("stops" in fill && fill.stops) fill.stops = fill.stops.map(withColorCopy);
        copied.fill = fill;
    }

    return copied;
}

export { copyStyle };
