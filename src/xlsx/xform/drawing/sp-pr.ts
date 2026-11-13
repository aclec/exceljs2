import type { StaticModel } from "../static-xform";

/** The shape properties of a picture, `<xdr:spPr>`: a rectangle, positioned by the anchor (rendered by a StaticXform). */
const spPr: StaticModel = {
    tag: "xdr:spPr",
    c: [
        {
            tag: "a:xfrm",
            c: [
                { tag: "a:off", $: { x: "0", y: "0" } },
                { tag: "a:ext", $: { cx: "0", cy: "0" } },
            ],
        },
        {
            tag: "a:prstGeom",
            $: { prst: "rect" },
            c: [{ tag: "a:avLst" }],
        },
    ],
};

export default spPr;
