/* eslint-disable -- TypeScript migration: remove with the @ts-nocheck below once the file is typed */
// @ts-nocheck
export default {
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
