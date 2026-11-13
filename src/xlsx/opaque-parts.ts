// Helpers shared by the parts carried opaque from a file read to the file written (pivot-parts.ts): the relationships part
// of a part, read with the ids checked. The resolution of the targets is chart-parts.ts resolvePartName.

import RelationshipsXform from "./xform/core/relationships-xform";
import type { ZipEntries } from "./chart-parts";
import type { RelationshipModel } from "../types";

// a relationship id as written back (xsd:ID, ASCII): anchored, a single character class (linear)
const REL_ID = /^[A-Za-z_][A-Za-z0-9_.-]*$/;

/** The relationships part of a part: `dir/_rels/name.rels`. */
export function relsPartName(partName: string): string {
    const slash = partName.lastIndexOf("/");
    return `${partName.slice(0, slash + 1)}_rels/${partName.slice(slash + 1)}.rels`;
}

/**
 * Reads the relationships of a part (inflating its relationships part once).
 *
 * @returns the relationships (`[]` without a relationships part), or undefined when the part is refused: malformed XML, an
 * id that is not an XML name or is repeated
 * @throws (rejects) when a decompression limit of the zip reader is exceeded
 */
export async function readPartRels(entries: ZipEntries, partName: string): Promise<RelationshipModel[] | undefined> {
    const entry = entries.get(relsPartName(partName));
    if (!entry) {
        return [];
    }
    const data = await entry.data();
    let rels: RelationshipModel[] | null | undefined;
    try {
        // a new xform per part: a failed parse leaves its state behind
        rels = await new RelationshipsXform().parseStream([data]);
    } catch {
        return undefined;
    }
    const ids = new Set<string>();
    for (const rel of rels ?? []) {
        if (!REL_ID.test(rel.Id) || ids.has(rel.Id)) {
            return undefined;
        }
        ids.add(rel.Id);
    }
    return rels ?? [];
}
