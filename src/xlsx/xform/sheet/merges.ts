import type { MergeCellModel } from "../../../types";
import Range from "../../../doc/range";

/**
 * The merged ranges of a worksheet being written, collected from its merged cells (worksheet-xform prepare passes it as
 * `options.merges` to the cell xforms): each merged cell expands the range of its master.
 */
class Merges {
    /** The range of each merge, by the address of its master (top left) cell, in the order the masters were added. */
    readonly merges = new Map<string, Range>();

    /**
     * Adds a merged cell: creates the range of its master, or expands it to the cell.
     *
     * @param merge - a merged cell (`master` is the address of the master cell)
     */
    add(merge: Pick<MergeCellModel, "address" | "master">): void {
        // merge is {address, master}
        const master = String(merge.master);
        const range = this.merges.get(master);
        if (range) {
            range.expandToAddress(merge.address);
        } else {
            this.merges.set(master, new Range(`${master}:${merge.address}`));
        }
    }

    /** The merged ranges (`A1:B2`), as the `<mergeCells>` of the worksheet. */
    get mergeCells(): string[] {
        return Array.from(this.merges.values(), (merge) => merge.range);
    }

    /**
     * @param address - a cell address (`B2`)
     * @returns the address of the master cell of the merged range holding the cell, or `undefined` when it is not merged
     */
    getMasterAddress(address: string): string | undefined {
        // if address has been merged, return its master's address
        for (const range of this.merges.values()) {
            if (range.contains(address)) {
                return range.tl;
            }
        }
        return undefined;
    }
}

export default Merges;
