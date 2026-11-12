// Finds the rectangle ({top, left, bottom, right}) that contains a cell among many, without scanning them all.
//
// A rectangle of height h and width w is filed under its size class (kr, kc): the smallest powers of two 2^kr >= h and
// 2^kc >= w. The grid is cut into blocks of 2^kr rows by 2^kc columns, and the rectangle is listed in the (at most 4) blocks
// it intersects. A lookup checks, for each size class in use, the one block of that class containing the cell. Rectangles
// that do not overlap are spread over the blocks (a block holds a few dozen of them at most), so a lookup costs about the
// number of size classes in use (at most 21 x 15), whatever the number of rectangles.
//
// When rectangles overlap, find() returns the one added first (the lowest order).

/** A rectangle of cells for a {@link RangeIndex}: 1-based row and column numbers, inclusive. */
export interface IndexedRange {
    top: number;
    left: number;
    bottom: number;
    right: number;
    /** Set by {@link RangeIndex.add}: the lowest wins between overlapping rectangles. */
    order?: number;
    /** Set by {@link RangeIndex.add}: [kr, kc]. */
    sizeClass?: readonly [number, number];
}

interface SizeClass {
    kr: number;
    kc: number;
    count: number;
}

// the smallest k such that 2^k >= size
const sizeClass = (size: number): number => (size <= 1 ? 0 : 32 - Math.clz32(size - 1));

// rows < 2^21, columns < 2^15, size classes < 32
const blockKey = (kr: number, kc: number, rowBlock: number, colBlock: number): number =>
    ((kr * 32 + kc) * 2 ** 21 + rowBlock) * 2 ** 15 + colBlock;

/**
 * An index of rectangles of cells, to find the one that contains a cell in about constant time (see above).
 * @typeParam R - The rectangles: add() sets their `order` and `sizeClass`.
 */
class RangeIndex<R extends IndexedRange = IndexedRange> {
    private _blocks: Map<number, R[]>; // blockKey => [range]
    private _classes: Map<number, SizeClass>; // kr * 32 + kc => {kr, kc, count}

    constructor() {
        this._blocks = new Map();
        this._classes = new Map();
    }

    /**
     * Adds a rectangle; its dimensions must not change until it is removed.
     * @param order - The lowest wins between overlapping rectangles.
     */
    add(range: R, order: number): void {
        range.order = order;
        const kr = sizeClass(range.bottom - range.top + 1);
        const kc = sizeClass(range.right - range.left + 1);
        range.sizeClass = [kr, kc];
        this._eachBlock(range, (key) => {
            const block = this._blocks.get(key);
            if (block) {
                block.push(range);
            } else {
                this._blocks.set(key, [range]);
            }
        });
        const classKey = kr * 32 + kc;
        const sizes = this._classes.get(classKey);
        if (sizes) {
            sizes.count++;
        } else {
            this._classes.set(classKey, { kr, kc, count: 1 });
        }
    }

    /**
     * Removes a rectangle added before, once: removing it again removes other rectangles of its blocks.
     * @throws TypeError when it was never added.
     */
    remove(range: R): void {
        this._eachBlock(range, (key) => {
            // an emptied block is kept: deleting and adding again the keys of a large Map is slow in V8
            const block = this._blocks.get(key)!;
            block.splice(block.indexOf(range), 1);
        });
        const [kr, kc] = classOf(range);
        const classKey = kr * 32 + kc;
        const sizes = this._classes.get(classKey)!;
        if (--sizes.count === 0) {
            this._classes.delete(classKey);
        }
    }

    /** The rectangle that contains the cell, the one with the lowest order when several do; undefined when none does. */
    find(row: number, col: number): R | undefined {
        let found: R | undefined;
        for (const { kr, kc } of this._classes.values()) {
            const block = this._blocks.get(blockKey(kr, kc, row >> kr, col >> kc));
            if (block) {
                for (let i = 0; i < block.length; i++) {
                    const range = block[i]!;
                    if (
                        row >= range.top &&
                        row <= range.bottom &&
                        col >= range.left &&
                        col <= range.right &&
                        // order is set by add()
                        (!found || range.order! < found.order!)
                    ) {
                        found = range;
                    }
                }
            }
        }
        return found;
    }

    // calls callback with the key of each block of its size class that the rectangle intersects
    private _eachBlock(range: R, callback: (key: number) => void): void {
        const [kr, kc] = classOf(range);
        for (let rowBlock = range.top >> kr; rowBlock <= range.bottom >> kr; rowBlock++) {
            for (let colBlock = range.left >> kc; colBlock <= range.right >> kc; colBlock++) {
                callback(blockKey(kr, kc, rowBlock, colBlock));
            }
        }
    }
}

// the size class that add() gave the rectangle
function classOf(range: IndexedRange): readonly [number, number] {
    if (!range.sizeClass) {
        throw new TypeError("RangeIndex: this range was not added");
    }
    return range.sizeClass;
}

export default RangeIndex;
