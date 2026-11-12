// Finds the rectangle ({top, left, bottom, right}) that contains a cell among many, without scanning them all.
//
// A rectangle of height h and width w is filed under its size class (kr, kc): the smallest powers of two 2^kr >= h and
// 2^kc >= w. The grid is cut into blocks of 2^kr rows by 2^kc columns, and the rectangle is listed in the (at most 4) blocks
// it intersects. A lookup checks, for each size class in use, the one block of that class containing the cell. Rectangles
// that do not overlap are spread over the blocks (a block holds a few dozen of them at most), so a lookup costs about the
// number of size classes in use (at most 21 x 15), whatever the number of rectangles.
//
// When rectangles overlap, find() returns the one added first (the lowest order).

// the smallest k such that 2^k >= size
const sizeClass = (size) => (size <= 1 ? 0 : 32 - Math.clz32(size - 1));

// rows < 2^21, columns < 2^15, size classes < 32
const blockKey = (kr, kc, rowBlock, colBlock) => ((kr * 32 + kc) * 2 ** 21 + rowBlock) * 2 ** 15 + colBlock;

class RangeIndex {
    constructor() {
        this._blocks = new Map(); // blockKey => [range]
        this._classes = new Map(); // kr * 32 + kc => {kr, kc, count}
    }

    // range: {top, left, bottom, right}, order: a number, the lowest wins between overlapping rectangles
    add(range, order) {
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

    remove(range) {
        this._eachBlock(range, (key) => {
            // an emptied block is kept: deleting and adding again the keys of a large Map is slow in V8
            const block = this._blocks.get(key);
            block.splice(block.indexOf(range), 1);
        });
        const [kr, kc] = range.sizeClass;
        const classKey = kr * 32 + kc;
        const sizes = this._classes.get(classKey);
        if (--sizes.count === 0) {
            this._classes.delete(classKey);
        }
    }

    find(row, col) {
        let found;
        this._classes.forEach(({ kr, kc }) => {
            const block = this._blocks.get(blockKey(kr, kc, row >> kr, col >> kc));
            if (block) {
                block.forEach((range) => {
                    if (
                        row >= range.top &&
                        row <= range.bottom &&
                        col >= range.left &&
                        col <= range.right &&
                        (!found || range.order < found.order)
                    ) {
                        found = range;
                    }
                });
            }
        });
        return found;
    }

    _eachBlock(range, callback) {
        const [kr, kc] = range.sizeClass;
        for (let rowBlock = range.top >> kr; rowBlock <= range.bottom >> kr; rowBlock++) {
            for (let colBlock = range.left >> kc; colBlock <= range.right >> kc; colBlock++) {
                callback(blockKey(kr, kc, rowBlock, colBlock));
            }
        }
    }
}

module.exports = RangeIndex;
