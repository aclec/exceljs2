# Changelog

Changes up to 4.4.0 are listed in the [Release History](README.md#release-history) section of the README.

Issue and pull request numbers (#NNNN) refer to the upstream repository, [exceljs/exceljs](https://github.com/exceljs/exceljs). Fixes were
re-implemented and tested here; upstream patches were used as bug reports only.

## 5.0.0

Read [UPGRADE-5.0.md](UPGRADE-5.0.md) before upgrading.

### Breaking changes

- Requires Node.js >= 22.12. Bun >= 1.4 is supported.
- The ES5 build (`exceljs/dist/es5`) was removed; the browser bundles target ES2020 and need no polyfills.
- Hardened parsing of untrusted files changes a few behaviours: grid bounds are enforced, merged cells are capped per worksheet, and large
  data-validation and defined-name ranges are kept as ranges. See [UPGRADE-5.0.md](UPGRADE-5.0.md#limits-when-reading-files).
- TypeScript typings corrected to match the runtime (for example `cell.col` / `cell.row` are numbers). See
  [UPGRADE-5.0.md](UPGRADE-5.0.md#typescript).

### Security

All fixes have regression tests in `spec/integration/security`.

- Merged cells: a merge covering the whole grid (`<mergeCell ref="A1:XFD1048576"/>`) exhausted the heap and aborted the process. The merged
  area is now capped at 1,048,576 cells per worksheet (#3114).
- Cell references: rows beyond 1,048,576 are rejected when decoding addresses, so that malformed references cannot become huge ranges
  (#3103, #3114).
- Merges: the overlap check was quadratic in the number of merges (CPU denial of service with many small merges); it is now proportional to
  the merged area (#2920).
- Data validations: a huge `sqref` was expanded to one object per cell (memory and CPU exhaustion). Large ranges are now kept as range keys;
  tokens that are not cell references are ignored (#2977).
- Defined names: a huge range was expanded to one entry per cell (memory exhaustion); ranges over 100,000 cells are kept as rectangles.
- Defined names: a sheet named `__proto__` in a defined name polluted `Object.prototype` (null-prototype maps are now used).
- Columns: `<col min max>` was not bounded, so `max="2000000000"` allocated two billion columns; it is clamped to the last Excel column.
- Images: parsing a drawing relationship `Target` used a regular expression with cubic complexity (ReDoS); it is now linear.
- Formulas over Excel's 8,192-character limit are rejected before quadratic regular-expression scans (shared-formula translation,
  `cell.dependencies`).
- `deepMerge` (used for cell notes) followed `__proto__`, `constructor` and `prototype` keys and polluted `Object.prototype` (#3083, #3101).
- CSV: cells named like `Object.prototype` members (`__proto__`, `constructor`, ...) were read as `Object.prototype` or functions instead of
  text (#3089).
- XML output: U+FFFE, U+FFFF and unpaired surrogates produced files that Excel and ExcelJS could not read back; they are stripped (#3113).
- Pivot tables: field names and values were written into XML without escaping (#2996).
- Streaming reader: temporary files (and their file descriptors) of deferred worksheets leaked when the iteration stopped early; the
  temporary directory is now always removed.

### Dependencies

- archiver ^5.0.0 → ^7.0.1
- dayjs ^1.8.34 → ^1.11.23
- fast-csv ^4.3.1 → ^5.0.8
- jszip ^3.10.1 → ^3.10.2
- saxes ^5.0.1 → ^6.0.0
- unzipper ^0.10.11 → ^0.12.5
- Removed `readable-stream` (replaced by `node:stream`), `tmp` (replaced by `fs.promises.mkdtemp`) and `uuid` (replaced by
  `crypto.randomUUID` / `crypto.getRandomValues`).
- Development: bun is the package manager (`bun.lock`); mocha 12, chai 6 and tsx replace mocha 7, chai 4 and ts-node; `bun audit` reports no
  known vulnerability. No dependency runs install scripts.

### Bug fixes

<!-- TODO(5.0.0): complete after merging the bug-fix branches. -->

- Streaming reader: worksheets and shared strings were silently dropped when the zip stream ended before all entries had been read (#2790,
  #2915, #3095).
- TypeScript typings:
    - `index.d.ts` references `@types/node`, so it compiles with TypeScript 6 and 7.
    - `xlsx.load()` and `Image.buffer` accept Node Buffers with TypeScript 5.7+.
    - `Address.col` / `Address.row` are numbers (#2720); `Row.dimensions` can be `null` (#2596); `Column.values` is an array and the
      `Row.values` getter returns an array (#2587).
    - `ImageEditAs` (#2983), all `cellIs` operators and optional rule `priority` (#2736), data bar colours (#2655), `DataValidationType`
      (#2697), optional `protect()` arguments (#2664), `null` table theme (#2061).
    - Streaming readers: optional `WorkbookReader` options (#1922), `WorksheetReader` properties and getters (#1901, #3097), EventEmitter
      methods.
    - The `csv.createInputStream()` typing was removed (it has thrown since 4.0). Comment typos (#3003, #3081).

### Documentation

- README: installation with any package manager, requirements, a "Security & limits" section, streaming reader defaults (#1531), tab colour
  example (#2577), formulas without a leading `=` (#3011).
- New [UPGRADE-5.0.md](UPGRADE-5.0.md) and [SECURITY.md](SECURITY.md).
