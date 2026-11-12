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
- Behaviour changes when reading and writing: internal hyperlinks are read as hyperlink cells and written as a `location`; style setters
  copy shared styles instead of modifying them; CSV whitespace-only cells, large integers and `parserOptions.headers`; table header
  defaults; streaming `WorksheetReader.id` and `sharedStrings: 'emit'` event shape. See [UPGRADE-5.0.md](UPGRADE-5.0.md).

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

Upstream numbers refer to issues / pull requests of exceljs/exceljs that reported the bug. Every fix was re-implemented in this fork (no
upstream PR code was merged) and has a regression test.

- Files reported as corrupt by Excel:
    - worksheet elements written in `CT_Worksheet` order, by both writers (`legacyDrawing` #2079; streaming writer #2685, #2803); `sheetPr`
      children order (#2807); data bar extension colour order (#2651);
    - multiple print areas written as one `_xlnm.Print_Area` name, whole-column/row print areas, apostrophes in sheet names (#1516, #3071);
    - `theme1.xml` always written (#2009); comment VML shape ids unique per sheet, both writers (#3070); unique picture `creationId`
      (#3092); the right image when pictures repeat (#3094);
    - empty rich-text runs (#2737); distinct rich-text shared strings in the streaming writer (#2588); dxf number format written as
      `[object Object]` (#2698);
    - internal hyperlinks (`#Sheet!A1`, `#DefinedName`) written as a `location` without an external relationship, by both writers (#2002);
      `inset="NaNmm,..."` written for notes whose VML textbox had no inset.
- Data lost or wrong on read:
    - rows/cells without an `r` attribute (#2962, both readers); `t="d"` ISO 8601 date cells (#2702, both readers);
    - falsy formula results (#3048); only numeric formula results converted to dates (#3111);
    - leading plain text of mixed rich-text strings (#3091); explicit `val="0"`/`"false"` boolean attributes (#2095, #2851, #3057);
    - hyperlink `location` (#2002); table columns after a column with child elements (#1936); table `filterColumn` matched by `colId`
      (#2978); filter buttons kept visible on round trip; `headerRowCount` default and no `autoFilter` without a header row (#1345); x14
      `expression`/`cellIs` conditional formats (#1767); `stopIfTrue` (#2736);
    - column width 9 when the sheet default differs, also when set after `worksheet.columns` (#3088); sheet names longer than 31 characters
      (#3087);
    - both readers parse `t="d"` values with the same rules (`T24:00`, offsets, years 0-99); hyperlinks on number or boolean cells keep
      their value type on write.
- Files that failed to load: unusual `docProps` parts (#2846); relationships to missing or absolute-path parts and VML notes (#3051, #3105);
  relationships without a `Target` (drawing, background image, worksheet); unmodelled table filter elements such as `dynamicFilter` (#2973);
  bare data bars and rules without formulae on write (#3019, #2883).
- Streaming reader: rejects instead of hanging or crashing on truncated/corrupt input and always removes its temp files (#2148); waits for
  `workbook.xml`/styles before parsing sheets stored earlier in the zip (#2790, #2915, #3095); UTF-8 characters split across chunks (#1665,
  #2894, #3102); column definitions, hidden rows, sheet format properties (#2800); absolute relationship targets (#2671); options passed to
  `read()`/`parse()` merged with the defaults; `sharedStrings: 'emit'` emits again.
- Document model: setting a style property on a cell no longer restyles every cell sharing a loaded style, while merged cells still share
  the master style (#2055, #2781); `spliceRows` stale rows (#1600); column keys / table names / CSV headers named like `Object.prototype`
  members (#2080), with `worksheet.tables` kept a plain object; assigning the model of a merged cell keeps its master, and releasing a merge
  reference twice no longer throws.
- CSV: whitespace-only cells stay strings (#2752); integers beyond `Number.MAX_SAFE_INTEGER` stay strings (#1743); `parserOptions.headers`
  support (#2080).

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
