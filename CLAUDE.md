# CLAUDE.md — exceljs (maintained fork, v5)

Read/write `.xlsx` and `.csv`. Fork of the unmaintained `exceljs/exceljs` (upstream frozen at `5bed18b`, 2024-01). Must run in **browsers**
(bundled `dist/`), **Node >= 22.12** and **Bun >= 1.4**. CommonJS source, no transpilation of `lib/`.

## Commands (bun only — never npm/yarn; no lockfile other than `bun.lock`)

| Task                   | Command                                                                                          |
| ---------------------- | ------------------------------------------------------------------------------------------------ |
| install                | `bun install --frozen-lockfile` (set `HUSKY=0` in worktrees/CI)                                  |
| tests, Node runtime    | `bun run test` (unit + integration + end-to-end)                                                 |
| tests, Bun runtime     | `bun run test:bun`                                                                               |
| single spec            | `bunx mocha --require spec/config/setup.js <file>`                                               |
| browser bundle / tests | `bun run build` then `bun run test:browser` (local Chrome, temp profile) and `bun run test:dist` |
| types                  | `bun run test:types` (tsc on `index.d.ts` + `spec/typescript`) and `bun run test:typescript`     |
| quality gates          | `bun run lint`, `bun run format:check`, `bun run audit`, `bun run check:install-scripts`         |

Done = all of the above green. CI (`.github/workflows/tests.yml`) runs the same on Node 22/24/26, Bun, Chrome.

## Layout

- `lib/doc/` — in-memory model: `workbook`, `worksheet`, `row`, `column`, `cell`, `range`, `table`, `defined-names`, `data-validations`.
- `lib/xlsx/xform/**` — XML ⇄ model. Each xform: `prepare` → `render(xmlStream, model)` on write; `parseOpen/parseText/parseClose` (SAX) →
  `reconcile` on read. `lib/xlsx/xlsx.js` orchestrates the zip parts. Element ORDER in rendered XML must follow ECMA-376 (e.g.
  `CT_Worksheet`), otherwise Excel reports a corrupt file — tests assert raw string order (chai-xml ignores sibling order).
- `lib/stream/xlsx/**` — streaming `WorkbookReader`/`WorkbookWriter`: **Node/Bun only** (fs). Its worksheet writer must keep the same
  element order as `worksheet-xform.js` render.
- `lib/csv/` (fast-csv), `lib/utils/` (col-cache, xml-stream, parse-sax with streaming `TextDecoder`, stream-buf, uuid).
- Entry points: `lib/exceljs.nodejs.js` (Node/Bun), `lib/exceljs.browser.js` / `lib/exceljs.bare.js` (bundles via `scripts/build.mjs`,
  esbuild + `scripts/browser-shims/`). Public types: `index.d.ts` (keep in sync with runtime).
- `spec/unit`, `spec/integration` (incl. `security/` hostile-input regressions), `spec/end-to-end`, `spec/browser`, `spec/typescript`.
- Code map (graphify, local and gitignored, scope = `lib/` via `.graphifyignore`): read `graphify-out/GRAPH_REPORT.md` before exploring
  unfamiliar areas; `graphify query "<question>"`, `graphify path "A" "B"`, `graphify explain "X"` for targeted lookups.
  `.husky/post-commit` and `post-checkout` rebuild it in the background. **Before relying on it, check freshness**: if `graphify-out/` is
  missing, or the report's "Built from commit" differs from `git rev-parse --short HEAD` (e.g. commits made with `HUSKY=0`), run
  `graphify update .` (AST only, a few seconds, no LLM).

## Invariants (do not break)

- **Untrusted input**: parsing must stay linear in input size. Bounds: rows ≤ 1,048,576, cols ≤ 16,384 (`colCache.MAX_*`), merged cells per
  sheet ≤ 1,048,576, formulas ≤ 8,192 chars; huge data-validation / defined-name ranges stay as rectangles (never expand per cell).
- Objects keyed by file/CSV data use `Object.create(null)` or `Map`; never `obj[key] =` with a raw data key on `{}`.
- XML built by string templates must go through `utils.xmlEncode` (strips chars invalid in XML 1.0). Prefer `xmlStream`.
- No backtracking-prone regex on attacker-controlled strings (anchor, or use `indexOf`/`slice`).
- Browser isolation: lib code outside `lib/stream/**` and the explicit exceptions in `eslint.config.js` must not require
  `fs`/`path`/`os`/`crypto`/`child_process`/network modules (lint-enforced).
- Supply chain: no install lifecycle scripts, `trustedDependencies` stays `[]`, actions pinned by SHA, `bunfig.toml` minimumReleaseAge.
  Never add a dependency without asking the maintainer.

## Upstream PRs policy

Upstream PR heads can be fetched as `upstream-pr/<N>` (`git fetch upstream '+refs/pull/*/head:refs/remotes/upstream-pr/*'`). They are **bug
reports only**: never merge, cherry-pick or `git apply` them, never copy their binary fixtures (several were hostile: package renames, fake
metadata, SSRF, postinstall hooks). Re-implement the fix, write the regression test first (must fail before the fix), cite
`Reported upstream: #N` in the commit body.

## Conventions

- Prettier config is the maintainer's (4 spaces, double quotes, width 140); ESLint = quality/security rules only. Pre-commit hook runs
  lint-staged.
- Tests: mocha + chai (`expect`, dirty-chai). Build xlsx fixtures in code with `spec/utils/xlsx-tamper.js` (`baseWorkbookBuffer` / `tamper`
  / `readPart`) — readable, reviewable, no new binary fixtures. Tests write only to `spec/out/` or a removed `mkdtemp` dir; network only on
  127.0.0.1.
- Breaking changes go to `UPGRADE-5.0.md` + `CHANGELOG.md`; security model in `SECURITY.md`.
- Session memory (local, gitignored): `tasks/todo.md` (plan/state), `tasks/lessons.md` (append `[YYYY-MM-DD] | what went wrong | rule`).
