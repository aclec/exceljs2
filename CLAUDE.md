# CLAUDE.md — @aclec/exceljs2 (maintained fork of exceljs, v5)

Published on npm as `@aclec/exceljs2` (`exceljs` on npm is the dead upstream). Read/write `.xlsx` and `.csv`. Fork of the unmaintained
`exceljs/exceljs` (upstream frozen at `5bed18b`, 2024-01). Must run in **browsers** (bundled `dist/`), **Node >= 22.12** and **Bun >= 1.4**.
TypeScript sources in `src/`; `lib/` is build output (tsc → CommonJS + `.d.ts`, gitignored), never edit it.

## Commands (bun only — never npm/yarn; no lockfile other than `bun.lock`)

| Task                   | Command                                                                                          |
| ---------------------- | ------------------------------------------------------------------------------------------------ |
| install                | `bun install --frozen-lockfile` (set `HUSKY=0` in worktrees/CI)                                  |
| compile src/ → lib/    | `bun run build:lib` (tsc -p tsconfig.build.json; `test`, `test:bun`, `test:typescript` run it)   |
| tests, Node runtime    | `bun run test` (unit + integration + end-to-end)                                                 |
| tests, Bun runtime     | `bun run test:bun`                                                                               |
| single spec            | `bun run build:lib` then `bunx mocha --require spec/config/setup.js <file>`                      |
| browser bundle / tests | `bun run build` then `bun run test:browser` (local Chrome, temp profile) and `bun run test:dist` |
| types                  | `bun run test:types` (tsc on `src/`, then `index.d.ts` + `spec/typescript`), `test:typescript`   |
| performance            | `bun run benchmark:suite` (`--lib <other>/lib` to compare two builds)                            |
| quality gates          | `bun run lint`, `check:tsdoc`, `check:readme`, `format:check`, `audit`, `check:install-scripts`  |

Done = all of the above green. CI (`.github/workflows/tests.yml`) runs the same on Node 22/24/26, Bun, Chrome.

## Layout

- `src/doc/` — in-memory model: `workbook`, `worksheet`, `row`, `column`, `cell`, `range`, `table`, `defined-names`, `data-validations`.
- `src/xlsx/xform/**` — XML ⇄ model. Each xform: `prepare` → `render(xmlStream, model)` on write; `parseOpen/parseText/parseClose` (SAX) →
  `reconcile` on read. `src/xlsx/xlsx.ts` orchestrates the zip parts. Element ORDER in rendered XML must follow ECMA-376 (e.g.
  `CT_Worksheet`), otherwise Excel reports a corrupt file — tests assert raw string order (chai-xml ignores sibling order).
- `src/stream/xlsx/**` — streaming `WorkbookReader`/`WorkbookWriter`: **Node/Bun only** (fs). Its worksheet writer must keep the same
  element order as `worksheet-xform.ts` render.
- `src/csv/` (fast-csv), `src/utils/` (col-cache, xml-stream, parse-sax with streaming `TextDecoder`, stream-buf, uuid).
- Entry points: `src/exceljs.nodejs.ts` (Node/Bun), `src/exceljs.browser.ts` / `src/exceljs.bare.ts` (bundled from `src/` by
  `scripts/build.mjs`, esbuild + `scripts/browser-shims/`). Entry points use `export =` so `require()` keeps its CommonJS shape. Public
  types: `index.d.ts` (hand written, the API contract; keep in sync with runtime). Internal shared types: `src/types/`.
- `spec/unit`, `spec/integration` (incl. `security/` hostile-input regressions), `spec/end-to-end`, `spec/browser`, `spec/typescript`.
- Code map (graphify, local and gitignored, scope = `src/` via `.graphifyignore`): read `graphify-out/GRAPH_REPORT.md` before exploring
  unfamiliar areas; `graphify query "<question>"`, `graphify path "A" "B"`, `graphify explain "X"` for targeted lookups.
  `.husky/post-commit`, `post-merge` and `post-checkout` rebuild it in the background. **Before relying on it, check freshness**: if
  `graphify-out/` is missing, or the report's "Built from commit" differs from `git rev-parse --short HEAD` (e.g. commits made with
  `HUSKY=0`), run `graphify update .` (AST only, a few seconds, no LLM).

## Invariants (do not break)

- **Untrusted input**: parsing must stay linear in input size. Bounds: rows ≤ 1,048,576, cols ≤ 16,384 (`colCache.MAX_*`), merged cells per
  sheet ≤ 1,048,576, formulas ≤ 8,192 chars; huge data-validation / defined-name ranges stay as rectangles (never expand per cell).
- Objects keyed by file/CSV data use `Object.create(null)` or `Map`; never `obj[key] =` with a raw data key on `{}`.
- XML built by string templates must go through `utils.xmlEncode` (strips chars invalid in XML 1.0). Prefer `xmlStream`.
- No backtracking-prone regex on attacker-controlled strings (anchor, or use `indexOf`/`slice`).
- Browser isolation: src code outside `src/stream/**` and the explicit exceptions in `eslint.config.js` must not require
  `fs`/`path`/`os`/`crypto`/`child_process`/network modules (lint-enforced).
- Supply chain: no install lifecycle scripts, `trustedDependencies` stays `[]`, actions pinned by SHA, `bunfig.toml` minimumReleaseAge.
  Never add a dependency without asking the maintainer.

## Repository scope — we work for this fork only

- **Our repository is `origin` = `git@github.com:aclec/exceljs2.git`; default branch `master`.** All work, branches, commits, pushes, issues
  and releases happen there, for this project only.
- The original `exceljs/exceljs` is **unmaintained and is not our project**: never push, open issues/PRs, comment, or send anything to it.
  The `upstream` remote is **read-only** (its push URL is set to `no_push_to_upstream`); it exists only to read old bug reports.
- Publishing: npm package `@aclec/exceljs2` (never `exceljs`). No GitHub CLI is used: integrate with plain git (`git merge --no-ff` into
  `master`, `git push origin`).

## Upstream PRs policy

Upstream PR heads can be fetched as `upstream-pr/<N>` (`git fetch upstream '+refs/pull/*/head:refs/remotes/upstream-pr/*'`). They are **bug
reports only**: never merge, cherry-pick or `git apply` them, never copy their binary fixtures (several were hostile: package renames, fake
metadata, SSRF, postinstall hooks). Re-implement the fix, write the regression test first (must fail before the fix), cite
`Reported upstream: #N` in the commit body.

## TypeScript

- tsconfig.base.json: strict + noUncheckedIndexedAccess + noImplicitOverride + noImplicitReturns + noFallthroughCasesInSwitch, target and
  lib ES2022 (no newer runtime APIs: the browser bundle has no polyfills). ESLint (typescript-eslint, type-aware) bans `any`, non-null (`!`)
  and unsafe type assertions: narrow with checks (clear `Error` on invalid input), or `// eslint-disable-next-line <rule> -- why`.
- `index.d.ts` is checked against `src/` at compile time (`src/types/doc-conformance.ts`, `stream-conformance.ts`): a mismatch fails `tsc`;
  an exception must state its reason. README code blocks are type-checked (`check:readme`).
- Modules use `import`/`export` (compiled to CommonJS). Public types are imported from the root `index.d.ts`
  (`import type { Style } from "../../index"`).
- Class fields use [[Define]] semantics: for a property set by a base constructor use `declare foo: T`, not `foo: T`.
- TSDoc on every exported symbol and public member (`bun run check:tsdoc`). Performance changes are measured with `bun run benchmark:suite`.

## Conventions

- Prettier config is the maintainer's (4 spaces, double quotes, width 140); ESLint = quality/security rules only. Pre-commit hook runs
  lint-staged.
- Tests: mocha + chai (`expect`, dirty-chai). Build xlsx fixtures in code with `spec/utils/xlsx-tamper.js` (`baseWorkbookBuffer` / `tamper`
  / `readPart`) — readable, reviewable, no new binary fixtures. Tests write only to `spec/out/` or a removed `mkdtemp` dir; network only on
  127.0.0.1.
- Breaking changes go to `UPGRADE-5.0.md` + `CHANGELOG.md`; security model in `SECURITY.md`.
- Session memory (local, gitignored): `tasks/todo.md` (plan/state), `tasks/lessons.md` (append `[YYYY-MM-DD] | what went wrong | rule`).
