# Roadmap: scoped `src/lib/**` line coverage → **60%** (excluding `ews-client.ts`)

This document is the **working plan** for milestone **M2** in the broader 80% program. It matches how we measure coverage today: `bun run test:coverage` → `coverage/lcov.info` → `node scripts/check-coverage-lib.mjs` (or `bun run verify:coverage:lib` with `COVERAGE_MIN_LINES_LIB`).

## 1. Scope and definitions

| Item | Detail |
|------|--------|
| **Included** | All instrumented lines under `src/lib/` per `SF:` records in `lcov.info` |
| **Excluded** | `src/lib/ews-client.ts` (SOAP; regression tests stay, but not in this gate) |
| **Excluded** | `src/lib/**/*.test.ts` (if they appear in lcov) |
| **Gate env** | `COVERAGE_MIN_LINES_LIB` in CI (`ci.yml` / `release.yml`) |
| **Backlog tool** | `bun run report:coverage:lib-gaps` (optionally `--top 25`) |

## 2. Baseline snapshot (run locally to refresh)

After `bun run test:coverage`:

```bash
COVERAGE_MIN_LINES_LIB=45 bun run verify:coverage:lib
node scripts/report-lib-coverage-gaps.mjs --top 20
```

**Interpretation:** Percentages move when source or tests change. Treat the **printed LH/LF** as the source of truth before bumping CI.

**Immediate hygiene:** If scoped % prints **below** the current CI floor (e.g. 45%) after a dependency or code change, either (a) add a small targeted test, or (b) set `COVERAGE_MIN_LINES_LIB` to `floor(measured %)` until back above the milestone—**never** leave main red without a tracked decision.

## 3. Target math (order-of-magnitude)

Let \(LF\) = total instrumented lines in scope, \(LH\) = hit lines. Target **60%** means:

\[
LH_{60} \approx \lceil 0.60 \times LF \rceil
\qquad
\Delta \approx LH_{60} - LH_{\text{now}}
\]

Example: if \(LF \approx 12.8\text{k}\) and \(LH \approx 5.75\text{k}\) (~45%), then \(\Delta \approx 1.9\text{k}\) **new hit lines**—a multi-PR effort. Milestones below split that into **trackable chunks** (~250–400 hit lines per step is a reasonable PR size if focused on one client module).

## 4. Trackable milestones (M2.0 → M2.4 → **60%**)

Use **integer** CI floors so the gate is unambiguous. After each milestone, update **both** workflows:

- `.github/workflows/ci.yml` — `COVERAGE_MIN_LINES_LIB`
- `.github/workflows/release.yml` — same

| ID | Gate `COVERAGE_MIN_LINES_LIB` | Purpose | Exit criteria |
|----|------------------------------|---------|----------------|
| **M2.0** | **45** (or **44** if baseline dipped—reconcile first) | Lock M1 / no backsliding | CI green; gap report archived in PR or run log |
| **M2.1** | **48** | First push into M2 | `verify:coverage:lib` green at 48; top gaps reviewed |
| **M2.2** | **52** | Mid-ramp | Green at 52; Todo/Outlook/Planner tests visibly expanded |
| **M2.3** | **56** | High line-yield clients | Green at 56; Excel + Copilot invoke batches progressed |
| **M2.4** | **60** | **M2 complete** | Green at **60**; document remaining gap to 80% for M3 |

**Rule:** Only raise the CI number when `test:coverage` + `verify:coverage:lib` pass locally **and** you intend to keep the branch green (no “step” jumps of more than ~4–5 points unless a large batch lands in one PR).

## 5. Workstreams (priority order)

Order by **uncovered lines × feasibility** (thin `fetch` wrappers first; avoid `graph-auth` until necessary).

### A. Large Graph clients (highest `LF − LH`)

1. **`todo-client.ts`** — extend existing [`todo-client.test.ts`](../src/lib/todo-client.test.ts): per-export success + Graph error JSON + one failure branch where cheap.
2. **`outlook-graph-client.ts`** — same pattern as [`outlook-graph-client.test.ts`](../src/lib/outlook-graph-client.test.ts).
3. **`planner-client.ts`** — same as [`planner-client.test.ts`](../src/lib/planner-client.test.ts).
4. **`copilot-graph-client.ts`** — keep **`mock.module` last** in the run order (e.g. `z-*.invoke.test.ts`); extend invoke suite for remaining wrappers.
5. **`graph-excel-client.ts`** — continue table-style `fetch` routing tests (many similar endpoints).

### B. Shared infrastructure

6. **`graph-client.ts`** — targeted tests for uncovered branches (pagination, retries, errors) in [`graph-client.test.ts`](../src/lib/graph-client.test.ts); avoid duplicating entire CLI flows.
7. **`graph-calendar-client.ts`** — remaining attachment / upload-session paths if still in gap list after bulk suite.

### C. Small high-leverage modules (quick wins between milestones)

8. **`webhook-server.ts`**, **`graph-subscriptions.ts`**, **`graph-directory.ts`**, **`places-client.ts`**, **`graph-event.ts`** — small files: few focused tests move % noticeably.

### D. Defer or isolate

- **`graph-auth.ts`** — interactive / env-heavy; schedule after A–C unless a safe `fetch`/cache mock pattern already exists.
- **`ews-client.ts`** — out of scope for this gate (per product decision).

## 6. Process (how we track work)

1. **Before a batch** — Run gap report; pick **one primary file** (or two small ones) per PR.
2. **PR checklist**
   - `bun run test`
   - `bun run test:coverage`
   - `COVERAGE_MIN_LINES_LIB=<next_floor> bun run verify:coverage:lib` (only if raising the gate in the same PR)
   - `bun run biome:check` on touched files (or full repo if policy requires)
3. **When raising CI** — Same PR must include workflow env bump **or** a follow-up PR immediately after merge (prefer same PR to avoid red main).
4. **Optional** — Paste `report:coverage:lib-gaps --top 15` into PR description for traceability.

## 7. Risks

| Risk | Mitigation |
|------|------------|
| Bun / OS lcov drift | Re-run coverage on Linux (CI) before arguing about 0.1% |
| `mock.module` leakage | Copilot-style suites stay in **`z-*`** files; prefer `fetch` mocks elsewhere |
| LF changes (new code) | Recompute milestone thresholds from fresh `lcov`; don’t chase old hit counts |

## 8. Progress note (2026-05-05)

**M2.4 (60%) is not complete yet.** After the coverage push in this window, scoped lib line coverage (same definition as §1) measured **~50.1%** — **8734** hit lines of **17417** total (excluding `ews-client.ts`). CI `COVERAGE_MIN_LINES_LIB` was raised to **50** (from 45) because `verify:coverage:lib` is green at that floor.

Largest remaining gaps (refresh with `bun run report:coverage:lib-gaps --top 15`): `copilot-graph-client.ts`, `graph-client.ts`, `outlook-graph-client.ts`, `planner-client.ts`, `graph-excel-client.ts`, `todo-client.ts`. Next step toward **60%** is more `fetch`/`callGraph` tests on those modules; `check-coverage-lib.mjs` now merges duplicate `SF:` blocks by taking the **maximum** hit count per source line when Bun emits multiple records for one file.

## 9. After 60%

Hand off to **M3** (60% → 80%): deeper `graph-client` + auth + edge cases; see the in-repo coverage scripts and [`check-coverage-lib.mjs`](../scripts/check-coverage-lib.mjs) (default min **80** when env unset).
