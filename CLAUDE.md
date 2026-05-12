# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-file React dashboard (`procurement-pipeline-tracker.jsx`) for tracking FAO Haiti procurement pipelines. It is a self-contained component with no build system or package.json — it is designed to be dropped into a React environment (e.g., a CodeSandbox, Next.js page, or similar) that already provides React, `lucide-react`, and `recharts`.

## Architecture

Everything lives in one ~1 500-line file. The key layers, top to bottom:

**Constants & stage definitions**
- `FAO_NAVY`, `FAO_BLUE`, `CATEGORY_COLORS` — brand tokens used throughout inline styles.
- `ITB_STAGES_BASE` / `RFP_STAGES_BASE` — ordered arrays of workflow stages with default durations. `buildStages(method, value, existingStages)` clones the right base, splices in extra review stages at value thresholds (≥$200K → RPC, ≥$500K → HQPC), and merges any persisted stage data.
- `SEED` — the hardcoded list of procurement records. This is the only data store; there is no backend or localStorage.

**Pure helpers**
- `computeProcurement(p)` — derives `remainingDays`, `estPODate`, buffer vs. target, and a `risk` object (`ok / watch / critical / neutral`) from a procurement record's stage array.
- Date utilities: `parseDate`, `toISODate`, `fmtDate`, `daysBetween`, `addDays`, `fmtUSD`, `fmtUSDShort`.

**Presentational atoms**
`StatusDot`, `RiskPill`, `VarianceTag`, `MethodBadge`, `CategoryDot`, `Row`, `KPI` — small, stateless display components using inline styles (no CSS files).

**Feature panels (stateful sub-components)**
- `CommentsPanel` — timestamped comment log with add/delete; state lives in parent.
- `StageEditor` — lets users set stage status and planned days; cascades "complete" status backwards through the pipeline.

**Views (rendered by the root component)**
- List view — sortable/filterable table of all procurements with expandable detail rows (stage timeline + comments + stage editor).
- Gantt view — horizontal bar chart per procurement showing remaining stages as colored segments.
- Analytics view — Recharts bar and pie charts summarising value by category and risk distribution.

**Root component** (`export default function ProcurementPipeline`)
- `useState` holds the full `procurements` array (initialized from `SEED` via `initializeProcurements()`).
- `useMemo` recomputes derived stats whenever `procurements` changes.
- `activeView` switches between "list", "gantt", and "analytics" tabs.
- Filtering state: `filterCategory`, `filterState`, `filterMethod`, `searchQuery`, `sortKey`, `sortDir`.
- `updateProcurement(updated)` is the single updater passed down to all child panels.

## Key domain concepts

- **Method**: `ITB` (Invitation to Bid, goods) vs `RFP` (Request for Proposals, services). Different stage sequences.
- **State**: `active | on_hold | cancelled | pre_pipeline | not_applicable` — controls risk pill rendering and stage status initialisation.
- **Stage status**: `complete | in_progress | not_started | blocked | skipped`.
- **Buffer**: calendar days between estimated PO date (sum of remaining planned days from today) and the `targetPO` field. Negative = overrun.

## Extending the data

To add or edit procurements, modify the `SEED` array (lines ~127–146). Each record needs: `id`, `pr`, `lot`, `intend`, `tender`, `category`, `method`, `estInitial`, `estPO`, `nBids`, `opening`, `closing`, `currentKey` (must match a key in the stage base), `narrative`, `targetPO` (ISO date string or `""`), `state`, `comments`.

To add a new procurement category, add an entry to `CATEGORY_COLORS`.
