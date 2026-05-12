# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Vite + React app for tracking the FAO Haiti procurement pipeline (project **OSRO/HAI/061/CHA**, deadline 19 Oct 2026). The UI lives almost entirely in one large file, `procurement-pipeline-tracker.jsx`. State persists to a Neon Postgres database via two Vercel serverless functions, and there is a Claude-powered PDF report generator.

Production: <https://haiti-procurement-tracker.vercel.app> (password-gated, hard-coded password in the source).

## Commands

```bash
npm run dev        # Vite dev server on :5173 — /api/* is proxied to production (see vite.config.js)
npm run dev:full   # vercel dev — serves the API routes locally too; needs .env.local with DATABASE_URL + ANTHROPIC_API_KEY
npm run build      # Vite production build
npm run preview    # Preview the production bundle locally
```

Deploys run automatically when `main` is pushed (Vercel GitHub integration). Do **not** use `vercel --prod` from the CLI — it bypasses git and silently drifts production away from `origin/main`. We hit this exact issue mid-session and had to re-commit a wedge of work on `main`.

## Architecture

### Frontend (single file: `procurement-pipeline-tracker.jsx`)

Top to bottom, the layers are:

**Constants & stage definitions**
- `FAO_NAVY`, `FAO_BLUE`, `CATEGORY_COLORS` — brand tokens, used in inline styles.
- `PROJECT_END = 2026-10-19`, `DISTRIBUTION_DAYS = 45` (default; user-editable in the Gantt header).
- `ITB_STAGES_BASE` / `RFP_STAGES_BASE` — ordered workflow stages with default durations.
- `buildStages(method, value, existingStages, deliveryDays)` clones the right base, splices in RPC (≥ $200K) or HQPC (≥ $500K) review stages, appends a `Delivery of goods` stage when `deliveryDays > 0` (used for `Inputs - *` categories, 21 days), and merges any persisted stage data by key.
- `SEED` — the hardcoded starting list. It only seeds the **database** on first load; after that the DB is the source of truth.

**Pure helpers**
- `computeProcurement(p)` — derives `remainingDays`, `estPODate`, buffer vs. target, `risk` (`ok | watch | critical | neutral`).
- `computeTimeline(p)` — derives absolute `stageStart` / `stageEnd` per stage, anchored to real bid dates (`closing` → `opening` → `ganttAnchor` → today). Supports per-stage `stageStartOverride` pins that stop cascade.
- Date utilities: `parseDate`, `toISODate`, `fmtDate`, `daysBetween`, `addDays`, `fmtUSD`, `fmtUSDShort`.

**Presentational atoms**
`StatusDot`, `RiskPill`, `VarianceTag`, `MethodBadge`, `CategoryDot`, `Row`, `KPI`, `SyncBadge`, `LegendDot` — small, stateless, inline-styled.

**Feature panels**
- `CommentsPanel` — timestamped log; state lives in the parent.
- `StageEditor` — set stage status, planned days, names. Drag-reorder via HTML5 DnD; per-stage start/end dates; insert/delete custom steps; `resetDefaults` matches by stage `key` (not index) so reorder+reset is safe.
- `GanttChart` — segmented bar per procurement, draggable to shift dates; tooltip on hover; amber distribution-window zone (length editable in header); per-procurement orange hatched **Buffer / delay** zone toggled by a checkbox in the row label.

**Views (tabs)**
- Pipeline — sortable/filterable table; expandable detail rows with stage timeline + comments + stage editor.
- Gantt — see above.
- Upcoming — agenda-style next-step view per procurement, grouped by week, with horizon filter (30/60/90d/All).
- Analytics — Recharts bar and pie charts; uses `estInitial` (not `estPO`, which is mostly 0).

**Root component** (`App`, default export)
- `useState` for `procurements`, `distributionDays`, `syncStatus`, filters, sort, modal flags, etc.
- One `useEffect` loads from `/api/state` on mount. If the DB is empty it seeds with `initializeProcurements()`. Stage arrays are re-passed through `buildStages` so newly-added default keys (like `delivery`) merge into existing records.
- One `useEffect` writes to `/api/state` with 600 ms debounce on any change to `procurements` or `distributionDays`. `hydratedRef` prevents the first render from triggering a save.
- `SyncBadge` in the header surfaces sync state (`loading | saving | synced | error | offline`).

### Backend (Vercel serverless functions in `api/`)

**`api/state.js`** — Node runtime. GET returns the single `procurements` row from the `app_state` table (JSONB blob with `procurements: [...]` and `distributionDays`). PUT upserts the body into the same row. Uses `@neondatabase/serverless`.

**`api/generate-report.js`** — Edge runtime. POSTs a pipeline snapshot to the Anthropic API (model is `claude-sonnet-4-5-20250929`) and returns a narrative string used as the executive summary in the PDF. Requires `ANTHROPIC_API_KEY`.

### Database (Neon Postgres)

One table:

```sql
CREATE TABLE app_state (
  id TEXT PRIMARY KEY,            -- only row used: 'procurements'
  data JSONB NOT NULL,
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);
```

Neon project is `haiti-procurement-db` (free tier, region `iad1`), connected to this Vercel project only. Provisioned via the Vercel marketplace integration; env vars (`DATABASE_URL`, `POSTGRES_URL`, etc.) are injected automatically.

### PDF report (`buildReportPdf` inside `procurement-pipeline-tracker.jsx`)

Dynamic-imports `jspdf` + `jspdf-autotable` (v5 — v3 used a different API and crashed) and the IBM Plex Sans TTFs from `src/fonts/`. The fonts are fetched as ArrayBuffer, base64-encoded, and registered in jsPDF's VFS; falls back to Helvetica if the fetch fails. Layout: navy header band → row of 4 KPI cards → two-column row with a risk-distribution donut + top-categories horizontal bar chart → executive summary text → at-risk table → upcoming PO milestones table.

## Key domain concepts

- **Method**: `ITB` (Invitation to Bid, goods) vs `RFP` (Request for Proposals, services). Different stage sequences.
- **State**: `active | on_hold | cancelled | pre_pipeline | not_applicable` — controls risk pill rendering and initial stage status.
- **Stage status**: `complete | in_progress | not_started | blocked | skipped`.
- **Risk levels**: `ok` (≥ 7d buffer or no target), `watch` (1–6d), `critical` (overrun), `neutral` (cancelled / pre-pipeline / not-applicable).
- **Buffer vs Delay**: two distinct things. *Buffer* in `computeProcurement` is the calendar gap between estimated PO date and `targetPO` (drives risk colour). *Buffer/delay* on the Gantt is a user-toggleable orange hatched zone drawn after the procurement's last stage, configured per procurement via `bufferEnabled` + `bufferDays`.

## Extending the data

`SEED` only matters on first load; after that, edit through the UI and the changes persist to Neon automatically. If you need to bulk-edit, write a one-off script that connects with `@neondatabase/serverless` and updates the JSONB blob (see the sanctions/stage-timing migration we ran mid-session for the pattern).

Each record needs: `id`, `pr`, `lot`, `intend`, `tender`, `category`, `method`, `estInitial`, `estPO`, `nBids`, `opening`, `closing`, `currentKey` (must match a key in the stage base), `narrative`, `targetPO` (ISO date string or `""`), `state`, `comments`. Optional persisted fields used by the UI: `bufferEnabled`, `bufferDays`, `nResponsive`, custom `stages` array (overrides defaults from `buildStages`).

To add a new procurement category, add an entry to `CATEGORY_COLORS`. If the category name starts with `Inputs`, the 21-day delivery stage is auto-appended.

## Environment variables

| Variable | Where used | How it's set |
|---|---|---|
| `DATABASE_URL` (+ other Neon vars) | `api/state.js` | Auto-injected by the Neon marketplace integration on Vercel; pulled locally with `vercel env pull .env.local`. |
| `ANTHROPIC_API_KEY` | `api/generate-report.js` | Set on Vercel for production/preview/development; also kept in `.env.local` for `vercel dev`. **Never commit it.** |

`.env*.local` and `.DS_Store` are gitignored. The Vite dev server proxies `/api/*` to production so `npm run dev` keeps working without `vercel dev` running.

## Pitfalls and lessons from earlier sessions

- **jsPDF + autoTable API:** `jspdf-autotable` v3.x exposes `doc.autoTable(opts)` only; v5 exposes the functional form `autoTable(doc, opts)`. The code uses the v5 form — keep the dep pinned to `^5`.
- **Vite dep cache:** after upgrading a lazy-imported library (jsPDF, etc.), clear `node_modules/.vite/` and restart, otherwise the browser keeps loading the old bundled module.
- **`vercel dev` env vars:** `vercel dev` only auto-loads `.env.local` for variables that **also** exist on the Vercel project. For local-only secrets, also export them in the shell (`set -a; . ./.env.local; set +a; vercel dev`).
- **Hooks order:** the login gate (`if (!authed) return <LoginGate />`) must come **after** every `useState` / `useMemo` / `useEffect` in `App`, or you violate the rules of hooks and the post-login render goes white.
- **Avoid `vercel --prod` from CLI.** It deploys local files outside of git, drifting production from `origin/main`. Always merge to `main` and let the GitHub integration auto-deploy.
