import React, { useState, useMemo, useEffect, useCallback, useRef } from "react";
import {
  ChevronDown,
  ChevronRight,
  CheckCircle2,
  Circle,
  Clock,
  Ban,
  PauseCircle,
  TrendingUp,
  TrendingDown,
  Minus,
  Filter,
  RotateCcw,
  MessageSquare,
  Send,
  List as ListIcon,
  GanttChartSquare,
  BarChart3,
  ArrowDown,
  ArrowUp,
  Plus,
  X,
  Download,
  GripVertical,
  CalendarDays,
  ChevronUp,
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
} from "recharts";

const FAO_NAVY = "#1A2E44";
const FAO_BLUE = "#009FDA";

const PROJECT_END = new Date("2026-10-19T00:00:00");
const DISTRIBUTION_DAYS = 45; // goods must be distributed within this window before project end
const DISTRIBUTION_START = new Date(PROJECT_END);
DISTRIBUTION_START.setDate(DISTRIBUTION_START.getDate() - DISTRIBUTION_DAYS);

const fontStack = {
  display: "'IBM Plex Serif', Georgia, serif",
  body: "'IBM Plex Sans', system-ui, sans-serif",
  mono: "'IBM Plex Mono', 'SF Mono', Menlo, monospace",
};

const CATEGORY_COLORS = {
  Livestock: "#0F766E",
  "Inputs - Seeds": "#65A30D",
  "Inputs - Feed": "#CA8A04",
  "Inputs - Fertilizer": "#9333EA",
  "Equipment - Agri": "#EA580C",
  "Equipment - ICT": "#0284C7",
  "Services - M&E": "#DB2777",
  Visibility: "#6B7280",
};

const ITB_STAGES_BASE = [
  { key: "pr_received", name: "PR received & validated", defaultDays: 2 },
  { key: "market_analysis", name: "Sourcing & market analysis", defaultDays: 7 },
  { key: "solicitation_prep", name: "Solicitation document prep", defaultDays: 5 },
  { key: "published", name: "Solicitation published (InTend)", defaultDays: 21 },
  { key: "clarifications", name: "Clarifications period", defaultDays: 0 },
  { key: "closing", name: "Bid closing & opening", defaultDays: 1 },
  { key: "tech_eval", name: "Technical evaluation", defaultDays: 10 },
  { key: "fin_eval", name: "Financial evaluation", defaultDays: 3 },
  { key: "sanctions", name: "Sanctions screening", defaultDays: 2 },
  { key: "lpc_award", name: "LPC award review", defaultDays: 7 },
  { key: "award", name: "Award decision & notification", defaultDays: 3 },
  { key: "po_create", name: "PO creation (GRMS)", defaultDays: 3 },
  { key: "po_signed", name: "PO signed & issued", defaultDays: 2 },
];

const RFP_STAGES_BASE = [
  { key: "pr_received", name: "PR received & validated", defaultDays: 2 },
  { key: "market_analysis", name: "Sourcing & market analysis", defaultDays: 7 },
  { key: "solicitation_prep", name: "Solicitation document prep", defaultDays: 7 },
  { key: "lpc_exante", name: "LPC ex-ante review (criteria)", defaultDays: 7 },
  { key: "published", name: "Solicitation published (InTend)", defaultDays: 30 },
  { key: "clarifications", name: "Clarifications period", defaultDays: 0 },
  { key: "closing", name: "Bid closing & opening", defaultDays: 1 },
  { key: "tech_eval", name: "Technical evaluation (env. 1)", defaultDays: 12 },
  { key: "fin_eval", name: "Financial evaluation (env. 2)", defaultDays: 4 },
  { key: "combined", name: "Combined scoring & ranking", defaultDays: 2 },
  { key: "sanctions", name: "Sanctions screening", defaultDays: 2 },
  { key: "lpc_award", name: "LPC award review", defaultDays: 7 },
  { key: "award", name: "Award decision & notification", defaultDays: 3 },
  { key: "negotiation", name: "Contract negotiation", defaultDays: 5 },
  { key: "contract_signed", name: "Contract signed", defaultDays: 3 },
];

function buildStages(method, value, existingStages) {
  const base = method === "RFP" ? RFP_STAGES_BASE : ITB_STAGES_BASE;
  let stages = base.map((s) => ({ ...s }));
  if (value >= 500000) {
    const idx = stages.findIndex((s) => s.key === "lpc_award");
    stages.splice(idx + 1, 0, { key: "hqpc", name: "HQPC review (≥ $500K)", defaultDays: 7 });
  } else if (value >= 200000) {
    const idx = stages.findIndex((s) => s.key === "lpc_award");
    stages.splice(idx + 1, 0, { key: "rpc", name: "RPC review ($200K–$500K)", defaultDays: 7 });
  }
  if (existingStages) {
    stages = stages.map((s) => {
      const ex = existingStages.find((e) => e.key === s.key);
      return ex
        ? { ...s, plannedDays: ex.plannedDays ?? s.defaultDays, status: ex.status ?? "not_started" }
        : { ...s, plannedDays: s.defaultDays, status: "not_started" };
    });
  } else {
    stages = stages.map((s) => ({ ...s, plannedDays: s.defaultDays, status: "not_started" }));
  }
  return stages;
}

function markProgressTo(stages, currentKey, currentStatus = "in_progress") {
  let reached = false;
  return stages.map((s) => {
    if (s.key === currentKey) {
      reached = true;
      return { ...s, status: currentStatus };
    }
    if (!reached) return { ...s, status: "complete" };
    return { ...s, status: "not_started" };
  });
}

const SEED = [
  // ── Active solicitations ────────────────────────────────────────────────
  { id: "4201282",    pr: "4201282", lot: "",      intend: "2026/FLHAI/FLHAI/136065", tender: "Baseline & Endline survey (TDR enquêtes)",           category: "Services - M&E",     method: "RFP", estInitial: 373412,  estPO: 0,  nBids: 1,    opening: "2026-04-02", closing: "2026-04-09", currentKey: "lpc_award",         narrative: "awarding LPC in progress",                                                                                                             targetPO: "2026-06-15", state: "active",       comments: [] },
  { id: "4201270",    pr: "4201270", lot: "",      intend: "2026/FLHAI/FLHAI/136142", tender: "Smartphones & power banks",                           category: "Equipment - ICT",    method: "ITB", estInitial: 40000,   estPO: 0,   nBids: 15,   opening: "2026-03-17", closing: "2026-03-31", currentKey: "solicitation_prep", narrative: "awaiting modified PR and confirmation on the way forward related to quantities",                                                       targetPO: "2026-07-15", state: "on_hold",      comments: [] },
  { id: "4201263",    pr: "4201263", lot: "",      intend: "2026/FLHAI/FLHAI/135970", tender: "Fertilizers UREA & NPK",                              category: "Inputs - Fertilizer",method: "ITB", estInitial: 39900,   estPO: 0,   nBids: 4,    opening: "2026-04-03", closing: "2026-04-09", currentKey: "tech_eval",         narrative: "Technical evaluation ongoing, one clarification is missing",                                                                            targetPO: "2026-06-15", state: "active",       comments: [] },
  { id: "4201264",    pr: "4201264", lot: "",      intend: "2026/FLHAI/FLHAI/135971", tender: "Vegetable seeds",                                     category: "Inputs - Seeds",     method: "ITB", estInitial: 742516,  estPO: 0,  nBids: 13,   opening: "2026-04-07", closing: "2026-04-24", currentKey: "tech_eval",         narrative: "offers sent for technical evaluation on 24.04.2026, technical questions shared with bidders, last deadline 13.5.2026",                   targetPO: "2026-07-15", state: "active",       comments: [] },
  { id: "4201265-L1", pr: "4201265", lot: "Lot 1", intend: "2026/FLHAI/FLHAI/135992", tender: "Goats — males & females (boucs, chèvres)",            category: "Livestock",          method: "ITB", estInitial: 3836700, estPO: 0, nBids: 7,    opening: "2026-04-20", closing: "2026-05-04", currentKey: "tech_eval",         narrative: "offers sent for technical evaluation on 5.5.2026",                                                                                      targetPO: "2026-08-01", state: "active",       comments: [] },
  { id: "4201265-L2", pr: "4201265", lot: "Lot 2", intend: "2026/FLHAI/FLHAI/135993", tender: "Rabbits (lapins)",                                    category: "Livestock",          method: "ITB", estInitial: 984375,  estPO: 0,  nBids: null, opening: "2026-04-20", closing: "2026-05-08", currentKey: "tech_eval",         narrative: "offers sent for tech. Evaluation 8.5.2026",                                                                                             targetPO: "2026-08-01", state: "active",       comments: [] },
  { id: "4201266",    pr: "4201266", lot: "",      intend: "2026/FLHAI/FLHAI/136002", tender: "Crop seeds (semences vivrières)",                     category: "Inputs - Seeds",     method: "ITB", estInitial: 2947875, estPO: 0, nBids: 12,   opening: "2026-04-08", closing: "2026-04-24", currentKey: "tech_eval",         narrative: "offers sent for technical evaluation on 24.04.2026, technical questions shared with bidders, last deadline 13.5.2026",                   targetPO: "2026-07-30", state: "active",       comments: [] },
  { id: "4201267",    pr: "4201267", lot: "",      intend: "2026/FLHAI/FLHAI/136036", tender: "Visibility material",                                 category: "Visibility",         method: "ITB", estInitial: 75675,   estPO: 0,   nBids: null, opening: "2026-05-11", closing: "2026-05-21", currentKey: "published",          narrative: "published",                                                                                                                             targetPO: "2026-07-15", state: "active",       comments: [] },
  { id: "4201268",    pr: "4201268", lot: "",      intend: "2026/FLHAI/FLHAI/136381", tender: "Agricultural tools",                                  category: "Equipment - Agri",   method: "ITB", estInitial: 3068400, estPO: 0, nBids: null, opening: "2026-04-15", closing: "2026-05-11", currentKey: "closing",           narrative: "Published",                                                                                                                             targetPO: "2026-08-15", state: "active",       comments: [] },
  { id: "4201269-L1", pr: "4201269", lot: "Lot 1", intend: "2026/FLHAI/FLHAI/136382", tender: "Roosters & laying hens",                              category: "Livestock",          method: "ITB", estInitial: 3802500, estPO: 0, nBids: null, opening: "2026-04-20", closing: "2026-05-08", currentKey: "tech_eval",         narrative: "offers sent for tech. Evaluation 8.5.2026",                                                                                             targetPO: "2026-08-01", state: "active",       comments: [] },
  { id: "4201269-L2", pr: "4201269", lot: "Lot 2", intend: "2026/FLHAI/FLHAI/136383", tender: "Poultry (ducks, guinea fowl, geese, turkey)",         category: "Livestock",          method: "ITB", estInitial: 3752250, estPO: 0, nBids: null, opening: "2026-04-20", closing: "2026-05-08", currentKey: "tech_eval",         narrative: "offers sent for tech. Evaluation 8.5.2026",                                                                                             targetPO: "2026-08-01", state: "active",       comments: [] },
  { id: "4201269-L3", pr: "4201269", lot: "Lot 3", intend: "2026/FLHAI/FLHAI/136039", tender: "Animal feed (poultry)",                               category: "Inputs - Feed",      method: "ITB", estInitial: 4408200, estPO: 0, nBids: null, opening: "",           closing: "",           currentKey: "solicitation_prep",  narrative: "ITB to be signed, intend down",                                                                                                         targetPO: "2026-08-15", state: "cancelled",    comments: [] },
  // ── Pipeline — no PR assigned yet ─────────────────────────────────────
  { id: "pipeline-antibiotics", pr: "No PR", lot: "", intend: "", tender: "Antibiotics & antiparasitics", category: "Inputs - Feed", method: "ITB", estInitial: 0, estPO: 0, nBids: null, opening: "", closing: "", currentKey: "pr_received", narrative: "To be confirmed.",                          targetPO: "", state: "pre_pipeline", comments: [] },
  { id: "pipeline-natcash",     pr: "No PR", lot: "", intend: "", tender: "NatCash",                      category: "Services - M&E", method: "ITB", estInitial: 0, estPO: 0, nBids: null, opening: "", closing: "", currentKey: "pr_received", narrative: "programme following-up with head of PAM",   targetPO: "", state: "pre_pipeline", comments: [] },
  { id: "pipeline-ordinateurs", pr: "??",    lot: "", intend: "", tender: "Ordinateurs",                  category: "Equipment - ICT", method: "ITB", estInitial: 0, estPO: 0, nBids: null, opening: "", closing: "", currentKey: "pr_received", narrative: "not with us",                               targetPO: "", state: "not_applicable", comments: [] },
  { id: "pipeline-warehouse",   pr: "No PR", lot: "", intend: "", tender: "Warehouse",                    category: "Equipment - Agri", method: "ITB", estInitial: 0, estPO: 0, nBids: null, opening: "", closing: "", currentKey: "pr_received", narrative: "NFF received, PR request sent on 8.5.2026", targetPO: "", state: "pre_pipeline", comments: [] },
];

function initializeProcurements() {
  return SEED.map((p) => {
    const stages = buildStages(p.method, p.estInitial);
    const withProgress = markProgressTo(
      stages,
      p.currentKey,
      p.state === "on_hold" || p.state === "cancelled" ? "blocked" : "in_progress"
    );
    return { ...p, stages: withProgress };
  });
}

const TODAY = new Date();
TODAY.setHours(0, 0, 0, 0);

const parseDate = (s) => {
  if (!s) return null;
  const d = new Date(s + "T00:00:00");
  return isNaN(d) ? null : d;
};
const toISODate = (d) => {
  if (!d) return "";
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
};
const fmtDate = (d) => {
  if (!d) return "—";
  return `${String(d.getDate()).padStart(2, "0")}.${String(d.getMonth() + 1).padStart(2, "0")}.${d.getFullYear()}`;
};
const daysBetween = (a, b) => (!a || !b ? null : Math.round((b - a) / 86400000));
const addDays = (d, n) => {
  const r = new Date(d);
  r.setDate(r.getDate() + n);
  return r;
};
const fmtUSD = (n) => (n == null || isNaN(n) ? "—" : "$" + n.toLocaleString("en-US"));
const fmtUSDShort = (n) => {
  if (n == null || isNaN(n)) return "—";
  if (n >= 1000000) return `$${(n / 1000000).toFixed(2)}M`;
  if (n >= 1000) return `$${(n / 1000).toFixed(0)}K`;
  return `$${n}`;
};

function computeProcurement(p) {
  const remainingDays = p.stages
    .filter((s) => s.status !== "complete" && s.status !== "skipped")
    .reduce((sum, s) => sum + (Number(s.plannedDays) || 0), 0);
  const estPODate = addDays(TODAY, remainingDays);
  const target = parseDate(p.targetPO);
  const buffer = target ? daysBetween(estPODate, target) : null;
  const variance = p.estPO && p.estInitial ? (p.estPO - p.estInitial) / p.estInitial : 0;
  let risk;
  if (p.state === "cancelled") risk = { level: "critical", label: "Cancelled" };
  else if (p.state === "on_hold") risk = { level: "watch", label: "On hold" };
  else if (p.state === "pre_pipeline") risk = { level: "neutral", label: "No PR yet" };
  else if (p.state === "not_applicable") risk = { level: "neutral", label: "Not with us" };
  else if (buffer == null) risk = { level: "neutral", label: "No target" };
  else if (buffer < 0) risk = { level: "critical", label: `${Math.abs(buffer)}d overrun` };
  else if (buffer <= 5) risk = { level: "watch", label: `${buffer}d buffer` };
  else risk = { level: "ok", label: `${buffer}d buffer` };
  return { remainingDays, estPODate, buffer, variance, risk };
}

const StatusDot = ({ status, size = 16 }) => {
  const map = {
    complete: { color: "#15803D", Icon: CheckCircle2 },
    in_progress: { color: FAO_BLUE, Icon: Clock },
    not_started: { color: "#94A3B8", Icon: Circle },
    blocked: { color: "#B45309", Icon: PauseCircle },
    skipped: { color: "#64748B", Icon: Ban },
  };
  const { color, Icon } = map[status] || map.not_started;
  return <Icon size={size} style={{ color }} strokeWidth={2.25} />;
};

const RiskPill = ({ risk, compact = false }) => {
  const styles = {
    ok: { bg: "#DCFCE7", fg: "#14532D", dot: "#16A34A", label: "On track" },
    watch: { bg: "#FEF3C7", fg: "#78350F", dot: "#D97706", label: "Watch" },
    critical: { bg: "#FEE2E2", fg: "#7F1D1D", dot: "#DC2626", label: "Critical" },
    neutral: { bg: "#F1F5F9", fg: "#475569", dot: "#94A3B8", label: "—" },
  };
  const s = styles[risk.level];
  return (
    <span
      className={`inline-flex items-center gap-2 rounded-full font-semibold whitespace-nowrap ${compact ? "px-2 py-0.5 text-[10px]" : "px-2.5 py-1 text-xs"}`}
      style={{ backgroundColor: s.bg, color: s.fg, fontFamily: fontStack.body, letterSpacing: "0.01em" }}
    >
      <span className="inline-block w-1.5 h-1.5 rounded-full" style={{ backgroundColor: s.dot }} />
      <span>{s.label}</span>
      {!compact && <><span className="opacity-60">·</span><span style={{ fontFamily: fontStack.mono }}>{risk.label}</span></>}
    </span>
  );
};

const VarianceTag = ({ variance }) => {
  if (Math.abs(variance) < 0.001) {
    return <span className="inline-flex items-center gap-1 text-xs" style={{ color: "#64748B", fontFamily: fontStack.mono }}><Minus size={12} /> 0.0%</span>;
  }
  const pct = (variance * 100).toFixed(1);
  const positive = variance > 0;
  const colors = Math.abs(variance) > 0.15
    ? { fg: "#B91C1C", Icon: positive ? TrendingUp : TrendingDown }
    : Math.abs(variance) > 0.05
    ? { fg: "#B45309", Icon: positive ? TrendingUp : TrendingDown }
    : { fg: "#15803D", Icon: positive ? TrendingUp : TrendingDown };
  const { fg, Icon } = colors;
  return (
    <span className="inline-flex items-center gap-1 text-xs font-medium" style={{ color: fg, fontFamily: fontStack.mono }}>
      <Icon size={12} strokeWidth={2.5} />{positive ? "+" : ""}{pct}%
    </span>
  );
};

const MethodBadge = ({ method }) => (
  <span className="inline-block px-2 py-0.5 rounded text-[10px] font-bold tracking-wider"
    style={{ backgroundColor: method === "RFP" ? "#EDE9FE" : "#DBEAFE", color: method === "RFP" ? "#5B21B6" : "#1E3A8A", fontFamily: fontStack.mono }}>
    {method}
  </span>
);

const CategoryDot = ({ category }) => (
  <span className="inline-block w-2 h-2 rounded-sm" style={{ backgroundColor: CATEGORY_COLORS[category] || "#94A3B8" }} title={category} />
);

const Row = ({ label, value, emphasis = "neutral" }) => {
  const color = emphasis === "good" ? "#15803D" : emphasis === "warn" ? "#B45309" : emphasis === "bad" ? "#B91C1C" : "#0F172A";
  return (
    <div className="flex items-center justify-between">
      <span style={{ color: "#64748B", fontFamily: fontStack.body }}>{label}</span>
      <span className="font-semibold" style={{ color, fontFamily: fontStack.mono }}>{value}</span>
    </div>
  );
};

const KPI = ({ label, value, sub, accent }) => (
  <div className="rounded-xl border p-5" style={{ borderColor: "#E2E8F0", backgroundColor: "white", borderTop: `3px solid ${accent}` }}>
    <div className="text-[10px] uppercase tracking-widest mb-2" style={{ color: "#64748B", fontFamily: fontStack.body, letterSpacing: "0.12em" }}>{label}</div>
    <div className="text-3xl font-bold" style={{ color: FAO_NAVY, fontFamily: fontStack.display }}>{value}</div>
    {sub && <div className="text-xs mt-1" style={{ color: "#64748B", fontFamily: fontStack.body }}>{sub}</div>}
  </div>
);

function CommentsPanel({ procurement, onUpdate }) {
  const [draft, setDraft] = useState("");
  const addComment = () => {
    if (!draft.trim()) return;
    onUpdate({ ...procurement, comments: [...(procurement.comments || []), { id: Date.now(), date: new Date().toISOString(), text: draft.trim() }] });
    setDraft("");
  };
  const deleteComment = (id) => {
    onUpdate({ ...procurement, comments: procurement.comments.filter((c) => c.id !== id) });
  };
  return (
    <div>
      <div className="flex items-center gap-2 mb-3">
        <MessageSquare size={14} style={{ color: FAO_NAVY }} />
        <h4 className="text-xs font-bold uppercase tracking-widest" style={{ color: FAO_NAVY, fontFamily: fontStack.body }}>
          Comments log ({procurement.comments?.length || 0})
        </h4>
      </div>
      <div className="space-y-2 mb-3 max-h-48 overflow-y-auto">
        {(procurement.comments || []).length === 0 && (
          <div className="text-xs italic px-3 py-2 rounded" style={{ color: "#94A3B8", backgroundColor: "#F8FAFC", fontFamily: fontStack.body }}>
            No comments yet. Use the field below to log decisions, blockers, or follow-ups.
          </div>
        )}
        {(procurement.comments || []).map((c) => (
          <div key={c.id} className="rounded border px-3 py-2 group" style={{ borderColor: "#E2E8F0", backgroundColor: "white" }}>
            <div className="flex items-center justify-between mb-1">
              <span className="text-[10px] uppercase tracking-wider font-semibold" style={{ color: FAO_BLUE, fontFamily: fontStack.mono }}>
                {fmtDate(new Date(c.date))} · {new Date(c.date).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}
              </span>
              <button onClick={() => deleteComment(c.id)} className="text-[10px] opacity-0 group-hover:opacity-100 transition" style={{ color: "#B91C1C", fontFamily: fontStack.body }}>
                Delete
              </button>
            </div>
            <div className="text-sm" style={{ color: "#0F172A", fontFamily: fontStack.body }}>{c.text}</div>
          </div>
        ))}
      </div>
      <div className="flex items-start gap-2">
        <textarea
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          onKeyDown={(e) => { if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) addComment(); }}
          placeholder="Add a comment… (⌘/Ctrl + Enter to submit)"
          rows={2}
          className="flex-1 px-3 py-2 rounded border text-sm"
          style={{ borderColor: "#CBD5E1", fontFamily: fontStack.body, color: "#0F172A", resize: "vertical" }}
        />
        <button onClick={addComment} disabled={!draft.trim()}
          className="inline-flex items-center gap-1 px-3 py-2 rounded text-sm font-medium text-white disabled:opacity-40"
          style={{ backgroundColor: FAO_BLUE, fontFamily: fontStack.body }}>
          <Send size={14} /> Post
        </button>
      </div>
    </div>
  );
}

function StageEditor({ procurement, onUpdate }) {
  const updateStage = (idx, patch) => onUpdate({ ...procurement, stages: procurement.stages.map((s, i) => (i === idx ? { ...s, ...patch } : s)) });
  const setStageStatus = (idx, status) => {
    let stages = procurement.stages.slice();
    if (status === "complete") {
      stages = stages.map((s, i) => (i <= idx ? { ...s, status: "complete" } : s));
    } else if (status === "in_progress") {
      stages = stages.map((s, i) => i < idx ? { ...s, status: "complete" } : i === idx ? { ...s, status: "in_progress" } : { ...s, status: "not_started" });
    } else {
      stages[idx] = { ...stages[idx], status };
    }
    onUpdate({ ...procurement, stages });
  };
  const resetDefaults = () => {
    const fresh = buildStages(procurement.method, procurement.estInitial);
    const merged = fresh.map((f) => { const ex = procurement.stages.find((s) => s.key === f.key); return { ...f, status: ex?.status || "not_started" }; });
    onUpdate({ ...procurement, stages: merged });
  };
  const [addingAfter, setAddingAfter] = useState(null);
  const [newStepName, setNewStepName] = useState("");
  const [dragSrc, setDragSrc] = useState(null);
  const [dragOver, setDragOver] = useState(null);

  const onDragStart = (i) => (e) => {
    setDragSrc(i);
    e.dataTransfer.effectAllowed = "move";
  };
  const onDragOver = (i) => (e) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
    if (i !== dragSrc) setDragOver(i);
  };
  const onDrop = (i) => (e) => {
    e.preventDefault();
    if (dragSrc === null || dragSrc === i) { setDragSrc(null); setDragOver(null); return; }
    const stages = [...procurement.stages];
    const [moved] = stages.splice(dragSrc, 1);
    stages.splice(i, 0, moved);
    onUpdate({ ...procurement, stages });
    setDragSrc(null);
    setDragOver(null);
  };
  const onDragEnd = () => { setDragSrc(null); setDragOver(null); };

  const insertCustomStep = (afterIdx) => {
    if (!newStepName.trim()) return;
    const newStage = { key: `custom_${Date.now()}`, name: newStepName.trim(), plannedDays: 1, status: "not_started", custom: true };
    const stages = [...procurement.stages];
    stages.splice(afterIdx + 1, 0, newStage);
    onUpdate({ ...procurement, stages });
    setAddingAfter(null);
    setNewStepName("");
  };

  const deleteCustomStep = (i) => {
    onUpdate({ ...procurement, stages: procurement.stages.filter((_, idx) => idx !== i) });
  };

  const computed = computeProcurement(procurement);
  const { stages: timedStages } = computeTimeline(procurement);

  const updateStageStart = (i, isoDate) => {
    const stages = procurement.stages.map((s, idx) => {
      if (idx === i) return { ...s, stageStartOverride: isoDate || undefined };
      if (idx > i) return { ...s, stageStartOverride: undefined }; // clear downstream pins so they cascade
      return s;
    });
    onUpdate({ ...procurement, stages });
  };

  const updateStageEnd = (i, isoDate) => {
    const start = timedStages[i]?.stageStart;
    const newEnd = parseDate(isoDate);
    if (!start || !newEnd) return;
    const newDays = Math.max(0, Math.round((newEnd - start) / 86400000));
    const stages = procurement.stages.map((s, idx) => {
      if (idx === i) return { ...s, plannedDays: newDays };
      if (idx > i) return { ...s, stageStartOverride: undefined };
      return s;
    });
    onUpdate({ ...procurement, stages });
  };

  return (
    <div className="grid grid-cols-12 gap-6">
      <div className="col-span-12 lg:col-span-7">
        <div className="flex items-center justify-between mb-3">
          <h4 className="text-xs font-bold uppercase tracking-widest" style={{ color: FAO_NAVY, fontFamily: fontStack.body }}>
            Standardized workflow ({procurement.method})
          </h4>
          <button onClick={resetDefaults} className="inline-flex items-center gap-1 text-xs px-2 py-1 rounded hover:bg-slate-200 transition" style={{ color: "#475569", fontFamily: fontStack.body }}>
            <RotateCcw size={12} /> Reset days to default
          </button>
        </div>
        <div className="rounded-lg overflow-hidden border" style={{ borderColor: "#E2E8F0" }}>
          <table className="w-full text-sm" style={{ fontFamily: fontStack.body }}>
            <thead>
              <tr className="text-[10px] uppercase tracking-wider" style={{ backgroundColor: "#F1F5F9", color: "#64748B" }}>
                <th className="w-6" />
                <th className="text-left px-3 py-2 font-semibold">#</th>
                <th className="text-left px-3 py-2 font-semibold">Stage</th>
                <th className="text-center px-3 py-2 font-semibold">Status</th>
                <th className="text-right px-3 py-2 font-semibold">Start</th>
                <th className="text-right px-3 py-2 font-semibold">End</th>
                <th className="text-right px-3 py-2 font-semibold">Days</th>
                <th className="w-14" />
              </tr>
            </thead>
            <tbody>
              {procurement.stages.map((s, i) => {
                const ts = timedStages[i];
                const isAdding = addingAfter === i;
                const isDragging = dragSrc === i;
                const isDropTarget = dragOver === i;
                return (
                  <React.Fragment key={s.key}>
                    <tr
                      draggable
                      onDragStart={onDragStart(i)}
                      onDragOver={onDragOver(i)}
                      onDrop={onDrop(i)}
                      onDragEnd={onDragEnd}
                      className="border-t group"
                      style={{
                        borderColor: isDropTarget ? FAO_BLUE : "#F1F5F9",
                        backgroundColor: isDragging ? "#F0F9FF" : isDropTarget ? "#EFF6FF" : s.status === "in_progress" ? "#EFF6FF" : "white",
                        opacity: isDragging ? 0.4 : 1,
                        outline: isDropTarget ? `2px solid ${FAO_BLUE}` : undefined,
                      }}>
                      <td className="pl-2 py-2 cursor-grab active:cursor-grabbing" style={{ color: "#CBD5E1" }}>
                        <GripVertical size={14} />
                      </td>
                      <td className="px-3 py-2 text-xs" style={{ color: "#94A3B8", fontFamily: fontStack.mono }}>{String(i + 1).padStart(2, "0")}</td>
                      <td className="px-3 py-2">
                        <div className="flex items-center gap-2">
                          <StatusDot status={s.status} />
                          <input value={s.name} onChange={(e) => updateStage(i, { name: e.target.value })}
                            className="flex-1 text-sm px-1.5 py-0.5 rounded border border-transparent hover:border-slate-200 focus:border-slate-300 bg-transparent"
                            style={{ fontFamily: fontStack.body, color: "#0F172A", minWidth: 0, outline: "none" }} />
                          {s.custom && <span className="text-[9px] px-1.5 py-0.5 rounded font-bold uppercase tracking-wider flex-shrink-0" style={{ backgroundColor: "#EDE9FE", color: "#5B21B6", fontFamily: fontStack.mono }}>custom</span>}
                        </div>
                      </td>
                      <td className="px-3 py-2 text-center">
                        <select value={s.status} onChange={(e) => setStageStatus(i, e.target.value)}
                          className="text-xs px-2 py-1 rounded border bg-white"
                          style={{ borderColor: "#CBD5E1", color: "#334155", fontFamily: fontStack.body }}>
                          <option value="not_started">Not started</option>
                          <option value="in_progress">In progress</option>
                          <option value="complete">Complete</option>
                          <option value="blocked">Blocked</option>
                          <option value="skipped">Skipped</option>
                        </select>
                      </td>
                      <td className="px-2 py-1.5">
                        <input type="date" value={ts?.stageStart ? toISODate(ts.stageStart) : ""}
                          onChange={(e) => updateStageStart(i, e.target.value)}
                          className="px-1.5 py-1 rounded border text-[11px] w-32"
                          style={{ borderColor: s.stageStartOverride ? FAO_BLUE : "#CBD5E1", fontFamily: fontStack.mono, color: "#0F172A", outline: "none" }} />
                      </td>
                      <td className="px-2 py-1.5">
                        <input type="date" value={ts?.stageEnd ? toISODate(ts.stageEnd) : ""}
                          onChange={(e) => updateStageEnd(i, e.target.value)}
                          className="px-1.5 py-1 rounded border text-[11px] w-32"
                          style={{ borderColor: "#CBD5E1", fontFamily: fontStack.mono, color: "#0F172A", outline: "none" }} />
                      </td>
                      <td className="px-3 py-2 text-right">
                        <input type="number" min="0" value={s.plannedDays}
                          onChange={(e) => updateStage(i, { plannedDays: Number(e.target.value) })}
                          disabled={s.status === "complete" || s.status === "skipped"}
                          className="w-14 text-right px-2 py-1 rounded border text-xs disabled:bg-slate-100 disabled:text-slate-400"
                          style={{ borderColor: "#CBD5E1", fontFamily: fontStack.mono, color: "#0F172A" }} />
                        <span className="ml-1 text-xs" style={{ color: "#94A3B8", fontFamily: fontStack.mono }}>d</span>
                      </td>
                      <td className="px-1 py-2">
                        <div className="flex items-center gap-0.5">
                          <button onClick={() => { setAddingAfter(isAdding ? null : i); setNewStepName(""); }}
                            title="Insert step after this one"
                            className="opacity-0 group-hover:opacity-100 transition p-1 rounded hover:bg-slate-100"
                            style={{ color: FAO_BLUE }}>
                            <Plus size={13} />
                          </button>
                          <button onClick={() => deleteCustomStep(i)} title="Remove step"
                            className="opacity-0 group-hover:opacity-100 transition p-1 rounded hover:bg-red-50"
                            style={{ color: "#DC2626" }}>
                            <X size={13} />
                          </button>
                        </div>
                      </td>
                    </tr>
                    {isAdding && (
                      <tr style={{ backgroundColor: "#F0F9FF" }}>
                        <td />
                        <td colSpan={5} className="px-3 py-2">
                          <div className="flex items-center gap-2">
                            <input autoFocus value={newStepName} onChange={(e) => setNewStepName(e.target.value)}
                              onKeyDown={(e) => { if (e.key === "Enter") insertCustomStep(i); if (e.key === "Escape") { setAddingAfter(null); setNewStepName(""); } }}
                              placeholder="Step name…"
                              className="flex-1 text-sm px-2 py-1 rounded border"
                              style={{ borderColor: FAO_BLUE, fontFamily: fontStack.body, color: "#0F172A", outline: "none" }} />
                            <button onClick={() => insertCustomStep(i)} disabled={!newStepName.trim()}
                              className="text-xs px-3 py-1 rounded font-medium text-white disabled:opacity-40"
                              style={{ backgroundColor: FAO_BLUE, fontFamily: fontStack.body }}>Add</button>
                            <button onClick={() => { setAddingAfter(null); setNewStepName(""); }}
                              className="text-xs px-2 py-1 rounded hover:bg-slate-200"
                              style={{ color: "#475569", fontFamily: fontStack.body }}>Cancel</button>
                          </div>
                        </td>
                        <td />
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      <div className="col-span-12 lg:col-span-5">
        <h4 className="text-xs font-bold uppercase tracking-widest mb-3" style={{ color: FAO_NAVY, fontFamily: fontStack.body }}>Plan vs. target</h4>
        <div className="rounded-lg border p-5 space-y-4" style={{ borderColor: "#E2E8F0", backgroundColor: "white" }}>
          <div>
            <div className="text-[10px] uppercase tracking-wider mb-2" style={{ color: "#64748B", fontFamily: fontStack.body }}>Values</div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs block mb-1" style={{ color: "#64748B", fontFamily: fontStack.body }}>Est. initial (PR)</label>
                <input type="number" value={procurement.estInitial}
                  onChange={(e) => onUpdate({ ...procurement, estInitial: Number(e.target.value) })}
                  className="w-full px-2 py-1.5 rounded border text-sm"
                  style={{ borderColor: "#CBD5E1", fontFamily: fontStack.mono, color: "#0F172A" }} />
              </div>
              <div>
                <label className="text-xs block mb-1" style={{ color: "#64748B", fontFamily: fontStack.body }}>Est. PO value</label>
                <input type="number" value={procurement.estPO}
                  onChange={(e) => onUpdate({ ...procurement, estPO: Number(e.target.value) })}
                  className="w-full px-2 py-1.5 rounded border text-sm"
                  style={{ borderColor: "#CBD5E1", fontFamily: fontStack.mono, color: "#0F172A" }} />
              </div>
            </div>
            <div className="mt-2 flex items-center justify-between text-xs">
              <span style={{ color: "#64748B", fontFamily: fontStack.body }}>Variance vs. PR</span>
              <VarianceTag variance={computed.variance} />
            </div>
          </div>
          <div className="border-t" style={{ borderColor: "#F1F5F9" }} />
          <div>
            <div className="text-[10px] uppercase tracking-wider mb-2" style={{ color: "#64748B", fontFamily: fontStack.body }}>Schedule</div>
            <label className="text-xs block mb-1" style={{ color: "#64748B", fontFamily: fontStack.body }}>Target PO date (manual)</label>
            <input type="date" value={procurement.targetPO || ""}
              onChange={(e) => onUpdate({ ...procurement, targetPO: e.target.value })}
              className="w-full px-2 py-1.5 rounded border text-sm"
              style={{ borderColor: "#CBD5E1", fontFamily: fontStack.mono, color: "#0F172A" }} />
            <div className="mt-3 space-y-2 text-sm">
              <Row label="Days remaining (planned)" value={`${computed.remainingDays}d`} />
              <Row label="Estimated PO date" value={fmtDate(computed.estPODate)} />
              <Row label="Buffer / overrun" value={computed.buffer == null ? "—" : `${computed.buffer}d`}
                emphasis={computed.buffer == null ? "neutral" : computed.buffer < 0 ? "bad" : computed.buffer <= 5 ? "warn" : "good"} />
            </div>
          </div>
          <div className="border-t" style={{ borderColor: "#F1F5F9" }} />
          <div>
            <label className="text-xs block mb-1" style={{ color: "#64748B", fontFamily: fontStack.body }}>Lifecycle state</label>
            <select value={procurement.state} onChange={(e) => onUpdate({ ...procurement, state: e.target.value })}
              className="w-full px-2 py-1.5 rounded border text-sm"
              style={{ borderColor: "#CBD5E1", color: "#0F172A", fontFamily: fontStack.body }}>
              <option value="active">Active</option>
              <option value="on_hold">On hold</option>
              <option value="pre_pipeline">Pre-pipeline (no PR)</option>
              <option value="not_applicable">Not applicable (not with us)</option>
              <option value="cancelled">Cancelled / re-issue</option>
              <option value="awarded">Awarded / closed</option>
            </select>
            <div className="mt-3"><RiskPill risk={computed.risk} /></div>
          </div>
        </div>
      </div>
    </div>
  );
}

const STAGE_STATUS_COLORS = {
  complete:    { fill: "#15803D", opacity: 0.75 },
  in_progress: { fill: "#009FDA", opacity: 0.90 },
  not_started: { fill: "#CBD5E1", opacity: 1.00 },
  blocked:     { fill: "#D97706", opacity: 0.85 },
  skipped:     { fill: "#94A3B8", opacity: 0.45 },
};

// Computes absolute start/end dates for every stage, anchored to real bid dates.
// Anchor priority: closing date → opening date → ganttAnchor field → TODAY at in-progress stage.
function computeTimeline(p) {
  const stages = p.stages;
  let anchorIdx = -1;
  let anchorDate = null;

  const closingDate = parseDate(p.closing);
  const openingDate = parseDate(p.opening);
  if (closingDate) {
    const i = stages.findIndex((s) => s.key === "closing");
    if (i !== -1) { anchorIdx = i; anchorDate = closingDate; }
  }
  if (anchorIdx === -1 && openingDate) {
    const i = stages.findIndex((s) => s.key === "published");
    if (i !== -1) { anchorIdx = i; anchorDate = openingDate; }
  }
  if (anchorIdx === -1) {
    anchorIdx = stages.findIndex((s) => s.status === "in_progress");
    if (anchorIdx === -1) anchorIdx = stages.findIndex((s) => s.status !== "complete");
    if (anchorIdx === -1) anchorIdx = 0;
    if (p.ganttAnchor) {
      const base = parseDate(p.ganttAnchor);
      const offset = stages.slice(0, anchorIdx).reduce((s, st) => s + (Number(st.plannedDays) || 0), 0);
      anchorDate = addDays(base, offset);
    } else {
      anchorDate = TODAY;
    }
  }

  // Build base timeline from anchor outward
  const out = stages.map((s) => ({ ...s, stageStart: null, stageEnd: null }));
  out[anchorIdx].stageStart = anchorDate;
  out[anchorIdx].stageEnd = addDays(anchorDate, Number(out[anchorIdx].plannedDays) || 0);
  for (let i = anchorIdx + 1; i < out.length; i++) {
    out[i].stageStart = out[i - 1].stageEnd;
    out[i].stageEnd = addDays(out[i].stageStart, Number(out[i].plannedDays) || 0);
  }
  for (let i = anchorIdx - 1; i >= 0; i--) {
    out[i].stageEnd = out[i + 1].stageStart;
    out[i].stageStart = addDays(out[i].stageEnd, -(Number(out[i].plannedDays) || 0));
  }

  // Apply per-stage start overrides: each override pins that stage and cascades
  // forward until the next pinned stage.
  for (let i = 0; i < out.length; i++) {
    const override = parseDate(stages[i].stageStartOverride);
    if (!override) continue;
    out[i].stageStart = override;
    out[i].stageEnd = addDays(override, Number(stages[i].plannedDays) || 0);
    for (let j = i + 1; j < out.length; j++) {
      if (stages[j].stageStartOverride) break; // stop at next explicit pin
      out[j].stageStart = out[j - 1].stageEnd;
      out[j].stageEnd = addDays(out[j].stageStart, Number(stages[j].plannedDays) || 0);
    }
  }

  return { stages: out, timelineStart: out.length ? out[0].stageStart : null, timelineEnd: out.length ? out[out.length - 1].stageEnd : null };
}

function GanttChart({ procurements, onUpdate }) {
  const timelines = useMemo(() => procurements.map((p) => computeTimeline(p)), [procurements]);

  const { minDate, maxDate } = useMemo(() => {
    const dates = timelines.flatMap((t) => [t.timelineStart, t.timelineEnd]).filter(Boolean);
    dates.push(TODAY);
    // Always include the project end + 14 days so the deadline and distribution window are visible
    dates.push(PROJECT_END);
    const min = new Date(Math.min(...dates.map((d) => d.getTime())));
    const max = new Date(Math.max(...dates.map((d) => d.getTime())));
    min.setDate(min.getDate() - 14);
    max.setDate(max.getDate() + 14);
    return { minDate: min, maxDate: max };
  }, [timelines]);

  const PX_PER_DAY = 5;
  const ROW_HEIGHT = 48;
  const LABEL_WIDTH = 280;
  const totalDays = Math.ceil((maxDate - minDate) / 86400000);
  const chartWidth = totalDays * PX_PER_DAY;

  const dateToX = useCallback((d) => (!d ? 0 : ((d - minDate) / 86400000) * PX_PER_DAY), [minDate]);

  const months = useMemo(() => {
    const out = [];
    const cursor = new Date(minDate);
    cursor.setDate(1);
    while (cursor <= maxDate) {
      out.push({ label: cursor.toLocaleDateString("en-US", { month: "short", year: "2-digit" }), x: dateToX(cursor) });
      cursor.setMonth(cursor.getMonth() + 1);
    }
    return out;
  }, [minDate, maxDate, dateToX]);

  const [drag, setDrag] = useState(null);

  const startDrag = (p) => (e) => {
    e.preventDefault();
    e.stopPropagation();
    const closingDate = parseDate(p.closing);
    const openingDate = parseDate(p.opening);
    let anchorType, originalAnchorDate;
    if (closingDate && p.stages.some((s) => s.key === "closing")) {
      anchorType = "closing"; originalAnchorDate = closingDate;
    } else if (openingDate && p.stages.some((s) => s.key === "published")) {
      anchorType = "opening"; originalAnchorDate = openingDate;
    } else {
      anchorType = "ganttAnchor";
      originalAnchorDate = parseDate(p.ganttAnchor) || TODAY;
    }
    setDrag({ id: p.id, startX: e.clientX, anchorType, originalAnchorDate });
  };

  useEffect(() => {
    if (!drag) return;
    const onMove = (e) => {
      const dx = e.clientX - drag.startX;
      const dayDelta = Math.round(dx / PX_PER_DAY);
      const proc = procurements.find((p) => p.id === drag.id);
      if (!proc) return;
      const newAnchorDate = addDays(drag.originalAnchorDate, dayDelta);
      const updates = { [drag.anchorType]: toISODate(newAnchorDate) };
      // Recompute timeline end to sync targetPO
      const { timelineEnd } = computeTimeline({ ...proc, ...updates });
      updates.targetPO = toISODate(timelineEnd);
      onUpdate({ ...proc, ...updates });
    };
    const onUp = () => setDrag(null);
    window.addEventListener("mousemove", onMove);
    window.addEventListener("mouseup", onUp);
    return () => { window.removeEventListener("mousemove", onMove); window.removeEventListener("mouseup", onUp); };
  }, [drag, procurements, onUpdate]);

  const todayX = dateToX(TODAY);
  const distStartX = dateToX(DISTRIBUTION_START);
  const projectEndX = dateToX(PROJECT_END);

  return (
    <div className="rounded-xl border bg-white overflow-hidden" style={{ borderColor: "#E2E8F0" }}>
      <div className="p-4 border-b flex items-center justify-between flex-wrap gap-3" style={{ borderColor: "#E2E8F0" }}>
        <div className="flex items-center gap-3">
          <GanttChartSquare size={18} style={{ color: FAO_NAVY }} />
          <div>
            <h3 className="text-sm font-bold" style={{ color: FAO_NAVY, fontFamily: fontStack.display }}>
              Timeline — each segment is a workflow stage · drag a bar to shift all dates
            </h3>
            <div className="text-[10px] mt-0.5" style={{ color: "#64748B", fontFamily: fontStack.body }}>
              Project ends <span className="font-semibold" style={{ color: "#DC2626" }}>19 Oct 2026</span> · distribution window starts <span className="font-semibold" style={{ color: "#D97706" }}>{fmtDate(DISTRIBUTION_START)}</span> · goods must be delivered before that date
            </div>
          </div>
        </div>
        <div className="flex items-center gap-3 text-[10px] flex-wrap" style={{ fontFamily: fontStack.body }}>
          <LegendDot color="#15803D" label="Complete" />
          <LegendDot color="#009FDA" label="In progress" />
          <LegendDot color="#CBD5E1" label="Not started" />
          <LegendDot color="#D97706" label="Blocked" />
          <span style={{ color: "#64748B" }}>|</span>
          <span style={{ color: "#DC2626" }}>▮ today</span>
          <span style={{ color: "#F97316" }}>▨ distribution</span>
          <span style={{ color: "#DC2626" }}>| project end</span>
        </div>
      </div>

      <div className="overflow-x-auto">
        <div style={{ width: LABEL_WIDTH + chartWidth, minWidth: "100%" }}>
          {/* Month header */}
          <div className="flex sticky top-0 z-10" style={{ backgroundColor: "#F8FAFC", borderBottom: "1px solid #E2E8F0", height: 32 }}>
            <div style={{ width: LABEL_WIDTH, borderRight: "1px solid #E2E8F0" }} />
            <div className="relative" style={{ width: chartWidth, height: 32 }}>
              {months.map((m, i) => (
                <div key={i} className="absolute top-0 h-full flex items-center"
                  style={{ left: m.x, fontFamily: fontStack.mono, color: "#64748B", fontSize: 11, borderLeft: "1px solid #E2E8F0", paddingLeft: 6, width: i < months.length - 1 ? months[i + 1].x - m.x : 80 }}>
                  {m.label}
                </div>
              ))}
              {/* Distribution window label in header */}
              {distStartX < chartWidth && (
                <div className="absolute top-0 h-full pointer-events-none" style={{ left: distStartX, width: Math.max(0, projectEndX - distStartX), backgroundColor: "#FEF3C7", opacity: 0.6 }} />
              )}
              {projectEndX > 0 && projectEndX < chartWidth && (
                <div className="absolute top-0 h-full flex items-center pointer-events-none" style={{ left: projectEndX - 1, width: 2, backgroundColor: "#DC2626" }}>
                  <span className="absolute text-[9px] font-bold whitespace-nowrap px-1" style={{ left: 4, color: "#DC2626", fontFamily: fontStack.mono, backgroundColor: "#F8FAFC" }}>END</span>
                </div>
              )}
            </div>
          </div>

          {procurements.map((p, idx) => {
            const { stages: timedStages, timelineEnd } = timelines[idx];
            const computed = computeProcurement(p);
            const endX = dateToX(timelineEnd);

            return (
              <div key={p.id} className="flex border-b"
                style={{ borderColor: "#F1F5F9", height: ROW_HEIGHT, backgroundColor: idx % 2 === 0 ? "white" : "#FAFBFC" }}>

                {/* Label column */}
                <div className="px-4 flex items-center gap-2 flex-shrink-0" style={{ width: LABEL_WIDTH, borderRight: "1px solid #E2E8F0" }}>
                  <CategoryDot category={p.category} />
                  <div className="min-w-0 flex-1">
                    <div className="text-xs font-semibold truncate" style={{ color: "#0F172A", fontFamily: fontStack.body }}>{p.tender}</div>
                    <div className="text-[10px]" style={{ color: "#64748B", fontFamily: fontStack.mono }}>
                      {p.pr}{p.lot && ` · ${p.lot}`}
                    </div>
                  </div>
                  <RiskPill risk={computed.risk} compact />
                </div>

                {/* Chart column */}
                <div className="relative cursor-grab active:cursor-grabbing" style={{ width: chartWidth, height: ROW_HEIGHT }}
                  onMouseDown={startDrag(p)}>

                  {/* Month grid lines */}
                  {months.map((m, i) => (
                    <div key={i} className="absolute top-0 h-full pointer-events-none" style={{ left: m.x, width: 1, backgroundColor: "#F1F5F9" }} />
                  ))}

                  {/* Distribution window */}
                  {distStartX < chartWidth && (
                    <div className="absolute top-0 h-full pointer-events-none"
                      style={{ left: distStartX, width: Math.max(0, projectEndX - distStartX), backgroundColor: "#FEF3C7", opacity: 0.55,
                        backgroundImage: "repeating-linear-gradient(45deg,transparent,transparent 5px,rgba(251,191,36,0.15) 5px,rgba(251,191,36,0.15) 10px)" }} />
                  )}
                  {/* Project end deadline */}
                  {projectEndX > 0 && projectEndX < chartWidth && (
                    <div className="absolute top-0 h-full pointer-events-none" style={{ left: projectEndX, width: 2, backgroundColor: "#DC2626", opacity: 0.85 }} />
                  )}
                  {/* Today line */}
                  <div className="absolute top-0 h-full pointer-events-none" style={{ left: todayX, width: 1.5, backgroundColor: "#DC2626", opacity: 0.5 }} />

                  {/* Stage segments */}
                  {timedStages.map((s, si) => {
                    const x = dateToX(s.stageStart);
                    const w = Math.max(dateToX(s.stageEnd) - x, s.plannedDays > 0 ? 1 : 0);
                    if (w === 0) return null;
                    const { fill, opacity } = STAGE_STATUS_COLORS[s.status] || STAGE_STATUS_COLORS.not_started;
                    const isActive = s.status === "in_progress";
                    return (
                      <div key={s.key} className="absolute pointer-events-none"
                        style={{
                          left: x, top: 10, width: w, height: 28,
                          backgroundColor: fill, opacity,
                          borderLeft: si === 0 ? `2px solid ${fill}` : "1px solid white",
                          borderRight: "1px solid white",
                          borderTop: isActive ? `2px solid ${fill}` : undefined,
                          borderBottom: isActive ? `2px solid ${fill}` : undefined,
                          borderRadius: si === 0 ? "3px 0 0 3px" : si === timedStages.length - 1 ? "0 3px 3px 0" : 0,
                          boxShadow: isActive ? `0 0 0 1px ${fill}` : undefined,
                        }}
                        title={`${s.name}\n${fmtDate(s.stageStart)} → ${fmtDate(s.stageEnd)} (${s.plannedDays}d)`}
                      />
                    );
                  })}

                  {/* End date label on drag */}
                  {drag && drag.id === p.id && (
                    <div className="absolute z-20 px-2 py-0.5 rounded text-[10px] text-white font-semibold pointer-events-none"
                      style={{ left: endX + 4, top: 14, backgroundColor: FAO_NAVY, fontFamily: fontStack.mono, whiteSpace: "nowrap" }}>
                      PO {fmtDate(timelineEnd)}
                    </div>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

const LegendDot = ({ color, label }) => (
  <span className="inline-flex items-center gap-1" style={{ color: "#475569" }}>
    <span className="inline-block w-2 h-2 rounded-sm" style={{ backgroundColor: color }} />{label}
  </span>
);

function getNextStep(p) {
  const stages = p.stages;
  // In-progress stage is the current active step
  const active = stages.find((s) => s.status === "in_progress");
  if (active) return { stage: active, isCurrent: true };
  // Otherwise first not_started after last complete
  const lastComplete = [...stages].reverse().find((s) => s.status === "complete");
  const lastCompleteIdx = lastComplete ? stages.lastIndexOf(lastComplete) : -1;
  const next = stages.slice(lastCompleteIdx + 1).find((s) => s.status === "not_started");
  return next ? { stage: next, isCurrent: false } : null;
}

function UpcomingActivities({ enriched }) {
  const HORIZON_DAYS = 60;
  const [horizon, setHorizon] = useState(HORIZON_DAYS);

  const activities = useMemo(() => {
    const items = [];
    enriched.forEach((p) => {
      if (p.state === "cancelled" || p.state === "not_applicable") return;
      const next = getNextStep(p);
      if (!next) return;
      const { stages: timed } = computeTimeline(p);
      const timedStage = timed.find((s) => s.key === next.stage.key);
      items.push({
        p,
        stage: next.stage,
        isCurrent: next.isCurrent,
        stageStart: timedStage?.stageStart || null,
        stageEnd: timedStage?.stageEnd || null,
        daysUntilStart: timedStage?.stageStart ? daysBetween(TODAY, timedStage.stageStart) : null,
        daysUntilEnd: timedStage?.stageEnd ? daysBetween(TODAY, timedStage.stageEnd) : null,
      });
    });
    return items
      .filter((a) => a.daysUntilEnd === null || a.daysUntilEnd <= horizon)
      .sort((a, b) => {
        const ad = a.stageEnd?.getTime() ?? Infinity;
        const bd = b.stageEnd?.getTime() ?? Infinity;
        return ad - bd;
      });
  }, [enriched, horizon]);

  // Group by week label relative to today
  const grouped = useMemo(() => {
    const groups = [];
    let currentLabel = null;
    activities.forEach((a) => {
      const d = a.stageEnd;
      let label;
      if (!d) {
        label = "No date set";
      } else {
        const diff = daysBetween(TODAY, d);
        if (diff < 0) label = "Overdue";
        else if (diff === 0) label = "Due today";
        else if (diff <= 7) label = "This week";
        else if (diff <= 14) label = "Next week";
        else {
          const monday = new Date(d);
          monday.setDate(d.getDate() - ((d.getDay() + 6) % 7));
          label = `Week of ${fmtDate(monday)}`;
        }
      }
      if (label !== currentLabel) { groups.push({ label, items: [] }); currentLabel = label; }
      groups[groups.length - 1].items.push(a);
    });
    return groups;
  }, [activities]);

  const urgencyStyle = (daysUntilEnd) => {
    if (daysUntilEnd === null) return { color: "#64748B", bg: "#F8FAFC", dot: "#94A3B8" };
    if (daysUntilEnd < 0)  return { color: "#B91C1C", bg: "#FEF2F2", dot: "#DC2626" };
    if (daysUntilEnd <= 3) return { color: "#B45309", bg: "#FFFBEB", dot: "#D97706" };
    if (daysUntilEnd <= 7) return { color: "#0369A1", bg: "#EFF6FF", dot: "#0284C7" };
    return { color: "#334155", bg: "white", dot: "#94A3B8" };
  };

  return (
    <div className="space-y-2">
      {/* Header bar */}
      <div className="rounded-xl border bg-white p-4 flex items-center justify-between flex-wrap gap-3" style={{ borderColor: "#E2E8F0" }}>
        <div className="flex items-center gap-3">
          <CalendarDays size={18} style={{ color: FAO_NAVY }} />
          <div>
            <h3 className="text-sm font-bold" style={{ color: FAO_NAVY, fontFamily: fontStack.display }}>Upcoming procurement activities</h3>
            <p className="text-xs mt-0.5" style={{ color: "#64748B", fontFamily: fontStack.body }}>Next pending step per active procurement · sorted by deadline</p>
          </div>
        </div>
        <div className="flex items-center gap-2 text-xs" style={{ fontFamily: fontStack.body, color: "#475569" }}>
          <span>Show next</span>
          {[30, 60, 90].map((d) => (
            <button key={d} onClick={() => setHorizon(d)}
              className="px-2.5 py-1 rounded-md border transition"
              style={{ backgroundColor: horizon === d ? FAO_NAVY : "white", color: horizon === d ? "white" : "#475569", borderColor: horizon === d ? FAO_NAVY : "#CBD5E1" }}>
              {d}d
            </button>
          ))}
          <button onClick={() => setHorizon(365)}
            className="px-2.5 py-1 rounded-md border transition"
            style={{ backgroundColor: horizon === 365 ? FAO_NAVY : "white", color: horizon === 365 ? "white" : "#475569", borderColor: horizon === 365 ? FAO_NAVY : "#CBD5E1" }}>
            All
          </button>
        </div>
      </div>

      {activities.length === 0 && (
        <div className="rounded-xl border bg-white p-8 text-center" style={{ borderColor: "#E2E8F0" }}>
          <p className="text-sm" style={{ color: "#94A3B8", fontFamily: fontStack.body }}>No upcoming activities in the next {horizon} days.</p>
        </div>
      )}

      {grouped.map((group) => (
        <div key={group.label}>
          {/* Week divider */}
          <div className="flex items-center gap-3 px-1 py-2">
            <span className="text-[10px] font-bold uppercase tracking-widest" style={{ color: "#94A3B8", fontFamily: fontStack.mono }}>{group.label}</span>
            <div className="flex-1 h-px" style={{ backgroundColor: "#E2E8F0" }} />
            <span className="text-[10px]" style={{ color: "#94A3B8", fontFamily: fontStack.mono }}>{group.items.length} action{group.items.length !== 1 ? "s" : ""}</span>
          </div>

          <div className="space-y-2">
            {group.items.map((a) => {
              const urg = urgencyStyle(a.daysUntilEnd);
              const computed = a.p.computed;
              return (
                <div key={`${a.p.id}-${a.stage.key}`}
                  className="rounded-xl border flex items-stretch overflow-hidden"
                  style={{ borderColor: "#E2E8F0", backgroundColor: urg.bg }}>

                  {/* Left accent */}
                  <div className="w-1 flex-shrink-0" style={{ backgroundColor: CATEGORY_COLORS[a.p.category] || "#94A3B8" }} />

                  {/* Main content */}
                  <div className="flex-1 px-4 py-3 grid gap-1" style={{ gridTemplateColumns: "1fr auto" }}>
                    <div>
                      {/* Procurement title */}
                      <div className="flex items-center gap-2 flex-wrap">
                        <span className="text-sm font-semibold" style={{ color: FAO_NAVY, fontFamily: fontStack.body }}>{a.p.tender}</span>
                        {a.p.lot && <span className="text-[10px] px-1.5 py-0.5 rounded" style={{ backgroundColor: "#F1F5F9", color: "#64748B", fontFamily: fontStack.mono }}>{a.p.lot}</span>}
                        <MethodBadge method={a.p.method} />
                        <CategoryDot category={a.p.category} />
                        <span className="text-xs" style={{ color: "#64748B", fontFamily: fontStack.body }}>{a.p.category}</span>
                      </div>
                      {/* Next step */}
                      <div className="flex items-center gap-2 mt-1.5">
                        <span className={`text-[9px] font-bold uppercase tracking-wider px-1.5 py-0.5 rounded`}
                          style={{ backgroundColor: a.isCurrent ? "#DBEAFE" : "#F1F5F9", color: a.isCurrent ? "#1E40AF" : "#475569", fontFamily: fontStack.mono }}>
                          {a.isCurrent ? "In progress" : "Up next"}
                        </span>
                        <span className="text-sm font-medium" style={{ color: "#0F172A", fontFamily: fontStack.body }}>{a.stage.name}</span>
                        <span className="text-xs" style={{ color: "#64748B", fontFamily: fontStack.mono }}>
                          {a.stageStart && a.stageEnd ? `${fmtDate(a.stageStart)} → ${fmtDate(a.stageEnd)} · ${a.stage.plannedDays}d` : "—"}
                        </span>
                      </div>
                    </div>

                    {/* Right: deadline + risk */}
                    <div className="flex flex-col items-end justify-between pl-4">
                      <RiskPill risk={computed.risk} compact />
                      {a.stageEnd && (
                        <div className="text-right mt-1">
                          <div className="text-xs font-semibold" style={{ color: urg.color, fontFamily: fontStack.mono }}>
                            {a.daysUntilEnd < 0 ? `${Math.abs(a.daysUntilEnd)}d overdue` : a.daysUntilEnd === 0 ? "Due today" : `${a.daysUntilEnd}d left`}
                          </div>
                          <div className="text-[10px]" style={{ color: "#94A3B8", fontFamily: fontStack.mono }}>due {fmtDate(a.stageEnd)}</div>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      ))}
    </div>
  );
}

function AnalyticsTab({ enriched }) {
  const byCategory = useMemo(() => {
    const map = {};
    enriched.forEach((p) => { if (p.state !== "cancelled") map[p.category] = (map[p.category] || 0) + p.estInitial; });
    return Object.entries(map).map(([category, value]) => ({ category, value })).sort((a, b) => b.value - a.value);
  }, [enriched]);

  const byStage = useMemo(() => {
    const map = {};
    enriched.forEach((p) => {
      const current = p.stages.find((s) => s.key === p.currentKey);
      const name = current ? current.name : "—";
      map[name] = (map[name] || 0) + 1;
    });
    return Object.entries(map).map(([stage, count]) => ({ stage, count })).sort((a, b) => b.count - a.count);
  }, [enriched]);

  const byRisk = useMemo(() => {
    const map = { ok: 0, watch: 0, critical: 0, neutral: 0 };
    enriched.forEach((p) => { map[p.computed.risk.level] = (map[p.computed.risk.level] || 0) + 1; });
    return [
      { name: "On track", value: map.ok, color: "#16A34A" },
      { name: "Watch", value: map.watch, color: "#D97706" },
      { name: "Critical", value: map.critical, color: "#DC2626" },
    ].filter((r) => r.value > 0);
  }, [enriched]);

  const byMethod = useMemo(() => {
    const map = {};
    enriched.forEach((p) => { if (p.state !== "cancelled") map[p.method] = (map[p.method] || 0) + p.estInitial; });
    return Object.entries(map).map(([name, value]) => ({ name, value, color: name === "RFP" ? "#5B21B6" : "#1E3A8A" }));
  }, [enriched]);

  const ChartCard = ({ title, subtitle, children, height = 280 }) => (
    <div className="rounded-xl border bg-white p-5" style={{ borderColor: "#E2E8F0" }}>
      <h3 className="text-sm font-bold mb-1" style={{ color: FAO_NAVY, fontFamily: fontStack.display }}>{title}</h3>
      {subtitle && <p className="text-xs mb-4" style={{ color: "#64748B", fontFamily: fontStack.body }}>{subtitle}</p>}
      <div style={{ width: "100%", height }}>{children}</div>
    </div>
  );

  const TooltipBox = ({ active, payload, label, format }) => {
    if (!active || !payload || !payload.length) return null;
    return (
      <div className="rounded border px-3 py-2 text-xs"
        style={{ backgroundColor: "white", borderColor: "#CBD5E1", fontFamily: fontStack.body, boxShadow: "0 4px 12px rgba(0,0,0,0.08)" }}>
        <div className="font-semibold" style={{ color: FAO_NAVY }}>{label || payload[0].name}</div>
        <div style={{ color: "#475569", fontFamily: fontStack.mono }}>{format ? format(payload[0].value) : payload[0].value}</div>
      </div>
    );
  };

  return (
    <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
      <ChartCard title="Value by category" subtitle="Sum of estimated PO values, excluding cancelled procurements" height={320}>
        <ResponsiveContainer>
          <BarChart data={byCategory} layout="vertical" margin={{ top: 5, right: 30, left: 10, bottom: 5 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#F1F5F9" horizontal={false} />
            <XAxis type="number" tickFormatter={(v) => fmtUSDShort(v)} tick={{ fill: "#64748B", fontSize: 10, fontFamily: fontStack.mono }} stroke="#CBD5E1" />
            <YAxis dataKey="category" type="category" width={130} tick={{ fill: "#0F172A", fontSize: 11, fontFamily: fontStack.body }} stroke="#CBD5E1" />
            <Tooltip content={<TooltipBox format={fmtUSD} />} cursor={{ fill: "#F8FAFC" }} />
            <Bar dataKey="value" radius={[0, 4, 4, 0]}>
              {byCategory.map((entry, i) => <Cell key={i} fill={CATEGORY_COLORS[entry.category] || "#94A3B8"} />)}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </ChartCard>

      <ChartCard title="Procurements by current stage" subtitle="Workload distribution across the procurement workflow" height={320}>
        <ResponsiveContainer>
          <BarChart data={byStage} margin={{ top: 5, right: 20, left: 0, bottom: 80 }}>
            <CartesianGrid strokeDasharray="3 3" stroke="#F1F5F9" vertical={false} />
            <XAxis dataKey="stage" angle={-35} textAnchor="end" interval={0} tick={{ fill: "#475569", fontSize: 10, fontFamily: fontStack.body }} stroke="#CBD5E1" height={80} />
            <YAxis allowDecimals={false} tick={{ fill: "#64748B", fontSize: 10, fontFamily: fontStack.mono }} stroke="#CBD5E1" />
            <Tooltip content={<TooltipBox />} cursor={{ fill: "#F8FAFC" }} />
            <Bar dataKey="count" fill={FAO_BLUE} radius={[4, 4, 0, 0]} />
          </BarChart>
        </ResponsiveContainer>
      </ChartCard>

      <ChartCard title="Risk distribution" subtitle="Count of procurements by current risk flag" height={260}>
        <ResponsiveContainer>
          <PieChart>
            <Pie data={byRisk} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={90} innerRadius={50} paddingAngle={2}
              label={({ name, value }) => `${name}: ${value}`} labelLine={false}
              style={{ fontFamily: fontStack.body, fontSize: 11 }}>
              {byRisk.map((entry, i) => <Cell key={i} fill={entry.color} />)}
            </Pie>
            <Tooltip content={<TooltipBox />} />
          </PieChart>
        </ResponsiveContainer>
      </ChartCard>

      <ChartCard title="Value by procurement method" subtitle="ITB vs RFP — excluding cancelled procurements" height={260}>
        <ResponsiveContainer>
          <PieChart>
            <Pie data={byMethod} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={90} innerRadius={50} paddingAngle={2}
              label={({ name, value }) => `${name}: ${fmtUSDShort(value)}`} labelLine={false}
              style={{ fontFamily: fontStack.body, fontSize: 11 }}>
              {byMethod.map((entry, i) => <Cell key={i} fill={entry.color} />)}
            </Pie>
            <Tooltip content={<TooltipBox format={fmtUSD} />} />
          </PieChart>
        </ResponsiveContainer>
      </ChartCard>
    </div>
  );
}

/* ============================================================
   Inline-editable narrative cell
   ============================================================ */
function EditableNarrative({ value, onChange }) {
  const [editing, setEditing] = useState(false);
  const [draft, setDraft] = useState(value);
  const ref = useRef(null);

  // Keep draft in sync if parent updates
  useEffect(() => { setDraft(value); }, [value]);

  // Auto-focus when entering edit mode
  useEffect(() => {
    if (editing && ref.current) {
      ref.current.focus();
      // Place cursor at end
      const len = ref.current.value.length;
      ref.current.setSelectionRange(len, len);
    }
  }, [editing]);

  const startEdit = (e) => {
    e.stopPropagation(); // don't toggle row expansion
    setDraft(value);
    setEditing(true);
  };

  const save = () => {
    setEditing(false);
    if (draft.trim() !== value) onChange(draft.trim());
  };

  if (editing) {
    return (
      <textarea
        ref={ref}
        value={draft}
        onChange={(e) => setDraft(e.target.value)}
        onBlur={save}
        onKeyDown={(e) => {
          e.stopPropagation();
          if (e.key === "Escape") { setDraft(value); setEditing(false); }
        }}
        onClick={(e) => e.stopPropagation()}
        rows={3}
        style={{
          width: "100%",
          minWidth: 220,
          padding: "6px 8px",
          borderRadius: 6,
          border: `1.5px solid ${FAO_BLUE}`,
          fontSize: 12,
          lineHeight: 1.5,
          fontFamily: fontStack.body,
          color: "#0F172A",
          resize: "vertical",
          outline: "none",
          backgroundColor: "white",
          boxShadow: `0 0 0 3px ${FAO_BLUE}22`,
        }}
      />
    );
  }

  return (
    <div
      onClick={startEdit}
      className="group cursor-text"
      title="Click to edit"
      style={{ minHeight: 36, minWidth: 200 }}
    >
      {value ? (
        <p
          className="text-xs leading-relaxed"
          style={{
            color: "#334155",
            fontFamily: fontStack.body,
            display: "-webkit-box",
            WebkitLineClamp: 3,
            WebkitBoxOrient: "vertical",
            overflow: "hidden",
          }}
        >
          {value}
        </p>
      ) : (
        <span className="text-xs italic" style={{ color: "#CBD5E1", fontFamily: fontStack.body }}>
          click to add…
        </span>
      )}
      <span
        className="block text-[10px] mt-0.5 opacity-0 group-hover:opacity-100 transition"
        style={{ color: FAO_BLUE, fontFamily: fontStack.body }}
      >
        ✎ click to edit
      </span>
    </div>
  );
}

function PipelineTable({ enriched, expandedId, setExpandedId, sortBy, sortDir, setSort, updateProcurement }) {
  const sortIndicator = (key) => sortBy !== key ? null : sortDir === "desc" ? <ArrowDown size={10} className="inline ml-1" /> : <ArrowUp size={10} className="inline ml-1" />;

  const SortableHeader = ({ keyName, label, align = "left" }) => (
    <th onClick={() => setSort(keyName)}
      className={`px-3 py-3 text-[10px] uppercase tracking-wider font-semibold cursor-pointer select-none transition text-${align}`}
      style={{ color: "white", backgroundColor: sortBy === keyName ? "#243B57" : undefined }}>
      {label} {sortIndicator(keyName)}
    </th>
  );

  return (
    <div className="rounded-xl border overflow-hidden bg-white" style={{ borderColor: "#E2E8F0" }}>
      <div className="overflow-x-auto">
        <table className="w-full text-sm" style={{ fontFamily: fontStack.body, minWidth: "1700px" }}>
          <thead>
            <tr style={{ backgroundColor: FAO_NAVY, color: "white" }}>
              <th className="w-8 px-3 py-3"></th>
              <SortableHeader keyName="pr" label="PR / Lot" />
              <th className="text-left px-3 py-3 text-[10px] uppercase tracking-wider font-semibold" style={{ color: "white" }}>Tender</th>
              <SortableHeader keyName="category" label="Category" />
              <SortableHeader keyName="method" label="Method" align="center" />
              <SortableHeader keyName="estInitial" label="Est. PR" align="right" />
              <SortableHeader keyName="estPO" label="Est. PO" align="right" />
              <th className="text-center px-3 py-3 text-[10px] uppercase tracking-wider font-semibold select-none" style={{ color: "white" }}>Δ</th>
              <SortableHeader keyName="currentStage" label="Current stage" />
              <th className="text-center px-3 py-3 text-[10px] uppercase tracking-wider font-semibold select-none" style={{ color: "white" }}>Bids rec.</th>
              <th className="text-center px-3 py-3 text-[10px] uppercase tracking-wider font-semibold select-none" style={{ color: "white" }}>Responsive</th>
              <th className="text-left px-3 py-3 text-[10px] uppercase tracking-wider font-semibold select-none" style={{ minWidth: 250, color: "white" }}>
                Status narrative
              </th>
              <SortableHeader keyName="targetPO" label="Target PO" align="center" />
              <SortableHeader keyName="estPODate" label="Est. PO" align="center" />
              <SortableHeader keyName="risk" label="Risk" />
            </tr>
          </thead>
          <tbody>
            {enriched.map((p, idx) => {
              const currentStage = p.stages.find((s) => s.key === p.currentKey);
              const isOpen = expandedId === p.id;
              return (
                <React.Fragment key={p.id}>
                  <tr onClick={() => setExpandedId(isOpen ? null : p.id)} className="cursor-pointer transition border-t"
                    style={{
                      backgroundColor: isOpen
                        ? "#F8FAFC"
                        : (p.state === "pre_pipeline" || p.state === "not_applicable")
                        ? idx % 2 === 0 ? "#FAFAFA" : "#F5F5F5"
                        : idx % 2 === 0 ? "white" : "#FAFBFC",
                      borderColor: "#F1F5F9",
                      opacity: (p.state === "pre_pipeline" || p.state === "not_applicable") ? 0.75 : 1,
                    }}>
                    <td className="px-3 py-3" style={{ color: "#64748B" }}>{isOpen ? <ChevronDown size={16} /> : <ChevronRight size={16} />}</td>
                    <td className="px-3 py-3" onClick={(e) => e.stopPropagation()}>
                      <input
                        value={p.pr}
                        onChange={(e) => updateProcurement({ ...p, pr: e.target.value })}
                        className="font-semibold text-sm w-full bg-transparent border border-transparent hover:border-slate-200 focus:border-slate-300 rounded px-1 py-0.5"
                        style={{ color: FAO_NAVY, fontFamily: fontStack.mono, outline: "none", minWidth: 80 }}
                      />
                      {p.lot && <div className="text-xs px-1" style={{ color: "#64748B", fontFamily: fontStack.mono }}>{p.lot}</div>}
                    </td>
                    <td className="px-3 py-3" style={{ maxWidth: "260px" }}>
                      <div className="text-sm font-medium leading-snug" style={{ color: "#0F172A" }}>{p.tender}</div>
                      <div className="text-xs mt-0.5 truncate" style={{ color: "#94A3B8", fontFamily: fontStack.mono }}>{p.intend}</div>
                    </td>
                    <td className="px-3 py-3">
                      <div className="flex items-center gap-2">
                        <CategoryDot category={p.category} />
                        <span className="text-xs" style={{ color: "#334155" }}>{p.category}</span>
                      </div>
                    </td>
                    <td className="px-3 py-3 text-center"><MethodBadge method={p.method} /></td>
                    <td className="px-3 py-3 text-right text-sm tabular-nums" style={{ color: "#334155", fontFamily: fontStack.mono }}>{fmtUSD(p.estInitial)}</td>
                    <td className="px-3 py-3 text-right text-sm font-semibold tabular-nums" style={{ color: FAO_NAVY, fontFamily: fontStack.mono }}>{fmtUSD(p.estPO)}</td>
                    <td className="px-3 py-3 text-center"><VarianceTag variance={p.computed.variance} /></td>
                    <td className="px-3 py-3">
                      <div className="text-sm font-medium flex items-center gap-2" style={{ color: "#0F172A" }}>
                        <StatusDot status={p.state === "on_hold" || p.state === "cancelled" ? "blocked" : currentStage?.status || "in_progress"} />
                        {currentStage?.name || "—"}
                      </div>
                      {p.comments && p.comments.length > 0 && (
                        <div className="text-[10px] mt-0.5 inline-flex items-center gap-1" style={{ color: "#64748B", fontFamily: fontStack.body }}>
                          <MessageSquare size={9} /> {p.comments.length}
                        </div>
                      )}
                    </td>
                    <td className="px-3 py-3 text-center text-sm tabular-nums" style={{ color: p.nBids != null && p.nBids < 3 ? "#B45309" : "#334155", fontFamily: fontStack.mono }}>
                      {p.nBids != null ? <>{p.nBids}{p.nBids < 3 && " ⚠"}</> : "—"}
                    </td>
                    <td className="px-3 py-3 text-center" onClick={(e) => e.stopPropagation()}>
                      <input
                        type="number" min="0"
                        value={p.nResponsive ?? ""}
                        placeholder="—"
                        onChange={(e) => updateProcurement({ ...p, nResponsive: e.target.value === "" ? null : Number(e.target.value) })}
                        className="w-14 text-center px-1 py-0.5 rounded border text-sm tabular-nums bg-transparent"
                        style={{ borderColor: "#CBD5E1", fontFamily: fontStack.mono, color: "#334155", outline: "none" }}
                      />
                    </td>
                    <td className="px-3 py-2" style={{ verticalAlign: "top", paddingTop: 10 }}>
                      <EditableNarrative
                        value={p.narrative}
                        onChange={(val) => updateProcurement({ ...p, narrative: val })}
                      />
                    </td>
                    <td className="px-3 py-3 text-center text-xs" style={{ color: "#475569", fontFamily: fontStack.mono }}>{fmtDate(parseDate(p.targetPO))}</td>
                    <td className="px-3 py-3 text-center text-xs font-semibold"
                      style={{ color: p.computed.buffer == null ? "#94A3B8" : p.computed.buffer < 0 ? "#B91C1C" : p.computed.buffer <= 5 ? "#B45309" : "#15803D", fontFamily: fontStack.mono }}>
                      {fmtDate(p.computed.estPODate)}
                    </td>
                    <td className="px-3 py-3"><RiskPill risk={p.computed.risk} /></td>
                  </tr>
                  {isOpen && (
                    <tr>
                      <td colSpan={15} className="p-0">
                        <div className="border-t p-6 space-y-6" style={{ backgroundColor: "#F8FAFC", borderColor: "#E2E8F0" }}>
                          <StageEditor procurement={p} onUpdate={updateProcurement} />
                          <div className="border-t pt-6" style={{ borderColor: "#E2E8F0" }}>
                            <CommentsPanel procurement={p} onUpdate={updateProcurement} />
                          </div>
                        </div>
                      </td>
                    </tr>
                  )}
                </React.Fragment>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ============================================================
   Add Procurement Modal
   ============================================================ */
const CATEGORIES = Object.keys(CATEGORY_COLORS);

function Field({ label, error, children }) {
  return (
    <div>
      <label className="text-xs block mb-1 font-medium" style={{ color: "#475569", fontFamily: fontStack.body }}>
        {label}
      </label>
      {children}
      {error && <p className="text-[10px] mt-1" style={{ color: "#B91C1C", fontFamily: fontStack.body }}>{error}</p>}
    </div>
  );
}

function AddProcurementModal({ onAdd, onClose }) {
  const [form, setForm] = useState({
    pr: "",
    lot: "",
    intend: "",
    tender: "",
    category: "Livestock",
    method: "ITB",
    estInitial: "",
    nBids: "",
    opening: "",
    closing: "",
    currentKey: "pr_received",
    narrative: "",
    targetPO: "",
    state: "active",
  });
  const [errors, setErrors] = useState({});

  const setField = (key, val) => {
    setForm((prev) => {
      const updated = { ...prev, [key]: val };
      // Reset currentKey when method changes so it stays valid
      if (key === "method") {
        updated.currentKey = "pr_received";
      }
      return updated;
    });
    setErrors((prev) => ({ ...prev, [key]: undefined }));
  };

  const availableStages = useMemo(
    () => buildStages(form.method, Number(form.estInitial) || 0),
    [form.method, form.estInitial]
  );

  const validate = () => {
    const errs = {};
    if (!form.pr.trim()) errs.pr = "PR number is required";
    if (!form.tender.trim()) errs.tender = "Tender description is required";
    if (!form.estInitial || isNaN(Number(form.estInitial)) || Number(form.estInitial) <= 0)
      errs.estInitial = "Must be a positive number";
    if (!form.targetPO) errs.targetPO = "Target PO date is required";
    return errs;
  };

  const handleAdd = () => {
    const errs = validate();
    if (Object.keys(errs).length > 0) { setErrors(errs); return; }

    const id = `${form.pr.trim()}${form.lot.trim() ? `-${form.lot.trim().replace(/\s+/g, "")}` : ""}-${Date.now()}`;
    const stages = buildStages(form.method, Number(form.estInitial));
    const currentStatus = form.state === "on_hold" || form.state === "cancelled" ? "blocked" : "in_progress";
    const withProgress = markProgressTo(stages, form.currentKey, currentStatus);

    onAdd({
      id,
      pr: form.pr.trim(),
      lot: form.lot.trim(),
      intend: form.intend.trim(),
      tender: form.tender.trim(),
      category: form.category,
      method: form.method,
      estInitial: Number(form.estInitial),
      estPO: Number(form.estInitial),
      nBids: form.nBids !== "" ? Number(form.nBids) : null,
      opening: form.opening,
      closing: form.closing,
      currentKey: form.currentKey,
      narrative: form.narrative.trim(),
      targetPO: form.targetPO,
      state: form.state,
      comments: [],
      stages: withProgress,
    });
    onClose();
  };

  // Determine HQPC/RPC label hint
  const val = Number(form.estInitial) || 0;
  const reviewHint = val >= 500000 ? "HQPC review (+7d) will be auto-inserted" : val >= 200000 ? "RPC review (+7d) will be auto-inserted" : null;

  return (
    <div
      className="fixed inset-0 z-50 flex items-start justify-center overflow-y-auto"
      style={{ backgroundColor: "rgba(15,23,42,0.55)", padding: "24px 16px" }}
      onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}
    >
      <div className="bg-white rounded-2xl w-full max-w-2xl shadow-2xl" style={{ marginTop: "auto", marginBottom: "auto" }}>
        {/* Modal header */}
        <div
          className="px-6 py-5 flex items-center justify-between border-b"
          style={{ borderColor: "#E2E8F0", borderRadius: "1rem 1rem 0 0", background: `linear-gradient(135deg, ${FAO_NAVY} 0%, #243B57 100%)` }}
        >
          <div>
            <div className="text-[10px] uppercase tracking-[0.25em] mb-1" style={{ color: FAO_BLUE, fontFamily: fontStack.body }}>
              New procurement
            </div>
            <h2 className="text-lg font-bold text-white" style={{ fontFamily: fontStack.display }}>
              Add to pipeline
            </h2>
          </div>
          <button
            onClick={onClose}
            className="rounded-full p-2 hover:bg-white hover:bg-opacity-10 transition"
            style={{ color: "white" }}
          >
            <X size={18} />
          </button>
        </div>

        {/* Form body */}
        <div className="px-6 py-6 space-y-5">
          {/* Row: PR + Lot */}
          <div className="grid grid-cols-2 gap-4">
            <Field label="PR number *" error={errors.pr}>
              <input
                type="text"
                value={form.pr}
                onChange={(e) => setField("pr", e.target.value)}
                placeholder="e.g. 4201290"
                className="w-full px-3 py-2 rounded-lg border text-sm"
                style={{ borderColor: errors.pr ? "#B91C1C" : "#CBD5E1", fontFamily: fontStack.mono, color: "#0F172A" }}
              />
            </Field>
            <Field label="Lot (optional)">
              <input
                type="text"
                value={form.lot}
                onChange={(e) => setField("lot", e.target.value)}
                placeholder="e.g. Lot 1"
                className="w-full px-3 py-2 rounded-lg border text-sm"
                style={{ borderColor: "#CBD5E1", fontFamily: fontStack.body, color: "#0F172A" }}
              />
            </Field>
          </div>

          {/* InTend */}
          <Field label="InTend reference (optional)">
            <input
              type="text"
              value={form.intend}
              onChange={(e) => setField("intend", e.target.value)}
              placeholder="e.g. 2026/FLHAI/FLHAI/136XXX"
              className="w-full px-3 py-2 rounded-lg border text-sm"
              style={{ borderColor: "#CBD5E1", fontFamily: fontStack.mono, color: "#0F172A" }}
            />
          </Field>

          {/* Tender name */}
          <Field label="Tender description *" error={errors.tender}>
            <input
              type="text"
              value={form.tender}
              onChange={(e) => setField("tender", e.target.value)}
              placeholder="e.g. Supply of vegetable seeds"
              className="w-full px-3 py-2 rounded-lg border text-sm"
              style={{ borderColor: errors.tender ? "#B91C1C" : "#CBD5E1", fontFamily: fontStack.body, color: "#0F172A" }}
            />
          </Field>

          {/* Category + Method */}
          <div className="grid grid-cols-2 gap-4">
            <Field label="Category">
              <select
                value={form.category}
                onChange={(e) => setField("category", e.target.value)}
                className="w-full px-3 py-2 rounded-lg border text-sm bg-white"
                style={{ borderColor: "#CBD5E1", color: "#0F172A", fontFamily: fontStack.body }}
              >
                {CATEGORIES.map((c) => <option key={c} value={c}>{c}</option>)}
              </select>
            </Field>
            <Field label="Procurement method">
              <div className="flex gap-2">
                {["ITB", "RFP"].map((m) => (
                  <button
                    key={m}
                    onClick={() => setField("method", m)}
                    className="flex-1 py-2 rounded-lg border text-sm font-bold tracking-wider transition"
                    style={{
                      backgroundColor: form.method === m ? (m === "RFP" ? "#5B21B6" : "#1E3A8A") : "white",
                      color: form.method === m ? "white" : m === "RFP" ? "#5B21B6" : "#1E3A8A",
                      borderColor: m === "RFP" ? "#5B21B6" : "#1E3A8A",
                      fontFamily: fontStack.mono,
                    }}
                  >
                    {m}
                  </button>
                ))}
              </div>
            </Field>
          </div>

          {/* Value + nBids */}
          <div className="grid grid-cols-2 gap-4">
            <Field label="Estimated initial value (USD) *" error={errors.estInitial}>
              <div className="relative">
                <span className="absolute left-3 top-1/2 -translate-y-1/2 text-sm" style={{ color: "#64748B", fontFamily: fontStack.mono }}>$</span>
                <input
                  type="number"
                  min="0"
                  value={form.estInitial}
                  onChange={(e) => setField("estInitial", e.target.value)}
                  placeholder="0"
                  className="w-full pl-7 pr-3 py-2 rounded-lg border text-sm"
                  style={{ borderColor: errors.estInitial ? "#B91C1C" : "#CBD5E1", fontFamily: fontStack.mono, color: "#0F172A" }}
                />
              </div>
              {reviewHint && (
                <p className="text-[10px] mt-1" style={{ color: "#6D28D9", fontFamily: fontStack.body }}>
                  ⚡ {reviewHint}
                </p>
              )}
            </Field>
            <Field label="Number of bids received (optional)">
              <input
                type="number"
                min="0"
                value={form.nBids}
                onChange={(e) => setField("nBids", e.target.value)}
                placeholder="Leave blank if unknown"
                className="w-full px-3 py-2 rounded-lg border text-sm"
                style={{ borderColor: "#CBD5E1", fontFamily: fontStack.mono, color: "#0F172A" }}
              />
            </Field>
          </div>

          {/* Opening + Closing */}
          <div className="grid grid-cols-2 gap-4">
            <Field label="Opening date (optional)">
              <input
                type="date"
                value={form.opening}
                onChange={(e) => setField("opening", e.target.value)}
                className="w-full px-3 py-2 rounded-lg border text-sm"
                style={{ borderColor: "#CBD5E1", fontFamily: fontStack.mono, color: "#0F172A" }}
              />
            </Field>
            <Field label="Closing date (optional)">
              <input
                type="date"
                value={form.closing}
                onChange={(e) => setField("closing", e.target.value)}
                className="w-full px-3 py-2 rounded-lg border text-sm"
                style={{ borderColor: "#CBD5E1", fontFamily: fontStack.mono, color: "#0F172A" }}
              />
            </Field>
          </div>

          {/* Target PO + State */}
          <div className="grid grid-cols-2 gap-4">
            <Field label="Target PO date *" error={errors.targetPO}>
              <input
                type="date"
                value={form.targetPO}
                onChange={(e) => setField("targetPO", e.target.value)}
                className="w-full px-3 py-2 rounded-lg border text-sm"
                style={{ borderColor: errors.targetPO ? "#B91C1C" : "#CBD5E1", fontFamily: fontStack.mono, color: "#0F172A" }}
              />
            </Field>
            <Field label="Lifecycle state">
              <select
                value={form.state}
                onChange={(e) => setField("state", e.target.value)}
                className="w-full px-3 py-2 rounded-lg border text-sm bg-white"
                style={{ borderColor: "#CBD5E1", color: "#0F172A", fontFamily: fontStack.body }}
              >
                <option value="active">Active</option>
                <option value="on_hold">On hold</option>
                <option value="pre_pipeline">Pre-pipeline (no PR)</option>
                <option value="not_applicable">Not applicable (not with us)</option>
                <option value="cancelled">Cancelled / re-issue</option>
                <option value="awarded">Awarded / closed</option>
              </select>
            </Field>
          </div>

          {/* Current stage */}
          <Field label="Current standardized stage">
            <select
              value={form.currentKey}
              onChange={(e) => setField("currentKey", e.target.value)}
              className="w-full px-3 py-2 rounded-lg border text-sm bg-white"
              style={{ borderColor: "#CBD5E1", color: "#0F172A", fontFamily: fontStack.body }}
            >
              {availableStages.map((s) => (
                <option key={s.key} value={s.key}>{s.name}</option>
              ))}
            </select>
            <p className="text-[10px] mt-1" style={{ color: "#64748B", fontFamily: fontStack.body }}>
              All stages before the selected one will be marked complete; the selected stage will be marked in-progress.
            </p>
          </Field>

          {/* Narrative */}
          <Field label="Narrative status / notes (optional)">
            <textarea
              value={form.narrative}
              onChange={(e) => setField("narrative", e.target.value)}
              rows={2}
              placeholder="e.g. Tech evaluation underway — awaiting clarification from 2 bidders"
              className="w-full px-3 py-2 rounded-lg border text-sm"
              style={{ borderColor: "#CBD5E1", fontFamily: fontStack.body, color: "#0F172A", resize: "vertical" }}
            />
          </Field>

          {/* Preview */}
          {form.pr && form.tender && form.estInitial && (
            <div className="rounded-lg p-4 border" style={{ backgroundColor: "#F8FAFC", borderColor: "#E2E8F0" }}>
              <div className="text-[10px] uppercase tracking-wider mb-2 font-semibold" style={{ color: FAO_NAVY, fontFamily: fontStack.body }}>
                Preview
              </div>
              <div className="flex items-center gap-3 flex-wrap">
                <span className="font-semibold text-sm" style={{ color: FAO_NAVY, fontFamily: fontStack.mono }}>{form.pr}{form.lot && ` / ${form.lot}`}</span>
                <MethodBadge method={form.method} />
                <span className="text-sm" style={{ color: "#0F172A", fontFamily: fontStack.body }}>{form.tender}</span>
                <span className="text-sm font-semibold" style={{ color: FAO_NAVY, fontFamily: fontStack.mono }}>{fmtUSD(Number(form.estInitial))}</span>
                <CategoryDot category={form.category} />
                <span className="text-xs" style={{ color: "#64748B", fontFamily: fontStack.body }}>{form.category}</span>
              </div>
              <div className="mt-2 text-xs" style={{ color: "#64748B", fontFamily: fontStack.body }}>
                {availableStages.length} stages · {reviewHint || "No RPC/HQPC review threshold triggered"} · starts at: <strong>{availableStages.find(s => s.key === form.currentKey)?.name || "—"}</strong>
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="px-6 py-4 border-t flex items-center justify-end gap-3" style={{ borderColor: "#E2E8F0" }}>
          <button
            onClick={onClose}
            className="px-4 py-2 rounded-lg border text-sm font-medium transition hover:bg-slate-50"
            style={{ borderColor: "#CBD5E1", color: "#475569", fontFamily: fontStack.body }}
          >
            Cancel
          </button>
          <button
            onClick={handleAdd}
            className="inline-flex items-center gap-2 px-5 py-2 rounded-lg text-sm font-semibold text-white transition"
            style={{ backgroundColor: FAO_NAVY, fontFamily: fontStack.body }}
          >
            <Plus size={15} /> Add to pipeline
          </button>
        </div>
      </div>
    </div>
  );
}

const PASSWORD = "fao-haiti-2026";

function LoginGate({ onAuth }) {
  const [value, setValue] = useState("");
  const [error, setError] = useState(false);
  const submit = () => {
    if (value === PASSWORD) { onAuth(); }
    else { setError(true); setValue(""); }
  };
  return (
    <div className="min-h-screen flex items-center justify-center" style={{ backgroundColor: "#F1F5F9" }}>
      <div className="w-full max-w-sm rounded-2xl border p-8 shadow-sm" style={{ backgroundColor: "white", borderColor: "#E2E8F0" }}>
        <div className="mb-6 text-center">
          <div className="text-xs uppercase tracking-widest mb-1" style={{ color: FAO_BLUE, fontFamily: fontStack.body }}>FAO Haiti · OCHA</div>
          <h1 className="text-2xl font-bold" style={{ color: FAO_NAVY, fontFamily: fontStack.display }}>Procurement Tracking</h1>
        </div>
        <div className="space-y-3">
          <input
            type="password"
            placeholder="Password"
            value={value}
            autoFocus
            onChange={(e) => { setValue(e.target.value); setError(false); }}
            onKeyDown={(e) => e.key === "Enter" && submit()}
            className="w-full px-4 py-2.5 rounded-lg border text-sm"
            style={{ borderColor: error ? "#DC2626" : "#CBD5E1", fontFamily: fontStack.body, outline: "none" }}
          />
          {error && <p className="text-xs" style={{ color: "#DC2626", fontFamily: fontStack.body }}>Incorrect password.</p>}
          <button
            onClick={submit}
            className="w-full py-2.5 rounded-lg text-sm font-semibold text-white"
            style={{ backgroundColor: FAO_NAVY, fontFamily: fontStack.body }}
          >
            Enter
          </button>
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const [authed, setAuthed] = useState(false);
  const [procurements, setProcurements] = useState(initializeProcurements);
  const [expandedId, setExpandedId] = useState(null);
  const [riskFilter, setRiskFilter] = useState("all");
  const [methodFilter, setMethodFilter] = useState("all");
  const [tab, setTab] = useState("pipeline");
  const [sortBy, setSortBy] = useState("estPO");
  const [sortDir, setSortDir] = useState("desc");
  const [showAdd, setShowAdd] = useState(false);

  const handleAdd = useCallback((newProc) => {
    setProcurements((prev) => [newProc, ...prev]);
  }, []);

  const enriched = useMemo(() => procurements.map((p) => ({ ...p, computed: computeProcurement(p) })), [procurements]);

  const filtered = useMemo(() => {
    const result = enriched.filter((p) => {
      if (methodFilter !== "all" && p.method !== methodFilter) return false;
      if (riskFilter !== "all" && p.computed.risk.level !== riskFilter) return false;
      return true;
    });
    const riskOrder = { critical: 0, watch: 1, ok: 2, neutral: 3 };
    const stageOrder = (p) => p.stages.findIndex((s) => s.key === p.currentKey);
    const cmp = (a, b) => {
      let av, bv;
      switch (sortBy) {
        case "pr": av = a.pr; bv = b.pr; break;
        case "category": av = a.category; bv = b.category; break;
        case "method": av = a.method; bv = b.method; break;
        case "estInitial": av = a.estInitial; bv = b.estInitial; break;
        case "estPO": av = a.estPO; bv = b.estPO; break;
        case "currentStage": av = stageOrder(a); bv = stageOrder(b); break;
        case "targetPO": av = parseDate(a.targetPO)?.getTime() || 0; bv = parseDate(b.targetPO)?.getTime() || 0; break;
        case "estPODate": av = a.computed.estPODate.getTime(); bv = b.computed.estPODate.getTime(); break;
        case "risk": av = riskOrder[a.computed.risk.level]; bv = riskOrder[b.computed.risk.level]; break;
        default: av = a.estPO; bv = b.estPO;
      }
      if (av < bv) return sortDir === "desc" ? 1 : -1;
      if (av > bv) return sortDir === "desc" ? -1 : 1;
      return 0;
    };
    return result.slice().sort(cmp);
  }, [enriched, methodFilter, riskFilter, sortBy, sortDir]);

  const totals = useMemo(() => {
    const formal = enriched.filter((p) => p.state !== "pre_pipeline" && p.state !== "not_applicable");
    const activeOnly = formal.filter((p) => p.state !== "cancelled");
    return {
      portfolioTotal: formal.reduce((s, p) => s + p.estInitial, 0),
      activeTotal: activeOnly.reduce((s, p) => s + p.estInitial, 0),
      cancelledTotal: formal.filter((p) => p.state === "cancelled").reduce((s, p) => s + p.estInitial, 0),
      poIssued: formal.reduce((s, p) => s + (p.estPO || 0), 0),
      countAll: formal.length,
      countActive: activeOnly.length,
      countCancelled: formal.length - activeOnly.length,
      countPipeline: enriched.filter((p) => p.state === "pre_pipeline" || p.state === "not_applicable").length,
      ok: enriched.filter((p) => p.computed.risk.level === "ok").length,
      watch: enriched.filter((p) => p.computed.risk.level === "watch").length,
      critical: enriched.filter((p) => p.computed.risk.level === "critical").length,
    };
  }, [enriched]);

  const updateProcurement = useCallback((updated) => {
    setProcurements((prev) => prev.map((p) => (p.id === updated.id ? updated : p)));
  }, []);

  const setSort = (key) => {
    if (sortBy === key) setSortDir((d) => (d === "desc" ? "asc" : "desc"));
    else { setSortBy(key); setSortDir("desc"); }
  };

  const exportToExcel = () => {
    const rows = enriched
      .filter((p) => p.state !== "pre_pipeline" && p.state !== "not_applicable")
      .map((p) => {
        const currentStage = p.stages.find((s) => s.key === p.currentKey);
        const { timelineStart, timelineEnd } = computeTimeline(p);
        return {
          "PR / Lot": p.lot ? `${p.pr} – ${p.lot}` : p.pr,
          "InTend Ref.": p.intend,
          "Tender": p.tender,
          "Category": p.category,
          "Method": p.method,
          "Est. PR ($)": p.estInitial,
          "Est. PO ($)": p.estPO || 0,
          "Variance %": p.estPO && p.estInitial ? `${(((p.estPO - p.estInitial) / p.estInitial) * 100).toFixed(1)}%` : "—",
          "State": p.state,
          "Current Stage": currentStage?.name || "—",
          "Bids Received": p.nBids ?? "",
          "Responsive": p.nResponsive ?? "",
          "Opening Date": fmtDate(parseDate(p.opening)),
          "Closing Date": fmtDate(parseDate(p.closing)),
          "Timeline Start": fmtDate(timelineStart),
          "Timeline End": fmtDate(timelineEnd),
          "Target PO": fmtDate(parseDate(p.targetPO)),
          "Est. PO Date": fmtDate(p.computed.estPODate),
          "Buffer (days)": p.computed.buffer ?? "",
          "Risk": p.computed.risk.label,
          "Status Narrative": p.narrative,
        };
      });

    const ws = XLSX.utils.json_to_sheet(rows);

    // Column widths
    const colWidths = [14, 30, 50, 20, 8, 14, 14, 10, 14, 30, 12, 12, 14, 14, 14, 14, 14, 14, 12, 20, 60];
    ws["!cols"] = colWidths.map((w) => ({ wch: w }));

    // Header row style
    const range = XLSX.utils.decode_range(ws["!ref"]);
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cell = ws[XLSX.utils.encode_cell({ r: 0, c: C })];
      if (cell) {
        cell.s = {
          font: { bold: true, color: { rgb: "FFFFFF" } },
          fill: { fgColor: { rgb: "1A2E44" } },
          alignment: { horizontal: "center", vertical: "center", wrapText: true },
          border: { bottom: { style: "thin", color: { rgb: "009FDA" } } },
        };
      }
    }

    // Alternate row fill + number format for money columns
    const moneyFmt = '#,##0';
    const moneyCols = [5, 6]; // Est. PR, Est. PO
    for (let R = 1; R <= range.e.r; R++) {
      const fill = R % 2 === 0 ? { fgColor: { rgb: "F1F5F9" } } : { fgColor: { rgb: "FFFFFF" } };
      for (let C = range.s.c; C <= range.e.c; C++) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        if (!ws[addr]) ws[addr] = { t: "z" };
        ws[addr].s = { fill, alignment: { vertical: "center", wrapText: C === 20 } };
        if (moneyCols.includes(C) && typeof ws[addr].v === "number") ws[addr].z = moneyFmt;
      }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Procurement Pipeline");

    // Summary sheet
    const summaryRows = [
      ["FAO Haiti – OCHA Procurement Tracking"],
      [`Exported: ${fmtDate(TODAY)}`],
      [],
      ["SUMMARY", ""],
      ["Portfolio Total (PR)", totals.portfolioTotal],
      ["Active Value (PR)", totals.activeTotal],
      ["Total PO Issued", totals.poIssued],
      ["On Track", totals.ok],
      ["Watch", totals.watch],
      ["Critical", totals.critical],
    ];
    const ws2 = XLSX.utils.aoa_to_sheet(summaryRows);
    ws2["!cols"] = [{ wch: 28 }, { wch: 18 }];
    ws2["A1"].s = { font: { bold: true, sz: 14, color: { rgb: "1A2E44" } } };
    ws2["A4"].s = { font: { bold: true } };
    ["B5", "B6", "B7"].forEach((addr) => { if (ws2[addr]) ws2[addr].z = '"$"#,##0'; });
    XLSX.utils.book_append_sheet(wb, ws2, "Summary");

    XLSX.writeFile(wb, `FAO-Haiti-Procurement-${toISODate(TODAY)}.xlsx`);
  };

  const TabButton = ({ id, label, Icon }) => (
    <button onClick={() => setTab(id)}
      className="inline-flex items-center gap-2 px-4 py-2 rounded-md text-sm font-medium transition"
      style={{ backgroundColor: tab === id ? FAO_NAVY : "transparent", color: tab === id ? "white" : "#475569", fontFamily: fontStack.body, border: tab === id ? `1px solid ${FAO_NAVY}` : "1px solid #E2E8F0" }}>
      <Icon size={14} />{label}
    </button>
  );

  const withPO = enriched.filter((p) => p.estPO > 0);
  const avgVariance = withPO.length ? (withPO.reduce((s, p) => s + p.computed.variance, 0) / withPO.length) * 100 : 0;

  if (!authed) return <LoginGate onAuth={() => setAuthed(true)} />;

  return (
    <div style={{ backgroundColor: "#F1F5F9", minHeight: "100vh" }}>
      <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Serif:wght@600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap" />

      <header className="px-8 py-6" style={{ background: `linear-gradient(135deg, ${FAO_NAVY} 0%, #243B57 100%)`, color: "white" }}>
        <div className="flex items-center justify-between flex-wrap gap-4">
          <div>
            <div className="text-[10px] uppercase tracking-[0.3em] mb-1" style={{ color: FAO_BLUE, fontFamily: fontStack.body }}>FAO Haiti · Country Office · OSRO/HAI/061/CHA</div>
            <h1 className="text-2xl font-bold leading-tight" style={{ fontFamily: fontStack.display, letterSpacing: "-0.01em" }}>Procurement Pipeline Tracker</h1>
            <div className="text-sm mt-1" style={{ color: "#CBD5E1", fontFamily: fontStack.body }}>Standardized stages · plan-vs-actual timeline · risk monitoring · PM: Costantino, Claudio (FLHAI)</div>
          </div>
          <div className="text-xs" style={{ color: "#94A3B8", fontFamily: fontStack.mono }}>{fmtDate(TODAY)}</div>
        </div>
      </header>

      <div className="px-8 py-6">
        <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-7 gap-4">
          <KPI label="Portfolio total (PR)" value={fmtUSDShort(totals.portfolioTotal)} sub={`${totals.countAll} solicitations (incl. cancelled) · ${totals.countPipeline} pipeline`} accent={FAO_NAVY} />
          <KPI label="Active value (PR)" value={fmtUSDShort(totals.activeTotal)} sub={`Excludes ${totals.countCancelled} cancelled (${fmtUSDShort(totals.cancelledTotal)})`} accent={FAO_BLUE} />
          <KPI label="PO issued" value={fmtUSDShort(totals.poIssued)} sub="Sum of est. PO values" accent="#0F766E" />
          <KPI label="On track" value={totals.ok} sub="🟢 within plan" accent="#16A34A" />
          <KPI label="Watch" value={totals.watch} sub="🟠 tight buffer / hold" accent="#D97706" />
          <KPI label="Critical" value={totals.critical} sub="🔴 overrun / cancelled" accent="#DC2626" />
          <KPI label="Avg. PR→PO variance" value={`${avgVariance.toFixed(1)}%`} sub="PO vs. PR estimate" accent="#9333EA" />
        </div>
      </div>

      <div className="px-8 pb-4 flex items-center gap-2 flex-wrap">
        <TabButton id="pipeline" label="Pipeline" Icon={ListIcon} />
        <TabButton id="gantt" label="Gantt" Icon={GanttChartSquare} />
        <TabButton id="analytics" label="Analytics" Icon={BarChart3} />
        <TabButton id="upcoming" label="Upcoming" Icon={CalendarDays} />

        <button
          onClick={() => setShowAdd(true)}
          className="inline-flex items-center gap-2 px-4 py-2 rounded-md text-sm font-medium text-white transition ml-2"
          style={{ backgroundColor: FAO_BLUE, fontFamily: fontStack.body, border: `1px solid ${FAO_BLUE}` }}
        >
          <Plus size={14} /> Add procurement
        </button>
        <button
          onClick={exportToExcel}
          className="inline-flex items-center gap-2 px-4 py-2 rounded-md text-sm font-medium transition"
          style={{ backgroundColor: "white", color: FAO_NAVY, fontFamily: fontStack.body, border: `1px solid #CBD5E1` }}
        >
          <Download size={14} /> Export Excel
        </button>
        {tab === "pipeline" && (
          <div className="ml-auto flex items-center gap-2 flex-wrap">
            <Filter size={14} style={{ color: "#64748B" }} />
            <select value={methodFilter} onChange={(e) => setMethodFilter(e.target.value)}
              className="px-3 py-1.5 rounded-md border text-sm bg-white"
              style={{ borderColor: "#CBD5E1", color: "#334155", fontFamily: fontStack.body }}>
              <option value="all">All methods</option>
              <option value="ITB">ITB only</option>
              <option value="RFP">RFP only</option>
            </select>
            <select value={riskFilter} onChange={(e) => setRiskFilter(e.target.value)}
              className="px-3 py-1.5 rounded-md border text-sm bg-white"
              style={{ borderColor: "#CBD5E1", color: "#334155", fontFamily: fontStack.body }}>
              <option value="all">All risk levels</option>
              <option value="ok">On track</option>
              <option value="watch">Watch</option>
              <option value="critical">Critical</option>
            </select>
            <div className="text-xs" style={{ color: "#64748B", fontFamily: fontStack.body }}>
              {filtered.length} of {enriched.length} · click column to sort
            </div>
          </div>
        )}
      </div>

      <div className="px-8 pb-12">
        {tab === "pipeline" && (
          <>
            <PipelineTable enriched={filtered} expandedId={expandedId} setExpandedId={setExpandedId}
              sortBy={sortBy} sortDir={sortDir} setSort={setSort} updateProcurement={updateProcurement} />
            <div className="mt-6 rounded-lg p-5 text-xs"
              style={{ backgroundColor: "white", border: "1px solid #E2E8F0", color: "#475569", fontFamily: fontStack.body, lineHeight: 1.6 }}>
              <div className="font-bold mb-2 text-[10px] uppercase tracking-widest" style={{ color: FAO_NAVY }}>Risk methodology</div>
              <p>
                Estimated PO date = today + Σ planned days for remaining (non-complete, non-skipped) stages. Buffer = target PO date − estimated PO date. A procurement is{" "}
                <span style={{ color: "#15803D", fontWeight: 600 }}>on track</span> with buffer &gt; 5 days,{" "}
                <span style={{ color: "#B45309", fontWeight: 600 }}>watch</span> with 0–5 days buffer or on-hold state, and{" "}
                <span style={{ color: "#B91C1C", fontWeight: 600 }}>critical</span> when overrun (negative buffer) or cancelled. RPC review (+7d, $200K–$500K) and HQPC review (+7d, ≥$500K) are inserted automatically by value.
              </p>
            </div>
          </>
        )}
        {tab === "gantt" && <GanttChart procurements={filtered} onUpdate={updateProcurement} />}
        {tab === "analytics" && <AnalyticsTab enriched={enriched} />}
        {tab === "upcoming" && <UpcomingActivities enriched={enriched} />}
      </div>

      {showAdd && <AddProcurementModal onAdd={handleAdd} onClose={() => setShowAdd(false)} />}
    </div>
  );
}
