import * as React from "react";
import {
  useReactTable,
  getCoreRowModel,
  getExpandedRowModel,
  flexRender,
  createColumnHelper,
  ExpandedState,
} from "@tanstack/react-table";
import * as ReactDOM from "react-dom";
import { TaskNode, updateNodeInTree } from "./buildTree";
import { FinancialTimeline } from "./FinancialTimeline";

interface Props {
  data:                 TaskNode[];
  onSave:               (changes: Record<string, Partial<TaskNode>>) => Promise<void>;
  onRefresh:            () => void;
  userId:               string;
  taskIds:              string[];
  latestApprovedBudget: number;
  projectId:            string;
}

const COST_CATEGORIES = [
  { value: 686490000, label: "Staff and Other Personnel Costs" },
  { value: 686490001, label: "Supplies, Commodities, and Materials" },
  { value: 686490002, label: "Equipment, Vehicles, and Furniture" },
  { value: 686490003, label: "Contractual Services" },
  { value: 686490004, label: "Travel" },
  { value: 686490005, label: "Indirect Costs" },
];

const FUNDING_SOURCES = [
  { value: 0, label: "Regular Budget" },
  { value: 1, label: "PK + Support Account" },
  { value: 2, label: "xB" },
  { value: 3, label: "10RCR (Cost Recovery)" },
  { value: 4, label: "20PCR (PK Cost Recovery)" },
];

const ALL_COLUMNS = [
  { id: "startDate",        label: "Start",              group: "schedule" },
  { id: "endDate",          label: "Finish",             group: "schedule" },
  { id: "pctDone",          label: "% Complete",         group: "schedule" },
  { id: "assignedTo",       label: "Assigned to",        group: "schedule" },
  { id: "quantity",         label: "Effort (h)",         group: "cost" },
  { id: "effortCompleted",  label: "Completed (h)",      group: "cost" },
  { id: "unitRate",         label: "Unit rate",          group: "cost" },
  { id: "totalPlannedCost", label: "Total planned",      group: "cost" },
  { id: "srcServiceName",   label: "Service",            group: "cost" },
  { id: "unit",             label: "Unit",               group: "cost" },
  { id: "plannedCost",      label: "Planned cost",       group: "cost" },
  { id: "fixedCost",        label: "Fixed cost",         group: "cost" },
  { id: "actualCost",       label: "Actual effort cost", group: "cost" },
  { id: "actualFixedCost",  label: "Actual fixed cost",  group: "cost" },
  { id: "totalActualCost",  label: "Total actual cost",  group: "cost" },
  { id: "remainingCost",    label: "Remaining",          group: "cost" },
  { id: "earnedValue",      label: "Earned value",       group: "cost" },
  { id: "fundingSource",    label: "Funding source",     group: "cost" },
  { id: "costCategory",     label: "Cost category",      group: "cost" },
] as const;

const DEFAULT_VISIBLE = new Set([
  // Schedule
  "startDate", "endDate", "pctDone", "assignedTo",
  // Effort
  "quantity", "effortCompleted",
  // Rate
  "unitRate",
  // Costs - keep it simple
  "totalPlannedCost", "totalActualCost", "remainingCost",
  // Classification
  "fundingSource", "costCategory",
]);

const DEFAULT_ORDER = [
  "select", "taskName",

  // Schedule
  "startDate", "endDate", "pctDone", "assignedTo",

  // Classification (set these first before costs)
  "costCategory", "fundingSource",

  // Service & Rate
  "srcServiceName", "unit", "unitRate",

  // Effort
  "quantity", "effortCompleted",

  // Planned costs (inputs → total)
  "plannedCost", "fixedCost", "totalPlannedCost",

  // Actual costs (inputs → total)
  "actualCost", "actualFixedCost", "totalActualCost",

  // Outcome metrics
  "remainingCost", "earnedValue",
];

interface SrcItem {
  id:             string;
  serviceId:      string;
  name:           string;
  price:          number;
  unit:           string;
  frequency:      string;
  fiscalYearId:   string | null;
  fiscalYearName: string | null;
  entityId:       string | null;
  entityName:     string | null;
}

interface EntityItem {
  id:   string;
  name: string;
}

interface FiscalYearItem {
  id:   string;
  name: string;
}

interface ResourceItem {
  id:   string;
  name: string;
}

// taskId → array of resource names
type TaskResourceMap = Record<string, ResourceItem[]>;

const col = createColumnHelper<TaskNode>();

function fmtCurrency(v: number): string {
  if (!v && v !== 0) return "—";
  return new Intl.NumberFormat(undefined, {
    style: "currency", currency: "USD", minimumFractionDigits: 2,
  }).format(v);
}

function fmtNumber(v: number): string {
  if (!v && v !== 0) return "—";
  return new Intl.NumberFormat("en-GB", {
    minimumFractionDigits: 2, maximumFractionDigits: 2,
  }).format(v);
}

function formatDate(raw: string): string {
  if (!raw) return "";
  const d = new Date(raw);
  if (isNaN(d.getTime())) return raw;
  return d.toLocaleDateString("en-GB", { day: "2-digit", month: "2-digit", year: "numeric" });
}

function BulkDropdown({ label, children }: { label: string; children: React.ReactNode }) {
  const [open, setOpen] = React.useState(false);
  const ref = React.useRef<HTMLDivElement>(null);

  React.useEffect(() => {
    if (!open) return;
    function handler(e: MouseEvent) {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    }
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [open]);

  return (
    <div ref={ref} style={{ position: "relative" }}>
      <button className="tg-btn" style={{ color: "#107c10", fontWeight: 500 }}
        onClick={() => setOpen(o => !o)}>
        {label}
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none"
          stroke="currentColor" strokeWidth="2.5">
          <polyline points="6 9 12 15 18 9"/>
        </svg>
      </button>
      {open && (
        <div style={{
          position: "absolute", top: "100%", left: 0, marginTop: 2,
          background: "white", border: "1px solid #e5e7eb", borderRadius: 4,
          boxShadow: "0 4px 16px rgba(0,0,0,0.12)", zIndex: 99999,
          minWidth: 220, padding: "4px 0",
        }}
          onClick={() => setOpen(false)}>
          {children}
        </div>
      )}
    </div>
  );
}

function BulkMenuItem({ onClick, children }: { onClick: () => void; children: React.ReactNode }) {
  return (
    <div onClick={onClick} style={{
      padding: "7px 14px", fontSize: 13, cursor: "pointer",
      color: "#1f2937", userSelect: "none",
    }}
      onMouseEnter={e => (e.currentTarget.style.background = "#f3f2f1")}
      onMouseLeave={e => (e.currentTarget.style.background = "white")}>
      {children}
    </div>
  );
}

function BulkUnitRateInput({ onApply }: { onApply: (rate: number) => void }) {
  const [value, setValue] = React.useState("");
  return (
    <div style={{ padding: "8px 10px" }} onClick={e => e.stopPropagation()}>
      <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 6, fontWeight: 600 }}>
        Enter rate per hour
      </div>
      <div style={{ display: "flex", gap: 6 }}>
        <input
          autoFocus
          type="text"
          inputMode="decimal"
          placeholder="e.g. 50.00"
          value={value}
          onChange={e => setValue(e.target.value)}
          onKeyDown={e => {
            if (e.key === "Enter") {
              const v = parseFloat(value);
              if (!isNaN(v) && v >= 0) { onApply(v); setValue(""); }
            }
          }}
          style={{
            width: 110, border: "1px solid #d1d5db", borderRadius: 3,
            padding: "5px 8px", fontSize: 13, outline: "none",
            fontFamily: "Segoe UI, system-ui, sans-serif",
          }}
        />
        <button
          onClick={() => {
            const v = parseFloat(value);
            if (!isNaN(v) && v >= 0) { onApply(v); setValue(""); }
          }}
          style={{
            padding: "5px 10px", background: "#107c10", color: "white",
            border: "none", borderRadius: 3, cursor: "pointer", fontSize: 13,
            fontWeight: 600,
          }}>
          Apply
        </button>
      </div>
    </div>
  );
}

function BulkServiceSearch({ items, onSelect }: {
  items: SrcItem[]; onSelect: (id: string) => void;
}) {
  const [query, setQuery] = React.useState("");
  const filtered = query.trim() === "" ? items
    : items.filter(s =>
        s.serviceId.toLowerCase().includes(query.toLowerCase()) ||
        s.name.toLowerCase().includes(query.toLowerCase()) ||
        (s.entityName ?? "").toLowerCase().includes(query.toLowerCase())
      );
  return (
    <React.Fragment>
      <input autoFocus placeholder="Search service..."
        value={query} onChange={e => setQuery(e.target.value)}
        style={{
          width: "100%", border: "1px solid #d1d5db", borderRadius: 2,
          padding: "4px 8px", fontSize: 13, outline: "none",
          boxSizing: "border-box", marginBottom: 4,
        }}
        onClick={e => e.stopPropagation()}
      />
      <div style={{ maxHeight: 220, overflowY: "auto" }}>
        {filtered.map(s => (
          <div key={s.id} onClick={() => onSelect(s.id)}
            style={{ padding: "6px 8px", fontSize: 13, cursor: "pointer", borderRadius: 2 }}
            onMouseEnter={e => (e.currentTarget.style.background = "#f3f2f1")}
            onMouseLeave={e => (e.currentTarget.style.background = "white")}>
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <span style={{ fontWeight: 500, fontSize: 14 }}>{s.serviceId}</span>
              <span style={{ fontSize: 11, color: "#0078d4" }}>{s.entityName}</span>
            </div>
            <div style={{ fontSize: 11, color: "#6b7280" }}>{s.name}</div>
          </div>
        ))}
      </div>
    </React.Fragment>
  );
}

function GaugeChart({ pct, color }: { pct: number; color: string }) {
  const r = 82, cx = 105, cy = 100;
  const angle = Math.PI + (pct / 100) * Math.PI;
  const x1 = cx - r, y1 = cy;
  const x2 = cx + r, y2 = cy;
  const px = cx + r * Math.cos(angle);
  const py = cy + r * Math.sin(angle);
  return (
    <svg width="210" height="118" viewBox="0 0 210 118">
      <path d={`M ${x1} ${y1} A ${r} ${r} 0 1 1 ${x2} ${y2}`}
        fill="none" stroke="#e5e7eb" strokeWidth="16" strokeLinecap="round"/>
      {pct > 0 && (
        <path d={`M ${x1} ${y1} A ${r} ${r} 0 0 0 ${px} ${py}`}
          fill="none" stroke={color} strokeWidth="16" strokeLinecap="round"/>
      )}
      <line x1={cx} y1={cy}
        x2={cx + (r - 12) * Math.cos(angle)}
        y2={cy + (r - 12) * Math.sin(angle)}
        stroke="#374151" strokeWidth="2.5" strokeLinecap="round"/>
      <circle cx={cx} cy={cy} r="5" fill="#374151"/>
      <text x="8" y="116" fontSize="10" fill="#9ca3af" fontFamily="Segoe UI, sans-serif">0%</text>
      <text x="172" y="116" fontSize="10" fill="#9ca3af" fontFamily="Segoe UI, sans-serif">100%</text>
      <text x={cx} y={cy - 16} fontSize="24" fontWeight="800" fill={color}
        textAnchor="middle" fontFamily="Segoe UI, sans-serif">
        {Math.min(pct, 999).toFixed(1)}%
      </text>
      <text x={cx} y={cy} fontSize="11" fill="#6b7280"
        textAnchor="middle" fontFamily="Segoe UI, sans-serif">spent</text>
    </svg>
  );
}

function DonutChart({ slices, size = 120 }: {
  slices: { label: string; value: number; color: string }[];
  size?: number;
}) {
  const total = slices.reduce((s, d) => s + d.value, 0);
  if (total === 0) return <p style={{ color: "#9ca3af", fontSize: 12 }}>No data</p>;
  const cx = size / 2, cy = size / 2, r = size / 2 - 10, ir = r * 0.55;
  let cumAngle = -Math.PI / 2;
  return (
    <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
      {slices.filter(s => s.value > 0).map((s, i) => {
        const angle = (s.value / total) * 2 * Math.PI;
        const x1 = cx + r * Math.cos(cumAngle);
        const y1 = cy + r * Math.sin(cumAngle);
        const x2 = cx + r * Math.cos(cumAngle + angle);
        const y2 = cy + r * Math.sin(cumAngle + angle);
        const ix1 = cx + ir * Math.cos(cumAngle);
        const iy1 = cy + ir * Math.sin(cumAngle);
        const ix2 = cx + ir * Math.cos(cumAngle + angle);
        const iy2 = cy + ir * Math.sin(cumAngle + angle);
        const large = angle > Math.PI ? 1 : 0;
        const path = `M ${ix1} ${iy1} L ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2} L ${ix2} ${iy2} A ${ir} ${ir} 0 ${large} 0 ${ix1} ${iy1} Z`;
        cumAngle += angle;
        return <path key={i} d={path} fill={s.color} stroke="white" strokeWidth="1.5"/>;
      })}
    </svg>
  );
}

const DONUT_COLORS = ["#1e40af","#0f766e","#d97706","#dc2626","#7c3aed","#0284c7","#16a34a","#9ca3af"];


function InfoPopover({ title, explanation, formula, thresholds }: {
  title: string;
  explanation: string;
  formula: string;
  thresholds: string;
}) {
  const [pos, setPos] = React.useState<{ top: number; left: number } | null>(null);
  const btnRef = React.useRef<HTMLButtonElement>(null);

  React.useEffect(() => {
    if (!pos) return;
    function handler(e: MouseEvent) {
      const portal = document.getElementById("tg-info-portal");
      if (portal && !portal.contains(e.target as Node) &&
          btnRef.current && !btnRef.current.contains(e.target as Node)) {
        setPos(null);
      }
    }
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [pos]);

  function handleClick(e: React.MouseEvent) {
    e.stopPropagation();
    if (pos) { setPos(null); return; }
    if (btnRef.current) {
      const rect = btnRef.current.getBoundingClientRect();
      const popWidth = 260;
      const left = Math.min(rect.left, window.innerWidth - popWidth - 12);
      setPos({ top: rect.bottom + 6, left });
    }
  }

  const portal = pos ? ReactDOM.createPortal(
    <div id="tg-info-portal" style={{
      position: "fixed", top: pos.top, left: pos.left,
      background: "white", border: "1px solid #e5e7eb",
      borderRadius: 8, boxShadow: "0 4px 16px rgba(0,0,0,0.12)",
      zIndex: 99999, width: 260, padding: "14px 16px",
    }}>
      <div style={{ fontSize: 13, fontWeight: 700, color: "#111827", marginBottom: 8 }}>{title}</div>
      <div style={{ fontSize: 12, color: "#374151", lineHeight: 1.6, marginBottom: 10 }}>{explanation}</div>
      <div style={{ background: "#f3f4f6", borderRadius: 6, padding: "8px 10px", marginBottom: 10 }}>
        <div style={{ fontSize: 10, fontWeight: 700, color: "#6b7280", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 4 }}>Formula</div>
        <div style={{ fontSize: 12, fontWeight: 600, color: "#1e40af", fontFamily: "monospace" }}>{formula}</div>
      </div>
      <div style={{ fontSize: 11, color: "#6b7280", lineHeight: 1.6, borderTop: "1px solid #f3f4f6", paddingTop: 8 }}>{thresholds}</div>
    </div>,
    document.body
  ) : null;

  return (
    <React.Fragment>
      <button
        ref={btnRef}
        onClick={handleClick}
        style={{
          background: "none", border: "none", cursor: "pointer",
          color: "#9ca3af", fontSize: 14, lineHeight: 1,
          padding: "0 0 0 5px", display: "inline-flex", alignItems: "center",
        }}
      >ⓘ</button>
      {portal}
    </React.Fragment>
  );
}

function SummaryPanel({ data, onClose, latestApprovedBudget }: { data: TaskNode[]; onClose: () => void; latestApprovedBudget: number }) {

  function flatLeaves(nodes: TaskNode[]): TaskNode[] {
    const out: TaskNode[] = [];
    function walk(n: TaskNode) {
      if (!n.subRows || n.subRows.length === 0) { out.push(n); return; }
      n.subRows.forEach(walk);
    }
    nodes.forEach(walk);
    return out;
  }
  const leaves = flatLeaves(data);

  const totalPlanned   = leaves.reduce((s, n) => s + (n.totalPlannedCost ?? 0), 0);
  const totalActual    = leaves.reduce((s, n) => s + (n.totalActualCost  ?? 0), 0);
  const totalRemaining = totalPlanned - totalActual;
  const totalEV        = leaves.reduce((s, n) => s + (n.earnedValue      ?? 0), 0);

  const weightedPct = data[0]?.pctDone ?? 0;

  const consumptionRate = (weightedPct > 0 && totalPlanned > 0)
    ? (totalActual / totalPlanned) / (weightedPct / 100)
    : null;
  const kpi1Status = consumptionRate === null ? null
    : consumptionRate <= 1.10 ? { label: "On Track",         color: "#16a34a", bg: "#f0fdf4" }
    : consumptionRate <= 1.30 ? { label: "Overrun Risk",     color: "#d97706", bg: "#fffbeb" }
    :                           { label: "High Overrun Risk", color: "#dc2626", bg: "#fef2f2" };

  const availablePct = totalPlanned > 0 ? (totalRemaining / totalPlanned) * 100 : 0;
  const kpi2Status = availablePct > 15
    ? { label: "High Balance Available", color: "#16a34a", bg: "#f0fdf4" }
    : availablePct >= 5
    ? { label: "Low Balance Available",  color: "#d97706", bg: "#fffbeb" }
    : { label: "Critical Balance",       color: "#dc2626", bg: "#fef2f2" };

  const pctConsumed = totalPlanned > 0 ? Math.min((totalActual / totalPlanned) * 100, 100) : 0;
  const gaugeColor  = kpi2Status.color;

  // KPI 3 — Needs PCR: triggered when totalPlanned exceeds latestApprovedBudget by ≥10%
  const pcrOverrun = latestApprovedBudget > 0
    ? (totalPlanned - latestApprovedBudget) / latestApprovedBudget
    : null;
  const pcrStatus = pcrOverrun === null
    ? { label: "Not Set",       color: "#6b7280", bg: "#f9fafb", triggered: false, notSet: true,  approver: "" }
    : pcrOverrun > 0.25
    ? { label: "PCR Required",  color: "#dc2626", bg: "#fef2f2", triggered: true,  notSet: false, approver: "Approving Authority" }
    : pcrOverrun > 0.10
    ? { label: "PCR Required",  color: "#d97706", bg: "#fffbeb", triggered: true,  notSet: false, approver: "Project Board" }
    : pcrOverrun >= 0
    ? { label: "Within Budget", color: "#16a34a", bg: "#f0fdf4", triggered: false, notSet: false, approver: "" }
    : { label: "Under Budget",  color: "#16a34a", bg: "#f0fdf4", triggered: false, notSet: false, approver: "" };

  const COST_CAT_MAP: Record<number, string> = {
    686490000: "Staff and Other Personnel Costs",
    686490001: "Supplies, Commodities, and Materials",
    686490002: "Equipment, Vehicles, and Furniture",
    686490003: "Contractual Services",
    686490004: "Travel",
    686490005: "Indirect Costs",
  };
  const FUNDING_MAP: Record<number, string> = {
    0: "Regular Budget", 1: "PK + Support Account", 2: "xB", 3: "10RCR", 4: "20PCR",
  };

  const byCategoryPlanned: Record<string, number> = {};
  const byCategoryActual:  Record<string, number> = {};
  leaves.forEach(n => {
    const label = n.costCategory != null ? (COST_CAT_MAP[n.costCategory] ?? "Other") : "Unassigned";
    byCategoryPlanned[label] = (byCategoryPlanned[label] ?? 0) + (n.totalPlannedCost ?? 0);
    byCategoryActual[label]  = (byCategoryActual[label]  ?? 0) + (n.totalActualCost  ?? 0);
  });

  const byFundingPlanned: Record<string, number> = {};
  const byFundingActual:  Record<string, number> = {};
  leaves.forEach(n => {
    const label = n.fundingSource != null ? (FUNDING_MAP[n.fundingSource] ?? "Other") : "Unassigned";
    byFundingPlanned[label] = (byFundingPlanned[label] ?? 0) + (n.totalPlannedCost ?? 0);
    byFundingActual[label]  = (byFundingActual[label]  ?? 0) + (n.totalActualCost  ?? 0);
  });

  const catRows  = Object.keys({ ...byCategoryPlanned, ...byCategoryActual }).map(label => ({
    label,
    planned: byCategoryPlanned[label] ?? 0,
    actual:  byCategoryActual[label]  ?? 0,
  }));

  const fundRows = Object.keys({ ...byFundingPlanned, ...byFundingActual }).map(label => ({
    label,
    planned: byFundingPlanned[label] ?? 0,
    actual:  byFundingActual[label]  ?? 0,
  }));

  const COST_CATS_FULL = [
    { value: 686490000, label: "Staff and Other Personnel Costs" },
    { value: 686490001, label: "Supplies, Commodities, and Materials" },
    { value: 686490002, label: "Equipment, Vehicles, and Furniture" },
    { value: 686490003, label: "Contractual Services" },
    { value: 686490004, label: "Travel" },
    { value: 686490005, label: "Indirect Costs" },
  ];

  type PidRow = { category: string; qty: number; unitCost: number; totalCost: number; funding: string; remarks: string };
  const pidRows: PidRow[] = [];
  COST_CATS_FULL.forEach(cat => {
    const catLeaves = leaves.filter(n =>
  n.costCategory === cat.value &&
  ((n.totalPlannedCost ?? 0) > 0 || (n.fixedCost ?? 0) > 0 || (n.actualFixedCost ?? 0) > 0)
);
if (catLeaves.length === 0) return;
    const fundingGroups: Record<string, TaskNode[]> = {};
    catLeaves.forEach(n => {
      const f = n.fundingSource != null ? (FUNDING_MAP[n.fundingSource] ?? "Other") : "Unassigned";
      if (!fundingGroups[f]) fundingGroups[f] = [];
      fundingGroups[f].push(n);
    });
    Object.entries(fundingGroups).forEach(([funding, tasks]) => {
      // Sub-group by unit rate so tasks with different rates stay separate
      const rateGroups: Record<string, TaskNode[]> = {};
      tasks.forEach(n => {
        const rateKey = String(n.unitRate ?? 0);
        if (!rateGroups[rateKey]) rateGroups[rateKey] = [];
        rateGroups[rateKey].push(n);
      });

      Object.entries(rateGroups).forEach(([rateKey, rateTasks]) => {
        const unitRate = Number(rateKey);
        const qty      = rateTasks.reduce((s, n) => s + (n.quantity ?? 0), 0);
        const fixedSum = rateTasks.reduce((s, n) => s + (n.fixedCost ?? 0), 0);
        const total    = rateTasks.reduce((s, n) => s + (n.totalPlannedCost ?? 0), 0);
        // If no effort hours but has fixed cost, show qty=1, unitCost=total
        const displayQty      = qty === 0 && fixedSum > 0 ? 1     : qty;
        const displayUnitCost = qty === 0 && fixedSum > 0 ? total  : unitRate;
        const remarks  = rateTasks.map(n => n.taskName).filter(Boolean).join(", ");
        pidRows.push({ category: cat.label, qty: displayQty, unitCost: displayUnitCost, totalCost: total, funding, remarks });
      });
    });
  });
  // After the COST_CATS_FULL.forEach block, add:
  const unassignedLeaves = leaves.filter(n =>
  n.costCategory == null &&
  ((n.totalPlannedCost ?? 0) > 0 || (n.fixedCost ?? 0) > 0 || (n.actualFixedCost ?? 0) > 0)
);
  if (unassignedLeaves.length > 0) {
    const fundingGroups: Record<string, TaskNode[]> = {};
    unassignedLeaves.forEach(n => {
      const f = n.fundingSource != null ? (FUNDING_MAP[n.fundingSource] ?? "Other") : "Unassigned";
      if (!fundingGroups[f]) fundingGroups[f] = [];
      fundingGroups[f].push(n);
    });
    Object.entries(fundingGroups).forEach(([funding, tasks]) => {
      const rateGroups: Record<string, TaskNode[]> = {};
      tasks.forEach(n => {
        const rateKey = String(n.unitRate ?? 0);
        if (!rateGroups[rateKey]) rateGroups[rateKey] = [];
        rateGroups[rateKey].push(n);
      });
      Object.entries(rateGroups).forEach(([rateKey, rateTasks]) => {
        const unitRate   = Number(rateKey);
        const qty        = rateTasks.reduce((s, n) => s + (n.quantity ?? 0), 0);
        const fixedSum   = rateTasks.reduce((s, n) => s + (n.fixedCost ?? 0), 0);
        const total      = rateTasks.reduce((s, n) => s + (n.totalPlannedCost ?? 0), 0);
        const displayQty      = qty === 0 && fixedSum > 0 ? 1     : qty;
        const displayUnitCost = qty === 0 && fixedSum > 0 ? total : unitRate;
        const remarks    = rateTasks.map(n => n.taskName).filter(Boolean).join(", ");
        pidRows.push({ category: "Unassigned", qty: displayQty, unitCost: displayUnitCost, totalCost: total, funding, remarks });
      });
    });
  }
  const pidTotal = pidRows.reduce((s, r) => s + r.totalCost, 0);

function GroupedHBarChart({ title, rows, totalPlanned, totalActual }: {
  title: string;
  rows: { label: string; planned: number; actual: number }[];
  totalPlanned: number;
  totalActual: number;
}) {
  const maxVal = Math.max(...rows.map(b => Math.max(b.planned, b.actual)), 1);
  return (
    <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: 8, padding: "18px 20px", flex: 1 }}>
      <div style={{ fontSize: 16, fontWeight: 700, color: "#111827", marginBottom: 16 }}>{title}</div>
      <div style={{ display: "flex", gap: 16, marginBottom: 14 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <div style={{ width: 12, height: 12, borderRadius: 2, background: "#4763a5" }}/>
          <span style={{ fontSize: 12, color: "#374151" }}>Planned</span>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
          <div style={{ width: 12, height: 12, borderRadius: 2, background: "#c07a2f" }}/>
          <span style={{ fontSize: 12, color: "#374151" }}>Actual</span>
        </div>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
        {rows.filter(b => b.planned > 0 || b.actual > 0).map((b, i) => (
          <div key={i}>
            <div style={{ fontSize: 13, fontWeight: 600, color: "#111827", marginBottom: 5 }}>{b.label}</div>
            <div style={{ marginBottom: 4 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ fontSize: 13, color: "#374151", width: 48, flexShrink: 0, fontWeight: 500 }}>Planned</span>
                <div style={{ flex: 1, height: 16, background: "#f3f4f6", borderRadius: 3, overflow: "hidden" }}>
                  <div style={{ height: "100%", width: `${(b.planned / maxVal) * 100}%`, background: "#4763a5", borderRadius: 3, transition: "width 0.4s ease" }}/>
                </div>
                <span style={{ fontSize: 13, fontWeight: 700, color: "#4763a5", width: 78, textAlign: "right", flexShrink: 0 }}>{fmtCurrency(b.planned)}</span>
              </div>
            </div>
            <div>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ fontSize: 13, color: "#374151", width: 48, flexShrink: 0, fontWeight: 500 }}>Actual</span>
                <div style={{ flex: 1, height: 16, background: "#f3f4f6", borderRadius: 3, overflow: "hidden" }}>
                  <div style={{ height: "100%", width: `${(b.actual / maxVal) * 100}%`, background: "#c07a2f", borderRadius: 3, transition: "width 0.4s ease" }}/>
                </div>
                <span style={{ fontSize: 13, fontWeight: 700, color: "#c07a2f", width: 78, textAlign: "right", flexShrink: 0 }}>{fmtCurrency(b.actual)}</span>
              </div>
            </div>
          </div>
        ))}
      </div>
      <div style={{ marginTop: 16, paddingTop: 12, borderTop: "1px solid #f3f4f6" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4 }}>
          <span style={{ fontSize: 13, fontWeight: 700, color: "#374151" }}>Total Planned</span>
          <span style={{ fontSize: 14, fontWeight: 800, color: "#4763a5" }}>{fmtCurrency(totalPlanned)}</span>
        </div>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span style={{ fontSize: 13, fontWeight: 700, color: "#374151" }}>Total Actual</span>
          <span style={{ fontSize: 14, fontWeight: 800, color: "#c07a2f" }}>{fmtCurrency(totalActual)}</span>
        </div>
      </div>
    </div>
  );
}

  // Power BI-style donut with callout lines
  function DonutWithLabels({ slices, title }: {
    slices: { label: string; value: number; color: string }[];
    title: string;
  }) {
    const total = slices.reduce((s, d) => s + d.value, 0);
    if (total === 0) return <p style={{ color: "#9ca3af", fontSize: 13 }}>No data</p>;

    const W = 450, H = 340;
    const cx = W / 2, cy = H / 2 - 10;
    const r = 105, ir = r * 0.52;
    let cumAngle = -Math.PI / 2;

    const segAngles = slices.filter(s => s.value > 0).map(s => {
      const a = (s.value / total) * 2 * Math.PI;
      const mid = cumAngle + a / 2;
      cumAngle += a;
      return { ...s, angle: a, mid };
    });

    return (
      <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: 8, padding: "16px", flex: 1 }}>
        <div style={{ fontSize: 16, fontWeight: 700, color: "#111827", marginBottom: 8 }}>{title}</div>
        <svg width="100%" viewBox={`0 0 ${W} ${H}`} style={{ overflow: "visible" }}>
          {/* Donut slices */}
          {(() => {
            let ca = -Math.PI / 2;
            return slices.filter(s => s.value > 0).map((s, i) => {
              const a = (s.value / total) * 2 * Math.PI;
              const x1 = cx + r  * Math.cos(ca),       y1 = cy + r  * Math.sin(ca);
              const x2 = cx + r  * Math.cos(ca + a),   y2 = cy + r  * Math.sin(ca + a);
              const ix1= cx + ir * Math.cos(ca),        iy1= cy + ir * Math.sin(ca);
              const ix2= cx + ir * Math.cos(ca + a),    iy2= cy + ir * Math.sin(ca + a);
              const large = a > Math.PI ? 1 : 0;
              const d = `M ${ix1} ${iy1} L ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2} L ${ix2} ${iy2} A ${ir} ${ir} 0 ${large} 0 ${ix1} ${iy1} Z`;
              ca += a;
              return <path key={i} d={d} fill={s.color} stroke="white" strokeWidth="2"/>;
            });
          })()}

          {/* Callout lines + labels */}
          {segAngles.filter(s => (s.value / total) > 0.03).map((s, i) => {
            const pct = ((s.value / total) * 100).toFixed(0);
            const lineR  = r + 12;
            const labelR = r + 28;
            const lx = cx + lineR * Math.cos(s.mid);
            const ly = cy + lineR * Math.sin(s.mid);
            const tx = cx + labelR * Math.cos(s.mid);
            const ty = cy + labelR * Math.sin(s.mid);
            const isRight = Math.cos(s.mid) >= 0;
            const endX = tx + (isRight ? 22 : -22);
            return (
              <g key={i}>
                <line x1={cx + r * Math.cos(s.mid)} y1={cy + r * Math.sin(s.mid)}
                  x2={lx} y2={ly} stroke={s.color} strokeWidth="1.5"/>
                <line x1={lx} y1={ly} x2={endX} y2={ty}
                  stroke={s.color} strokeWidth="1.5"/>
                <text x={isRight ? endX + 4 : endX - 4} y={ty - 6}
                  fontSize="11" fontWeight="700" fill="#111827"
                  textAnchor={isRight ? "start" : "end"}
                  fontFamily="Segoe UI, sans-serif">
                  {pct}%
                </text>
                <text x={isRight ? endX + 4 : endX - 4} y={ty + 7}
                  fontSize="10" fill="#374151"
                  textAnchor={isRight ? "start" : "end"}
                  fontFamily="Segoe UI, sans-serif">
                  {s.label}
                </text>
                <text x={isRight ? endX + 4 : endX - 4} y={ty + 19}
                  fontSize="10" fill="#6b7280"
                  textAnchor={isRight ? "start" : "end"}
                  fontFamily="Segoe UI, sans-serif">
                  {fmtCurrency(s.value)}
                </text>
              </g>
            );
          })}

          {/* Centre total */}
          <text x={cx} y={cy - 6} fontSize="12" fill="#6b7280" textAnchor="middle"
            fontFamily="Segoe UI, sans-serif">Total</text>
          <text x={cx} y={cy + 12} fontSize="13" fontWeight="700" fill="#111827"
            textAnchor="middle" fontFamily="Segoe UI, sans-serif">
            {fmtCurrency(total)}
          </text>
        </svg>
      </div>
    );
  }

  const [pidCopied, setPidCopied] = React.useState(false);
  function copyPidTable() {
    const header = "Category\tQty\tUnit Cost\tTotal Cost\tFunding Source\tRemarks";
    const rows = pidRows.map(r =>
      `${r.category}\t${r.qty.toFixed(0)}\t${fmtCurrency(r.unitCost)}\t${fmtCurrency(r.totalCost)}\t${r.funding}\t${r.remarks}`
    );
    const total = `Total\t\t\t${fmtCurrency(pidTotal)}\t\t`;
    navigator.clipboard.writeText([header, ...rows, total].join("\n")).then(() => {
      setPidCopied(true);
      setTimeout(() => setPidCopied(false), 2000);
    });
  }

  // Gauge — bigger, standalone
  function BigGauge() {
  const r = 110, cx = 145, cy = 145;
  const angle = Math.PI + (pctConsumed / 100) * Math.PI;
  const x1 = cx - r, y1 = cy;
  const x2 = cx + r, y2 = cy;
  const px = cx + r * Math.cos(angle);
  const py = cy + r * Math.sin(angle);

  return (
    <svg width="290" height="175" viewBox="0 0 290 175">
      {/* Gray background */}
      <path d={`M ${x1} ${y1} A ${r} ${r} 0 1 1 ${x2} ${y2}`}
        fill="none" stroke="#e5e7eb" strokeWidth="16" strokeLinecap="round"/>
      {/* Green fill */}
      {pctConsumed > 0 && (
        <path d={`M ${x1} ${y1} A ${r} ${r} 0 0 1 ${px} ${py}`}
          fill="none" stroke={gaugeColor} strokeWidth="16" strokeLinecap="round"/>
      )}
      {/* Needle */}
      <line x1={cx} y1={cy}
        x2={cx + (r - 12) * Math.cos(angle)}
        y2={cy + (r - 12) * Math.sin(angle)}
        stroke="#1f2937" strokeWidth="3.5" strokeLinecap="round"/>
      <circle cx={cx} cy={cy} r="7" fill="#1f2937"/>
      <text x="8" y="172" fontSize="13" fill="#6b7280" fontFamily="Segoe UI, sans-serif">0%</text>
      <text x="244" y="172" fontSize="13" fill="#6b7280" fontFamily="Segoe UI, sans-serif">100%</text>
      <text x={cx} y={cy - 50} fontSize="13" fill="#374151"
        textAnchor="middle" fontFamily="Segoe UI, sans-serif" fontWeight="700">SPENT</text>
      <text x={cx} y={cy - 14} fontSize="36" fontWeight="800" fill={gaugeColor}
        textAnchor="middle" fontFamily="Segoe UI, sans-serif">
        {Math.min(pctConsumed, 999).toFixed(1)}%
      </text>
    </svg>
  );
}
  return (
    <div className="tg-summary-panel">
      <div className="tg-summary-header">
        <span className="tg-summary-title">📊 Budget Summary</span>
        <button className="tg-summary-close" onClick={onClose}>✕</button>
      </div>

      {/* ── Row 1: KPI cards — full width, side by side ─────────── */}
      <div style={{ display: "flex", gap: 14, padding: "16px 18px 0" }}>

        {/* KPI 1 — Budget Remaining Status */}
        <div style={{ background: kpi2Status.bg, border: `2px solid ${kpi2Status.color}`, borderRadius: 10, padding: "18px 20px", flex: 1 }}>
          <div style={{ fontSize: 15, color: "#111827", fontWeight: 800, letterSpacing: "0.01em", display: "flex", alignItems: "center" }}>
            Budget Remaining Status
            <InfoPopover
              title="Budget Remaining Status"
              explanation="How much budget is still available to spend."
              formula="Available = Total Planned − Total Spent"
              thresholds="🟢 > 15% = High Balance Available · 🟡 5–15% = Low Balance Available · 🔴 < 5% = Critical Balance"
            />
          </div>
          <div style={{ marginTop: 10, display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 22, height: 22, borderRadius: "50%", background: kpi2Status.color, flexShrink: 0 }}/>
            <span style={{ fontSize: 20, fontWeight: 800, color: kpi2Status.color }}>{kpi2Status.label}</span>
          </div>
          <div style={{ fontSize: 14, color: "#111827", marginTop: 10, fontWeight: 600 }}>
            {availablePct.toFixed(1)}% remaining
          </div>
          <div style={{ fontSize: 13, color: "#6b7280", marginTop: 5 }}>
            {">"}15% sufficient · 5–15% low · {"<"}5% overrun
          </div>
        </div>

        {/* KPI 2 — Consumption Rate */}
        <div style={{ background: kpi1Status?.bg ?? "#f9fafb", border: `2px solid ${kpi1Status?.color ?? "#e5e7eb"}`, borderRadius: 10, padding: "18px 20px", flex: 1 }}>
          <div style={{ fontSize: 15, color: "#111827", fontWeight: 800, letterSpacing: "0.01em", display: "flex", alignItems: "center" }}>
            Consumption Rate
            <InfoPopover
              title="Consumption Rate"
              explanation="Are you spending at the right pace for the work completed? 1.0 = perfect balance."
              formula="Rate = (Spent ÷ Budget) ÷ (% Complete ÷ 100)"
              thresholds="🟢 ≤ 1.10 = On Track · 🟡 ≤ 1.30 = At Risk · 🔴 > 1.30 = High Risk"
            />
          </div>
          <div style={{ marginTop: 10, display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 22, height: 22, borderRadius: "50%", background: kpi1Status?.color ?? "#9ca3af", flexShrink: 0 }}/>
            <span style={{ fontSize: 20, fontWeight: 800, color: kpi1Status?.color ?? "#374151" }}>
              {kpi1Status?.label ?? "N/A"}
            </span>
          </div>
          <div style={{ fontSize: 14, color: "#111827", marginTop: 10, fontWeight: 600 }}>
            Progress: {weightedPct.toFixed(1)}% complete
          </div>
          <div style={{ fontSize: 13, color: "#6b7280", marginTop: 5 }}>
            Rate: {consumptionRate !== null ? consumptionRate.toFixed(2) : "—"} (≤1.10 on track)
          </div>
        </div>

        {/* KPI 3 — Needs PCR */}
        <div style={{ background: pcrStatus.bg, border: `2px solid ${pcrStatus.color}`, borderRadius: 10, padding: "18px 20px", flex: 1 }}>
          <div style={{ fontSize: 15, color: "#111827", fontWeight: 800, display: "flex", alignItems: "center" }}>
            Needs PCR?
            <InfoPopover
              title="Needs PCR?"
              explanation="A PCR is needed when planned cost exceeds the approved budget by 10% or more."
              formula="Overrun % = (Planned − Approved) ÷ Approved × 100"
              thresholds="🟢 0–10% = Within Budget · 🟡 >10–25% = PCR → Project Board · 🔴 >25% = PCR → Approving Authority"
            />
          </div>
          <div style={{ marginTop: 8, display: "flex", alignItems: "center", gap: 10 }}>
            <div style={{ width: 22, height: 22, borderRadius: "50%", background: pcrStatus.color, flexShrink: 0 }}/>
            <div>
              <span style={{ fontSize: 20, fontWeight: 800, color: pcrStatus.color }}>{pcrStatus.label}</span>
              {pcrStatus.approver && (
                <div style={{ fontSize: 12, fontWeight: 700, color: pcrStatus.color, marginTop: 2 }}>
                  → Submit to {pcrStatus.approver}
                </div>
              )}
            </div>
          </div>
          {!pcrStatus.notSet && (
            <div style={{ marginTop: 10, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6 }}>
              <div>
                <div style={{ fontSize: 10, color: "#6b7280", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em" }}>Approved Budget</div>
                <div style={{ fontSize: 13, fontWeight: 700, color: "#111827", marginTop: 2 }}>{fmtCurrency(latestApprovedBudget)}</div>
              </div>
              <div>
                <div style={{ fontSize: 10, color: "#6b7280", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em" }}>Total Planned</div>
                <div style={{ fontSize: 13, fontWeight: 700, color: "#111827", marginTop: 2 }}>{fmtCurrency(totalPlanned)}</div>
              </div>
              <div style={{ gridColumn: "1 / -1" }}>
                <div style={{ fontSize: 10, color: "#6b7280", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em" }}>Variance</div>
                <div style={{ fontSize: 13, fontWeight: 700, color: pcrStatus.color, marginTop: 2 }}>
                  {fmtCurrency(totalPlanned - latestApprovedBudget)} ({pcrOverrun !== null ? (pcrOverrun * 100).toFixed(1) : "0.0"}%)
                </div>
              </div>
            </div>
          )}
          {pcrStatus.notSet && (
            <div style={{ fontSize: 12, color: "#9ca3af", marginTop: 8 }}>
              No approved budget set yet. Once the project budget is approved and recorded, this KPI will show whether a Project Change Request is needed.
            </div>
          )}
        </div>
      </div>

      {/* ── Row 2: Gauge + Value Cards + EVM Metrics ─────────────── */}
      {(() => {
        const cpi = totalActual > 0 ? totalEV / totalActual : null;
        const eac = cpi && cpi > 0 ? totalPlanned / cpi : null;
        const etc = totalPlanned - totalEV;
        const cpiColor = cpi === null ? "#6b7280"
          : cpi >= 0.95 ? "#16a34a"
          : cpi >= 0.80 ? "#d97706"
          : "#dc2626";
        const eacColor = eac === null ? "#6b7280"
          : eac <= totalPlanned ? "#16a34a"
          : eac <= totalPlanned * 1.10 ? "#d97706"
          : "#dc2626";
        const etcColor = etc <= totalRemaining ? "#16a34a"
          : etc <= totalRemaining * 1.10 ? "#d97706"
          : "#dc2626";
        return (
          <div style={{ display: "flex", gap: 14, padding: "14px 18px 0", alignItems: "stretch" }}>

            {/* Gauge */}
            <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", flex: "0 0 300px", background: "white", border: "1px solid #e5e7eb", borderRadius: 10, padding: "16px 8px 12px" }}>
              <div style={{ fontSize: 14, fontWeight: 700, color: "#111827", marginBottom: 4 }}>Budget Consumption (% Spent)</div>
              <BigGauge />
            </div>

            {/* Value cards + EVM metrics stacked */}
            <div style={{ display: "flex", flex: 1, gap: 14 }}>

              {/* Total Budget */}
            <div style={{ display: "flex", flexDirection: "column", gap: 10, flex: 1 }}>
                <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: 10, padding: "16px 20px", flex: 1, minHeight: 100, display: "flex", flexDirection: "column", justifyContent: "space-between" }}>
                  <div style={{ fontSize: 11, color: "#374151", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em" }}>Total Budget</div>
                  <div style={{ fontSize: 26, fontWeight: 800, color: "#111827", marginTop: 6 }}>{fmtCurrency(totalPlanned)}</div>
                  <div style={{ fontSize: 12, color: "#6b7280", marginTop: 4 }}>{weightedPct.toFixed(1)}% complete</div>
                </div>
                <div style={{ background: "white", border: `1px solid ${eacColor}`, borderRadius: 10, padding: "16px 20px", flex: 1, minHeight: 100, display: "flex", flexDirection: "column", justifyContent: "space-between" }}>
                  <div style={{ fontSize: 11, color: "#6b7280", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em", display: "flex", alignItems: "center" }}>
                    EAC — Forecast Final Cost
                    <InfoPopover
                      title="EAC — Estimate at Completion"
                      explanation="Forecasted total cost if current spending efficiency continues."
                      formula="EAC = Total Budget ÷ CPI"
                      thresholds="🟢 ≤ Budget = On track · 🟡 ≤ +10% = Slight overrun · 🔴 > +10% = Overrun"
                    />
                  </div>
                  <div style={{ fontSize: 26, fontWeight: 800, color: eacColor, marginTop: 6 }}>{eac !== null ? fmtCurrency(eac) : "—"}</div>
                  <div style={{ fontSize: 12, color: "#6b7280", marginTop: 4 }}>
                    {eac !== null ? `${eac <= totalPlanned ? "Under" : "Over"} budget by ${fmtCurrency(Math.abs(eac - totalPlanned))}` : "No spend yet"}
                  </div>
                </div>
              </div>

              {/* Spent */}
              <div style={{ display: "flex", flexDirection: "column", gap: 10, flex: 1 }}>
                <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: 10, padding: "16px 20px", flex: 1, minHeight: 100, display: "flex", flexDirection: "column", justifyContent: "space-between" }}>
                  <div style={{ fontSize: 11, color: "#374151", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em" }}>Spent</div>
                  <div style={{ fontSize: 26, fontWeight: 800, color: "#0f766e", marginTop: 6 }}>{fmtCurrency(totalActual)}</div>
                  <div style={{ fontSize: 12, color: "#6b7280", marginTop: 4 }}>EV: {fmtCurrency(totalEV)}</div>
                </div>
                <div style={{ background: "white", border: `1px solid ${cpiColor}`, borderRadius: 10, padding: "16px 20px", flex: 1, minHeight: 100, display: "flex", flexDirection: "column", justifyContent: "space-between" }}>
                  <div style={{ fontSize: 11, color: "#6b7280", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em", display: "flex", alignItems: "center" }}>
                    CPI — Cost Performance
                    <InfoPopover
                      title="CPI — Cost Performance Index"
                      explanation="For every $1 spent, how much work value was delivered? Unlike Consumption Rate (which compares spending pace to progress), CPI uses Earned Value — the monetary worth of work completed. CPI < 1.0 means you are paying more than the work is worth."
                      formula="CPI = Earned Value (EV) ÷ Actual Cost"
                      thresholds="🟢 ≥ 0.95 = On budget · 🟡 0.80–0.95 = At risk · 🔴 < 0.80 = Overrun"
                    />
                  </div>
                  <div style={{ fontSize: 26, fontWeight: 800, color: cpiColor, marginTop: 6 }}>{cpi !== null ? cpi.toFixed(2) : "—"}</div>
                  <div style={{ fontSize: 12, color: "#6b7280", marginTop: 4 }}>
                    {cpi === null ? "No spend yet" : cpi >= 0.95 ? "On budget" : cpi >= 0.80 ? "Slight overrun" : "Significant overrun"}
                    {cpi !== null ? ` · $1 spent = ${fmtCurrency(cpi)} value` : ""}
                  </div>
                </div>
              </div>

              {/* Available */}
              <div style={{ display: "flex", flexDirection: "column", gap: 10, flex: 1 }}>
                <div style={{ background: "white", border: `2px solid ${gaugeColor}`, borderRadius: 10, padding: "16px 20px", flex: 1, minHeight: 100, display: "flex", flexDirection: "column", justifyContent: "space-between" }}>
                  <div style={{ fontSize: 11, color: "#374151", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em" }}>Available</div>
                  <div style={{ fontSize: 26, fontWeight: 800, color: gaugeColor, marginTop: 6 }}>{fmtCurrency(totalRemaining)}</div>
                  <div style={{ fontSize: 12, color: "#6b7280", marginTop: 4 }}>{availablePct.toFixed(1)}% of budget</div>
                </div>
                <div style={{ background: "white", border: `1px solid ${etcColor}`, borderRadius: 10, padding: "16px 20px", flex: 1, minHeight: 100, display: "flex", flexDirection: "column", justifyContent: "space-between" }}>
                  <div style={{ fontSize: 11, color: "#6b7280", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em", display: "flex", alignItems: "center" }}>
                    ETC — Cost to Finish
                    <InfoPopover
                      title="ETC — Estimate to Complete"
                      explanation="Money still needed to finish remaining work. If ETC > Available, you have a shortfall."
                      formula="ETC = Total Budget − Earned Value (EV)"
                      thresholds="🟢 ETC ≤ Available = Covered · 🔴 ETC > Available = Shortfall"
                    />
                  </div>
                  <div style={{ fontSize: 26, fontWeight: 800, color: etcColor, marginTop: 6 }}>{fmtCurrency(etc)}</div>
                  <div style={{ fontSize: 12, color: "#6b7280", marginTop: 4 }}>
                    {etc <= totalRemaining ? "Covered by available budget" : `Shortfall of ${fmtCurrency(etc - totalRemaining)}`}
                  </div>
                </div>
              </div>

            </div>
          </div>
        );
      })()}

      {/* ── Row 2: Horizontal Bar Charts ────────────────────────── */}
      <div style={{ display: "flex", gap: 14, padding: "0 18px 16px" }}>
        <GroupedHBarChart
          title="Planned vs Actual by Category"
          rows={catRows}
          totalPlanned={totalPlanned}
          totalActual={totalActual}
        />
        <GroupedHBarChart
          title="Planned vs Actual by Funding Source"
          rows={fundRows}
          totalPlanned={totalPlanned}
          totalActual={totalActual}
        />
      </div>

      {/* ── Row 3: Budget Table ──────────────────────────────────── */}
      <div style={{ padding: "0 18px 24px" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
          <div style={{ fontSize: 15, fontWeight: 700, color: "#111827" }}>Budget Table</div>
          <button className="tg-btn" style={{ fontSize: 12, height: 28, padding: "0 12px", border: "1px solid #e5e7eb", color: pidCopied ? "#16a34a" : "#374151" }}
            onClick={copyPidTable}>
            {pidCopied ? "✓ Copied!" : "📋 Copy"}
          </button>
        </div>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: "#1f5c8b" }}>
              {["Category", "Qty", "Unit Cost", "Total Cost", "Funding", "Remarks"].map(h => (
                <th key={h} style={{ padding: "8px 10px", color: "white", fontWeight: 700, fontSize: 13, textAlign: h === "Category" || h === "Funding" || h === "Remarks" ? "left" : "right", borderRight: "1px solid #2d6fa0", whiteSpace: "nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {pidRows.length === 0 ? (
              <tr><td colSpan={5} style={{ padding: "14px", color: "#6b7280", textAlign: "center", fontSize: 13 }}>No budget data — assign Cost Categories and Funding Sources to tasks</td></tr>
            ) : (
              pidRows.map((r, i) => (
                <tr key={i} style={{ background: i % 2 === 0 ? "white" : "#f9fafb" }}>
                  <td style={{ padding: "8px 10px", borderBottom: "1px solid #f3f4f6", color: "#111827", fontSize: 13 }}>{r.category}</td>
                  <td style={{ padding: "8px 10px", textAlign: "right", borderBottom: "1px solid #f3f4f6", color: "#111827", fontSize: 13 }}>{r.qty.toFixed(0)}</td>
                  <td style={{ padding: "8px 10px", textAlign: "right", borderBottom: "1px solid #f3f4f6", color: "#111827", fontSize: 13 }}>{fmtCurrency(r.unitCost)}</td>
                  <td style={{ padding: "8px 10px", textAlign: "right", borderBottom: "1px solid #f3f4f6", color: "#111827", fontSize: 13 }}>{fmtCurrency(r.totalCost)}</td>
                  <td style={{ padding: "8px 10px", borderBottom: "1px solid #f3f4f6", color: "#374151", fontSize: 13 }}>{r.funding}</td>
                  <td style={{ padding: "8px 10px", borderBottom: "1px solid #f3f4f6", color: "#6b7280", fontSize: 12, maxWidth: 350, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.remarks}</td>
                </tr>
              ))
            )}
            <tr style={{ background: "white" }}>
              <td colSpan={3} style={{ padding: "9px 10px", borderTop: "2px solid #e5e7eb", fontWeight: 700, fontSize: 13, color: "#111827" }}>Total</td>
              <td style={{ padding: "9px 10px", textAlign: "right", borderTop: "2px solid #e5e7eb", color: "#4f46e5", fontSize: 14, fontWeight: 400 }}>{fmtCurrency(pidTotal)}</td>
              <td style={{ borderTop: "2px solid #e5e7eb" }}/>
              <td style={{ borderTop: "2px solid #e5e7eb" }}/>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ColTooltip({ label, tip }: { label: string; tip: string }) {
  const [pos, setPos] = React.useState<{ top: number; left: number } | null>(null);
  const ref = React.useRef<HTMLSpanElement>(null);

  function handleMouseEnter() {
    if (ref.current) {
      const rect = ref.current.getBoundingClientRect();
      setPos({ top: rect.bottom + 6, left: rect.left });
    }
  }

  function handleMouseLeave() {
    setPos(null);
  }

  return (
    <React.Fragment>
      <span ref={ref} onMouseEnter={handleMouseEnter} onMouseLeave={handleMouseLeave}
        style={{ display: "inline-flex", alignItems: "center", cursor: "default" }}>
        {label}
      </span>
      {pos && ReactDOM.createPortal(
        <div style={{
          position: "fixed",
          top: pos.top,
          left: pos.left,
          background: "#ffffcc",
          color: "#1f2937",
          fontSize: 12,
          fontWeight: 400,
          lineHeight: 1.5,
          padding: "4px 8px",
          border: "1px solid #999",
          borderRadius: 2,
          width: 220,
          whiteSpace: "normal",
          pointerEvents: "none",
          zIndex: 99999,
          boxShadow: "1px 1px 3px rgba(0,0,0,0.2)",
        }}>
          {tip}
        </div>,
        document.body
      )}
    </React.Fragment>
  );
}

function RefreshIcon() {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none"
      stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <polyline points="23 4 23 10 17 10"/>
      <polyline points="1 20 1 14 7 14"/>
      <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/>
    </svg>
  );
}

const CSS = `
.tg-wrap {
    font-family: 'Segoe UI', system-ui, sans-serif;
    font-size: 13px; display: flex; flex-direction: column;
    height: 100%; background: #fff; color: #1f2937;
    overflow: hidden; margin: 0 15px 0 0;
    border: 1px solid #e5e7eb; border-radius: 4px;
  }
.tg-wrap input, .tg-wrap select {
    font-family: 'Segoe UI', system-ui, sans-serif !important;
    font-size: 13px !important;
  }
  .tg-input {
    font-size: 14px !important;
  }
  .tg-select { cursor: cell !important; }
  .tg-select:focus { cursor: default !important; }
  .tg-input  { cursor: cell !important; }
  .tg-input:focus { cursor: text !important; }
  .tg-input-wrap { cursor: cell !important; }
  .tg-input-wrap:focus-within { cursor: text !important; }
  .tg-toolbar {
    background: #fff; border-bottom: 2px solid #d0d0d0;
    padding: 0 12px; display: flex; align-items: center;
    gap: 2px; flex-shrink: 0; height: 54px;
    position: sticky; top: 0; z-index: 10;
  }
  .tg-title {
    font-weight: 600; font-size: 13px; color: #1f2937;
    display: flex; align-items: center; gap: 6px; margin-right: 8px;
  }
  .tg-divider { width: 1px; height: 16px; background: #e5e7eb; margin: 0 6px; }
  .tg-btn {
    display: inline-flex; align-items: center; gap: 4px;
    padding: 4px 14px; border-radius: 4px; border: none;
    font-size: 14px; cursor: pointer; font-weight: 400;
    background: transparent; color: #323130; transition: background 0.1s;
    height: 42px; white-space: nowrap;
  }
  .tg-btn:hover         { background: #f3f2f1; }
  .tg-btn-save          { background: #0078d4; color: white; font-size: 14px; }
  .tg-btn-save:hover    { background: #106ebe; }
  .tg-btn-save:disabled { background: #c7e0f4; cursor: not-allowed; color: white; }
  .tg-btn-discard       { color: #d13438; }
  .tg-btn-discard:hover { background: #fde7e9; }
  .tg-badge {
    background: #0078d4; color: white; border-radius: 10px;
    padding: 1px 6px; font-size: 10px; font-weight: 700;
  }
  .tg-saved  { color: #16a34a; font-size: 12px; font-weight: 500; }
  .tg-error  { color: #dc2626; font-size: 11px; padding: 4px 12px; background: #fef2f2; }
  .tg-scroll { overflow: auto; flex: 1; }
  .tg-table  { border-collapse: collapse; min-width: 100%; }
.tg-table th {
    background: #ffffff; color: #323130; font-weight: 400;
    font-size: 12px; text-transform: none !important;
    padding: 10px 12px; border-bottom: 1px solid #e5e7eb;
    border-right: 1px solid #f3f4f6; text-align: left;
    position: sticky; top: 0; z-index: 3; white-space: nowrap; user-select: none;
    overflow: hidden;
  }
  .tg-table th[data-pinned="true"] {
    z-index: 5;
  }
  .tg-resize-handle {
    position: absolute;
    right: 0; top: 0; bottom: 0; width: 5px;
    cursor: col-resize; user-select: none; touch-action: none;
    background: transparent; z-index: 3;
  }
  .tg-resize-handle:hover { background: #0078d4; opacity: 0.5; }
  .tg-resize-handle.is-resizing { background: #0078d4; opacity: 1; }
  .tg-table th.th-right { text-align: right; }
  .tg-table th:last-child { border-right: none; }
  .tg-expand-btn {
    background: none; border: none; cursor: pointer; padding: 0; margin-right: 4px;
    color: #605e5c; display: inline-flex; align-items: center; justify-content: center;
    width: 16px; height: 16px; border-radius: 2px; flex-shrink: 0; transition: all 0.1s;
  }
  .tg-expand-btn:hover { background: #edebe9; color: #323130; }
  .tg-cell-text { font-size: 14px; border-radius: 3px; padding: 1px 3px; }
  .tg-cell-right { text-align: right; display: block; width: 100%; }
  .tg-dash { color: #d1d5db; display: block; text-align: center; }
  .tg-progress-wrap { display: flex; align-items: center; gap: 8px; }
  .tg-progress-track {
    flex: 1; height: 11px; background: #edebe9; border-radius: 2px; overflow: hidden;
  }
  .tg-progress-fill  { height: 100%; border-radius: 0px; transition: width 0.3s; }
  .tg-progress-label { font-size: 14px; color: #605e5c; width: 42px; text-align: right; flex-shrink: 0; }
.tg-select {
    font-size: 13px; border: 1px solid transparent;
    padding: 4px 4px; background: transparent; width: 100%;
    color: #1f2937; cursor: pointer; max-width: 220px;
    height: 100%; appearance: auto; border-radius: 0;
  }
  .tg-select:hover { border: 1px solid #107c10; background: white; }
  .tg-select:focus { outline: none; border: 2px solid #107c10; background: white; }
.tg-input-wrap {
    display: flex; align-items: center; justify-content: flex-end;
    border: 1px solid transparent; background: transparent;
    width: 100%; margin-left: auto; height: 100%;
    box-sizing: border-box;
  }
.tg-input-wrap:hover { background: white; border: 1px solid #107c10; cursor: text; }
  .tg-input-wrap:focus-within {
    border: 2px solid #107c10; background: white;
  }
  .tg-input { cursor: text; }
  .tg-input-symbol {
    padding: 2px 5px; background: transparent; color: #605e5c;
    font-size: 13px !important; flex-shrink: 0;
  }
  .tg-input {
    font-size: 13px !important; border: none; padding: 2px 6px;
    background: transparent; width: 100%; text-align: right; color: #1f2937;
    outline: none; height: 100%; box-sizing: border-box;
    -webkit-appearance: none; -moz-appearance: textfield;
    font-family: 'Segoe UI', system-ui, sans-serif;
  }
  .tg-input::-webkit-outer-spin-button,
  .tg-input::-webkit-inner-spin-button { -webkit-appearance: none; margin: 0; }
  .tg-footer {
    padding: 4px 12px; background: #fff; border-top: 1px solid #e5e7eb;
    font-size: 11px; color: #a19f9d; display: flex;
    justify-content: space-between; flex-shrink: 0;
    position: sticky; bottom: 0; z-index: 10;
  }
    .tg-col-panel {
    position: absolute;
    background: white;
    border: 1px solid #e5e7eb;
    border-radius: 4px;
    box-shadow: 0 4px 16px rgba(0,0,0,0.12);
    z-index: 99999;
    min-width: 220px;
    padding: 8px 0;
    display: flex;
    flex-direction: column;
  }
  .tg-col-panel-header {
    padding: 6px 12px 4px;
    font-size: 11px;
    font-weight: 600;
    color: #605e5c;
    text-transform: uppercase;
    letter-spacing: 0.05em;
  }
  .tg-col-item {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 5px 12px;
    cursor: pointer;
    font-size: 13px;
    color: #1f2937;
    user-select: none;
  }
  .tg-col-item:hover { background: #f3f2f1; }
  .tg-col-item input[type="checkbox"] { 
    width: 14px; height: 14px; cursor: pointer; accent-color: #0078d4;
  }
  .tg-col-divider {
    height: 1px; background: #f3f4f6; margin: 4px 0;
  }
  .tg-col-footer {
    padding: 6px 12px 2px;
    display: flex;
    gap: 6px;
  }
  .tg-col-footer button {
    font-size: 12px; padding: 2px 8px; border-radius: 3px;
    border: 1px solid #e5e7eb; cursor: pointer; background: white;
    color: #323130;
  }
  .tg-col-footer button:hover { background: #f3f2f1; }
  .tg-table.is-resizing { cursor: col-resize; user-select: none; }
  .tg-table.is-resizing td { pointer-events: none; }
  .tg-table th[data-pinned="true"],
  .tg-table td[data-pinned="true"] {
    border-right: 2px solid #d1d5db !important;
  }
  .tg-table td:not([data-pinned="true"]) {
    position: relative;
    z-index: 0;
  }
  .tg-table th:not([data-pinned="true"]) {
    position: relative;
    z-index: 2;
  }
  .tg-th-dragging   { opacity: 0.4; cursor: grabbing; }
  .tg-th-drop-left  { border-left: 2px solid #0078d4 !important; }
.tg-th-drop-right { border-right: 2px solid #0078d4 !important; }
  .tg-table th { cursor: grab; }
  .tg-table th.th-nodrag { cursor: default; }
  .tg-col-tip { position: relative; display: inline-flex; align-items: center; }
  .tg-col-tip::after {
    content: attr(data-tip);
    position: absolute;
    background: #ffffcc;
    color: #1f2937;
    font-size: 12px;
    font-weight: 400;
    line-height: 1.5;
    padding: 4px 6px;
    border: 1px solid #999;
    border-radius: 2px;
    width: 200px;
    white-space: normal;
    pointer-events: none;
    opacity: 0;
    transition: opacity 0.15s ease;
    z-index: 99999;
    box-shadow: 1px 1px 3px rgba(0,0,0,0.2);
    top: 100%;
    left: 0;
    margin-top: 4px;
  }
  .tg-col-tip:hover::after { opacity: 1; transition-delay: 0.8s; }
  .tg-filter-panel {
    position: absolute; top: 0; right: 0; bottom: 0;
    width: 340px; background: white; z-index: 100;
    border-left: 1px solid #e5e7eb;
    box-shadow: -4px 0 16px rgba(0,0,0,0.10);
    display: flex; flex-direction: column;
    animation: slideIn 0.2s ease;
  }
  @keyframes slideIn {
    from { transform: translateX(300px); }
    to   { transform: translateX(0); }
  }
  .tg-filter-header {
    display: flex; align-items: center; justify-content: space-between;
    padding: 14px 16px; border-bottom: 1px solid #e5e7eb; flex-shrink: 0;
  }
  .tg-filter-title { font-size: 16px; font-weight: 600; color: #1f2937; }
  .tg-filter-clearall {
    font-size: 13px; color: #107c10; cursor: pointer; font-weight: 500;
    background: none; border: none; padding: 0;
  }
  .tg-filter-clearall:hover { text-decoration: underline; }
  .tg-filter-close {
    background: none; border: none; cursor: pointer; color: #605e5c;
    font-size: 18px; line-height: 1; padding: 0 0 0 8px;
  }
  .tg-filter-close:hover { color: #1f2937; }
  .tg-filter-search {
    margin: 12px 16px; border: 2px solid #107c10; border-radius: 2px;
    padding: 6px 10px; font-size: 13px; width: calc(100% - 32px);
    box-sizing: border-box; outline: none;
    font-family: 'Segoe UI', system-ui, sans-serif;
  }
  .tg-filter-body { overflow-y: auto; flex: 1; padding-bottom: 16px; }
  .tg-filter-section { border-bottom: 1px solid #f3f4f6; }
  .tg-filter-section-header {
    display: flex; align-items: center; justify-content: space-between;
    padding: 10px 16px; cursor: pointer; user-select: none;
    font-size: 14px; font-weight: 600; color: #1f2937;
  }
  .tg-filter-section-header:hover { background: #f9fafb; }
  .tg-filter-item {
    display: flex; align-items: center; gap: 8px;
    padding: 6px 16px 6px 24px; cursor: pointer;
    font-size: 14px; color: #374151; overflow: hidden;
  }
  .tg-filter-item span:first-of-type {
    overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex: 1;
  }
  .tg-filter-item:hover { background: #f3f2f1; }
  .tg-filter-item input[type="checkbox"] {
    width: 14px; height: 14px; cursor: pointer; accent-color: #107c10;
  }
  .tg-filter-count { margin-left: auto; color: #374151; font-size: 14px; font-weight: 600; min-width: 28px; text-align: right; }
  .tg-filter-badge {
    background: #107c10; color: white; border-radius: 10px;
    padding: 1px 6px; font-size: 10px; font-weight: 700; margin-left: 4px;
  }
    .tg-checkbox-cell {
    width: 32px; padding: 0 8px; text-align: center;
    border-right: 1px solid #f9fafb;
  }
  .tg-checkbox-cell input[type="checkbox"] {
    width: 14px; height: 14px; cursor: pointer; accent-color: #107c10;
  }
  .tg-row-selected td { background: #f0fdf4 !important; }
  .tg-selection-badge {
    font-size: 12px; color: #107c10; font-weight: 500;
    display: flex; align-items: center; gap: 6px;
  }
  .tg-avatars { display: flex; flex-direction: column; gap: 3px; padding: 2px 0; }
  .tg-avatar-row { display: flex; align-items: center; gap: 6px; }
  .tg-avatar {
    width: 22px; height: 22px; border-radius: 50%;
    background: #0078d4; color: white;
    font-size: 9px; font-weight: 700;
    display: flex; align-items: center; justify-content: center;
    flex-shrink: 0; letter-spacing: 0.02em;
  }
  .tg-avatar-name { font-size: 12px; color: #374151; white-space: nowrap; 
    overflow: hidden; text-overflow: ellipsis; max-width: 140px; }
  .tg-summary-panel {
    position: absolute; top: 0; left: 0; bottom: 0; right: 0;
    background: #f9fafb; z-index: 100;
    border-right: 1px solid #e5e7eb;
    box-shadow: 4px 0 16px rgba(0,0,0,0.10);
    display: flex; flex-direction: column;
    animation: slideInLeft 0.2s ease;
    overflow-y: auto;
  }
  @keyframes slideInLeft {
    from { transform: translateX(-100%); }
    to   { transform: translateX(0); }
  }
  .tg-summary-header {
    display: flex; align-items: center; justify-content: space-between;
    padding: 14px 16px; border-bottom: 1px solid #e5e7eb;
    flex-shrink: 0; background: white;
  }
  .tg-summary-title { font-size: 16px; font-weight: 600; color: #1f2937; }
  .tg-summary-close {
    background: none; border: none; cursor: pointer; color: #605e5c;
    font-size: 18px; line-height: 1; padding: 0;
  }
  .tg-summary-close:hover { color: #1f2937; }
  .tg-kpi-grid {
    display: grid; grid-template-columns: 1fr 1fr;
    gap: 10px; padding: 14px 16px;
  }
  .tg-kpi-card {
    background: white; border: 1px solid #e5e7eb; border-radius: 6px;
    padding: 12px 14px; display: flex; flex-direction: column; gap: 4px;
  }
  .tg-kpi-label { font-size: 11px; color: #6b7280; font-weight: 500; text-transform: uppercase; letter-spacing: 0.04em; }
  .tg-kpi-value { font-size: 20px; font-weight: 700; color: #111827; }
  .tg-kpi-sub   { font-size: 11px; color: #9ca3af; }
  .tg-chart-section { padding: 0 16px 16px; }
  .tg-chart-title { font-size: 13px; font-weight: 600; color: #374151; margin-bottom: 8px; margin-top: 14px; }
`;

function ChevronRight() {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none"
      stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
      <polyline points="9 18 15 12 9 6" />
    </svg>
  );
}
function ChevronDown() {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none"
      stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
      <polyline points="6 9 12 15 18 9" />
    </svg>
  );
}
function GridIcon() {
  return (
    <svg style={{ width: 16, height: 16, color: "#6366f1" }} viewBox="0 0 24 24" fill="none"
      stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/>
      <rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/>
    </svg>
  );
}

function ColumnsIcon() {
  return (
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none"
      stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <rect x="3" y="3" width="5" height="18"/><rect x="10" y="3" width="5" height="18"/>
      <rect x="17" y="3" width="5" height="18"/>
    </svg>
  );
}

function ProgressCell({ value }: { value: number }) {
  const pct   = Math.min(100, Math.max(0, Number(value) || 0));
  const color = "#496945";
  return (
    <div className="tg-progress-wrap">
      <div className="tg-progress-track">
        <div className="tg-progress-fill" style={{ width: `${pct}%`, background: color }} />
      </div>
      <span className="tg-progress-label">{pct.toFixed(0)}%</span>
    </div>
  );
}

const AVATAR_COLORS = [
  "#0F4C81","#1B6B3A","#6B3FA0","#B45309",
  "#0E7490","#7C3D12","#1E3A5F","#3D6B35",
];

function getInitials(name: string): string {
  const parts = name.trim().split(/\s+/);
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

function getColor(name: string): string {
  let hash = 0;
  for (let i = 0; i < name.length; i++) hash = name.charCodeAt(i) + ((hash << 5) - hash);
  return AVATAR_COLORS[Math.abs(hash) % AVATAR_COLORS.length];
}

function ResourceCell({ resources }: { resources: ResourceItem[] }) {
  if (!resources || resources.length === 0) {
    return <span className="tg-dash">—</span>;
  }

 // Single assignee — show avatar + full name
  if (resources.length === 1) {
    const r = resources[0];
    return (
      <div className="tg-avatar-row">
        <div className="tg-avatar" style={{ background: getColor(r.name), width: 30, height: 30, fontSize: 11 }}>
          {getInitials(r.name)}
        </div>
        <span className="tg-avatar-name">{r.name}</span>
      </div>
    );
  }

  // Multiple assignees — stacked avatars with count
  const stackWidth = 30 + (resources.length - 1) * 20;
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
      <div style={{ position: "relative", width: stackWidth, height: 30, flexShrink: 0 }}>
        {resources.map((r, i) => (
          <div
            key={r.id}
            title={r.name}
            className="tg-avatar"
            style={{
              background: getColor(r.name),
              width: 30, height: 30, fontSize: 11,
              position: "absolute",
              top: 0, left: i * 20,
              border: "2px solid white",
              zIndex: resources.length - i,
              cursor: "default",
            }}
          >
            {getInitials(r.name)}
          </div>
        ))}
      </div>
      <span style={{ fontSize: 11, color: "#6b7280" }}>
        {resources.length} people
      </span>
    </div>
  );
}

// Currency input with £ symbol prefix
function CurrencyInput({ value, onChange, disabled }: {
  value: number; onChange: (v: number) => void; disabled?: boolean;
}) {
  const [local, setLocal] = React.useState(value === 0 ? "" : String(value));
  const [focused, setFocused] = React.useState(false);
  React.useEffect(() => { setLocal(value === 0 ? "" : String(value)); }, [value]);

  function commit() {
    const n = parseFloat(local);
    const resolved = isNaN(n) ? 0 : n;
    if (resolved !== value) {
      onChange(resolved);
    }
    setFocused(false);
  }

  const displayValue = focused
    ? local
    : value === 0 ? "" : new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value);

  return (
    <div className="tg-input-wrap">
      <span className="tg-input-symbol">$</span>
      <input
        className="tg-input"
        type={focused ? "number" : "text"}
        min={0} step="0.01"
        placeholder="0.00"
        value={displayValue}
        disabled={disabled}
        onChange={e => setLocal(e.target.value)}
        onFocus={() => setFocused(true)}
        onBlur={commit}
        onKeyDown={e => { if (e.key === "Enter") commit(); }}
      />
    </div>
  );
}

function NumberInput({ value, onChange }: {
  value: number; onChange: (v: number) => void;
}) {
  const [local, setLocal] = React.useState(value === 0 ? "" : String(value));
  React.useEffect(() => { setLocal(value === 0 ? "" : String(value)); }, [value]);

  function commit() {
    const n = parseFloat(local);
    onChange(isNaN(n) ? 0 : n);
  }

  return (
    <div className="tg-input-wrap" style={{ width: 80 }}>
      <input
        className="tg-input"
        type="number" min={0} step="0.01"
        placeholder="0"
        value={local}
        onChange={e => setLocal(e.target.value)}
        onBlur={commit}
        onKeyDown={e => { if (e.key === "Enter") commit(); }}
      />
    </div>
  );
}

// Dash for summary rows in input columns
function SummaryValue({ value, currency = false }: { value: number; currency?: boolean }) {
  if (!value && value !== 0) return <span className="tg-dash">—</span>;
  return (
    <span className="tg-cell-right" style={{ fontWeight: 600 }}>
      {currency ? fmtCurrency(value) : fmtNumber(value)}
    </span>
  );
}

function ServiceCombobox({ value, items, onChange }: {
  value: string | null;
  items: SrcItem[];
  onChange: (id: string, item: SrcItem) => void;
}) {
  const [open, setOpen]       = React.useState(false);
  const [query, setQuery]     = React.useState("");
  const [hovered, setHovered] = React.useState(false);
  const triggerRef            = React.useRef<HTMLDivElement>(null);
  const [dropPos, setDropPos] = React.useState({ top: 0, left: 0, width: 0 });

  function openDropdown() {
    if (triggerRef.current) {
      const rect = triggerRef.current.getBoundingClientRect();
      setDropPos({
        top:   rect.bottom + window.scrollY,
        left:  rect.left   + window.scrollX,
        width: rect.width,
      });
    }
    setOpen(true);
    setQuery("");
  }

  // Close on outside click
  React.useEffect(() => {
    if (!open) return;
    function handler(e: MouseEvent) {
      const target = e.target as Node;
      const portal = document.getElementById("tg-service-portal");
      if (
        triggerRef.current && !triggerRef.current.contains(target) &&
        portal && !portal.contains(target)
      ) {
        setOpen(false);
        setQuery("");
      }
    }
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [open]);

  const selected = items.find(s => s.id === value);
  const filtered = query.trim() === ""
    ? items
    : items.filter(s => {
        const q = query.toLowerCase();
        return (
          s.serviceId.toLowerCase().includes(q)              ||
          s.name.toLowerCase().includes(q)                   ||
          (s.fiscalYearName ?? "").toLowerCase().includes(q) ||
          (s.entityName     ?? "").toLowerCase().includes(q)
        );
      });

  const dropdown = open ? (
    <div
      id="tg-service-portal"
      style={{
        position: "absolute",
        top:      dropPos.top,
        left:     dropPos.left,
        width:    Math.max(dropPos.width, 300),
        background: "white",
        border: "1px solid #e5e7eb",
        borderRadius: 4,
        boxShadow: "0 4px 16px rgba(0,0,0,0.14)",
        zIndex: 99999,
        maxHeight: 320,
        display: "flex",
        flexDirection: "column",
      }}
    >
      <div style={{ padding: "6px 8px", borderBottom: "1px solid #f3f4f6" }}>
        <input
          autoFocus
          placeholder="Search by service, FY or entity..."
          value={query}
          onChange={e => setQuery(e.target.value)}
          style={{
            width: "100%", border: "1px solid #d1d5db", borderRadius: 2,
            padding: "6px 8px", fontSize: 14, outline: "none",
            boxSizing: "border-box",
          }}
          onClick={e => e.stopPropagation()}
        />
      </div>
      <div style={{ overflowY: "auto", maxHeight: 260 }}>
        <div
          onClick={() => { onChange("", {} as SrcItem); setOpen(false); setQuery(""); }}
          style={{
            padding: "5px 12px", fontSize: 12, cursor: "pointer",
            color: "#9ca3af", borderBottom: "1px solid #f9fafb",
          }}
          onMouseEnter={e => (e.currentTarget.style.background = "#f9fafb")}
          onMouseLeave={e => (e.currentTarget.style.background = "white")}
        >
          — clear selection —
        </div>
        {filtered.length === 0 ? (
          <div style={{ padding: "8px 12px", color: "#9ca3af", fontSize: 12 }}>
            No services found
          </div>
        ) : (
          filtered.map(s => (
            <div
              key={s.id}
              onClick={() => { onChange(s.id, s); setOpen(false); setQuery(""); }}
              style={{
                padding: "8px 12px", fontSize: 14, cursor: "pointer",
                background: s.id === value ? "#eff6ff" : "white",
                color: s.id === value ? "#0078d4" : "#1f2937",
                borderLeft: s.id === value ? "3px solid #0078d4" : "3px solid transparent",
              }}
              onMouseEnter={e => {
                if (s.id !== value) (e.currentTarget as HTMLDivElement).style.background = "#f9fafb";
              }}
              onMouseLeave={e => {
                (e.currentTarget as HTMLDivElement).style.background =
                  s.id === value ? "#eff6ff" : "white";
              }}
            >
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontWeight: 500, fontSize: 12 }}>{s.serviceId}</span>
                <span style={{ fontSize: 13, color: "#0078d4" }}>
                  {[s.fiscalYearName, s.entityName].filter(Boolean).join(" · ")}
                </span>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 1 }}>
                <span style={{ fontSize: 13, color: "#6b7280" }}>{s.name}</span>
                <span style={{ fontSize: 13, color: "#374151", fontWeight: 500 }}>
                  {fmtCurrency(s.price)}/{s.unit}
                </span>
              </div>
            </div>
          ))
        )}
      </div>
    </div>
  ) : null;

  return (
    <React.Fragment>
      <div
        ref={triggerRef}
        onClick={openDropdown}
        onMouseDown={e => e.stopPropagation()}
        onMouseEnter={() => setHovered(true)}
        onMouseLeave={() => setHovered(false)}
        style={{
          display: "flex", alignItems: "center", justifyContent: "space-between",
          border: open ? "2px solid #107c10" : hovered ? "1px solid #107c10" : "1px solid transparent",
          borderRadius: 0, padding: "2px 6px",
          background: open || hovered ? "white" : "transparent", cursor: "cell", fontSize: 14,
          color: selected ? "#1f2937" : "#9ca3af",
          minHeight: 38, height: 38,
          userSelect: "none", width: "100%", maxWidth: 260,
          boxSizing: "border-box",
        }}
      >
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1 }}>
          {selected
            ? `${selected.serviceId} – ${selected.name}${selected.fiscalYearName ? ` (${selected.fiscalYearName})` : ""}`
            : "— select service —"}
        </span>
<svg width="17" height="17" viewBox="0 0 24 24" fill="none"
          stroke="#323130" strokeWidth="2.5" style={{ flexShrink: 0, marginLeft: 4 }}>
          <polyline points="6 9 12 15 18 9"/>
        </svg>
      </div>
      {dropdown && typeof document !== "undefined"
        ? ReactDOM.createPortal(dropdown, document.body)
        : null}
    </React.Fragment>
  );
}

const TaskDetailPanel = React.memo(function TaskDetailPanel({
  task,
  resources,
  srcItems,
  onSave,
  onClose,
}: {
  task: TaskNode;
  resources: ResourceItem[];
  srcItems: SrcItem[];
  onSave: (recordId: string, updates: Partial<TaskNode>) => Promise<void>;
  onClose: () => void;
}) {
  const [local, setLocal] = React.useState<Partial<TaskNode>>({});
  const [saving, setSaving] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);
  const [saved, setSaved] = React.useState(false);

  const [localUnitRate, setLocalUnitRate] = React.useState<string>(task.unitRate ? String(task.unitRate) : "");
  const [localFixedCost, setLocalFixedCost] = React.useState(task.fixedCost ? String(task.fixedCost) : "");
  const [localActualFixed, setLocalActualFixed] = React.useState(task.actualFixedCost ? String(task.actualFixedCost) : "");
  const [savedUnitRate, setSavedUnitRate] = React.useState<string | null>(null);
  const [savedFixedCost, setSavedFixedCost] = React.useState<string | null>(null);
  const [savedActualFixed, setSavedActualFixed] = React.useState<string | null>(null);
  const unitRateRef = React.useRef<HTMLInputElement>(null);
  const fixedCostRef = React.useRef<HTMLInputElement>(null);
  const actualFixedRef = React.useRef<HTMLInputElement>(null);

  React.useEffect(() => {
  setLocalUnitRate(task.unitRate ? String(task.unitRate) : "");
  setLocalFixedCost(task.fixedCost ? String(task.fixedCost) : "");
  setLocalActualFixed(task.actualFixedCost ? String(task.actualFixedCost) : "");
  setLocal({});
}, [task.recordId]);

  // Merge task with local edits
  const current = { ...task, ...local };

  function recalcLocal(overrides: Partial<TaskNode>): Partial<TaskNode> {
    const qty         = overrides.quantity        ?? current.quantity        ?? 0;
    const rate        = overrides.unitRate        ?? current.unitRate        ?? 0;
    const fixed       = overrides.fixedCost       ?? current.fixedCost       ?? 0;
    const completed   = overrides.effortCompleted ?? current.effortCompleted ?? 0;
    const actualFixed = overrides.actualFixedCost ?? current.actualFixedCost ?? 0;
    const pct         = overrides.pctDone         ?? current.pctDone         ?? 0;
    const planned     = qty * rate;
    const actual      = completed * rate;
    const total       = planned + fixed;
    const totalActual = actual + actualFixed;
    const remaining   = total - totalActual;
    const ev          = (pct / 100) * total;
    return {
      ...overrides,
      plannedCost:      planned,
      actualCost:       actual,
      totalPlannedCost: total,
      totalActualCost:  totalActual,
      remainingCost:    remaining,
      earnedValue:      ev,
    };
  }

  function commitUnitRate() {}
  function commitFixedCost() {}
  function commitActualFixed() {}
  
  function onSrcChange(srcId: string) {
    const src = srcItems.find(s => s.id === srcId);
    if (!src) return;
    setLocalUnitRate(String(src.price));
    setLocal(prev => ({
      ...prev,
      ...recalcLocal({
        ...prev,
        srcServiceId:   src.id,
        srcServiceName: src.name,
        unitRate:       src.price,
        unit:           src.unit,
        frequency:      src.frequency,
      }),
    }));
  }

  async function handleSave() {
    const unitRateVal = parseFloat(unitRateRef.current?.value ?? "") || 0;
    const fixedVal = parseFloat(fixedCostRef.current?.value ?? "") || 0;
    const actualFixedVal = parseFloat(actualFixedRef.current?.value ?? "") || 0;
    const finalLocal = { ...local, ...recalcLocal({ ...local, unitRate: unitRateVal, fixedCost: fixedVal, actualFixedCost: actualFixedVal }) };

    setSaving(true);
    setError(null);
    try {
      // Capture the new values before re-render resets them
      setSavedUnitRate(unitRateRef.current?.value ?? null);
      setSavedFixedCost(fixedCostRef.current?.value ?? null);
      setSavedActualFixed(actualFixedRef.current?.value ?? null);
      await onSave(task.recordId, finalLocal);
      setSaved(true);
      setTimeout(() => onClose(), 1000);
    } catch (e: any) {
      setError(e.message ?? "Save failed");
    } finally {
      setSaving(false);
    }
  }

  // Label + value layout helper
  function Field({ label, children }: { label: string; children: React.ReactNode }) {
    return (
      <div style={{ marginBottom: 14 }}>
        <div style={{
          fontSize: 14, fontWeight: 600, color: "#1f2937",
          letterSpacing: "0em", marginBottom: 6,
        }}>{label}</div>
        {children}
      </div>
    );
  }

  function ReadOnlyField({ label, value }: { label: string; value: string }) {
    return (
      <Field label={label}>
        <div style={{
          fontSize: 14, color: "#111827", padding: "7px 10px",
          border: "1px solid #8a8886", borderRadius: 4,
        }}>{value || "—"}</div>
      </Field>
    );
  }

  function MetricCard({ label, value, color }: { label: string; value: string; color?: string }) {
    return (
      <div style={{
        background: "white", border: "1px solid #e5e7eb", borderRadius: 8,
        padding: "10px 14px", flex: 1,
      }}>
        <div style={{ fontSize: 13, fontWeight: 600, color: "#6b7280", letterSpacing: "0em" }}>{label}</div>
        <div style={{ fontSize: 15, fontWeight: 700, color: color ?? "#111827", marginTop: 4 }}>{value}</div>
      </div>
    );
  }

  const inputStyle: React.CSSProperties = {
    width: "100%", border: "1px solid #8a8886", borderRadius: 4,
    padding: "8px 10px", fontSize: 14, color: "#111827",
    boxSizing: "border-box", outline: "none",
    fontFamily: "Segoe UI, system-ui, sans-serif",
  };

  const selectStyle: React.CSSProperties = {
    ...inputStyle, border: "1px solid #8a8886", background: "white", cursor: "pointer", appearance: "auto" as any,
  };

  return (
    <React.Fragment>
      {/* Backdrop — dims and locks the grid */}
      <div style={{
        position: "absolute", inset: 0, background: "rgba(0,0,0,0.15)",
        zIndex: 50,
      }} />

      {/* Panel */}
      <div style={{
        position: "absolute", top: 0, right: 0, bottom: 0,
        width: 440, background: "white", zIndex: 51,
        borderLeft: "1px solid #e5e7eb",
        boxShadow: "-4px 0 20px rgba(0,0,0,0.12)",
        display: "flex", flexDirection: "column",
        animation: "slideIn 0.2s ease",
      }}>

        {/* Header */}
        <div style={{
          padding: "14px 16px", borderBottom: "2px solid #d0d0d0",
          display: "flex", alignItems: "flex-start", justifyContent: "space-between",
          flexShrink: 0, background: "white",
        }}>
          <div>
            <div style={{ fontSize: 14, fontWeight: 400, color: "#6b7280", letterSpacing: "0.05em", marginBottom: 4 }}>Task Details</div>
            <div style={{ fontSize: 23, fontWeight: 400, color: "#111827", lineHeight: 1.3 }}>{task.taskName}</div>
          </div>
          <button onClick={onClose} style={{
            background: "none", border: "none", cursor: "pointer",
            color: "#605e5c", fontSize: 18, lineHeight: 1, padding: "0 0 0 8px", flexShrink: 0,
          }}>✕</button>
        </div>

        {/* Scrollable body */}
        <div style={{ flex: 1, overflowY: "auto", padding: "16px" }}>

          {/* Read-only info */}
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 16 }}>
            <ReadOnlyField label="Start"       value={formatDate(task.startDate)} />
            <ReadOnlyField label="Finish"      value={formatDate(task.endDate)} />
            <ReadOnlyField label="% Complete"  value={`${(task.pctDone ?? 0).toFixed(1)}%`} />
            <div style={{ marginBottom: 14 }}>
              <div style={{ fontSize: 14, fontWeight: 600, color: "#1f2937", marginBottom: 6 }}>Assigned to</div>
              <div style={{ padding: "4px 0" }}>
                <ResourceCell resources={resources} />
              </div>
            </div>
          </div>

          <div style={{ height: 1, background: "#f3f4f6", margin: "0 0 16px" }} />

          {/* Editable fields */}
          <Field label="Funding Source">
            <select style={selectStyle}
              value={current.fundingSource ?? ""}
              onChange={e => { const val = e.target.value === "" ? null : Number(e.target.value); setLocal(prev => ({ ...prev, fundingSource: val })); }}>
              <option value="">— select —</option>
              {FUNDING_SOURCES.map(f => (
                <option key={f.value} value={f.value}>{f.label}</option>
              ))}
            </select>
          </Field>

          <Field label="Cost Category">
            <select style={selectStyle}
              value={current.costCategory ?? ""}
              onChange={e => { const val = e.target.value === "" ? null : Number(e.target.value); setLocal(prev => ({ ...prev, costCategory: val })); }}>
              <option value="">— select —</option>
              {COST_CATEGORIES.map(c => (
                <option key={c.value} value={c.value}>{c.label}</option>
              ))}
            </select>
          </Field>

          <Field label="Service">
            <div style={{ width: "100%" }}>
              <ServiceCombobox
                value={current.srcServiceId ?? null}
                items={srcItems}
                onChange={(id) => onSrcChange(id)}
              />
            </div>
          </Field>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10 }}>
            <Field label="Effort (h)">
              <input style={{ ...inputStyle, background: "#f9fafb", color: "#6b7280" }}
                value={current.quantity ?? ""} readOnly />
            </Field>
            <Field label="Unit">
              <input style={{ ...inputStyle, background: "#f9fafb", color: "#6b7280" }}
                value={current.unit ?? ""} readOnly />
            </Field>
            <Field label="Unit Rate">
              <input ref={unitRateRef} style={inputStyle} 
                type="text" inputMode="decimal"
                defaultValue={savedUnitRate ?? localUnitRate}
                key={savedUnitRate ?? localUnitRate}
                onBlur={commitUnitRate}
                onKeyDown={e => { if (e.key === "Enter") { e.preventDefault(); commitUnitRate(); } }}
              />
            </Field>
          </div>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            <Field label="Fixed Cost ($)">
             <input ref={fixedCostRef} style={inputStyle}
                type="text" inputMode="decimal"
                defaultValue={savedFixedCost ?? localFixedCost}
                key={savedFixedCost ?? localFixedCost}
                onBlur={commitFixedCost}
                onKeyDown={e => { if (e.key === "Enter") { e.preventDefault(); commitFixedCost(); } }}
              />
            </Field>
            <Field label="Actual Fixed Cost ($)">
              <input ref={actualFixedRef} style={inputStyle}
                type="text" inputMode="decimal"
                defaultValue={savedActualFixed ?? localActualFixed}
                key={savedActualFixed ?? localActualFixed}
                onBlur={commitActualFixed}
                onKeyDown={e => { if (e.key === "Enter") { e.preventDefault(); commitActualFixed(); } }}
              />
            </Field>
          </div>

          <div style={{ height: 1, background: "#f3f4f6", margin: "4px 0 16px" }} />

          {/* Computed metrics */}
          <div style={{ fontSize: 14, fontWeight: 600, color: "#1f2937", letterSpacing: "0em", marginBottom: 10 }}>Computed</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
            <MetricCard label="Planned Cost"     value={fmtCurrency(current.totalPlannedCost ?? 0)} color="#4f46e5" />
            <MetricCard label="Actual Cost"      value={fmtCurrency(current.totalActualCost  ?? 0)} color="#0f766e" />
            <MetricCard label="Remaining"        value={fmtCurrency(current.remainingCost    ?? 0)} color={(current.remainingCost ?? 0) < 0 ? "#dc2626" : "#374151"} />
            <MetricCard label="Earned Value"     value={fmtCurrency(current.earnedValue      ?? 0)} color="#374151" />
          </div>

          {error && (
            <div style={{ background: "#fef2f2", border: "1px solid #fca5a5", borderRadius: 6, padding: "10px 12px", fontSize: 13, color: "#dc2626", marginTop: 12 }}>
              ⚠ {error}
            </div>
          )}
        </div>

        {/* Footer */}
        <div style={{
          padding: "12px 16px", borderTop: "1px solid #e5e7eb",
          display: "flex", gap: 8, flexShrink: 0, background: "white",
        }}>
          {saved ? (
            <div style={{ flex: 1, textAlign: "center", color: "#16a34a", fontWeight: 600, fontSize: 14 }}>
              ✓ Saved successfully
            </div>
          ) : (
            <React.Fragment>
              <button onClick={onClose} style={{
                flex: 1, padding: "8px 0", border: "1px solid #e5e7eb",
                borderRadius: 4, background: "white", cursor: "pointer",
                fontSize: 14, color: "#374151", fontWeight: 500,
              }}>Cancel</button>
              <button onClick={handleSave} disabled={saving} style={{
                flex: 2, padding: "8px 0", border: "none",
                borderRadius: 4, background: saving ? "#c7e0f4" : "#0078d4",
                cursor: saving ? "not-allowed" : "pointer",
                fontSize: 14, color: "white", fontWeight: 600,
              }}>{saving ? "Saving…" : "Save"}</button>
            </React.Fragment>
          )}
        </div>
      </div>
    </React.Fragment>
  );
});

export function TaskGrid({ data: initialData, onSave, onRefresh, userId, taskIds, latestApprovedBudget, projectId }: Props) {
  const [data, setData]             = React.useState<TaskNode[]>(initialData);
  //const [expanded, setExpanded]     = React.useState<ExpandedState>({ "0": true });
  const [allExpanded, setAllExpanded] = React.useState(false);
  const expandKey = `tg-expand-${userId}`;
  const [expanded, setExpanded] = React.useState<ExpandedState>(() => {
    try {
      const stored = localStorage.getItem(`tg-expand-${userId}`);
      if (stored) return JSON.parse(stored);
    } catch {}
    return {};
  });
  const [pending, setPending]       = React.useState<Record<string, Partial<TaskNode>>>({});
  const [saving, setSaving]         = React.useState(false);
  const [savedMsg, setSavedMsg]     = React.useState(false);
  const [hoveredRow, setHoveredRow] = React.useState<string | null>(null);
  const [srcItems, setSrcItems]     = React.useState<SrcItem[]>([]);
  const [entityItems, setEntityItems] = React.useState<EntityItem[]>([]);
  const [fiscalYearItems, setFiscalYearItems] = React.useState<FiscalYearItem[]>([]);
  const [taskResources, setTaskResources] = React.useState<TaskResourceMap>({});
  const [loadError, setLoadError]   = React.useState<string | null>(null);
  
// Column visibility
  const storageKey = `tg-columns-${userId}`;
  const [visibleCols, setVisibleCols] = React.useState<Set<string>>(() => {
    try {
      const stored = localStorage.getItem(storageKey);
      if (stored) {
        const arr = JSON.parse(stored) as string[];
        return new Set(arr);
      }
    } catch {}
    return new Set(DEFAULT_VISIBLE);
  });
  const orderKey = `tg-order-${userId}`;
  const [columnOrder, setColumnOrder] = React.useState<string[]>(() => {
    try {
      const stored = localStorage.getItem(orderKey);
      if (stored) return JSON.parse(stored);
    } catch {}
    return DEFAULT_ORDER;
  });
  const [dragColId,   setDragColId]   = React.useState<string | null>(null);
  const [dropColId,   setDropColId]   = React.useState<string | null>(null);
  const [colPanelOpen, setColPanelOpen]   = React.useState(false);
  const sizingKey = `tg-sizing-${userId}`;
  const [columnSizing, setColumnSizing] = React.useState<Record<string, number>>(() => {
    try {
      const stored = localStorage.getItem(sizingKey);
      if (stored) return JSON.parse(stored);
    } catch {}
    return {};
  });
  const [detailTask, setDetailTask] = React.useState<TaskNode | null>(null);
  const frozenDetailTask = React.useRef<TaskNode | null>(null);
  if (detailTask && frozenDetailTask.current?.recordId !== detailTask.recordId) {
    frozenDetailTask.current = detailTask;
  }

  // Persist sizing changes
  React.useEffect(() => {
    try { localStorage.setItem(sizingKey, JSON.stringify(columnSizing)); } catch {}
  }, [columnSizing]);
  const [colPanelPos,  setColPanelPos]    = React.useState({ top: 0, left: 0 });
  const colBtnRef = React.useRef<HTMLButtonElement>(null);

    const [refreshCount, setRefreshCount] = React.useState(0);

    const [filterOpen, setFilterOpen]         = React.useState(false);
    const [filterKeyword, setFilterKeyword]   = React.useState("");
    const [filterFunding, setFilterFunding]   = React.useState<Set<number>>(new Set());
    const [filterCategory, setFilterCategory] = React.useState<Set<number>>(new Set());
    const [filterAssignee, setFilterAssignee] = React.useState<Set<string>>(new Set());
    const [filterService, setFilterService] = React.useState<Set<string>>(new Set());
    const [selectedRows, setSelectedRows] = React.useState<string[]>([]);
    const [lastSelectedId, setLastSelectedId]   = React.useState<string | null>(null);
    const [summaryOpen, setSummaryOpen] = React.useState(false);
    const [timelineOpen, setTimelineOpen] = React.useState(false);
    const [sectionOpen, setSectionOpen] = React.useState<Record<string, boolean>>({
      funding: true, category: true, service: true, assignee: true,
    });

    React.useEffect(() => {
        if (Object.keys(pending).length === 0) {
        setData(initialData);
        setRefreshCount(c => c + 1);
        }
    }, [initialData]);

// Expand first TWO levels on initial load only if no stored state
  React.useEffect(() => {
    if (initialData.length === 0) return;
    try {
      const stored = localStorage.getItem(expandKey);
      if (stored) return; // user has a saved state, don't override
    } catch {}
    const twoLevels: Record<string, boolean> = {};
    initialData.forEach((rootNode, i) => {
      twoLevels[String(i)] = true;
      rootNode.subRows?.forEach((_, j) => {
        twoLevels[`${i}.${j}`] = true;
      });
    });
    setExpanded(twoLevels);
  }, [initialData.length > 0]);

  // Persist expand state on every change
  React.useEffect(() => {
    try { localStorage.setItem(expandKey, JSON.stringify(expanded)); } catch {}
  }, [expanded]);

  React.useEffect(() => {
    // Load Entities
    fetch("/api/data/v9.2/pmo_entities?$select=pmo_entityid,pmo_name&$filter=statecode eq 0&$orderby=pmo_name asc")
      .then(r => {
        if (!r.ok) throw new Error(`Entity load failed: ${r.status} ${r.statusText}`);
        return r.json();
      })
      .then(d => {
        const items: EntityItem[] = (d.value || []).map((r: any) => ({
          id:   r.pmo_entityid,
          name: r.pmo_name,
        }));
        setEntityItems(items);
      })
      .catch(e => {
        console.error("[TaskGrid] Entity load error:", e);
        // Non-fatal — entity names just won't show
      });

    // Load Fiscal Years
    fetch("/api/data/v9.2/pmo_fiscalyears?$select=pmo_fiscalyearid,pmo_id&$filter=statecode eq 0&$orderby=pmo_id asc")
      .then(r => r.ok ? r.json() : { value: [] })
      .then(d => {
        const items: FiscalYearItem[] = (d.value || []).map((r: any) => ({
          id:   r.pmo_fiscalyearid,
          name: r.pmo_id,
        }));
        setFiscalYearItems(items);
      })
      .catch(e => console.error("[TaskGrid] FY load error:", e));

    // Load SRC with FY and Entity GUIDs
    fetch("/api/data/v9.2/pmo_serviceratecards?$select=pmo_serviceratecardid,pmo_serviceid,pmo_servicename,pmo_price,pmo_unit,pmo_frequency,_pmo_fiscalyear_value,_pmo_entity_value&$filter=statecode eq 0&$orderby=pmo_serviceid asc&$top=500")
      .then(r => {
        if (!r.ok) throw new Error(`SRC load failed: ${r.status} ${r.statusText}`);
        return r.json();
      })
      .then(d => {
        const items: SrcItem[] = (d.value || []).map((r: any) => ({
          id:             r.pmo_serviceratecardid,
          serviceId:      r.pmo_serviceid,
          name:           r.pmo_servicename,
          price:          r.pmo_price,
          unit:           r.pmo_unit,
          frequency:      r.pmo_frequency,
          fiscalYearId:   r._pmo_fiscalyear_value ?? null,
          fiscalYearName: null,
          entityId:       r._pmo_entity_value     ?? null,
          entityName:     null,
        }));
        setSrcItems(items);
      })
      .catch(e => {
        console.error("[TaskGrid] SRC load error:", e);
        setLoadError("Could not load Service Rate Card: " + e.message);
      });
  }, []);

  // Resolve FY and Entity names client-side
  const resolvedSrcItems = React.useMemo(() => {
    return srcItems.map(s => ({
      ...s,
      fiscalYearName: fiscalYearItems.find(f => f.id === s.fiscalYearId)?.name ?? null,
      entityName:     entityItems.find(e => e.id === s.entityId)?.name ?? null,
    }));
  }, [srcItems, entityItems, fiscalYearItems]);

// Load resource assignments using known task IDs
  React.useEffect(() => {
    if (!taskIds || taskIds.length === 0) return;

    // Build OData filter: _msdyn_taskid_value eq 'id1' or _msdyn_taskid_value eq 'id2'...
    // Dataverse supports up to ~100 conditions — chunk if needed
    const chunk = taskIds.slice(0, 80);
    const filter = chunk
      .map(id => `_msdyn_taskid_value eq ${id}`)
      .join(" or ");

    Promise.all([
      fetch(`/api/data/v9.2/msdyn_resourceassignments?$select=_msdyn_taskid_value,_msdyn_bookableresourceid_value&$filter=${encodeURIComponent(filter)}&$top=500`)
        .then(r => r.ok ? r.json() : { value: [] }),
      fetch(`/api/data/v9.2/bookableresources?$select=bookableresourceid,name`)
        .then(r => r.ok ? r.json() : { value: [] }),
    ]).then(([assignments, resources]) => {
      const resourceMap: Record<string, string> = {};
      (resources.value || []).forEach((r: any) => {
        resourceMap[r.bookableresourceid] = r.name;
      });

      const taskMap: TaskResourceMap = {};
      (assignments.value || []).forEach((a: any) => {
        const taskId     = a._msdyn_taskid_value;
        const resourceId = a._msdyn_bookableresourceid_value;
        const name       = resourceMap[resourceId];
        if (!taskId || !name) return;
        if (!taskMap[taskId]) taskMap[taskId] = [];
        if (!taskMap[taskId].find(r => r.id === resourceId)) {
          taskMap[taskId].push({ id: resourceId, name });
        }
      });
      setTaskResources(taskMap);
    }).catch(e => console.error("[TaskGrid] Resource load error:", e));
    }, [taskIds.join(","), refreshCount]);

  // Close column panel on outside click
  React.useEffect(() => {
    if (!colPanelOpen) return;
    function handler(e: MouseEvent) {
      const panel = document.getElementById("tg-col-panel");
      if (
        colBtnRef.current && !colBtnRef.current.contains(e.target as Node) &&
        panel && !panel.contains(e.target as Node)
      ) {
        setColPanelOpen(false);
      }
    }
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [colPanelOpen]);

  function updateField(recordId: string, field: keyof TaskNode, value: unknown) {
    setData(prev => updateNodeInTree(prev, recordId, field, value));
    setPending(p => ({ ...p, [recordId]: { ...(p[recordId] ?? {}), [field]: value } }));
  }

  function applyUpdates(recordId: string, updates: Partial<TaskNode>) {
    setData(prev => {
      let next = prev;
      for (const [field, val] of Object.entries(updates)) {
        next = updateNodeInTree(next, recordId, field as keyof TaskNode, val);
      }
      return next;
    });
    setPending(p => ({ ...p, [recordId]: { ...(p[recordId] ?? {}), ...updates } }));
  }

function recalc(row: TaskNode, overrides: Partial<TaskNode>): Partial<TaskNode> {
    const qty         = overrides.quantity        ?? row.quantity        ?? 0;
    const rate        = overrides.unitRate        ?? row.unitRate        ?? 0;
    const fixed       = overrides.fixedCost       ?? row.fixedCost       ?? 0;
    const completed   = overrides.effortCompleted ?? row.effortCompleted ?? 0;
    const actualFixed = overrides.actualFixedCost ?? row.actualFixedCost ?? 0;
    const pct         = overrides.pctDone         ?? row.pctDone         ?? 0;
    console.log("[recalc] pct=", pct, "fixed=", fixed, "rate=", rate, "qty=", qty);
    const planned     = qty * rate;
    const actual      = completed * rate;  // actualCost = effortCompleted × unitRate
    const total       = planned + fixed;
    const totalActual = actual + actualFixed;
    const remaining   = total - totalActual;
    const ev          = (pct / 100) * total;
    return {
      ...overrides,
      plannedCost:      planned,
      actualCost:       actual,
      totalPlannedCost: total,
      totalActualCost:  totalActual,
      remainingCost:    remaining,
      earnedValue:      ev,
    };
  }

  function onSrcChange(row: TaskNode, srcId: string) {
    const src = srcItems.find(s => s.id === srcId);
    if (!src) return;
    applyUpdates(row.recordId, recalc(row, {
      srcServiceId:   src.id,
      srcServiceName: src.name,
      unitRate:       src.price,
      unit:           src.unit,
      frequency:      src.frequency,
    }));
  }

  function onQuantityChange(row: TaskNode, qty: number) {
    applyUpdates(row.recordId, recalc(row, { quantity: qty }));
  }

function onFixedCostChange(row: TaskNode, fixed: number) {
  setData(prev => {
    function findNode(nodes: TaskNode[]): TaskNode | null {
      for (const n of nodes) {
        if (n.recordId === row.recordId) return n;
        if (n.subRows) { const f = findNode(n.subRows); if (f) return f; }
      }
      return null;
    }
    const current = findNode(prev) ?? row;
    const updates = recalc(current, { fixedCost: fixed });
    setPending(p => ({ ...p, [row.recordId]: { ...(p[row.recordId] ?? {}), ...updates } }));
    let next = prev;
    for (const [field, val] of Object.entries(updates)) {
      next = updateNodeInTree(next, row.recordId, field as keyof TaskNode, val);
    }
    return next;
  });
}

function onActualFixedCostChange(row: TaskNode, actualFixed: number) {
  setData(prev => {
    function findNode(nodes: TaskNode[]): TaskNode | null {
      for (const n of nodes) {
        if (n.recordId === row.recordId) return n;
        if (n.subRows) { const f = findNode(n.subRows); if (f) return f; }
      }
      return null;
    }
    const current = findNode(prev) ?? row;
    const updates = recalc(current, { actualFixedCost: actualFixed });
    setPending(p => ({ ...p, [row.recordId]: { ...(p[row.recordId] ?? {}), ...updates } }));
    let next = prev;
    for (const [field, val] of Object.entries(updates)) {
      next = updateNodeInTree(next, row.recordId, field as keyof TaskNode, val);
    }
    return next;
  });
}

// Flatten all leaf nodes for counting
  function flattenLeaves(nodes: TaskNode[]): TaskNode[] {
    const result: TaskNode[] = [];
    function walk(n: TaskNode) {
      if (!n.subRows || n.subRows.length === 0) { result.push(n); return; }
      n.subRows.forEach(walk);
    }
    nodes.forEach(walk);
    return result;
  }

  const allLeaves = React.useMemo(() => flattenLeaves(data), [data]);

  const fundingCounts = React.useMemo(() => {
    const counts: Record<number, number> = {};
    allLeaves.forEach(n => {
      if (n.fundingSource != null) {
        counts[n.fundingSource] = (counts[n.fundingSource] ?? 0) + 1;
      }
    });
    return counts;
  }, [allLeaves]);

  const categoryCounts = React.useMemo(() => {
    const counts: Record<number, number> = {};
    allLeaves.forEach(n => {
      if (n.costCategory != null) {
        counts[n.costCategory] = (counts[n.costCategory] ?? 0) + 1;
      }
    });
    return counts;
  }, [allLeaves]);

  const serviceCounts = React.useMemo(() => {
    const counts: Record<string, number> = {};
    allLeaves.forEach(n => {
      if (n.srcServiceName) {
        counts[n.srcServiceName] = (counts[n.srcServiceName] ?? 0) + 1;
      }
    });
    return counts;
  }, [allLeaves]);

  const assigneeCounts = React.useMemo(() => {
    const counts: Record<string, number> = {};
    allLeaves.forEach(n => {
      const resources = taskResources[n.recordId] ?? [];
      resources.forEach(r => {
        counts[r.name] = (counts[r.name] ?? 0) + 1;
      });
    });
    return counts;
  }, [allLeaves, taskResources]);

  const totalActiveFilters =
    filterFunding.size + filterCategory.size +
    filterAssignee.size + filterService.size + (filterKeyword.trim() ? 1 : 0);

  function isLeafVisible(node: TaskNode): boolean {
    if (filterKeyword.trim()) {
      if (!node.taskName.toLowerCase().includes(filterKeyword.toLowerCase())) return false;
    }
    if (filterFunding.size > 0) {
      if (node.fundingSource == null || !filterFunding.has(node.fundingSource)) return false;
    }
    if (filterCategory.size > 0) {
      if (node.costCategory == null || !filterCategory.has(node.costCategory)) return false;
    }
    if (filterAssignee.size > 0) {
      const names = (taskResources[node.recordId] ?? []).map(r => r.name);
      if (!names.some(n => filterAssignee.has(n))) return false;
    }
    if (filterService.size > 0) {
      if (!node.srcServiceName || !filterService.has(node.srcServiceName)) return false;
    }
    return true;
  }

function getVisibleLeafIds(): string[] {
    return table.getRowModel().rows
      .filter(r => !r.original.isSummary && (totalActiveFilters === 0 || isLeafVisible(r.original)))
      .map(r => r.original.recordId);
  }

  function getChildLeafIds(node: TaskNode): string[] {
    const result: string[] = [];
    function walk(n: TaskNode) {
      if (!n.subRows || n.subRows.length === 0) { result.push(n.recordId); return; }
      n.subRows.forEach(walk);
    }
    walk(node);
    return result;
  }

  function handleRowSelect(row: any, e: React.MouseEvent) {
    const recordId = row.original.recordId;
    const isSummary = row.original.isSummary;

    // Summary row — select all children
    if (isSummary) {
      const childIds = getChildLeafIds(row.original);
      setSelectedRows(prev => {
      const allSelected = childIds.every(id => prev.includes(id));
      return allSelected
        ? prev.filter(id => !childIds.includes(id))
        : [...new Set([...prev, ...childIds])];
    });
      return;
    }

    // Shift+click — select range
    if (e.shiftKey && lastSelectedId) {
      const allIds = getVisibleLeafIds();
      const fromIdx = allIds.indexOf(lastSelectedId);
      const toIdx   = allIds.indexOf(recordId);
      if (fromIdx !== -1 && toIdx !== -1) {
        const [start, end] = fromIdx < toIdx ? [fromIdx, toIdx] : [toIdx, fromIdx];
        const rangeIds = allIds.slice(start, end + 1);
        setSelectedRows(prev => [...new Set([...prev, ...rangeIds])]);
        return;
      }
    }

    // Normal click — toggle
    setSelectedRows(prev =>
      prev.includes(recordId) ? prev.filter(id => id !== recordId) : [...prev, recordId]
    );
    setLastSelectedId(recordId);
  }

function getVisibleSelectedIds(): string[] {
    const visibleIds = new Set(getVisibleLeafIds());
    return selectedRows.filter(id => visibleIds.has(id));
  }

function applyBulkSRC(srcId: string) {
    const src = srcItems.find(s => s.id === srcId);
    if (!src) return;
    getVisibleSelectedIds().forEach(recordId => {
      const node = table.getRowModel().rows.find(r => r.original.recordId === recordId)?.original;
      if (!node) return;
      applyUpdates(recordId, recalc(node, {
        srcServiceId: src.id, srcServiceName: src.name,
        unitRate: src.price, unit: src.unit, frequency: src.frequency,
      }));
    });
  }

function applyBulkFunding(value: number) {
    getVisibleSelectedIds().forEach(recordId => updateField(recordId, "fundingSource", value));
  }

  function applyBulkUnitRate(rate: number) {
    getVisibleSelectedIds().forEach(recordId => {
      const node = table.getRowModel().rows.find(r => r.original.recordId === recordId)?.original;
      if (!node) return;
      applyUpdates(recordId, recalc(node, { unitRate: rate }));
    });
  }

  function applyBulkCategory(value: number) {
    getVisibleSelectedIds().forEach(recordId => updateField(recordId, "costCategory", value));
  }

function clearAllFilters() {
    setFilterKeyword("");
    setFilterFunding(new Set());
    setFilterCategory(new Set());
    setFilterAssignee(new Set());
    setFilterService(new Set());
  }


  
  const columns = React.useMemo(() => [

col.display({
    id: "select",
    header: () => {
  const allIds = flattenLeaves(data).map(n => n.recordId);
  const isAllSelected = allIds.length > 0 && allIds.every(id => selectedRows.includes(id));
  return (
    <div style={{ display: "flex", justifyContent: "center", alignItems: "center", height: "100%" }}
      onClick={() => setSelectedRows(isAllSelected ? [] : allIds)}>
      <div style={{
        width: 10, height: 10,
        border: isAllSelected ? "1.5px solid #107c10" : "1.5px solid #8a8886",
        borderRadius: 2,
        background: isAllSelected ? "#107c10" : "white",
        cursor: "pointer",
        display: "flex", alignItems: "center", justifyContent: "center",
        flexShrink: 0,
      }}>
        {isAllSelected && (
          <svg width="10" height="10" viewBox="0 0 24 24" fill="none"
            stroke="white" strokeWidth="3.5" strokeLinecap="round" strokeLinejoin="round">
            <polyline points="20 6 9 17 4 12"/>
          </svg>
        )}
      </div>
    </div>
  );
},
    size: 32,
    cell: function SelectCell({ row }) {
      if (row.original.isSummary) {
        const childIds = getChildLeafIds(row.original);
        const allSelected = childIds.length > 0 && childIds.every(id => selectedRows.includes(id));
        const someSelected = childIds.some(id => selectedRows.includes(id));
return (
          <div
            onClick={() => {
              const allSel = childIds.every(id => selectedRows.includes(id));
              setSelectedRows(allSel
                ? selectedRows.filter(id => !childIds.includes(id))
                : [...new Set([...selectedRows, ...childIds])]);
            }}
            style={{
              width: 10, height: 10,
              border: allSelected ? "1.5px solid #107c10" : someSelected ? "1.5px solid #107c10" : "1.5px solid #8a8886",
              borderRadius: 2,
              background: allSelected ? "#107c10" : someSelected ? "white" : "white",
              cursor: "pointer",
              display: "flex", alignItems: "center", justifyContent: "center",
              flexShrink: 0, margin: "0 auto",
            }}>
            {allSelected && (
              <svg width="10" height="10" viewBox="0 0 24 24" fill="none"
                stroke="white" strokeWidth="3.5" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="20 6 9 17 4 12"/>
              </svg>
            )}
            {someSelected && !allSelected && (
              <svg width="8" height="8" viewBox="0 0 24 24" fill="none"
                stroke="#107c10" strokeWidth="4" strokeLinecap="round">
                <line x1="5" y1="12" x2="19" y2="12"/>
              </svg>
            )}
          </div>
        );
      }
      const isSelected = selectedRows.includes(row.original.recordId);
      return (
        <div
          onClick={e => { e.stopPropagation(); handleRowSelect(row, e as any); }}
          style={{
            width: 10, height: 10,
            border: isSelected ? "1.5px solid #107c10" : "1.5px solid #8a8886",
            borderRadius: 2,
            background: isSelected ? "#107c10" : "white",
            cursor: "pointer",
            display: "flex", alignItems: "center", justifyContent: "center",
            flexShrink: 0, margin: "0 auto",
          }}>
          {isSelected && (
            <svg width="10" height="10" viewBox="0 0 24 24" fill="none"
              stroke="white" strokeWidth="3.5" strokeLinecap="round" strokeLinejoin="round">
              <polyline points="20 6 9 17 4 12"/>
            </svg>
          )}
        </div>
      );
    },
  }),

  // ── Schedule ──────────────────────────────────────────────────────────────
  col.accessor("taskName", {
    header: "Task name", size: 240,
    enableResizing: true,
    cell: function TaskNameCell({ row, getValue }) {
      const isSummary = row.original.isSummary;
      const isRowHovered = hoveredRow === row.id;
      return (
        <div style={{
          paddingLeft: row.depth * 20, display: "flex", alignItems: "center",
          fontWeight: isSummary ? 600 : 400,
          color: isSummary ? "#111827" : "#374151",
        }}>
          {row.getCanExpand() ? (
            <button className="tg-expand-btn" onClick={row.getToggleExpandedHandler()}>
              {row.getIsExpanded() ? <ChevronDown /> : <ChevronRight />}
            </button>
          ) : (
            <span style={{ width: 22, display: "inline-block", flexShrink: 0 }} />
          )}
          <span className="tg-cell-text" style={{
            textDecoration: row.original.pctDone >= 100 ? "line-through" : "none",
            color: row.original.pctDone >= 100 ? "#605e5c" : "inherit",
          }}>{String(getValue() ?? "")}</span>
          {!isSummary && (
            <button
              onClick={e => { e.stopPropagation(); setDetailTask(row.original); }}
              style={{
                background: "none", border: "none", cursor: "pointer",
                color: "#107c10", fontSize: 14, lineHeight: 1,
                padding: "0 0 0 6px", display: "inline-flex", alignItems: "center",
                marginLeft: "auto",
                opacity: isRowHovered ? 1 : 0.75,
                transition: "opacity 0.15s",
                pointerEvents: "auto",
              }}
              title="Open task details"
            >ⓘ</button>
          )}
        </div>
      );
    },
  }),
  col.accessor("startDate", {
    header: "Start", size: 90,
    cell: function StartCell({ getValue }) {
      return <span className="tg-cell-text">{formatDate(String(getValue() ?? ""))}</span>;
    },
  }),
  col.accessor("endDate", {
    header: "Finish", size: 90,
    cell: function EndCell({ getValue }) {
      return <span className="tg-cell-text">{formatDate(String(getValue() ?? ""))}</span>;
    },
  }),
  col.accessor("pctDone", {
    header: "% Complete", size: 120,
    cell: function PctCell({ getValue }) {
      return <ProgressCell value={Number(getValue() ?? 0)} />;
    },
  }),

  col.display({
    id: "assignedTo",
    header: () => <ColTooltip label="Assigned To" tip="Team members assigned to this task. Read-only — managed via the project scheduling engine, not editable in this grid." />,
    size: 180,
    cell: function AssignedCell({ row }) {
      if (row.original.isSummary) return <span className="tg-dash">—</span>;
      const resources = taskResources[row.original.recordId] ?? [];
      return <ResourceCell resources={resources} />;
    },
  }),

  // ── Cost — input fields hidden on summary rows ────────────────────────────
  col.accessor("fundingSource", {
    header: () => <ColTooltip label="Funding Source" tip="The budget line funding this task. Regular Budget = core UN budget · PK + Support Account = peacekeeping support · xB = extrabudgetary · 10RCR = cost recovery · 20PCR = peacekeeping cost recovery." />,
    size: 160,
    cell: function FundingCell({ row }) {
      if (row.original.isSummary) return <span className="tg-dash">—</span>;
      return (
        <select className="tg-select"
          value={row.original.fundingSource ?? ""}
          onChange={e => updateField(row.original.recordId, "fundingSource",
            e.target.value === "" ? null : Number(e.target.value))}>
          <option value="">— select —</option>
          {FUNDING_SOURCES.map(f => (
            <option key={f.value} value={f.value}>{f.label}</option>
          ))}
        </select>
      );
    },
  }),
  col.accessor("costCategory", {
    header: () => <ColTooltip label="Cost Category" tip="UN cost classification for this task. Staff = personnel costs · Supplies = consumables · Equipment = assets and vehicles · Contractual = external services · Travel = mission travel · Indirect = overhead costs." />,
    size: 160,
    cell: function CatCell({ row }) {
      // Summary: show dash — no editing, no dropdown
      if (row.original.isSummary) return <span className="tg-dash">—</span>;
      return (
        <select className="tg-select"
          value={row.original.costCategory ?? ""}
          onChange={e => updateField(row.original.recordId, "costCategory",
            e.target.value === "" ? null : Number(e.target.value))}>
          <option value="">— select —</option>
          {COST_CATEGORIES.map(c => (
            <option key={c.value} value={c.value}>{c.label}</option>
          ))}
        </select>
      );
    },
  }),
col.accessor("srcServiceName", {
  header: () => <ColTooltip label="Service" tip="Links this task to a Service Rate Card entry. Selecting a service automatically fills the Unit Rate and Unit fields. Search by service code, name, fiscal year, or entity." />,
  size: 220,
  cell: function SrcCell({ row }) {
    if (row.original.isSummary) return <span className="tg-dash">—</span>;
    return (
        <ServiceCombobox
            value={row.original.srcServiceId}
            items={resolvedSrcItems}
            onChange={(id, src) => onSrcChange(row.original, id)}
        />
    );
  },
}),
col.accessor("quantity", {
  header: () => <ColTooltip label="Effort (h)" tip="Total planned hours assigned to this task. Drives planned cost when multiplied by the unit rate. (Planned Cost = Effort × Unit Rate)" />,
  size: 75,
  cell: function QtyCell({ row }) {
    if (row.original.isSummary) return <span className="tg-dash">—</span>;
    return (
      <span className="tg-cell-right" style={{ color: "#6b7280" }}>
        {fmtNumber(row.original.quantity)}
      </span>
    );
  },
}),
col.accessor("effortCompleted", {
    header: () => <ColTooltip label="Completed (h)" tip="Actual hours worked on this task so far. Used to calculate actual effort cost. (Actual Effort Cost = Completed (h) × Unit Rate)" />,
    size: 90,
    cell: function CompletedCell({ row }) {
      if (row.original.isSummary) return <span className="tg-dash">—</span>;
      return (
        <span className="tg-cell-right" style={{ color: "#6b7280" }}>
          {fmtNumber(row.original.effortCompleted ?? 0)}
        </span>
      );
    },
  }),
  col.accessor("unit", {
    header: "Unit", size: 60,
    cell: function UnitCell({ row }) {
      // Summary: show dash
      if (row.original.isSummary) return <span className="tg-dash">—</span>;
      return <span className="tg-cell-text">{String(row.original.unit ?? "")}</span>;
    },
  }),
  col.accessor("unitRate", {
    header: () => <ColTooltip label="Unit Rate" tip="Cost per hour for this task. Pulled automatically from the Service Rate Card when a service is selected. Editable directly if needed." />,
    size: 95,
    cell: function RateCell({ row }) {
      if (row.original.isSummary) return <span className="tg-dash">—</span>;
      return (
        <CurrencyInput
          value={row.original.unitRate ?? 0}
          onChange={v => applyUpdates(row.original.recordId, recalc(row.original, { unitRate: v }))}
        />
      );
    },
  }),
  col.accessor("plannedCost", {
    header: () => <ColTooltip label="Planned Cost" tip="Effort-based planned cost only. Does not include fixed costs. (Planned Cost = Effort (h) × Unit Rate)" />,
    size: 100,
    cell: function PlannedCell({ row }) {
      return (
        <span className="tg-cell-right"
          style={{ fontWeight: row.original.isSummary ? 700 : 500 }}>
          {fmtCurrency(Number(row.original.plannedCost ?? 0))}
        </span>
      );
    },
  }),
  col.accessor("fixedCost", {
    header: () => <ColTooltip label="Fixed Cost" tip="A one-off cost independent of effort. Examples: equipment, travel, venue hire. (Total Planned = Planned Cost + Fixed Cost)" />,
    size: 100,
    cell: function FixedCell({ row }) {
      // Summary: show rolled-up total, read-only
      if (row.original.isSummary) {
        return (
          <span className="tg-cell-right" style={{ fontWeight: 600 }}>
            {fmtCurrency(row.original.fixedCost)}
          </span>
        );
      }
      return <CurrencyInput value={row.original.fixedCost}
        onChange={v => onFixedCostChange(row.original, v)} />;
    },
  }),
  col.accessor("totalPlannedCost", {
    header: () => <ColTooltip label="Total Planned" tip="Full budgeted cost combining effort-based cost and fixed costs. Primary budget figure used in KPIs. (Total Planned = Planned Cost + Fixed Cost)" />,
    size: 110,
    cell: function TotalCell({ row }) {
      return (
        <span className="tg-cell-right"
          style={{ fontWeight: row.original.isSummary ? 700 : 500, color: "#4f46e5" }}>
          {fmtCurrency(Number(row.original.totalPlannedCost ?? 0))}
        </span>
      );
    },
  }),
col.accessor("actualCost", {
    header: () => <ColTooltip label="Actual Effort Cost" tip="Cost of hours actually worked. Calculated automatically. (Actual Effort Cost = Completed (h) × Unit Rate)" />,
    size: 120,
    cell: function ActualCell({ row }) {
      return (
        <span className="tg-cell-right"
          style={{ fontWeight: row.original.isSummary ? 700 : 500 }}>
          {fmtCurrency(Number(row.original.actualCost ?? 0))}
        </span>
      );
    },
  }),
  col.accessor("actualFixedCost", {
    header: () => <ColTooltip label="Actual Fixed Cost" tip="Fixed costs already incurred. Enter manually as invoices or expenses are confirmed. (Total Actual = Actual Effort Cost + Actual Fixed Cost)" />,
    size: 120,
    cell: function ActualFixedCell({ row }) {
      if (row.original.isSummary) {
        return (
          <span className="tg-cell-right" style={{ fontWeight: 600 }}>
            {fmtCurrency(row.original.actualFixedCost)}
          </span>
        );
      }
      return <CurrencyInput value={row.original.actualFixedCost ?? 0}
        onChange={v => onActualFixedCostChange(row.original, v)} />;
    },
  }),
  col.accessor("totalActualCost", {
    header: () => <ColTooltip label="Total Actual Cost" tip="Total spent to date combining effort and confirmed fixed costs. (Total Actual = Actual Effort Cost + Actual Fixed Cost)" />,
    size: 120,
    cell: function TotalActualCell({ row }) {
      return (
        <span className="tg-cell-right" style={{
          fontWeight: row.original.isSummary ? 700 : 500,
          color: "#0f766e",
        }}>
          {fmtCurrency(Number(row.original.totalActualCost ?? 0))}
        </span>
      );
    },
  }),
  col.accessor("remainingCost", {
    header: () => <ColTooltip label="Remaining" tip="Budget still available for this task. Negative means the task has overrun. (Remaining = Total Planned − Total Actual Cost)" />,
    size: 100,
    cell: function RemCell({ row }) {
      const v = Number(row.original.remainingCost ?? 0);
      return (
        <span className="tg-cell-right" style={{
          color: v < 0 ? "#dc2626" : "#374151",
          fontWeight: row.original.isSummary ? 600 : 400,
        }}>
          {fmtCurrency(v)}
        </span>
      );
    },
  }),
  col.accessor("earnedValue", {
  header: () => <ColTooltip label="Earned Value" tip="Monetary value of work actually completed. If EV is below actual cost, you are spending more than the work is worth. (EV = % Complete × Total Planned)" />,
  size: 95,
  cell: function EvCell({ row }) {
    return (
      <span className="tg-cell-right"
        style={{ fontWeight: row.original.isSummary ? 600 : 400 }}>
        {fmtCurrency(Number(row.original.earnedValue ?? 0))}
      </span>
    );
  },
}),
], [resolvedSrcItems, taskResources, selectedRows]);

const columnVisibility = React.useMemo(() => {
    const vis: Record<string, boolean> = {};
    ALL_COLUMNS.forEach(c => { vis[c.id] = visibleCols.has(c.id); });
    return vis;
  }, [visibleCols]);

const table = useReactTable({
    data, columns,
    state: { expanded, columnVisibility, columnOrder, columnSizing },
    onExpandedChange:     setExpanded,
    onColumnOrderChange:  setColumnOrder,
    onColumnSizingChange: sizing => {
      setColumnSizing(sizing as Record<string, number>);
    },
    columnResizeMode:     "onChange",
    getSubRows:           row => row.subRows,
    getCoreRowModel:      getCoreRowModel(),
    getExpandedRowModel:  getExpandedRowModel(),
});

  const changesCount = Object.keys(pending).length;

    async function handleSave() {
        setSaving(true);
        try {
        await onSave(pending);
        setPending({});
        setSavedMsg(true);
        setTimeout(() => setSavedMsg(false), 2500);
        // Auto-refresh after save so Dataverse calculated fields update
        setTimeout(() => onRefresh(), 1500);
        } finally {
        setSaving(false);
        }
    }

  function toggleExpandAll() {
  if (allExpanded) {
    setExpanded({});
    setAllExpanded(false);
  } else {
    const all: Record<string, boolean> = {};
    function walk(rows: any[]) {
      rows.forEach((r: any) => {
        if (r.getCanExpand()) { all[r.id] = true; walk(r.subRows ?? []); }
      });
    }
    walk(table.getRowModel().rows);
    setExpanded(all);
    setAllExpanded(true);
  }
}

function toggleColumn(id: string) {
    setVisibleCols(prev => {
      const next = new Set(prev);
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }
      try { localStorage.setItem(storageKey, JSON.stringify([...next])); } catch {}
      return next;
    });
  }

  function openColPanel() {
    if (colBtnRef.current) {
      const rect = colBtnRef.current.getBoundingClientRect();
      setColPanelPos({ top: rect.bottom + window.scrollY, left: rect.left + window.scrollX });
    }
    setColPanelOpen(o => !o);
  }

  function showAllColumns() {
    const all = new Set(ALL_COLUMNS.map(c => c.id));
    setVisibleCols(all);
    try { localStorage.setItem(storageKey, JSON.stringify([...all])); } catch {}
  }

  function reorderColumn(dragId: string, dropId: string) {
    if (dragId === dropId || dragId === "taskName" || dragId === "select") return;
        const currentOrder = columnOrder.length > 0
        ? columnOrder
        : table.getAllLeafColumns().map(c => c.id);
        const next = [...currentOrder];
        const fromIdx = next.indexOf(dragId);
        const toIdx   = next.indexOf(dropId);
        if (fromIdx === -1 || toIdx === -1) return;
        next.splice(fromIdx, 1);
        next.splice(toIdx, 0, dragId);
        // Ensure taskName always first
        const tnIdx = next.indexOf("taskName");
        if (tnIdx > 0) { next.splice(tnIdx, 1); next.unshift("taskName"); }
        const selIdx = next.indexOf("select");
        if (selIdx > 0) { next.splice(selIdx, 1); next.unshift("select"); }
        setColumnOrder(next);
        try { localStorage.setItem(orderKey, JSON.stringify(next)); } catch {}
    }

    const rightCols = new Set(["plannedCost","fixedCost","totalPlannedCost",
    "actualCost","actualFixedCost","totalActualCost","remainingCost","earnedValue","unitRate","quantity"]);

  const handleDetailSave = React.useCallback(async (recordId: string, updates: Partial<TaskNode>) => {
    await onSave({ [recordId]: updates });
    onRefresh();
  }, []);

  const handleDetailClose = React.useCallback(() => {
    setDetailTask(null);
  }, []);

  return (
    <div className="tg-wrap">
      <style>{CSS}</style>

      <div className="tg-toolbar">
  <div style={{ width: 6 }} />
  <button className="tg-btn" onClick={toggleExpandAll}>
    {allExpanded ? (
      <React.Fragment>
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none"
          stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <polyline points="6 9 12 15 18 9"/>
        </svg>
        Collapse all
      </React.Fragment>
    ) : (
      <React.Fragment>
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none"
          stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <polyline points="9 18 15 12 9 6"/>
        </svg>
        Expand all
      </React.Fragment>
    )}
  </button>
  <button className="tg-btn" onClick={onRefresh}>
    <RefreshIcon />Refresh
  </button>
  <button ref={colBtnRef} className="tg-btn" onClick={openColPanel}>
    <ColumnsIcon />Columns
  </button>

  <button className="tg-btn" onClick={() => setSummaryOpen(o => !o)}>
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none"
      stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/>
      <line x1="6" y1="20" x2="6" y2="14"/>
    </svg>
    Summary
    {summaryOpen && <span className="tg-filter-badge">●</span>}
  </button>
  {/* Timeline button hidden — msdyn_plannedwork contains remaining work not planned work (Microsoft platform limitation) */}
  <button className="tg-btn" onClick={() => setTimelineOpen(o => !o)}>
  <svg width="14" height="14" viewBox="0 0 24 24" fill="none"
    stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <rect x="3" y="4" width="18" height="18" rx="2" ry="2"/>
    <line x1="16" y1="2" x2="16" y2="6"/>
    <line x1="8" y1="2" x2="8" y2="6"/>
    <line x1="3" y1="10" x2="21" y2="10"/>
  </svg>
  Financial Plan
  {timelineOpen && <span className="tg-filter-badge">●</span>}
</button>
  <button className="tg-btn" onClick={() => setFilterOpen(o => !o)}>
    <svg width="14" height="14" viewBox="0 0 24 24" fill="none"
      stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/>
    </svg>
    Filter
    {totalActiveFilters > 0 && (
      <span className="tg-filter-badge">{totalActiveFilters}</span>
    )}
  </button>

<div style={{ flex: 1 }} />
  {selectedRows.length > 0 && (
    <React.Fragment>
      <div className="tg-divider" />
      <span style={{ fontSize: 12, color: "#107c10", fontWeight: 600, whiteSpace: "nowrap" }}>
        ✓ {getVisibleSelectedIds().length} {getVisibleSelectedIds().length === 1 ? "row" : "rows"} selected (visible)
      </span>

      <BulkDropdown label="Funding">
        {FUNDING_SOURCES.map(f => (
          <BulkMenuItem key={f.value} onClick={() => applyBulkFunding(f.value)}>
            {f.label}
          </BulkMenuItem>
        ))}
      </BulkDropdown>

      <BulkDropdown label="Category">
        {COST_CATEGORIES.map(c => (
          <BulkMenuItem key={c.value} onClick={() => applyBulkCategory(c.value)}>
            {c.label}
          </BulkMenuItem>
        ))}
      </BulkDropdown>

      <BulkDropdown label="Unit Rate">
        <BulkUnitRateInput onApply={rate => { applyBulkUnitRate(rate); }} />
      </BulkDropdown>

      <BulkDropdown label="Service">
        <div style={{ padding: "6px 8px", borderBottom: "1px solid #f3f4f6" }}>
          <BulkServiceSearch items={resolvedSrcItems} onSelect={applyBulkSRC} />
        </div>
      </BulkDropdown>

      <button className="tg-btn" style={{ color: "#6b7280", fontSize: 12 }}
        onClick={() => setSelectedRows([])}>
        ✕ Clear
      </button>
    </React.Fragment>
  )}
  {changesCount > 0 && !savedMsg && (
    <React.Fragment>
      <span style={{ fontSize: 12, color: "#605e5c" }}>
        <span className="tg-badge">{changesCount}</span>
        {" "}{changesCount === 1 ? "unsaved change" : "unsaved changes"}
      </span>
      <button className="tg-btn tg-btn-discard"
        onClick={() => { setData(initialData); setPending({}); }}>
        ✕ Discard
      </button>
      <button className="tg-btn tg-btn-save" onClick={handleSave} disabled={saving}>
        {saving ? "Saving…" : "Save"}
      </button>
    </React.Fragment>
  )}
  {savedMsg && <span className="tg-saved">✓ Saved</span>}
  {/* Column visibility panel */}
      {colPanelOpen && typeof document !== "undefined" && ReactDOM.createPortal(
        <div id="tg-col-panel" className="tg-col-panel"
          style={{ top: colPanelPos.top, left: colPanelPos.left, position: "absolute" }}>

          <div className="tg-col-panel-header">Schedule</div>
          {ALL_COLUMNS.filter(c => c.group === "schedule").map(c => (
            <label key={c.id} className="tg-col-item">
              <input type="checkbox"
                checked={visibleCols.has(c.id)}
                onChange={() => toggleColumn(c.id)}
              />
              {c.label}
            </label>
          ))}

          <div className="tg-col-divider" />
          <div className="tg-col-panel-header">Cost</div>
          {ALL_COLUMNS.filter(c => c.group === "cost").map(c => (
            <label key={c.id} className="tg-col-item">
              <input type="checkbox"
                checked={visibleCols.has(c.id)}
                onChange={() => toggleColumn(c.id)}
              />
              {c.label}
            </label>
          ))}

          <div className="tg-col-divider" />
            <div className="tg-col-footer">
                <button onClick={showAllColumns}>Show all</button>
                <button onClick={() => {
                    setColumnOrder(DEFAULT_ORDER);
                    try { localStorage.setItem(orderKey, JSON.stringify(DEFAULT_ORDER)); } catch {}
                }}>Reset order</button>
                <button onClick={() => {
                    setColumnSizing({});
                    try { localStorage.removeItem(sizingKey); } catch {}
                }}>Reset widths</button>
                <button onClick={() => setColPanelOpen(false)}>Close</button>
            </div>
        </div>,
        document.body
      )}
</div>

      {loadError && <div className="tg-error">⚠ {loadError}</div>}

      {detailTask && (
        <TaskDetailPanel
          task={detailTask}
          resources={taskResources[detailTask.recordId] ?? []}
          srcItems={resolvedSrcItems}
          onSave={handleDetailSave}
          onClose={handleDetailClose}
        />
      )}

      {summaryOpen && (
        <SummaryPanel data={data} onClose={() => setSummaryOpen(false)} latestApprovedBudget={latestApprovedBudget} />
      )}
      {filterOpen && (
        <div className="tg-filter-panel">
          <div className="tg-filter-header">
            <span className="tg-filter-title">Filter Tasks</span>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              {totalActiveFilters > 0 && (
                <button className="tg-filter-clearall" onClick={clearAllFilters}>
                  Clear All
                </button>
              )}
              <button className="tg-filter-close" onClick={() => setFilterOpen(false)}>✕</button>
            </div>
          </div>

          <input
            className="tg-filter-search"
            placeholder="Filter by keyword..."
            value={filterKeyword}
            onChange={e => setFilterKeyword(e.target.value)}
          />

          <div className="tg-filter-body">

            {/* Funding Source */}
            <div className="tg-filter-section">
              <div className="tg-filter-section-header"
                onClick={() => setSectionOpen(p => ({ ...p, funding: !p.funding }))}>
                <span>Funding Source {filterFunding.size > 0 && <span className="tg-filter-badge">{filterFunding.size}</span>}</span>
                {sectionOpen.funding ? <ChevronDown /> : <ChevronRight />}
              </div>
              {sectionOpen.funding && FUNDING_SOURCES.map(f => (
                <label key={f.value} className="tg-filter-item">
                  <input type="checkbox"
                    checked={filterFunding.has(f.value)}
                    onChange={() => setFilterFunding(prev => {
                      const next = new Set(prev);
                      next.has(f.value) ? next.delete(f.value) : next.add(f.value);
                      return next;
                    })}
                  />
                  <span>{f.label}</span>
                  <span className="tg-filter-count">({fundingCounts[f.value] ?? 0})</span>
                </label>
              ))}
            </div>

            {/* Cost Category */}
            <div className="tg-filter-section">
              <div className="tg-filter-section-header"
                onClick={() => setSectionOpen(p => ({ ...p, category: !p.category }))}>
                <span>Cost Category {filterCategory.size > 0 && <span className="tg-filter-badge">{filterCategory.size}</span>}</span>
                {sectionOpen.category ? <ChevronDown /> : <ChevronRight />}
              </div>
              {sectionOpen.category && COST_CATEGORIES.map(c => (
                <label key={c.value} className="tg-filter-item">
                  <input type="checkbox"
                    checked={filterCategory.has(c.value)}
                    onChange={() => setFilterCategory(prev => {
                      const next = new Set(prev);
                      next.has(c.value) ? next.delete(c.value) : next.add(c.value);
                      return next;
                    })}
                  />
                  <span>{c.label}</span>
                  <span className="tg-filter-count">({categoryCounts[c.value] ?? 0})</span>
                </label>
              ))}
            </div>
            {/* Service */}
            <div className="tg-filter-section">
              <div className="tg-filter-section-header"
                onClick={() => setSectionOpen(p => ({ ...p, service: !p.service }))}>
                <span>Service {filterService.size > 0 && <span className="tg-filter-badge">{filterService.size}</span>}</span>
                {sectionOpen.service ? <ChevronDown /> : <ChevronRight />}
              </div>
              {sectionOpen.service && Object.entries(serviceCounts).map(([name, count]) => (
                <label key={name} className="tg-filter-item">
                  <input type="checkbox"
                    checked={filterService.has(name)}
                    onChange={() => setFilterService(prev => {
                      const next = new Set(prev);
                      next.has(name) ? next.delete(name) : next.add(name);
                      return next;
                    })}
                  />
                  <span>{name}</span>
                  <span className="tg-filter-count">({count})</span>
                </label>
              ))}
            </div>
            {/* Assignee */}
            <div className="tg-filter-section">
              <div className="tg-filter-section-header"
                onClick={() => setSectionOpen(p => ({ ...p, assignee: !p.assignee }))}>
                <span>Assignee {filterAssignee.size > 0 && <span className="tg-filter-badge">{filterAssignee.size}</span>}</span>
                {sectionOpen.assignee ? <ChevronDown /> : <ChevronRight />}
              </div>
              {sectionOpen.assignee && Object.entries(assigneeCounts).map(([name, count]) => (
                <label key={name} className="tg-filter-item">
                  <input type="checkbox"
                    checked={filterAssignee.has(name)}
                    onChange={() => setFilterAssignee(prev => {
                      const next = new Set(prev);
                      next.has(name) ? next.delete(name) : next.add(name);
                      return next;
                    })}
                  />
                  <span>{name}</span>
                  <span className="tg-filter-count">({count})</span>
                </label>
              ))}
            </div>

          </div>
        </div>
      )}

      <div className="tg-scroll" style={{ overflow: "scroll", flex: 1, minHeight: 0 }}>
         <table className="tg-table">
          <thead style={{ position: "sticky", top: 0, zIndex: 3 }}>
            {table.getHeaderGroups().map(hg => (
              <tr key={hg.id}>
{hg.headers.map(h => {
                  const isLocked  = h.column.id === "taskName";
                  const isDragging = dragColId === h.column.id;
                  const isDropTarget = dropColId === h.column.id;
                  return (
                    <th key={h.id}
                      className={[
                        rightCols.has(h.column.id) ? "th-right" : "",
                        isLocked    ? "th-nodrag"     : "",
                        isDragging  ? "tg-th-dragging" : "",
                        isDropTarget && !isLocked ? "tg-th-drop-left" : "",
                      ].filter(Boolean).join(" ")}
                      style={{
                          width: h.getSize(),
                          minWidth: h.getSize(),
                      }}
                      draggable={!isLocked}
                      onDragStart={e => {
                        if (isLocked || h.column.getIsResizing()) return;
                        setDragColId(h.column.id);
                        e.dataTransfer.effectAllowed = "move";
                      }}
                      onDragOver={e => {
                        e.preventDefault();
                        e.dataTransfer.dropEffect = "move";
                        if (!isLocked && h.column.id !== dragColId) {
                          setDropColId(h.column.id);
                        }
                      }}
                      onDragLeave={() => setDropColId(null)}
                      onDrop={e => {
                        e.preventDefault();
                        if (dragColId && !isLocked) {
                          reorderColumn(dragColId, h.column.id);
                        }
                        setDragColId(null);
                        setDropColId(null);
                      }}
                      onDragEnd={() => {
                        setDragColId(null);
                        setDropColId(null);
                      }}
                    >
                      {flexRender(h.column.columnDef.header, h.getContext())}
                      {h.column.getCanResize() && (
                        <div
                          className={`tg-resize-handle${h.column.getIsResizing() ? " is-resizing" : ""}`}
                          onMouseDown={e => {
                            e.stopPropagation();
                            e.preventDefault();
                            const startX = e.clientX;
                            const startSize = h.column.getSize();
                            function onMove(ev: MouseEvent) {
                              const delta = ev.clientX - startX;
                              const newSize = Math.max(40, startSize + delta);
                              h.column.parent?.columnDef ?? null;
                              table.setColumnSizing(prev => ({
                                ...prev,
                                [h.column.id]: newSize,
                              }));
                            }
                            function onUp() {
                              document.removeEventListener("mousemove", onMove);
                              document.removeEventListener("mouseup", onUp);
                              try {
                                localStorage.setItem(sizingKey, JSON.stringify(
                                  table.getState().columnSizing
                                ));
                              } catch {}
                            }
                            document.addEventListener("mousemove", onMove);
                            document.addEventListener("mouseup", onUp);
                          }}
                        />
                      )}
                    </th>
                  );
                })}
              </tr>
            ))}
          </thead>
          <tbody>
            {table.getRowModel().rows.filter(row => {
              if (totalActiveFilters === 0) return true;
              if (row.original.isSummary) return true;
              return isLeafVisible(row.original);
            }).map(row => {
              const isSummary = row.original.isSummary;
              const isHovered = hoveredRow === row.id;
              const isChanged = !!pending[row.original.recordId];

                const tdBase: React.CSSProperties = {
                padding: "0 10px",
                borderBottom: "1px solid #d8d8d8",
                verticalAlign: "middle",
                height: 38,
                fontWeight: isSummary ? 600 : 400,
                color: isSummary ? "#111827" : "#374151",
                background: isHovered ? "#f3f2f1" : isSummary ? "#fafafa" : "#ffffff",
                overflow: "hidden",
                textOverflow: "ellipsis",
                whiteSpace: "nowrap",
                maxWidth: 0,
                };

              return (
                <tr key={row.id}
                  className={selectedRows.includes(row.original.recordId) ? "tg-row-selected" : ""}
                  onMouseEnter={() => { if (!detailTask) setHoveredRow(row.id); }}
                  onMouseLeave={() => { if (!detailTask) setHoveredRow(null); }}
                  >
                  {row.getVisibleCells().map((cell, i) => (
                  <td key={cell.id} style={{
                      ...tdBase,
                      borderRight: i === row.getVisibleCells().length - 1
                        ? "none" : "1px solid #f9fafb",
                      borderLeft: i === 0 && isChanged
                        ? "3px solid #4f46e5" : undefined,
                    }}>
                      {flexRender(cell.column.columnDef.cell, cell.getContext())}
                    </td>
                  ))}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div className="tg-footer">
        <span>{table.getRowModel().rows.length} rows visible</span>
        <span></span>
      </div>
      {timelineOpen && (
      <FinancialTimeline
        data={data}
        isLeafVisible={isLeafVisible}
        hasActiveFilters={totalActiveFilters > 0}
        onClose={() => setTimelineOpen(false)}
        projectId={projectId}
        taskIds={taskIds}
      />
      )}
    </div>
  );
}