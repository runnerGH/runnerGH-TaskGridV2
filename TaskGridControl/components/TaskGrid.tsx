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

interface Props {
  data:                  TaskNode[];
  onSave:                (changes: Record<string, Partial<TaskNode>>) => Promise<void>;
  onRefresh:             () => void;
  userId:                string;
  taskIds:               string[];
  latestApprovedBudget:  number;
}

const COST_CATEGORIES = [
  { value: 847020000, label: "Staff and Other Personnel Costs" },
  { value: 847020001, label: "Supplies, Commodities, and Materials" },
  { value: 847020002, label: "Equipment, Vehicles, and Furniture" },
  { value: 847020003, label: "Contractual Services" },
  { value: 847020004, label: "Travel" },
  { value: 847020005, label: "Indirect Costs" },
];

const FUNDING_SOURCES = [
  { value: 0, label: "Regular Budget" },
  { value: 1, label: "Support Account" },
  { value: 2, label: "xB" },
  { value: 3, label: "10RCR (Cost Recovery)" },
  { value: 4, label: "20PCR (PK Cost Recovery)" },
];

const ALL_COLUMNS = [
  { id: "startDate",        label: "Start",              group: "schedule" },
  { id: "endDate",          label: "Finish",             group: "schedule" },
  { id: "pctDone",          label: "% Complete",         group: "schedule" },
  { id: "assignedTo",       label: "Assigned to",        group: "schedule" },
  { id: "fundingSource",    label: "Funding source",     group: "cost" },
  { id: "costCategory",     label: "Cost category",      group: "cost" },
  { id: "srcServiceName",   label: "Service",            group: "cost" },
  { id: "quantity",         label: "Effort (h)",         group: "cost" },
  { id: "effortCompleted",  label: "Completed (h)",      group: "cost" },
  { id: "unit",             label: "Unit",               group: "cost" },
  { id: "unitRate",         label: "Unit rate",          group: "cost" },
  { id: "plannedCost",      label: "Planned cost",       group: "cost" },
  { id: "fixedCost",        label: "Fixed cost",         group: "cost" },
  { id: "totalPlannedCost", label: "Total planned",      group: "cost" },
  { id: "actualCost",       label: "Actual effort cost", group: "cost" },
  { id: "actualFixedCost",  label: "Actual fixed cost",  group: "cost" },
  { id: "totalActualCost",  label: "Total actual cost",  group: "cost" },
  { id: "remainingCost",    label: "Remaining",          group: "cost" },
  { id: "earnedValue",      label: "Earned value",       group: "cost" },
] as const;

const DEFAULT_VISIBLE = new Set(ALL_COLUMNS.map(c => c.id));

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
  return d.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
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
              <span style={{ fontWeight: 500, fontSize: 12 }}>{s.serviceId}</span>
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
  const largeArc = pct > 50 ? 1 : 0;
  return (
    <svg width="210" height="118" viewBox="0 0 210 118">
      <path d={`M ${x1} ${y1} A ${r} ${r} 0 1 1 ${x2} ${y2}`}
        fill="none" stroke="#e5e7eb" strokeWidth="16" strokeLinecap="round"/>
      {pct > 0 && (
        <path d={`M ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${px} ${py}`}
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

const DONUT_COLORS = ["#4f46e5","#0f766e","#d97706","#dc2626","#7c3aed","#0284c7","#16a34a","#9ca3af"];

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

  const weightedPct = totalPlanned > 0
    ? leaves.reduce((s, n) => s + ((n.pctDone ?? 0) * (n.totalPlannedCost ?? 0)), 0) / totalPlanned
    : 0;

  const consumptionRate = (weightedPct > 0 && totalPlanned > 0)
    ? (totalActual / totalPlanned) / (weightedPct / 100)
    : null;
  const kpi1Status = consumptionRate === null ? null
    : consumptionRate <= 1.10 ? { label: "On Track",         color: "#16a34a", bg: "#f0fdf4" }
    : consumptionRate <= 1.30 ? { label: "Overrun Risk",     color: "#d97706", bg: "#fffbeb" }
    :                           { label: "High Overrun Risk", color: "#dc2626", bg: "#fef2f2" };

  const availablePct = totalPlanned > 0 ? (totalRemaining / totalPlanned) * 100 : 0;
  const kpi2Status = availablePct > 15
    ? { label: "Sufficient",     color: "#16a34a", bg: "#f0fdf4" }
    : availablePct >= 5
    ? { label: "Low Balance",    color: "#d97706", bg: "#fffbeb" }
    : { label: "Budget Overrun", color: "#dc2626", bg: "#fef2f2" };

  const pctConsumed = totalPlanned > 0 ? Math.min((totalActual / totalPlanned) * 100, 100) : 0;
  const gaugeColor  = kpi2Status.color;

  // KPI 3 — Needs PCR: triggered when totalPlanned exceeds latestApprovedBudget by ≥10%
  const pcrOverrun = latestApprovedBudget > 0
    ? (totalPlanned - latestApprovedBudget) / latestApprovedBudget
    : null;
  const pcrStatus = pcrOverrun === null
    ? { label: "Not Set", color: "#6b7280", bg: "#f9fafb", triggered: false, notSet: true }
    : pcrOverrun >= 0.10
    ? { label: "PCR Required",  color: "#dc2626", bg: "#fef2f2", triggered: true,  notSet: false }
    : pcrOverrun >= 0
    ? { label: "Within Budget", color: "#16a34a", bg: "#f0fdf4", triggered: false, notSet: false }
    : { label: "Under Budget",  color: "#16a34a", bg: "#f0fdf4", triggered: false, notSet: false };

  const COST_CAT_MAP: Record<number, string> = {
    847020000: "Staff", 847020001: "Supplies", 847020002: "Equipment",
    847020003: "Contractual", 847020004: "Travel", 847020005: "Indirect",
  };
  const FUNDING_MAP: Record<number, string> = {
    0: "Regular Budget", 1: "Support Account", 2: "xB", 3: "10RCR", 4: "20PCR",
  };

  const byCategory: Record<string, number> = {};
  leaves.forEach(n => {
    const label = n.costCategory != null ? (COST_CAT_MAP[n.costCategory] ?? "Other") : "Unassigned";
    byCategory[label] = (byCategory[label] ?? 0) + (n.totalActualCost ?? 0);
  });

  const byFunding: Record<string, number> = {};
  leaves.forEach(n => {
    const label = n.fundingSource != null ? (FUNDING_MAP[n.fundingSource] ?? "Other") : "Unassigned";
    byFunding[label] = (byFunding[label] ?? 0) + (n.totalPlannedCost ?? 0);
  });

  const DONUT_COLORS = ["#4f46e5","#0f766e","#d97706","#dc2626","#7c3aed","#0284c7","#16a34a","#9ca3af"];
  const catSlices  = Object.entries(byCategory).map(([label, value], i) => ({ label, value, color: DONUT_COLORS[i % DONUT_COLORS.length] }));
  const fundSlices = Object.entries(byFunding).map(([label, value], i) => ({ label, value, color: DONUT_COLORS[i % DONUT_COLORS.length] }));

  const COST_CATS_FULL = [
    { value: 847020000, label: "Staff and Other Personnel Costs" },
    { value: 847020001, label: "Supplies, Commodities, and Materials" },
    { value: 847020002, label: "Equipment, Vehicles, and Furniture" },
    { value: 847020003, label: "Contractual Services" },
    { value: 847020004, label: "Travel" },
    { value: 847020005, label: "Indirect Costs" },
  ];

  type PidRow = { category: string; qty: number; unitCost: number; totalCost: number; funding: string };
  const pidRows: PidRow[] = [];
  COST_CATS_FULL.forEach(cat => {
    const catLeaves = leaves.filter(n => n.costCategory === cat.value);
    if (catLeaves.length === 0) return;
    const fundingGroups: Record<string, TaskNode[]> = {};
    catLeaves.forEach(n => {
      const f = n.fundingSource != null ? (FUNDING_MAP[n.fundingSource] ?? "Other") : "Unassigned";
      if (!fundingGroups[f]) fundingGroups[f] = [];
      fundingGroups[f].push(n);
    });
    Object.entries(fundingGroups).forEach(([funding, tasks]) => {
      const qty      = tasks.reduce((s, n) => s + (n.quantity ?? 0), 0);
      const total    = tasks.reduce((s, n) => s + (n.totalPlannedCost ?? 0), 0);
      const unitCost = qty > 0 ? total / qty : 0;
      pidRows.push({ category: cat.label, qty, unitCost, totalCost: total, funding });
    });
  });
  const pidTotal = pidRows.reduce((s, r) => s + r.totalCost, 0);

function HBarChart({ title, bars, total }: {
    title: string;
    bars: { label: string; value: number; color: string }[];
    total: number;
  }) {
    const maxVal = Math.max(...bars.map(b => b.value), 1);
    return (
      <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: 8, padding: "18px 20px", flex: 1 }}>
        <div style={{ fontSize: 16, fontWeight: 700, color: "#111827", marginBottom: 16 }}>{title}</div>
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          {bars.filter(b => b.value > 0).map((b, i) => {
            const pct = total > 0 ? (b.value / total) * 100 : 0;
            const barPct = (b.value / maxVal) * 100;
            return (
              <div key={i}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 5 }}>
                  <span style={{ fontSize: 13, fontWeight: 600, color: "#111827" }}>{b.label}</span>
                  <div style={{ display: "flex", gap: 10, alignItems: "baseline" }}>
                    <span style={{ fontSize: 13, fontWeight: 700, color: "#111827" }}>{fmtCurrency(b.value)}</span>
                    <span style={{ fontSize: 12, color: "#6b7280", minWidth: 36, textAlign: "right" }}>{pct.toFixed(1)}%</span>
                  </div>
                </div>
                <div style={{ height: 20, background: "#f3f4f6", borderRadius: 4, overflow: "hidden" }}>
                  <div style={{
                    height: "100%", width: `${barPct}%`,
                    background: b.color, borderRadius: 4,
                    transition: "width 0.4s ease",
                  }}/>
                </div>
              </div>
            );
          })}
        </div>
        <div style={{ marginTop: 16, paddingTop: 12, borderTop: "1px solid #f3f4f6", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span style={{ fontSize: 13, fontWeight: 700, color: "#374151" }}>Total</span>
          <span style={{ fontSize: 15, fontWeight: 800, color: "#111827" }}>{fmtCurrency(total)}</span>
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
    const header = "Category\tQty\tUnit Cost\tTotal Cost\tFunding Source";
    const rows = pidRows.map(r =>
      `${r.category}\t${r.qty.toFixed(2)}\t${fmtCurrency(r.unitCost)}\t${fmtCurrency(r.totalCost)}\t${r.funding}`
    );
    const total = `Total\t\t\t${fmtCurrency(pidTotal)}\t`;
    navigator.clipboard.writeText([header, ...rows, total].join("\n")).then(() => {
      setPidCopied(true);
      setTimeout(() => setPidCopied(false), 2000);
    });
  }

  // Gauge — bigger, standalone
  function BigGauge() {
    const r = 110, cx = 145, cy = 130;
    const angle = Math.PI + (pctConsumed / 100) * Math.PI;
    const x1 = cx - r, y1 = cy;
    const x2 = cx + r, y2 = cy;
    const px = cx + r * Math.cos(angle);
    const py = cy + r * Math.sin(angle);
    const largeArc = pctConsumed > 50 ? 1 : 0;
    return (
      <svg width="290" height="162" viewBox="0 0 290 162">
        <path d={`M ${x1} ${y1} A ${r} ${r} 0 1 1 ${x2} ${y2}`}
          fill="none" stroke="#e5e7eb" strokeWidth="20" strokeLinecap="round"/>
        {pctConsumed > 0 && (
          <path d={`M ${x1} ${y1} A ${r} ${r} 0 ${largeArc} 1 ${px} ${py}`}
            fill="none" stroke={gaugeColor} strokeWidth="20" strokeLinecap="round"/>
        )}
        <line x1={cx} y1={cy}
          x2={cx + (r - 16) * Math.cos(angle)}
          y2={cy + (r - 16) * Math.sin(angle)}
          stroke="#1f2937" strokeWidth="3.5" strokeLinecap="round"/>
        <circle cx={cx} cy={cy} r="7" fill="#1f2937"/>
        <text x="8" y="158" fontSize="13" fill="#6b7280" fontFamily="Segoe UI, sans-serif">0%</text>
        <text x="244" y="158" fontSize="13" fill="#6b7280" fontFamily="Segoe UI, sans-serif">100%</text>
        <text x={cx} y={cy - 58} fontSize="13" fill="#374151"
          textAnchor="middle" fontFamily="Segoe UI, sans-serif" fontWeight="700">SPENT</text>
        <text x={cx} y={cy - 16} fontSize="36" fontWeight="800" fill={gaugeColor}
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

      {/* ── Row 1: KPIs + Gauge + Value Cards ───────────────────── */}
<div style={{ display: "flex", gap: 14, padding: "16px 18px", alignItems: "stretch" }}>

        {/* KPI cards */}
        <div style={{ display: "flex", flexDirection: "column", gap: 12, flex: "0 0 250px" }}>

          {/* KPI 3 — Needs PCR */}
          <div style={{ background: pcrStatus.bg, border: `2px solid ${pcrStatus.color}`, borderRadius: 10, padding: "14px 20px" }}>
              <div style={{ fontSize: 15, color: "#111827", fontWeight: 800 }}>Needs PCR?</div>
              <div style={{ marginTop: 8, display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ width: 18, height: 18, borderRadius: "50%", background: pcrStatus.color, flexShrink: 0 }}/>
                <span style={{ fontSize: 18, fontWeight: 800, color: pcrStatus.color }}>{pcrStatus.label}</span>
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
                  Set pmo_latestapprovedbudget on the project record to enable this KPI.
                </div>
              )}
            </div>

          <div style={{ background: kpi1Status?.bg ?? "#f9fafb", border: `2px solid ${kpi1Status?.color ?? "#e5e7eb"}`, borderRadius: 10, padding: "18px 20px", flex: 1 }}>
            <div style={{ fontSize: 15, color: "#111827", fontWeight: 800, letterSpacing: "0.01em" }}>Consumption Rate</div>
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

          <div style={{ background: kpi2Status.bg, border: `2px solid ${kpi2Status.color}`, borderRadius: 10, padding: "18px 20px", flex: 1 }}>
            <div style={{ fontSize: 15, color: "#111827", fontWeight: 800, letterSpacing: "0.01em" }}>Budget Remaining Status</div>
            <div style={{ marginTop: 10, display: "flex", alignItems: "center", gap: 10 }}>
              <div style={{ width: 22, height: 22, borderRadius: "50%", background: kpi2Status.color, flexShrink: 0 }}/>
              <span style={{ fontSize: 20, fontWeight: 800, color: kpi2Status.color }}>
                {kpi2Status.label}
              </span>
            </div>
            <div style={{ fontSize: 14, color: "#111827", marginTop: 10, fontWeight: 600 }}>
              {availablePct.toFixed(1)}% remaining
            </div>
            <div style={{ fontSize: 13, color: "#6b7280", marginTop: 5 }}>
              {">"}15% sufficient · 5–15% low · {"<"}5% overrun
            </div>
          </div>
        </div>

        {/* Gauge — centre */}
        <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", flex: "0 0 300px", background: "white", border: "1px solid #e5e7eb", borderRadius: 10, padding: "16px 8px 12px" }}>
          <div style={{ fontSize: 14, fontWeight: 700, color: "#111827", marginBottom: 4 }}>Budget Consumption (% Spent)</div>
          <BigGauge />
        </div>

        {/* Value cards */}
        <div style={{ display: "flex", flexDirection: "column", gap: 10, flex: 1 }}>
          <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: 10, padding: "18px 20px", flex: 1 }}>
            <div style={{ fontSize: 12, color: "#374151", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em" }}>Total Budget</div>
            <div style={{ fontSize: 28, fontWeight: 800, color: "#111827", marginTop: 8 }}>{fmtCurrency(totalPlanned)}</div>
            <div style={{ fontSize: 13, color: "#6b7280", marginTop: 6 }}>{leaves.length} tasks</div>
          </div>
          <div style={{ background: "white", border: "1px solid #e5e7eb", borderRadius: 10, padding: "18px 20px", flex: 1 }}>
            <div style={{ fontSize: 12, color: "#374151", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em" }}>Spent</div>
            <div style={{ fontSize: 28, fontWeight: 800, color: "#0f766e", marginTop: 8 }}>{fmtCurrency(totalActual)}</div>
            <div style={{ fontSize: 13, color: "#6b7280", marginTop: 6 }}>EV: {fmtCurrency(totalEV)}</div>
          </div>
          <div style={{ background: "white", border: `2px solid ${gaugeColor}`, borderRadius: 10, padding: "18px 20px", flex: 1 }}>
            <div style={{ fontSize: 12, color: "#374151", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em" }}>Available</div>
            <div style={{ fontSize: 28, fontWeight: 800, color: gaugeColor, marginTop: 8 }}>{fmtCurrency(totalRemaining)}</div>
            <div style={{ fontSize: 13, color: "#6b7280", marginTop: 6 }}>{availablePct.toFixed(1)}% of budget</div>
          </div>
        </div>
      </div>

      {/* ── Row 2: Horizontal Bar Charts ────────────────────────── */}
      <div style={{ display: "flex", gap: 14, padding: "0 18px 16px" }}>
        <HBarChart
          title="Actual Spend by Category"
          bars={catSlices.map(s => ({ label: s.label, value: s.value, color: s.color }))}
          total={totalActual}
        />
        <HBarChart
          title="Planned Budget by Funding Source"
          bars={fundSlices.map(s => ({ label: s.label, value: s.value, color: s.color }))}
          total={totalPlanned}
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
              {["Category", "Qty", "Unit Cost", "Total Cost", "Funding"].map(h => (
                <th key={h} style={{ padding: "8px 10px", color: "white", fontWeight: 700, fontSize: 13, textAlign: h === "Category" || h === "Funding" ? "left" : "right", borderRight: "1px solid #2d6fa0", whiteSpace: "nowrap" }}>{h}</th>
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
                </tr>
              ))
            )}
            <tr style={{ background: "white" }}>
              <td colSpan={3} style={{ padding: "9px 10px", borderTop: "2px solid #e5e7eb", fontWeight: 700, fontSize: 13, color: "#111827" }}>Total</td>
              <td style={{ padding: "9px 10px", textAlign: "right", borderTop: "2px solid #e5e7eb", color: "#4f46e5", fontSize: 14, fontWeight: 400 }}>{fmtCurrency(pidTotal)}</td>
              <td style={{ borderTop: "2px solid #e5e7eb" }}/>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
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
    background: #fff; border-bottom: 1px solid #e5e7eb;
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
  .tg-cell-text { font-size: 13px; border-radius: 3px; padding: 1px 3px; }
  .tg-cell-right { text-align: right; display: block; width: 100%; }
  .tg-dash { color: #d1d5db; display: block; text-align: center; }
  .tg-progress-wrap { display: flex; align-items: center; gap: 8px; }
  .tg-progress-track {
    flex: 1; height: 8px; background: #edebe9; border-radius: 2px; overflow: hidden;
  }
  .tg-progress-fill  { height: 100%; border-radius: 2px; transition: width 0.3s; }
  .tg-progress-label { font-size: 12px; color: #605e5c; width: 38px; text-align: right; flex-shrink: 0; }
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
  .tg-filter-panel {
    position: absolute; top: 0; right: 0; bottom: 0;
    width: 300px; background: white; z-index: 100;
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
    font-size: 13px; font-weight: 600; color: #1f2937;
  }
  .tg-filter-section-header:hover { background: #f9fafb; }
  .tg-filter-item {
    display: flex; align-items: center; gap: 8px;
    padding: 6px 16px 6px 24px; cursor: pointer;
    font-size: 13px; color: #374151;
  }
  .tg-filter-item:hover { background: #f3f2f1; }
  .tg-filter-item input[type="checkbox"] {
    width: 14px; height: 14px; cursor: pointer; accent-color: #107c10;
  }
  .tg-filter-count { margin-left: auto; color: #9ca3af; font-size: 12px; }
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
    position: absolute; top: 0; left: 0; bottom: 0;
    width: 960px; background: #f9fafb; z-index: 100;
    border-right: 1px solid #e5e7eb;
    box-shadow: 4px 0 16px rgba(0,0,0,0.10);
    display: flex; flex-direction: column;
    animation: slideInLeft 0.2s ease;
    overflow-y: auto;
  }
  @keyframes slideInLeft {
    from { transform: translateX(-960px); }
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
  const color = pct === 100 ? "#22c55e" : pct > 50 ? "#4f46e5" : "#f59e0b";
  return (
    <div className="tg-progress-wrap">
      <div className="tg-progress-track">
        <div className="tg-progress-fill" style={{ width: `${pct}%`, background: color }} />
      </div>
      <span className="tg-progress-label">{pct.toFixed(1)}%</span>
    </div>
  );
}

const AVATAR_COLORS = [
  "#0078d4","#107c10","#8764b8","#d83b01",
  "#038387","#b4009e","#004e8c","#498205",
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
  return (
    <div className="tg-avatars">
      {resources.map(r => (
        <div key={r.id} className="tg-avatar-row">
          <div className="tg-avatar" style={{ background: getColor(r.name) }}>
            {getInitials(r.name)}
          </div>
          <span className="tg-avatar-name">{r.name}</span>
        </div>
      ))}
    </div>
  );
}

// Currency input with £ symbol prefix
function CurrencyInput({ value, onChange, disabled }: {
  value: number; onChange: (v: number) => void; disabled?: boolean;
}) {
  const [local, setLocal] = React.useState(value === 0 ? "" : String(value));
  React.useEffect(() => { setLocal(value === 0 ? "" : String(value)); }, [value]);

  function commit() {
    const n = parseFloat(local);
    onChange(isNaN(n) ? 0 : n);
  }

return (
    <div className="tg-input-wrap">
      <span className="tg-input-symbol">$</span>
      <input
        className="tg-input"
        type="number" min={0} step="0.01"
        placeholder="0.00"
        value={local}
        disabled={disabled}
        onChange={e => setLocal(e.target.value)}
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
            padding: "4px 8px", fontSize: 13, outline: "none",
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
                padding: "6px 12px", fontSize: 13, cursor: "pointer",
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
                <span style={{ fontSize: 11, color: "#0078d4" }}>
                  {[s.fiscalYearName, s.entityName].filter(Boolean).join(" · ")}
                </span>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 1 }}>
                <span style={{ fontSize: 11, color: "#6b7280" }}>{s.name}</span>
                <span style={{ fontSize: 11, color: "#374151", fontWeight: 500 }}>
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
        onMouseEnter={() => setHovered(true)}
        onMouseLeave={() => setHovered(false)}
        style={{
          display: "flex", alignItems: "center", justifyContent: "space-between",
          border: open ? "2px solid #107c10" : hovered ? "1px solid #107c10" : "1px solid transparent",
          borderRadius: 0, padding: "2px 6px",
          background: open || hovered ? "white" : "transparent", cursor: "pointer", fontSize: 13,
          color: selected ? "#1f2937" : "#9ca3af",
          minHeight: 0, height: "100%",
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

export function TaskGrid({ data: initialData, onSave, onRefresh, userId, taskIds, latestApprovedBudget }: Props) {
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
    return [];
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
      fiscalYearName: null, // FY resolved from SRC name display only
      entityName:     entityItems.find(e => e.id === s.entityId)?.name ?? null,
    }));
  }, [srcItems, entityItems]);

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
      .filter(r => !r.original.isSummary)
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

function applyBulkSRC(srcId: string) {
    const src = srcItems.find(s => s.id === srcId);
    if (!src) return;
    selectedRows.forEach(recordId => {
      const node = table.getRowModel().rows.find(r => r.original.recordId === recordId)?.original;
      if (!node) return;
      applyUpdates(recordId, recalc(node, {
        srcServiceId: src.id, srcServiceName: src.name,
        unitRate: src.price, unit: src.unit, frequency: src.frequency,
      }));
    });
  }

function applyBulkFunding(value: number) {
    selectedRows.forEach(recordId => updateField(recordId, "fundingSource", value));
  }

  function applyBulkCategory(value: number) {
    selectedRows.forEach(recordId => updateField(recordId, "costCategory", value));
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
    header: () => (
      <div style={{ display: "flex", justifyContent: "center", alignItems: "center" }}>
      <input type="checkbox"
        style={{ width: 12, height: 12, cursor: "pointer", accentColor: "#107c10" }}
        checked={selectedRows.length > 0 && getVisibleLeafIds().every(id => selectedRows.includes(id))}
        onChange={e => {
          const ids = getVisibleLeafIds();
          setSelectedRows(e.target.checked ? ids : []);
        }}
      />
      </div>
    ),
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
              width: 9, height: 9,
              border: allSelected ? "2px solid #107c10" : someSelected ? "2px solid #107c10" : "2px solid #d1d5db",
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
            width: 9, height: 9,
            border: isSelected ? "2px solid #107c10" : "2px solid #d1d5db",
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
            color: row.original.pctDone >= 100 ? "#9ca3af" : "inherit",
          }}>{String(getValue() ?? "")}</span>
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
    header: "Assigned to",
    size: 180,
    cell: function AssignedCell({ row }) {
      if (row.original.isSummary) return <span className="tg-dash">—</span>;
      const resources = taskResources[row.original.recordId] ?? [];
      return <ResourceCell resources={resources} />;
    },
  }),

  // ── Cost — input fields hidden on summary rows ────────────────────────────
  col.accessor("fundingSource", {
    header: "Funding source", size: 160,
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
    header: "Cost category", size: 160,
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
  header: "Service", size: 220,
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
  header: "Effort (h)", size: 75,
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
    header: "Completed (h)", size: 90,
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
    header: "Unit rate", size: 95,
    cell: function RateCell({ row }) {
      // Summary: show dash
      if (row.original.isSummary) return <span className="tg-dash">—</span>;
      return (
        <span className="tg-cell-right">
          {fmtCurrency(Number(row.original.unitRate ?? 0))}
        </span>
      );
    },
  }),
  col.accessor("plannedCost", {
    header: "Planned cost", size: 100,
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
    header: "Fixed cost", size: 100,
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
    header: "Total planned", size: 110,
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
    header: "Actual effort cost", size: 120,
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
    header: "Actual fixed cost", size: 120,
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
    header: "Total actual cost", size: 120,
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
    header: "Remaining", size: 100,
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
  header: "Earned value", size: 95,
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
    const all = new Set(DEFAULT_VISIBLE);
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
        ✓ {selectedRows.length} {selectedRows.length === 1 ? "row" : "rows"} selected
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
                    setColumnOrder([]);
                    try { localStorage.removeItem(orderKey); } catch {}
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
                borderBottom: "1px solid #f3f4f6",
                verticalAlign: "middle",
                height: 38,   // ← was 36
                fontWeight: isSummary ? 600 : 400,
                color: isSummary ? "#111827" : "#374151",
                background: isHovered ? "#f3f2f1" : isSummary ? "#fafafa" : "#ffffff",
                };

              return (
                <tr key={row.id}
                  className={selectedRows.includes(row.original.recordId) ? "tg-row-selected" : ""}
                  onMouseEnter={() => setHoveredRow(row.id)}
                  onMouseLeave={() => setHoveredRow(null)}
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
        <span>Double-click any cell to edit</span>
      </div>
    </div>
  );
}