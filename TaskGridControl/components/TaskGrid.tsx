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
  data:      TaskNode[];
  onSave:    (changes: Record<string, Partial<TaskNode>>) => Promise<void>;
  onRefresh: () => void;
  userId:    string;
  taskIds:   string[];
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
    padding: 0 8px; display: flex; align-items: center;
    gap: 2px; flex-shrink: 0; height: 44px;
    position: sticky; top: 0; z-index: 10;
  }
  .tg-title {
    font-weight: 600; font-size: 13px; color: #1f2937;
    display: flex; align-items: center; gap: 6px; margin-right: 8px;
  }
  .tg-divider { width: 1px; height: 16px; background: #e5e7eb; margin: 0 6px; }
  .tg-btn {
    display: inline-flex; align-items: center; gap: 4px;
    padding: 4px 12px; border-radius: 4px; border: none;
    font-size: 14px; cursor: pointer; font-weight: 400;
    background: transparent; color: #323130; transition: background 0.1s;
    height: 32px; white-space: nowrap;
  }
  .tg-btn:hover         { background: #edebe9; }
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

export function TaskGrid({ data: initialData, onSave, onRefresh, userId, taskIds }: Props) {
  const [data, setData]             = React.useState<TaskNode[]>(initialData);
  //const [expanded, setExpanded]     = React.useState<ExpandedState>({ "0": true });
  const [allExpanded, setAllExpanded] = React.useState(false);
  const [expanded, setExpanded] = React.useState<ExpandedState>({});
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

    React.useEffect(() => {
        if (Object.keys(pending).length === 0) {
        setData(initialData);
        setRefreshCount(c => c + 1);
        }
    }, [initialData]);

    // Expand first TWO levels on initial load
    React.useEffect(() => {
      if (initialData.length === 0) return;
      const twoLevels: Record<string, boolean> = {};
      // Root rows are "0", "1", etc. Their children are "0.0", "0.1", "1.0", etc.
      initialData.forEach((rootNode, i) => {
        twoLevels[String(i)] = true;                          // level 1
        rootNode.subRows?.forEach((_, j) => {
          twoLevels[`${i}.${j}`] = true;                     // level 2
        });
      });
      setExpanded(twoLevels);
    }, [initialData.length > 0]);

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

  const columns = React.useMemo(() => [

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
], [resolvedSrcItems, taskResources]);

const columnVisibility = React.useMemo(() => {
    const vis: Record<string, boolean> = {};
    ALL_COLUMNS.forEach(c => { vis[c.id] = visibleCols.has(c.id); });
    return vis;
  }, [visibleCols]);

const table = useReactTable({
    data, columns,
    state: { expanded, columnVisibility, columnOrder },
    onExpandedChange:     setExpanded,
    onColumnOrderChange:  setColumnOrder,
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
        if (dragId === dropId || dragId === "taskName") return;
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
        setColumnOrder(next);
        try { localStorage.setItem(orderKey, JSON.stringify(next)); } catch {}
    }

    const rightCols = new Set(["plannedCost","fixedCost","totalPlannedCost",
    "actualCost","actualFixedCost","totalActualCost","remainingCost","earnedValue","unitRate","quantity"]);

  return (
    <div className="tg-wrap">
      <style>{CSS}</style>

      <div className="tg-toolbar">
  <span className="tg-title">
    <GridIcon />Task Grid
  </span>
  <div className="tg-divider" />
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
  <div style={{ flex: 1 }} />
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
                <button onClick={() => setColPanelOpen(false)}>Close</button>
            </div>
        </div>,
        document.body
      )}
</div>

      {loadError && <div className="tg-error">⚠ {loadError}</div>}

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
                        if (isLocked) return;
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
                    </th>
                  );
                })}
              </tr>
            ))}
          </thead>
          <tbody>
            {table.getRowModel().rows.map(row => {
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
                  onMouseEnter={() => setHoveredRow(row.id)}
                  onMouseLeave={() => setHoveredRow(null)}>
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