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
}

const COST_CATEGORIES = [
  { value: 847020000, label: "Staff and Other Personnel Costs" },
  { value: 847020001, label: "Supplies, Commodities, and Materials" },
  { value: 847020002, label: "Equipment, Vehicles, and Furniture" },
  { value: 847020003, label: "Contractual Services" },
  { value: 847020004, label: "Travel" },
  { value: 847020005, label: "Indirect Costs" },
];

const ALL_COLUMNS = [
  { id: "startDate",        label: "Start",          group: "schedule" },
  { id: "endDate",          label: "Finish",          group: "schedule" },
  { id: "pctDone",          label: "% Complete",      group: "schedule" },
  { id: "costCategory",     label: "Cost category",   group: "cost" },
  { id: "fiscalYearName",   label: "FY",              group: "cost" },
  { id: "srcServiceName",   label: "Service",         group: "cost" },
  { id: "quantity",         label: "Effort (h)",      group: "cost" },
  { id: "unit",             label: "Unit",            group: "cost" },
  { id: "unitRate",         label: "Unit rate",       group: "cost" },
  { id: "plannedCost",      label: "Planned cost",    group: "cost" },
  { id: "fixedCost",        label: "Fixed cost",      group: "cost" },
  { id: "totalPlannedCost", label: "Total planned",   group: "cost" },
  { id: "actualCost",       label: "Actual cost",     group: "cost" },
  { id: "remainingCost",    label: "Remaining",       group: "cost" },
  { id: "earnedValue",      label: "Earned value",    group: "cost" },
] as const;

const DEFAULT_VISIBLE = new Set(ALL_COLUMNS.map(c => c.id));

interface SrcItem {
  id:          string;
  serviceId:   string;
  name:        string;
  price:       number;
  unit:        string;
  frequency:   string;
  fiscalYearId: string | null;
}

interface FyItem {
  id:   string;
  name: string;
}

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
  }
  .tg-toolbar {
    background: #fff; border-bottom: 1px solid #e5e7eb;
    padding: 0 8px; display: flex; align-items: center;
    gap: 2px; flex-shrink: 0; height: 44px;
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
    position: sticky; top: 0; z-index: 2; white-space: nowrap; user-select: none;
  }
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
    font-size: 13px; border: 1px solid #d1d5db; border-radius: 2px;
    padding: 2px 4px; background: white; width: 100%;
    color: #1f2937; cursor: pointer; max-width: 220px;
  }
  .tg-select:focus { outline: none; border-color: #0078d4; }
  .tg-input-wrap {
    display: flex; align-items: center; justify-content: flex-end;
    border: 1px solid #d1d5db; border-radius: 2px; background: white;
    overflow: hidden; width: 110px; margin-left: auto;
  }
  .tg-input-symbol {
    padding: 2px 5px; background: #f3f4f6; color: #605e5c;
    font-size: 13px; border-right: 1px solid #e5e7eb; flex-shrink: 0;
  }
  .tg-input {
    font-size: 13px; border: none; padding: 2px 6px;
    background: white; width: 100%; text-align: right; color: #1f2937;
    outline: none;
  }
  .tg-footer {
    padding: 4px 12px; background: #fff; border-top: 1px solid #e5e7eb;
    font-size: 11px; color: #a19f9d; display: flex;
    justify-content: space-between; flex-shrink: 0;
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
  const [open, setOpen]     = React.useState(false);
  const [query, setQuery]   = React.useState("");
  const triggerRef          = React.useRef<HTMLDivElement>(null);
  const [dropPos, setDropPos] = React.useState({ top: 0, left: 0, width: 0 });

  // Position the dropdown based on trigger position
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
    : items.filter(s =>
        s.serviceId.toLowerCase().includes(query.toLowerCase()) ||
        s.name.toLowerCase().includes(query.toLowerCase())
      );

  // Portal dropdown — rendered outside scroll container
  const dropdown = open ? (
    <div
      id="tg-service-portal"
      style={{
        position: "absolute",
        top:      dropPos.top,
        left:     dropPos.left,
        width:    Math.max(dropPos.width, 280),
        background: "white",
        border: "1px solid #e5e7eb",
        borderRadius: 4,
        boxShadow: "0 4px 16px rgba(0,0,0,0.14)",
        zIndex: 99999,
        maxHeight: 300,
        display: "flex",
        flexDirection: "column",
      }}
    >
      <div style={{ padding: "6px 8px", borderBottom: "1px solid #f3f4f6" }}>
        <input
          autoFocus
          placeholder="Search services..."
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
      <div style={{ overflowY: "auto", maxHeight: 240 }}>
        {/* Clear option */}
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
                if (s.id !== value) (e.currentTarget as HTMLDivElement).style.background =
                  s.id === value ? "#eff6ff" : "white";
              }}
            >
              <div style={{ fontWeight: 500, fontSize: 12 }}>{s.serviceId}</div>
              <div style={{ fontSize: 11, color: "#6b7280", marginTop: 1 }}>{s.name}</div>
            </div>
          ))
        )}
      </div>
    </div>
  ) : null;

  return (
    <React.Fragment>
      {/* Trigger button */}
      <div
        ref={triggerRef}
        onClick={openDropdown}
        style={{
          display: "flex", alignItems: "center", justifyContent: "space-between",
          border: "1px solid #d1d5db", borderRadius: 2, padding: "2px 6px",
          background: "white", cursor: "pointer", fontSize: 13,
          color: selected ? "#1f2937" : "#9ca3af", minHeight: 26,
          userSelect: "none", width: "100%", maxWidth: 260,
          boxSizing: "border-box",
        }}
      >
        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1 }}>
          {selected ? `${selected.serviceId} – ${selected.name}` : "— select service —"}
        </span>
        <svg width="12" height="12" viewBox="0 0 24 24" fill="none"
          stroke="#6b7280" strokeWidth="2.5" style={{ flexShrink: 0, marginLeft: 4 }}>
          <polyline points="6 9 12 15 18 9"/>
        </svg>
      </div>

      {/* Portal — appended to document.body to escape overflow:hidden */}
      {dropdown && typeof document !== "undefined"
        ? ReactDOM.createPortal(dropdown, document.body)
        : null}
    </React.Fragment>
  );
}

export function TaskGrid({ data: initialData, onSave, onRefresh, userId }: Props) {
  const [data, setData]             = React.useState<TaskNode[]>(initialData);
  //const [expanded, setExpanded]     = React.useState<ExpandedState>({ "0": true });
  const [allExpanded, setAllExpanded] = React.useState(false);
  const [expanded, setExpanded] = React.useState<ExpandedState>({});
  const [pending, setPending]       = React.useState<Record<string, Partial<TaskNode>>>({});
  const [saving, setSaving]         = React.useState(false);
  const [savedMsg, setSavedMsg]     = React.useState(false);
  const [hoveredRow, setHoveredRow] = React.useState<string | null>(null);
  const [srcItems, setSrcItems]     = React.useState<SrcItem[]>([]);
  const [fyItems, setFyItems]       = React.useState<FyItem[]>([]);
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
  const [colPanelOpen, setColPanelOpen]   = React.useState(false);
  const [colPanelPos,  setColPanelPos]    = React.useState({ top: 0, left: 0 });
  const colBtnRef = React.useRef<HTMLButtonElement>(null);

  //React.useEffect(() => { setData(initialData); }, [initialData]);
    // Expand first level on initial load
    React.useEffect(() => {
    if (initialData.length === 0) return;
    const firstLevel: Record<string, boolean> = {};
    // Root rows in react-table are indexed 0, 1, 2...
    initialData.forEach((_, i) => { firstLevel[String(i)] = true; });
    setExpanded(firstLevel);
    }, [initialData.length > 0]);

  React.useEffect(() => {
    // Load Fiscal Years
    fetch("/api/data/v9.2/pmo_fiscalyear1s?$select=pmo_fiscalyear1id,pmo_id&$orderby=pmo_id asc")
      .then(r => {
        if (!r.ok) throw new Error(`FY load failed: ${r.status} ${r.statusText}`);
        return r.json();
      })
      .then(d => {
        const items: FyItem[] = (d.value || []).map((r: any) => {
            // Try both casings defensively
            const id   = r.pmo_fiscalyear1id || r.pmo_FiscalYear1Id || "";
            const name = r.pmo_id            || r.pmo_ID            || "";
            return { id, name };
        });
        setFyItems(items);
        })
      .catch(e => {
        console.error("[TaskGrid] FY load error:", e);
        setLoadError("Could not load Fiscal Years: " + e.message);
      });

    // Load SRC — include fiscal year lookup
    fetch("/api/data/v9.2/pmo_serviceratecards?$select=pmo_serviceratecardid,pmo_serviceid,pmo_servicename,pmo_price,pmo_unit,pmo_frequency&$orderby=pmo_serviceid asc&$top=500")
      .then(r => {
        if (!r.ok) throw new Error(`SRC load failed: ${r.status} ${r.statusText}`);
        return r.json();
      })
      .then(d => {
        const items: SrcItem[] = (d.value || []).map((r: any) => ({
        id:           r.pmo_serviceratecardid,
        serviceId:    r.pmo_serviceid,
        name:         r.pmo_servicename,
        price:        r.pmo_price,
        unit:         r.pmo_unit,
        frequency:    r.pmo_frequency,
        fiscalYearId: null,
        }));
        setSrcItems(items);
      })
      .catch(e => {
        console.error("[TaskGrid] SRC load error:", e);
        setLoadError("Could not load Service Rate Card: " + e.message);
      });
  }, []);

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
    const qty      = overrides.quantity         ?? row.quantity         ?? 0;
    const rate     = overrides.unitRate         ?? row.unitRate         ?? 0;
    const fixed    = overrides.fixedCost        ?? row.fixedCost        ?? 0;
    const actual   = overrides.actualCost       ?? row.actualCost       ?? 0;
    const pct      = overrides.pctDone          ?? row.pctDone          ?? 0;
    const planned  = qty * rate;
    const total    = planned + fixed;
    const remaining = total - actual;
    const ev       = (pct / 100) * total;
    return {
      ...overrides,
      plannedCost:      planned,
      totalPlannedCost: total,
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
    // Find the latest version of this row from current state
    function findNode(nodes: TaskNode[]): TaskNode | null {
      for (const n of nodes) {
        if (n.recordId === row.recordId) return n;
        if (n.subRows) { const f = findNode(n.subRows); if (f) return f; }
      }
      return null;
    }
    const current  = findNode(prev) ?? row;
    const planned  = current.plannedCost  ?? 0;
    const actual   = current.actualCost   ?? 0;
    const total    = planned + fixed;
    const remaining = total - actual;
    const ev       = (current.pctDone / 100) * total;

    const updates: Partial<TaskNode> = {
      fixedCost: fixed, totalPlannedCost: total,
      remainingCost: remaining, earnedValue: ev,
    };
    setPending(p => ({ ...p, [row.recordId]: { ...(p[row.recordId] ?? {}), ...updates } }));
    let next = prev;
    for (const [field, val] of Object.entries(updates)) {
      next = updateNodeInTree(next, row.recordId, field as keyof TaskNode, val);
    }
    return next;
  });
}

function onActualCostChange(row: TaskNode, actual: number) {
  setData(prev => {
    function findNode(nodes: TaskNode[]): TaskNode | null {
      for (const n of nodes) {
        if (n.recordId === row.recordId) return n;
        if (n.subRows) { const f = findNode(n.subRows); if (f) return f; }
      }
      return null;
    }
    const current   = findNode(prev) ?? row;
    const total     = current.totalPlannedCost ?? 0;
    const remaining = total - actual;
    const ev        = (current.pctDone / 100) * total;

    const updates: Partial<TaskNode> = {
      actualCost: actual, remainingCost: remaining, earnedValue: ev,
    };
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
          <span className="tg-cell-text">{String(getValue() ?? "")}</span>
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

  // ── Cost — input fields hidden on summary rows ────────────────────────────
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
  col.accessor("fiscalYearName", {
    header: "FY", size: 70,
    cell: function FyCell({ row }) {
      // Summary: show dash
      if (row.original.isSummary) return <span className="tg-dash">—</span>;
      return (
        <select className="tg-select" style={{ maxWidth: 80 }}
          value={row.original.fiscalYearId ?? ""}
          onChange={e => {
            const id = e.target.value;
            const fy = fyItems.find(f => f.id === id);
            applyUpdates(row.original.recordId, {
              fiscalYearId:   id,
              fiscalYearName: fy?.name ?? "",
            });
          }}>
          <option value="">—</option>
          {fyItems.map(f => (
            <option key={f.id} value={f.id}>{f.name}</option>
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
        items={srcItems}
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
    header: "Actual cost", size: 100,
    cell: function ActualCell({ row }) {
      // Summary: show rolled-up total, read-only
      if (row.original.isSummary) {
        return (
          <span className="tg-cell-right" style={{ fontWeight: 600 }}>
            {fmtCurrency(row.original.actualCost)}
          </span>
        );
      }
      return <CurrencyInput value={row.original.actualCost}
        onChange={v => onActualCostChange(row.original, v)} />;
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
], [srcItems, fyItems]);

const columnVisibility = React.useMemo(() => {
    const vis: Record<string, boolean> = {};
    ALL_COLUMNS.forEach(c => { vis[c.id] = visibleCols.has(c.id); });
    return vis;
  }, [visibleCols]);

  const table = useReactTable({
    data, columns,
    state: { expanded, columnVisibility },
    onExpandedChange: setExpanded,
    getSubRows:          row => row.subRows,
    getCoreRowModel:     getCoreRowModel(),
    getExpandedRowModel: getExpandedRowModel(),
  });

  const changesCount = Object.keys(pending).length;

  async function handleSave() {
    setSaving(true);
    try {
      await onSave(pending);
      setPending({});
      setSavedMsg(true);
      setTimeout(() => setSavedMsg(false), 2500);
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

  const rightCols = new Set(["plannedCost","fixedCost","totalPlannedCost",
    "actualCost","remainingCost","earnedValue","unitRate","quantity"]);

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
            <button onClick={() => setColPanelOpen(false)}>Close</button>
          </div>
        </div>,
        document.body
      )}
</div>

      {loadError && <div className="tg-error">⚠ {loadError}</div>}

      <div className="tg-scroll">
        <table className="tg-table">
          <thead>
            {table.getHeaderGroups().map(hg => (
              <tr key={hg.id}>
                {hg.headers.map(h => (
                  <th key={h.id}
                    className={rightCols.has(h.column.id) ? "th-right" : ""}
                    style={{ width: h.getSize(), minWidth: h.getSize() }}>
                    {flexRender(h.column.columnDef.header, h.getContext())}
                  </th>
                ))}
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