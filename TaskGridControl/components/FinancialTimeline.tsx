/**
 * FinancialTimeline.tsx
 *
 * Collapsible budget-phasing panel below the task grid.
 *
 * EFFORT COST  → derived from msdyn_plannedwork (scheduling engine daily buckets)
 *                Falls back to proportional calendar-day split when null.
 * FIXED COST   → always proportional by calendar days.
 *
 * HOW TO INTEGRATE
 * ─────────────────────────────────────────────────
 *   <FinancialTimeline
 *     data={data}
 *     isLeafVisible={isLeafVisible}
 *     hasActiveFilters={totalActiveFilters > 0}
 *     onClose={() => setTimelineOpen(false)}
 *     projectId={projectId}
 *     taskIds={taskIds}
 *   />
 */

import * as React from "react";
import { TaskNode } from "./buildTree";

// ─── Brand palette ────────────────────────────────────────────────────────────
const C = {
  planned:      "#5C9057",
  actual:       "#234A21",
  remaining:    "#709D6A",
  cumPlanned:   "#1E3C1C",
  cumActual:    "#97B891",
  nowBg:        "#D2E1D0",
  nowText:      "#234A21",
  nowMarker:    "#2D672A",
  totalBg:      "#EBF2EA",
  totalBorder:  "#5C9057",
  totalText:    "#1E3C1C",
  headerBorder: "#ABC6A6",
  rowDivider:   "#F0F5EF",
  sepLine:      "#BED3BA",
  panelBg:      "#FAFCFA",
  pillActiveBg: "#EBF2EA",
  negativeText: "#C0392B",
  negativeBg:   "#FEF2F2",
};

// ─── Types ────────────────────────────────────────────────────────────────────

type Granularity = "monthly" | "quarterly" | "yearly";
type RowKey      = "planned" | "actual" | "remaining" | "cumPlanned" | "cumActual";

interface PeriodData {
  label:   string;
  isoKey:  string;
  isToday: boolean;
  isPast:  boolean;
}

interface WorkSlot {
  start: Date;
  end:   Date;
  hours: number;
}

interface AssignmentData {
  taskId:      string;
  plannedWork: WorkSlot[] | null;
  effortTotal: number;
  resourceIds: string[];  // bookable resource IDs assigned to this task
}

interface Props {
  data:             TaskNode[];
  isLeafVisible:    (node: TaskNode) => boolean;
  hasActiveFilters: boolean;
  onClose:          () => void;
  projectId:        string;
  taskIds:          string[];
}

// ─── Row definitions ──────────────────────────────────────────────────────────

const ROW_DEFS: { key: RowKey; label: string }[] = [
  { key: "planned",    label: "Planned"            },
  { key: "actual",     label: "Actual"             },
  { key: "remaining",  label: "Remaining"          },
  { key: "cumPlanned", label: "Cumulative Planned"  },
  { key: "cumActual",  label: "Cumulative Actual"   },
];

const ROW_COLORS: Record<RowKey, string> = {
  planned:    C.planned,
  actual:     C.actual,
  remaining:  C.remaining,
  cumPlanned: C.cumPlanned,
  cumActual:  C.cumActual,
};

const DEFAULT_VISIBLE_ROWS = new Set<RowKey>(
  ["planned", "actual", "remaining", "cumPlanned", "cumActual"]
);

// ─── Helpers ──────────────────────────────────────────────────────────────────

function fmtCurrency(v: number): string {
  if (v === 0) return "—";
  return new Intl.NumberFormat("en-US", {
    style: "currency", currency: "USD",
    minimumFractionDigits: 0, maximumFractionDigits: 0,
  }).format(v);
}

function parseDate(raw: string | null | undefined): Date | null {
  if (!raw) return null;
  const d = new Date(raw);
  return isNaN(d.getTime()) ? null : d;
}

function parseDvDate(raw: string): Date {
  const ms = parseInt(raw.replace(/\/Date\((\d+)\)\//, "$1"), 10);
  return new Date(ms);
}

function parsePlannedWork(raw: string | null): WorkSlot[] | null {
  if (!raw) return null;
  try {
    const arr = JSON.parse(raw) as { Start: string; End: string; Hours: number }[];
    return arr.map(s => ({
      start: parseDvDate(s.Start),
      end:   parseDvDate(s.End),
      hours: s.Hours,
    }));
  } catch {
    return null;
  }
}

function periodBounds(key: string, granularity: Granularity): { from: Date; to: Date } {
  if (granularity === "yearly") {
    const y = parseInt(key, 10);
    return { from: new Date(y, 0, 1), to: new Date(y + 1, 0, 1) };
  }
  if (granularity === "quarterly") {
    const [y, q]     = key.split("-Q").map(Number);
    const startMonth = (q - 1) * 3;
    return { from: new Date(y, startMonth, 1), to: new Date(y, startMonth + 3, 1) };
  }
  const [y, m] = key.split("-").map(Number);
  return { from: new Date(y, m - 1, 1), to: new Date(y, m, 1) };
}

function periodKey(date: Date, granularity: Granularity): string {
  const y = date.getFullYear();
  const m = date.getMonth();
  if (granularity === "yearly")    return String(y);
  if (granularity === "quarterly") return `${y}-Q${Math.floor(m / 3) + 1}`;
  return `${y}-${String(m + 1).padStart(2, "0")}`;
}

function periodLabel(key: string, granularity: Granularity): string {
  if (granularity === "yearly")    return key;
  if (granularity === "quarterly") return key.replace("-", " ");
  const [y, m] = key.split("-").map(Number);
  return new Date(y, m - 1, 1).toLocaleDateString("en-GB", { month: "short", year: "2-digit" });
}

function flatLeaves(nodes: TaskNode[]): TaskNode[] {
  const out: TaskNode[] = [];
  function walk(n: TaskNode) {
    if (!n.subRows || n.subRows.length === 0) { out.push(n); return; }
    n.subRows.forEach(walk);
  }
  nodes.forEach(walk);
  return out;
}

/**
 * Distribute a task's costs across periods.
 * Returns 6 arrays — planned/actual totals + their effort/fixed breakdowns.
 * Planned is NEVER affected by pct or actual fields.
 */
function distributeCost(
  task:        TaskNode,
  periods:     PeriodData[],
  granularity: Granularity,
  assignment:  AssignmentData | null
): {
  plannedByPeriod:       number[];
  plannedEffortByPeriod: number[];
  plannedFixedByPeriod:  number[];
  actualByPeriod:        number[];
  actualEffortByPeriod:  number[];
  actualFixedByPeriod:   number[];
} {
  const n = periods.length;
  const plannedByPeriod       = new Array(n).fill(0);
  const plannedEffortByPeriod = new Array(n).fill(0);
  const plannedFixedByPeriod  = new Array(n).fill(0);
  const actualByPeriod        = new Array(n).fill(0);
  const actualEffortByPeriod  = new Array(n).fill(0);
  const actualFixedByPeriod   = new Array(n).fill(0);

  // Use fresh Date objects — never mutate shared references
  const startRaw = parseDate(task.startDate);
  const endRaw   = parseDate(task.endDate);

  if (!startRaw || !endRaw) {
    return { plannedByPeriod, plannedEffortByPeriod, plannedFixedByPeriod,
             actualByPeriod, actualEffortByPeriod, actualFixedByPeriod };
  }

  // Snap to day boundaries using NEW Date objects
  const start = new Date(startRaw);
  const end   = new Date(endRaw);
  start.setHours(0, 0, 0, 0);
  end.setHours(23, 59, 59, 999);

  const taskSpan   = Math.max(1, end.getTime() - start.getTime());
  const bounds     = periods.map(p => periodBounds(p.isoKey, granularity));
  const unitRate   = task.unitRate  ?? 0;
  const fixedCost  = task.fixedCost ?? 0;
  const planned    = task.totalPlannedCost ?? 0;

  // ── PLANNED: Fixed cost — always proportional ─────────────────────────────
  if (fixedCost > 0) {
    periods.forEach((_, i) => {
      const { from, to } = bounds[i];
      const pFrom = Math.max(start.getTime(), from.getTime());
      const pTo   = Math.min(end.getTime(),   to.getTime());
      if (pFrom <= pTo) {
        const v = ((pTo - pFrom) / taskSpan) * fixedCost;
        plannedByPeriod[i]      += v;
        plannedFixedByPeriod[i] += v;
      }
    });
  }

  // ── PLANNED: Effort cost — exact slots or proportional fallback ───────────
  const plannedWork = assignment?.plannedWork ?? null;
  if (plannedWork && plannedWork.length > 0 && unitRate > 0) {
    plannedWork.forEach(slot => {
      const slotKey = periodKey(slot.start, granularity);
      const idx     = periods.findIndex(p => p.isoKey === slotKey);
      if (idx !== -1) {
        const v = slot.hours * unitRate;
        plannedByPeriod[idx]       += v;
        plannedEffortByPeriod[idx] += v;
      }
    });
  } else {
    const effortCost = planned - fixedCost;
    if (effortCost > 0) {
      periods.forEach((_, i) => {
        const { from, to } = bounds[i];
        const pFrom = Math.max(start.getTime(), from.getTime());
        const pTo   = Math.min(end.getTime(),   to.getTime());
        if (pFrom <= pTo) {
          const v = ((pTo - pFrom) / taskSpan) * effortCost;
          plannedByPeriod[i]       += v;
          plannedEffortByPeriod[i] += v;
        }
      });
    }
  }

  // ── ACTUAL: distribute actualCost + actualFixedCost proportionally ────────
  // Uses % complete to determine how far through the task actuals have occurred.
  // actualCost and actualFixedCost come directly from TaskNode — no date mutation.
  const pct         = Math.min(100, Math.max(0, task.pctDone ?? 0)) / 100;
  const actualEffort = task.actualCost      ?? 0;  // effortCompleted × unitRate
  const actualFixed  = task.actualFixedCost ?? 0;
  const totalActual  = actualEffort + actualFixed;

  if (totalActual > 0 && pct > 0) {
    // Actual end = start + pct × taskSpan (fresh calculation, no mutation)
    const actualEndMs  = start.getTime() + pct * taskSpan;
    const actualSpanMs = Math.max(1, actualEndMs - start.getTime());

    periods.forEach((_, i) => {
      const { from, to } = bounds[i];
      const aFrom = Math.max(start.getTime(), from.getTime());
      const aTo   = Math.min(actualEndMs,     to.getTime());
      if (aFrom < aTo) {
        const ratio = (aTo - aFrom) / actualSpanMs;
        actualByPeriod[i]       += ratio * totalActual;
        actualEffortByPeriod[i] += ratio * actualEffort;
        actualFixedByPeriod[i]  += ratio * actualFixed;
      }
    });
  }

  return {
    plannedByPeriod, plannedEffortByPeriod, plannedFixedByPeriod,
    actualByPeriod,  actualEffortByPeriod,  actualFixedByPeriod,
  };
}

// ─── Component ────────────────────────────────────────────────────────────────

export function FinancialTimeline({
  data, isLeafVisible, hasActiveFilters, onClose, projectId, taskIds,
}: Props) {
  const [granularity, setGranularity]         = React.useState<Granularity>("monthly");
  const [visibleRows, setVisibleRows]         = React.useState<Set<RowKey>>(new Set(DEFAULT_VISIBLE_ROWS));
  const [height, setHeight]                   = React.useState(260);
  const [plannedExpanded, setPlannedExpanded]     = React.useState(false);
  const [actualExpanded, setActualExpanded]       = React.useState(false);
  const [remainingExpanded, setRemainingExpanded] = React.useState(false);

  const [assignments, setAssignments]         = React.useState<Map<string, AssignmentData>>(new Map());
  const [loadingAssignments, setLoadingAssignments] = React.useState(false);
  const [exactCount, setExactCount]           = React.useState(0);

  const resizerRef = React.useRef<HTMLDivElement>(null);

  // ── Fetch msdyn_plannedwork ───────────────────────────────────────────────
  React.useEffect(() => {
    if (!taskIds || taskIds.length === 0) return;
    setLoadingAssignments(true);
    const chunk  = taskIds.slice(0, 80);
    const filter = chunk.map(id => `_msdyn_taskid_value eq ${id}`).join(" or ");

    fetch(
      `/api/data/v9.2/msdyn_resourceassignments` +
      `?$select=_msdyn_taskid_value,msdyn_plannedwork,msdyn_effort,_msdyn_bookableresourceid_value` +
      `&$filter=${encodeURIComponent(filter)}&$top=500`
    )
      .then(r => r.ok ? r.json() : { value: [] })
      .then(d => {
        const map = new Map<string, AssignmentData>();
        let exact = 0;
        (d.value || []).forEach((a: any) => {
          const taskId      = a._msdyn_taskid_value as string;
          const rawWork     = a.msdyn_plannedwork   as string | null;
          const effortTotal = a.msdyn_effort        as number ?? 0;
          const plannedWork = parsePlannedWork(rawWork);
          if (plannedWork) exact++;
          const resourceId = a._msdyn_bookableresourceid_value as string | null;
          if (map.has(taskId)) {
            const ex = map.get(taskId)!;
            if (plannedWork && ex.plannedWork) ex.plannedWork.push(...plannedWork);
            else if (plannedWork && !ex.plannedWork) ex.plannedWork = plannedWork;
            ex.effortTotal += effortTotal;
            if (resourceId && !ex.resourceIds.includes(resourceId)) {
              ex.resourceIds.push(resourceId);
            }
          } else {
            map.set(taskId, { taskId, plannedWork, effortTotal, resourceIds: resourceId ? [resourceId] : [] });
          }
        });
        setAssignments(map);
        setExactCount(exact);
        setLoadingAssignments(false);
      })
      .catch(() => setLoadingAssignments(false));
  }, [taskIds.join(",")]);

  // ── Drag-to-resize ────────────────────────────────────────────────────────
  function onResizerMouseDown(e: React.MouseEvent) {
    e.preventDefault();
    const startY = e.clientY;
    const startH = height;
    function onMove(ev: MouseEvent) {
      setHeight(Math.min(600, Math.max(140, startH + (startY - ev.clientY))));
    }
    function onUp() {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
    }
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  }

  // ── Leaves ────────────────────────────────────────────────────────────────
  const leaves = React.useMemo(() => {
    const all = flatLeaves(data);
    return hasActiveFilters ? all.filter(isLeafVisible) : all;
  }, [data, isLeafVisible, hasActiveFilters]);

  // ── Periods ───────────────────────────────────────────────────────────────
  const periods = React.useMemo((): PeriodData[] => {
    if (leaves.length === 0) return [];
    let minDate: Date | null = null;
    let maxDate: Date | null = null;
    leaves.forEach(n => {
      const s = parseDate(n.startDate);
      const e = parseDate(n.endDate);
      if (s && (!minDate || s < minDate)) minDate = s;
      if (e && (!maxDate || e > maxDate)) maxDate = e;
    });
    if (!minDate || !maxDate) return [];
    const resolvedMin: Date = minDate;
    const resolvedMax: Date = maxDate;
    const today    = new Date();
    const todayKey = periodKey(today, granularity);
    const keys: string[] = [];
    const cursor = new Date(resolvedMin);
    if (granularity === "monthly")   cursor.setDate(1);
    if (granularity === "quarterly") { cursor.setDate(1); cursor.setMonth(Math.floor(cursor.getMonth() / 3) * 3); }
    if (granularity === "yearly")    { cursor.setMonth(0); cursor.setDate(1); }
    let i = 0;
    while (cursor <= resolvedMax && i++ < 200) {
      keys.push(periodKey(cursor, granularity));
      if (granularity === "monthly")   cursor.setMonth(cursor.getMonth() + 1);
      if (granularity === "quarterly") cursor.setMonth(cursor.getMonth() + 3);
      if (granularity === "yearly")    cursor.setFullYear(cursor.getFullYear() + 1);
    }
    return [...new Set(keys)].map(key => {
      const { from } = periodBounds(key, granularity);
      return { label: periodLabel(key, granularity), isoKey: key,
               isToday: key === todayKey, isPast: from < today && key !== todayKey };
    });
  }, [leaves, granularity]);

  // ── Aggregate ─────────────────────────────────────────────────────────────
  const agg = React.useMemo(() => {
    const pp  = new Array(periods.length).fill(0); // planned total
    const pe  = new Array(periods.length).fill(0); // planned effort
    const pf  = new Array(periods.length).fill(0); // planned fixed
    const pa  = new Array(periods.length).fill(0); // actual total
    const ae  = new Array(periods.length).fill(0); // actual effort
    const af  = new Array(periods.length).fill(0); // actual fixed

    leaves.forEach(task => {
      const assignment = assignments.get(task.recordId) ?? null;
      // Get working days for this task's resources — use first resource's calendar
      const r = distributeCost(task, periods, granularity, assignment);
      r.plannedByPeriod.forEach((v, i)       => { pp[i] += v; });
      r.plannedEffortByPeriod.forEach((v, i) => { pe[i] += v; });
      r.plannedFixedByPeriod.forEach((v, i)  => { pf[i] += v; });
      r.actualByPeriod.forEach((v, i)        => { pa[i] += v; });
      r.actualEffortByPeriod.forEach((v, i)  => { ae[i] += v; });
      r.actualFixedByPeriod.forEach((v, i)   => { af[i] += v; });
    });

    return { pp, pe, pf, pa, ae, af };
  }, [leaves, periods, granularity, assignments]);

  // ── Derived totals ────────────────────────────────────────────────────────
  const totals = React.useMemo(() => {
    const totalPlanned        = agg.pp.reduce((s, v) => s + v, 0);
    const totalActual         = agg.pa.reduce((s, v) => s + v, 0);
    const totalPlannedEffort  = agg.pe.reduce((s, v) => s + v, 0);
    const totalPlannedFixed   = agg.pf.reduce((s, v) => s + v, 0);
    const totalActualEffort   = agg.ae.reduce((s, v) => s + v, 0);
    const totalActualFixed    = agg.af.reduce((s, v) => s + v, 0);
    const totalRemaining      = totalPlanned - totalActual;
    const totalRemEffort      = totalPlannedEffort - totalActualEffort;
    const totalRemFixed       = totalPlannedFixed  - totalActualFixed;
    let cumP = 0, cumA = 0;
    const cumPlanned  = agg.pp.map(v => { cumP += v; return cumP; });
    const cumActual   = agg.pa.map(v => { cumA += v; return cumA; });
    const remaining   = agg.pp.map((v, i) => v - agg.pa[i]);
    const remEffort   = agg.pe.map((v, i) => v - agg.ae[i]);
    const remFixed    = agg.pf.map((v, i) => v - agg.af[i]);
    return {
      totalPlanned, totalActual, totalRemaining,
      totalPlannedEffort, totalPlannedFixed,
      totalActualEffort, totalActualFixed,
      totalRemEffort, totalRemFixed,
      cumPlanned, cumActual, remaining, remEffort, remFixed,
    };
  }, [agg]);

  // ── Styling helpers ───────────────────────────────────────────────────────
  const maxPlanned = Math.max(...agg.pp, 1);

  function cellBg(value: number, isToday: boolean): string {
    if (value <= 0.5) return isToday ? C.nowBg : "transparent";
    const ratio = Math.min(1, value / maxPlanned);
    const alpha = 0.08 + ratio * 0.22;
    return isToday
      ? `rgba(92, 144, 87, ${(alpha + 0.08).toFixed(2)})`
      : `rgba(92, 144, 87, ${alpha.toFixed(2)})`;
  }

  function cellStyle(value: number, isNegative = false, isToday = false, overrideColor?: string): React.CSSProperties {
    const negative = isNegative && value < 0;
    return {
      textAlign:   "right", padding: "6px 12px", fontSize: 14, fontWeight: 400,
      color:       negative ? C.negativeText : value <= 0.5 ? "#ABC6A6" : (overrideColor ?? "#192E17"),
      background:  negative ? C.negativeBg : cellBg(Math.abs(value), isToday),
      whiteSpace:  "nowrap",
      borderRight: `1px solid ${C.rowDivider}`,
      borderBottom:`1px solid ${C.rowDivider}`,
    };
  }

  function subCellStyle(value: number, isNegative = false, isToday = false, overrideColor?: string): React.CSSProperties {
    return {
      ...cellStyle(value, isNegative, isToday, overrideColor),
      fontSize:   13,
      fontStyle:  "italic",
      background: isNegative && value < 0 ? C.negativeBg : "#FAFCFA",
    };
  }

  function totalCellStyle(value: number, isNegative = false, overrideColor?: string): React.CSSProperties {
    const negative = isNegative && value < 0;
    return {
      textAlign: "right", padding: "6px 12px", fontSize: 14, fontWeight: 700,
      color:      negative ? C.negativeText : (overrideColor ?? C.totalText),
      background: negative ? C.negativeBg  : C.totalBg,
      borderLeft: `2px solid ${C.totalBorder}`,
      borderBottom: `1px solid ${C.rowDivider}`,
      whiteSpace: "nowrap",
      position:   "sticky" as const, right: 0, zIndex: 1,
    };
  }

  function subTotalStyle(value: number, isNegative = false, overrideColor?: string): React.CSSProperties {
    return {
      ...totalCellStyle(value, isNegative, overrideColor),
      fontSize: 13, fontStyle: "italic", opacity: 0.9,
    };
  }

  function subLabelStyle(accentColor: string): React.CSSProperties {
    return {
      ...rowLabelStyle(accentColor, true),
      paddingLeft: 28, fontSize: 13, fontWeight: 400,
      borderLeft: `3px solid ${accentColor}55`,
      background: "#FAFCFA",
    };
  }

  function toggleRow(key: RowKey) {
    setVisibleRows(prev => {
      const next = new Set(prev);
      next.has(key) ? next.delete(key) : next.add(key);
      return next;
    });
  }

  // ── Render ────────────────────────────────────────────────────────────────
  return (
    <div style={{
      flexShrink: 0, borderTop: `2px solid ${C.headerBorder}`,
      background: "#fff", display: "flex", flexDirection: "column", position: "relative",
    }}>

      {/* Resize handle */}
      <div ref={resizerRef} onMouseDown={onResizerMouseDown}
        style={{ height: 5, cursor: "row-resize", background: "transparent", flexShrink: 0 }}
        onMouseEnter={e => (e.currentTarget.style.background = C.planned)}
        onMouseLeave={e => (e.currentTarget.style.background = "transparent")}
      />

      {/* ── Header ──────────────────────────────────────────────────────── */}
      <div style={{
        display: "flex", alignItems: "center", gap: 8, padding: "0 14px",
        height: 46, borderBottom: `1px solid ${C.headerBorder}`,
        flexShrink: 0, background: C.panelBg,
      }}>
        <span style={{ fontSize: 13, fontWeight: 600, color: C.actual, marginRight: 4 }}>
          📅 Financial Timeline
        </span>

        {!loadingAssignments && (
          <span style={{
            fontSize: 11, fontWeight: 600,
            color:      exactCount > 0 ? C.actual : "#9A9A9A",
            background: exactCount > 0 ? C.nowBg  : "#F5F5F5",
            border:     `1px solid ${exactCount > 0 ? C.remaining : "#D4D4D4"}`,
            borderRadius: 10, padding: "1px 7px",
          }}>
            {exactCount > 0 ? `⚡ ${exactCount} exact` : "~ proportional"}
          </span>
        )}
        {loadingAssignments && <span style={{ fontSize: 11, color: "#9A9A9A" }}>loading…</span>}

        {hasActiveFilters && (
          <span style={{
            fontSize: 11, color: C.actual, fontWeight: 600,
            background: C.nowBg, border: `1px solid ${C.remaining}`,
            borderRadius: 10, padding: "1px 7px",
          }}>filtered</span>
        )}

        <div style={{ width: 1, height: 16, background: C.headerBorder, margin: "0 2px" }} />

        {(["monthly", "quarterly", "yearly"] as Granularity[]).map(g => (
          <button key={g} onClick={() => setGranularity(g)} style={{
            padding: "3px 11px", fontSize: 13,
            fontWeight:   granularity === g ? 700 : 400,
            border:       granularity === g ? `1.5px solid ${C.planned}` : `1px solid ${C.headerBorder}`,
            borderRadius: 12,
            background:   granularity === g ? C.pillActiveBg : "white",
            color:        granularity === g ? C.actual : "#626162",
            cursor: "pointer", fontFamily: "'Segoe UI', system-ui, sans-serif",
          }}>
            {g.charAt(0).toUpperCase() + g.slice(1)}
          </button>
        ))}

        <div style={{ width: 1, height: 16, background: C.headerBorder, margin: "0 2px" }} />

        {ROW_DEFS.map(r => (
          <button key={r.key} onClick={() => toggleRow(r.key)} style={{
            display: "flex", alignItems: "center", gap: 5,
            padding: "3px 10px", fontSize: 13,
            fontWeight:   visibleRows.has(r.key) ? 600 : 400,
            border:       visibleRows.has(r.key) ? `1.5px solid ${ROW_COLORS[r.key]}` : `1px solid ${C.headerBorder}`,
            borderRadius: 12,
            background:   visibleRows.has(r.key) ? `${ROW_COLORS[r.key]}1A` : "white",
            color:        visibleRows.has(r.key) ? ROW_COLORS[r.key] : "#9A9A9A",
            cursor: "pointer", transition: "all 0.15s",
            fontFamily: "'Segoe UI', system-ui, sans-serif",
          }}>
            <span style={{
              width: 7, height: 7, borderRadius: "50%",
              background: visibleRows.has(r.key) ? ROW_COLORS[r.key] : C.headerBorder,
              flexShrink: 0,
            }} />
            {r.label}
          </button>
        ))}

        <div style={{ flex: 1 }} />
        <button onClick={onClose} style={{
          background: "none", border: "none", cursor: "pointer",
          color: C.remaining, fontSize: 16, lineHeight: 1, padding: 2,
        }}>✕</button>
      </div>

      {/* ── Table ───────────────────────────────────────────────────────── */}
      <div style={{ overflow: "auto", height, flexShrink: 0 }}>
        {periods.length === 0 ? (
          <div style={{
            display: "flex", alignItems: "center", justifyContent: "center",
            height: "100%", color: C.remaining, fontSize: 14,
          }}>No tasks with start/end dates to display.</div>
        ) : (
          <table style={{
            borderCollapse: "collapse", minWidth: "100%",
            fontSize: 14, tableLayout: "auto",
            fontFamily: "'Segoe UI', system-ui, sans-serif",
          }}>
            <thead>
              <tr style={{ background: "#fff", position: "sticky", top: 0, zIndex: 2 }}>
                <th style={{
                  position: "sticky", left: 0, zIndex: 3, background: "#fff",
                  padding: "7px 14px",
                  borderBottom: `2px solid ${C.headerBorder}`,
                  borderRight:  `2px solid ${C.headerBorder}`,
                  minWidth: 180, whiteSpace: "nowrap",
                }} />
                {periods.map(p => (
                  <th key={p.isoKey} style={{
                    padding: "7px 12px", fontWeight: 600, fontSize: 14,
                    color:        p.isToday ? C.nowText : p.isPast ? C.cumActual : C.actual,
                    textAlign:    "right",
                    borderBottom: `2px solid ${C.headerBorder}`,
                    borderRight:  `1px solid ${C.rowDivider}`,
                    whiteSpace:   "nowrap",
                    background:   p.isToday ? C.nowBg : "#fff",
                    minWidth:     96,
                  }}>
                    {p.label}
                    {p.isToday && (
                      <div style={{
                        fontSize: 9, fontWeight: 700, color: C.nowMarker,
                        letterSpacing: "0.06em", textTransform: "uppercase", marginTop: 1,
                      }}>▲ NOW</div>
                    )}
                  </th>
                ))}
                <th style={{
                  padding: "7px 12px", fontWeight: 700, fontSize: 14, color: C.totalText,
                  textAlign: "right",
                  borderBottom: `2px solid ${C.headerBorder}`,
                  borderLeft:   `2px solid ${C.totalBorder}`,
                  background:   C.totalBg, whiteSpace: "nowrap",
                  minWidth: 110, position: "sticky", right: 0, zIndex: 2,
                }}>Total</th>
              </tr>
            </thead>

            <tbody>

              {/* ── PLANNED ─────────────────────────────────────────────── */}
              {visibleRows.has("planned") && (
                <React.Fragment>
                  <tr>
                    <td style={{ ...rowLabelStyle(C.planned), cursor: "pointer" }}
                      onClick={() => setPlannedExpanded(x => !x)}>
                      <span style={{ marginRight: 6, fontSize: 10 }}>{plannedExpanded ? "▼" : "▶"}</span>
                      Planned
                    </td>
                    {agg.pp.map((v, i) => (
                      <td key={i} style={cellStyle(v, false, periods[i].isToday, C.actual)}>
                        {v > 0.5 ? fmtCurrency(Math.round(v)) : "—"}
                      </td>
                    ))}
                    <td style={totalCellStyle(totals.totalPlanned, false, C.planned)}>
                      {fmtCurrency(Math.round(totals.totalPlanned))}
                    </td>
                  </tr>
                  {plannedExpanded && (
                    <React.Fragment>
                      <tr>
                        <td style={subLabelStyle(C.planned)}>↳ Effort cost</td>
                        {agg.pe.map((v, i) => (
                          <td key={i} style={subCellStyle(v, false, periods[i].isToday, C.planned)}>
                            {v > 0.5 ? fmtCurrency(Math.round(v)) : "—"}
                          </td>
                        ))}
                        <td style={subTotalStyle(totals.totalPlannedEffort, false, C.planned)}>
                          {fmtCurrency(Math.round(totals.totalPlannedEffort))}
                        </td>
                      </tr>
                      <tr>
                        <td style={subLabelStyle(C.remaining)}>↳ Fixed cost</td>
                        {agg.pf.map((v, i) => (
                          <td key={i} style={subCellStyle(v, false, periods[i].isToday, C.remaining)}>
                            {v > 0.5 ? fmtCurrency(Math.round(v)) : "—"}
                          </td>
                        ))}
                        <td style={subTotalStyle(totals.totalPlannedFixed, false, C.remaining)}>
                          {fmtCurrency(Math.round(totals.totalPlannedFixed))}
                        </td>
                      </tr>
                    </React.Fragment>
                  )}
                </React.Fragment>
              )}

              {/* ── ACTUAL ──────────────────────────────────────────────── */}
              {visibleRows.has("actual") && (
                <React.Fragment>
                  <tr>
                    <td style={{ ...rowLabelStyle(C.actual), cursor: "pointer" }}
                      onClick={() => setActualExpanded(x => !x)}>
                      <span style={{ marginRight: 6, fontSize: 10 }}>{actualExpanded ? "▼" : "▶"}</span>
                      Actual
                    </td>
                    {agg.pa.map((v, i) => (
                      <td key={i} style={cellStyle(v, false, periods[i].isToday, C.actual)}>
                        {v > 0.5 ? fmtCurrency(Math.round(v)) : "—"}
                      </td>
                    ))}
                    <td style={totalCellStyle(totals.totalActual, false, C.actual)}>
                      {fmtCurrency(Math.round(totals.totalActual))}
                    </td>
                  </tr>
                  {actualExpanded && (
                    <React.Fragment>
                      <tr>
                        <td style={subLabelStyle(C.actual)}>↳ Effort cost</td>
                        {agg.ae.map((v, i) => (
                          <td key={i} style={subCellStyle(v, false, periods[i].isToday, C.actual)}>
                            {v > 0.5 ? fmtCurrency(Math.round(v)) : "—"}
                          </td>
                        ))}
                        <td style={subTotalStyle(totals.totalActualEffort, false, C.actual)}>
                          {fmtCurrency(Math.round(totals.totalActualEffort))}
                        </td>
                      </tr>
                      <tr>
                        <td style={subLabelStyle(C.remaining)}>↳ Fixed cost</td>
                        {agg.af.map((v, i) => (
                          <td key={i} style={subCellStyle(v, false, periods[i].isToday, C.remaining)}>
                            {v > 0.5 ? fmtCurrency(Math.round(v)) : "—"}
                          </td>
                        ))}
                        <td style={subTotalStyle(totals.totalActualFixed, false, C.remaining)}>
                          {fmtCurrency(Math.round(totals.totalActualFixed))}
                        </td>
                      </tr>
                    </React.Fragment>
                  )}
                </React.Fragment>
              )}

              {/* ── REMAINING ───────────────────────────────────────────── */}
              {visibleRows.has("remaining") && (
                <React.Fragment>
                  <tr>
                    <td style={{ ...rowLabelStyle(C.remaining), cursor: "pointer" }}
                      onClick={() => setRemainingExpanded(x => !x)}>
                      <span style={{ marginRight: 6, fontSize: 10 }}>{remainingExpanded ? "▼" : "▶"}</span>
                      Remaining
                    </td>
                    {totals.remaining.map((v, i) => (
                      <td key={i} style={cellStyle(v, true, periods[i].isToday)}>
                        {Math.abs(v) > 0.5 ? fmtCurrency(Math.round(v)) : "—"}
                      </td>
                    ))}
                    <td style={totalCellStyle(totals.totalRemaining, true, C.remaining)}>
                      {fmtCurrency(Math.round(totals.totalRemaining))}
                    </td>
                  </tr>
                  {remainingExpanded && (
                    <React.Fragment>
                      <tr>
                        <td style={subLabelStyle(C.planned)}>↳ Effort remaining</td>
                        {totals.remEffort.map((v, i) => (
                          <td key={i} style={subCellStyle(v, true, periods[i].isToday, C.planned)}>
                            {Math.abs(v) > 0.5 ? fmtCurrency(Math.round(v)) : "—"}
                          </td>
                        ))}
                        <td style={subTotalStyle(totals.totalRemEffort, true, C.planned)}>
                          {fmtCurrency(Math.round(totals.totalRemEffort))}
                        </td>
                      </tr>
                      <tr>
                        <td style={subLabelStyle(C.remaining)}>↳ Fixed remaining</td>
                        {totals.remFixed.map((v, i) => (
                          <td key={i} style={subCellStyle(v, true, periods[i].isToday, C.remaining)}>
                            {Math.abs(v) > 0.5 ? fmtCurrency(Math.round(v)) : "—"}
                          </td>
                        ))}
                        <td style={subTotalStyle(totals.totalRemFixed, true, C.remaining)}>
                          {fmtCurrency(Math.round(totals.totalRemFixed))}
                        </td>
                      </tr>
                    </React.Fragment>
                  )}
                </React.Fragment>
              )}

              {/* ── Separator ───────────────────────────────────────────── */}
              {(visibleRows.has("planned") || visibleRows.has("actual") || visibleRows.has("remaining")) &&
               (visibleRows.has("cumPlanned") || visibleRows.has("cumActual")) && (
                <tr>
                  <td colSpan={periods.length + 2}
                    style={{ height: 2, background: C.sepLine, padding: 0 }} />
                </tr>
              )}

              {/* ── CUMULATIVE PLANNED ──────────────────────────────────── */}
              {visibleRows.has("cumPlanned") && (
                <tr>
                  <td style={rowLabelStyle(C.cumPlanned, true)}>Cumulative Planned</td>
                  {totals.cumPlanned.map((v, i) => (
                    <td key={i} style={{
                      ...cellStyle(v, false, periods[i].isToday, C.cumPlanned),
                      fontStyle: "italic", fontWeight: 500,
                    }}>
                      {v > 0.5 ? fmtCurrency(Math.round(v)) : "—"}
                    </td>
                  ))}
                  <td style={{ ...totalCellStyle(totals.totalPlanned, false, C.cumPlanned), fontStyle: "italic" }}>
                    {fmtCurrency(Math.round(totals.totalPlanned))}
                  </td>
                </tr>
              )}

              {/* ── CUMULATIVE ACTUAL ───────────────────────────────────── */}
              {visibleRows.has("cumActual") && (
                <tr>
                  <td style={rowLabelStyle(C.cumActual, true)}>Cumulative Actual</td>
                  {totals.cumActual.map((v, i) => (
                    <td key={i} style={{
                      ...cellStyle(v, false, periods[i].isToday, C.cumActual),
                      fontStyle: "italic", fontWeight: 500,
                    }}>
                      {v > 0.5 ? fmtCurrency(Math.round(v)) : "—"}
                    </td>
                  ))}
                  <td style={{ ...totalCellStyle(totals.totalActual, false, C.cumActual), fontStyle: "italic" }}>
                    {fmtCurrency(Math.round(totals.totalActual))}
                  </td>
                </tr>
              )}

            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

// ─── Row label cell ───────────────────────────────────────────────────────────

function rowLabelStyle(accentColor: string, italic = false): React.CSSProperties {
  return {
    position:    "sticky", left: 0, zIndex: 1, background: "#fff",
    padding:     "6px 14px", fontWeight: 600, fontSize: 14, color: accentColor,
    borderRight: "2px solid #BED3BA", borderBottom: "1px solid #F0F5EF",
    borderLeft:  `3px solid ${accentColor}`, whiteSpace: "nowrap",
    fontStyle:   italic ? "italic" : "normal", minWidth: 180,
    fontFamily:  "'Segoe UI', system-ui, sans-serif",
  };
}