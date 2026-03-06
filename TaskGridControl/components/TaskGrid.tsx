import * as React from "react";
import {
  useReactTable, getCoreRowModel, getExpandedRowModel,
  flexRender, createColumnHelper, ExpandedState,
} from "@tanstack/react-table";
import { TaskNode, updateNodeInTree } from "./buildTree";
import { EditableCell } from "./EditableCell";

interface Props {
  data: TaskNode[];
  onSave: (changes: Record<string, Partial<TaskNode>>) => Promise<void>;
}

const col = createColumnHelper<TaskNode>();

const CSS = `
  .tg-wrap { font-family: 'Segoe UI', system-ui, sans-serif; font-size: 13px;
             display: flex; flex-direction: column; height: 100%; background: #f8fafc; }
  .tg-toolbar { background: white; border-bottom: 1px solid #e2e8f0;
                padding: 6px 12px; display: flex; align-items: center; gap: 6px; flex-shrink: 0; }
  .tg-title   { font-weight: 700; font-size: 14px; color: #1e293b; margin-right: 8px; }
  .tg-btn     { display: inline-flex; align-items: center; gap: 4px; padding: 4px 10px;
                border-radius: 4px; border: 1px solid #e2e8f0; font-size: 12px;
                cursor: pointer; font-weight: 500; background: white; color: #374151; }
  .tg-btn:hover         { background: #f1f5f9; }
  .tg-btn-save          { background: #2563eb; color: white; border-color: #2563eb; }
  .tg-btn-save:hover    { background: #1d4ed8; }
  .tg-btn-save:disabled { background: #93c5fd; cursor: not-allowed; }
  .tg-btn-discard       { color: #dc2626; border-color: #fca5a5; }
  .tg-btn-discard:hover { background: #fef2f2; }
  .tg-badge   { background: #3b82f6; color: white; border-radius: 10px;
                padding: 1px 6px; font-size: 10px; font-weight: 700; }
  .tg-scroll  { overflow: auto; flex: 1; }
  .tg-table   { border-collapse: collapse; width: 100%; }
  .tg-table th { background: #f1f5f9; color: #475569; font-weight: 600; font-size: 11px;
                 text-transform: uppercase; letter-spacing: 0.05em; padding: 7px 10px;
                 border-bottom: 2px solid #e2e8f0; border-right: 1px solid #e2e8f0;
                 text-align: left; position: sticky; top: 0; z-index: 2; white-space: nowrap; }
  .tg-table th:last-child, .tg-table td:last-child { border-right: none; }
  .tg-table td { padding: 4px 10px; border-bottom: 1px solid #f1f5f9;
                 border-right: 1px solid #f1f5f9; vertical-align: middle; height: 32px; }
  .tg-table tr:hover td  { background: #f0f7ff !important; }
  .tg-summary td         { background: #f8f7ff; }
  .tg-changed td:first-child { border-left: 3px solid #3b82f6; }
  .tg-expand-btn { background: none; border: none; cursor: pointer; padding: 0 3px 0 0;
                   color: #64748b; font-size: 10px; }
  .editable-cell         { border-radius: 2px; }
  .editable-cell:hover   { background: #dbeafe; }
  .tg-footer  { padding: 5px 12px; background: white; border-top: 1px solid #e2e8f0;
                font-size: 11px; color: #94a3b8; display: flex; justify-content: space-between; flex-shrink: 0; }
  .tg-saved   { color: #16a34a; font-size: 12px; font-weight: 500; }
`;

export function TaskGrid({ data: initialData, onSave }: Props) {
  const [data, setData]         = React.useState<TaskNode[]>(initialData);
  const [expanded, setExpanded] = React.useState<ExpandedState>({ "0": true });
  const [pending, setPending]   = React.useState<Record<string, Partial<TaskNode>>>({});
  const [saving, setSaving]     = React.useState(false);
  const [savedMsg, setSavedMsg] = React.useState(false);

  React.useEffect(() => { setData(initialData); }, [initialData]);

  function updateCell(recordId: string, field: keyof TaskNode, value: unknown) {
    setData(prev => updateNodeInTree(prev, recordId, field, value));
    setPending(p => ({ ...p, [recordId]: { ...(p[recordId] ?? {}), [field]: value } }));
  }

  const columns = React.useMemo(() => [
    col.accessor("taskName", {
      header: "Task Name",
      size: 320,
      cell: function TaskNameCell({ row, getValue, column, table }) {
        const isSummary = row.original.isSummary;
        return (
          <div style={{
            paddingLeft: row.depth * 18, display: "flex", alignItems: "center",
            fontWeight: isSummary ? 600 : 400, color: isSummary ? "#1e293b" : "#334155",
          }}>
            {row.getCanExpand() ? (
              <button className="tg-expand-btn" onClick={row.getToggleExpandedHandler()}>
                {row.getIsExpanded() ? "▾" : "▸"}
              </button>
            ) : (
              <span style={{ width: 14, display: "inline-block" }} />
            )}
            <span style={{
              width: 7, height: 7, flexShrink: 0, marginRight: 7, display: "inline-block",
              borderRadius: isSummary ? 2 : "50%",
              background: isSummary ? "#6366f1" : "#94a3b8",
            }} />
            <EditableCell getValue={getValue} row={row} column={column} table={table}
              editable={!isSummary} type="text" />
          </div>
        );
      },
    }),
    col.accessor("assignee",  { header: "Assignee", size: 110,
      cell: function AssigneeCell(p) { return <EditableCell {...p} editable={!p.row.original.isSummary} type="text" />; } }),
    col.accessor("startDate", { header: "Start",    size: 110,
      cell: function StartCell(p)    { return <EditableCell {...p} editable={!p.row.original.isSummary} type="date" />; } }),
    col.accessor("endDate",   { header: "Finish",   size: 110,
      cell: function EndCell(p)      { return <EditableCell {...p} editable={!p.row.original.isSummary} type="date" />; } }),
    col.accessor("pctDone",   { header: "% Done",   size: 130,
      cell: function PctCell(p)      { return <EditableCell {...p} editable={!p.row.original.isSummary} type="number" />; } }),
  ], []);

  const table = useReactTable({
    data,
    columns,
    state: { expanded },
    onExpandedChange: setExpanded,
    getSubRows: function(row) { return row.subRows; },
    getCoreRowModel: getCoreRowModel(),
    getExpandedRowModel: getExpandedRowModel(),
    meta: { updateCell },
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

function expandAll() {
  const all: Record<string, boolean> = {};
  function walk(rows: any[]) {
    rows.forEach((r: any) => {
      if (r.getCanExpand()) {
        all[r.id] = true;
        walk(r.subRows ?? []);
      }
    });
  }
  walk(table.getRowModel().rows);
  setExpanded(all);
}

  return (
    <div className="tg-wrap">
      <style>{CSS}</style>
      <div className="tg-toolbar">
        <span className="tg-title">📋 Task Grid</span>
        <button className="tg-btn" onClick={expandAll}>▾ Expand All</button>
        <button className="tg-btn" onClick={() => setExpanded({})}>▸ Collapse All</button>
        <div style={{ flex: 1 }} />
        {changesCount > 0 && !savedMsg && (
          <React.Fragment>
            <span style={{ fontSize: 12, color: "#64748b" }}>
              <span className="tg-badge">{changesCount}</span>
              {" "}{changesCount === 1 ? "unsaved change" : "unsaved changes"}
            </span>
            <button className="tg-btn tg-btn-discard"
              onClick={() => { setData(initialData); setPending({}); }}>
              ✕ Discard
            </button>
            <button className="tg-btn tg-btn-save" onClick={handleSave} disabled={saving}>
              {saving ? "Saving…" : "💾 Save"}
            </button>
          </React.Fragment>
        )}
        {savedMsg && <span className="tg-saved">✓ Saved</span>}
      </div>
      <div className="tg-scroll">
        <table className="tg-table">
          <thead>
            {table.getHeaderGroups().map(hg => (
              <tr key={hg.id}>
                {hg.headers.map(h => (
                  <th key={h.id} style={{ width: h.getSize(), minWidth: h.getSize() }}>
                    {flexRender(h.column.columnDef.header, h.getContext())}
                  </th>
                ))}
              </tr>
            ))}
          </thead>
          <tbody>
            {table.getRowModel().rows.map(row => (
              <tr key={row.id} className={[
                row.original.isSummary         ? "tg-summary" : "",
                pending[row.original.recordId] ? "tg-changed" : "",
              ].join(" ")}>
                {row.getVisibleCells().map(cell => (
                  <td key={cell.id}>
                    {flexRender(cell.column.columnDef.cell, cell.getContext())}
                  </td>
                ))}
              </tr>
            ))}
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