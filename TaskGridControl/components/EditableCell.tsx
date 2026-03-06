import * as React from "react";

interface Props {
  getValue: () => unknown;
  row: any;
  column: any;
  table: any;
  editable?: boolean;
  type?: "text" | "date" | "number";
}

export function EditableCell({ getValue, row, column, table, editable = true, type = "text" }: Props) {
  const initial = getValue() as string | number;
  const [editing, setEditing] = React.useState(false);
  const [value, setValue]     = React.useState(initial);
  const ref                   = React.useRef<HTMLInputElement>(null);

  React.useEffect(() => { setValue(initial); }, [initial]);
  React.useEffect(() => { if (editing && ref.current) ref.current.focus(); }, [editing]);

  function commit() {
    setEditing(false);
    if (value !== initial) {
      table.options.meta?.updateCell(row.original.recordId, column.id, value);
    }
  }

  function onKeyDown(e: React.KeyboardEvent) {
    if (e.key === "Enter")  commit();
    if (e.key === "Escape") { setValue(initial); setEditing(false); }
  }

  if (!editing) {
    if (type === "number") {
      const pct = Number(value) || 0;
      return (
        <div onDoubleClick={() => editable && setEditing(true)}
          style={{ cursor: editable ? "text" : "default", display: "flex", alignItems: "center", gap: 6 }}>
          <div style={{ flex: 1, height: 6, background: "#e2e8f0", borderRadius: 3, overflow: "hidden" }}>
            <div style={{
              width: `${pct}%`, height: "100%", borderRadius: 3,
              background: pct === 100 ? "#22c55e" : pct > 50 ? "#3b82f6" : "#f59e0b",
            }} />
          </div>
          <span style={{ fontSize: 11, color: "#64748b", width: 30, textAlign: "right" }}>{pct}%</span>
        </div>
      );
    }
    return (
      <div onDoubleClick={() => editable && setEditing(true)}
        className={editable ? "editable-cell" : ""}
        style={{ cursor: editable ? "text" : "default", minHeight: 20, padding: "0 2px" }}>
        {String(value ?? "")}
      </div>
    );
  }

  return (
    <input ref={ref} type={type}
      min={type === "number" ? 0 : undefined}
      max={type === "number" ? 100 : undefined}
      value={value as string}
      onChange={e => setValue(type === "number" ? Number(e.target.value) : e.target.value)}
      onBlur={commit} onKeyDown={onKeyDown}
      style={{ width: "100%", border: "1px solid #3b82f6", borderRadius: 3,
               padding: "1px 4px", fontSize: 12, outline: "none", boxSizing: "border-box" }}
    />
  );
}