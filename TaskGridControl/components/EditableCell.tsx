import * as React from "react";

interface Props {
  getValue: () => unknown;
  row: any;
  column: any;
  table: any;
  editable?: boolean;
  type?: "text" | "date" | "number";
  className?: string;
}

export function EditableCell({
  getValue, row, column, table,
  editable = true,
  type = "text",
  className = "tg-cell-text",
}: Props) {
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
    return (
      <div
        onDoubleClick={() => editable && setEditing(true)}
        className={className}
        title={editable ? "Double-click to edit" : undefined}
      >
        {String(value ?? "")}
      </div>
    );
  }

  return (
    <input
      ref={ref}
      type={type}
      min={type === "number" ? 0   : undefined}
      max={type === "number" ? 100 : undefined}
      value={value as string}
      onChange={e => setValue(type === "number" ? Number(e.target.value) : e.target.value)}
      onBlur={commit}
      onKeyDown={onKeyDown}
      style={{
        width: "100%", border: "1px solid #4f46e5", borderRadius: 3,
        padding: "1px 4px", fontSize: 12, outline: "none", boxSizing: "border-box",
      }}
    />
  );
}