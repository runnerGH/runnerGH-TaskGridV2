export interface TaskNode {
  // Core fields
  id:               number;
  recordId:         string;
  taskName:         string;
  parentTaskId:     number | null;
  parentGuid?:      string | null;
  startDate:        string;
  endDate:          string;
  pctDone:          number;
  isSummary:        boolean;
  assignee:         string;

  // Cost fields
  costCategory:     number | null;
  srcServiceId:     string | null;
  srcServiceName:   string | null;
  fiscalYearId:     string | null;
  fiscalYearName:   string | null;
  quantity:         number;
  unitRate:         number;
  unit:             string;
  frequency:        string;
  plannedCost:      number;
  fixedCost:        number;
  totalPlannedCost: number;
  actualCost:       number;
  remainingCost:    number;
  earnedValue:      number;

  subRows?: TaskNode[];
}

export function buildTree(flat: TaskNode[]): TaskNode[] {
  const map  = new Map<number, TaskNode>();
  const roots: TaskNode[] = [];

  flat.forEach(r => map.set(r.id, { ...r, subRows: [] }));

  map.forEach(node => {
    if (node.parentTaskId !== null && map.has(node.parentTaskId)) {
      map.get(node.parentTaskId)!.subRows!.push(node);
    } else {
      roots.push(node);
    }
  });

  map.forEach(node => {
    if (node.subRows && node.subRows.length === 0) delete node.subRows;
  });

  return roots;
}

export function updateNodeInTree(
  nodes: TaskNode[],
  recordId: string,
  field: keyof TaskNode,
  value: unknown
): TaskNode[] {
  return nodes.map(n => {
    if (n.recordId === recordId) return { ...n, [field]: value };
    if (n.subRows)               return { ...n, subRows: updateNodeInTree(n.subRows, recordId, field, value) };
    return n;
  });
}

// ── Cost rollup ───────────────────────────────────────────────────────────────
// Recursively sum cost fields from children up to summary rows
export function rollupCosts(nodes: TaskNode[]): TaskNode[] {
  return nodes.map(node => {
    if (!node.subRows || node.subRows.length === 0) return node;

    const rolledUp = rollupCosts(node.subRows);

    const sum = (field: keyof TaskNode) =>
      rolledUp.reduce((acc, child) => acc + (Number(child[field]) || 0), 0);

    return {
      ...node,
      subRows:          rolledUp,
      quantity:         sum("quantity"),
      plannedCost:      sum("plannedCost"),
      fixedCost:        sum("fixedCost"),
      totalPlannedCost: sum("totalPlannedCost"),
      actualCost:       sum("actualCost"),
      remainingCost:    sum("remainingCost"),
      earnedValue:      sum("earnedValue"),
    };
  });
}