export interface TaskNode {
  id: number;
  recordId: string;
  taskName: string;
  parentTaskId: number | null;
  parentGuid?: string | null;    // ← add this
  startDate: string;
  endDate: string;
  pctDone: number;
  isSummary: boolean;
  assignee: string;
  subRows?: TaskNode[];
}

export function buildTree(flat: TaskNode[]): TaskNode[] {
  const map = new Map<number, TaskNode>();
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
    if (n.recordId === recordId)  return { ...n, [field]: value };
    if (n.subRows)                return { ...n, subRows: updateNodeInTree(n.subRows, recordId, field, value) };
    return n;
  });
}