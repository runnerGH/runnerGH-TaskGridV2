import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { TaskGrid } from "./components/TaskGrid";
import { TaskNode, rollupCosts } from "./components/buildTree";

interface FlatTask {
  recordId:         string;
  taskName:         string;
  parentGuid:       string | null;
  startDate:        string;
  endDate:          string;
  pctDone:          number;
  isSummary:        boolean;
  assignee:         string;
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
}

function buildTreeByGuid(flat: FlatTask[]): TaskNode[] {
  const map   = new Map<string, TaskNode>();
  const roots: TaskNode[] = [];

  flat.forEach(r => {
    const node: TaskNode = {
      id:               0,
      recordId:         r.recordId,
      taskName:         r.taskName,
      parentTaskId:     null,
      parentGuid:       r.parentGuid,
      startDate:        r.startDate,
      endDate:          r.endDate,
      pctDone:          r.pctDone,
      isSummary:        r.isSummary,
      assignee:         r.assignee,
      costCategory:     r.costCategory,
      srcServiceId:     r.srcServiceId,
      srcServiceName:   r.srcServiceName,
      fiscalYearId:     r.fiscalYearId,
      fiscalYearName:   r.fiscalYearName,
      quantity:         r.quantity,
      unitRate:         r.unitRate,
      unit:             r.unit,
      frequency:        r.frequency,
      plannedCost:      r.plannedCost,
      fixedCost:        r.fixedCost,
      totalPlannedCost: r.totalPlannedCost,
      actualCost:       r.actualCost,
      remainingCost:    r.remainingCost,
      earnedValue:      r.earnedValue,
      subRows:          [],
    };
    map.set(r.recordId, node);
  });

  flat.forEach(r => {
    const node = map.get(r.recordId)!;
    if (r.parentGuid && map.has(r.parentGuid)) {
      map.get(r.parentGuid)!.subRows!.push(node);
    } else {
      roots.push(node);
    }
  });

  map.forEach(node => {
    if (node.subRows && node.subRows.length === 0) delete node.subRows;
  });

  return roots;
}

export class TaskGridControl
  implements ComponentFramework.StandardControl<IInputs, IOutputs> {

  private container!: HTMLDivElement;
  private notifyOutputChanged!: () => void;

public init(
  context: ComponentFramework.Context<IInputs>,
  notifyOutputChanged: () => void,
  _state: ComponentFramework.Dictionary,
  container: HTMLDivElement
): void {
  this.container = container;
  this.notifyOutputChanged = notifyOutputChanged;
  context.parameters.TaskDataSet.paging.setPageSize(5000);
}

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    if (!this.container) return;

    const dataset = context.parameters.TaskDataSet;
    if (dataset.loading) return;

    if (dataset.paging.hasNextPage) {
      dataset.paging.loadNextPage();
      return;
    }

const flat: FlatTask[] = dataset.sortedRecordIds.map((id: string) => {
  const rec = dataset.records[id];

  const parentRef  = rec.getValue("parentTaskId") as any;
  const parentGuid = parentRef?.id?.guid ?? parentRef?.id ?? null;

  const srcRef         = rec.getValue("srcService") as any;
  const srcServiceId   = srcRef?.id?.guid ?? srcRef?.id ?? null;
  const srcServiceName = srcRef?.name ?? null;

  const fyRef          = rec.getValue("fiscalYear") as any;
  const fiscalYearId   = fyRef?.id?.guid ?? fyRef?.id ?? null;
  const fiscalYearName = fyRef?.name ?? null;

  const quantity         = Number(rec.getValue("quantity")         ?? 0);
  const unitRate         = Number(rec.getValue("unitRate")         ?? 0);
  const fixedCost        = Number(rec.getValue("fixedCost")        ?? 0);
  const actualCost       = Number(rec.getValue("actualCost")       ?? 0);
  const pctDone          = Number(rec.getValue("pctDone")          ?? 0) * 100;

  // Read plannedCost from Dataverse calculated field
  const plannedCost      = Number(rec.getValue("plannedCost")      ?? 0);
  const totalPlannedCost = Number(rec.getValue("totalPlannedCost") ?? 0);
  const remainingCost    = Number(rec.getValue("remainingCost")  ?? 0);
  const earnedValue      = Number(rec.getValue("earnedValue")    ?? 0);

  return {
    recordId:         id,
    taskName:         String(rec.getValue("taskName")         ?? ""),
    parentGuid:       parentGuid,
    startDate:        String(rec.getValue("startDate")        ?? ""),
    endDate:          String(rec.getValue("endDate")          ?? ""),
    pctDone,
    isSummary:        rec.getValue("isSummary")               == 1,
    assignee:         "",
    costCategory:     rec.getValue("costCategory") != null
                        ? Number(rec.getValue("costCategory")) : null,
    srcServiceId,
    srcServiceName,
    fiscalYearId,
    fiscalYearName,
    quantity,
    unitRate,
    unit:             String(rec.getValue("unit")             ?? ""),
    frequency:        String(rec.getValue("frequency")        ?? ""),
    plannedCost,
    fixedCost,
    totalPlannedCost,
    actualCost,
    remainingCost,
    earnedValue,
  };
});

    const tree       = buildTreeByGuid(flat);
    const rolledUp   = rollupCosts(tree);

    // Extract project ID from first task record
    const firstId = context.parameters.TaskDataSet.sortedRecordIds[0];
    const projectId = firstId
      ? String((context.parameters.TaskDataSet.records[firstId] as any)
          .raw?._msdyn_project_value ?? "")
      : "";

    // Collect all task record IDs for resource assignment filtering
    const taskIds = context.parameters.TaskDataSet.sortedRecordIds;

    ReactDOM.render(
      React.createElement(TaskGrid, {
        data:      rolledUp,
        onSave:    this.saveToDataverse.bind(this),
        onRefresh: () => context.parameters.TaskDataSet.refresh(),
        userId:    context.userSettings.userId,
        taskIds,
      }),
      this.container
    );
  }

private async saveToDataverse(
  changes: Record<string, Partial<TaskNode>>
): Promise<void> {
  const entityPluralName = "msdyn_projecttasks";

  const requests = Object.entries(changes).map(async ([recordId, fields]) => {
    const payload: Record<string, unknown> = {};

    if (fields.taskName   !== undefined) payload["msdyn_subject"]    = fields.taskName;
    if (fields.unitRate   !== undefined) payload["pmo_unitrate"]     = fields.unitRate;
    if (fields.fixedCost  !== undefined) payload["pmo_fixedcost"]    = fields.fixedCost;
    if (fields.actualCost !== undefined) payload["cred8_actualcost"] = fields.actualCost;
    if (fields.unit       !== undefined) payload["pmo_unit"]         = fields.unit;
    if (fields.frequency  !== undefined) payload["pmo_frequency"]    = fields.frequency;

    if (fields.costCategory !== undefined) {
      payload["cred8_costcategory"] = fields.costCategory;
    }
    if (fields.srcServiceId !== undefined && fields.srcServiceId !== null) {
      payload["pmo_SRCService@odata.bind"] =
        `/pmo_serviceratecards(${fields.srcServiceId})`;
    }
    if (fields.fiscalYearId !== undefined && fields.fiscalYearId !== null) {
      payload["pmo_fiscalyear@odata.bind"] =
        `/pmo_fiscalyear1s(${fields.fiscalYearId})`;
    }

    const response = await fetch(
      `/api/data/v9.2/${entityPluralName}(${recordId})`,
      {
        method: "PATCH",
        headers: {
          "Content-Type": "application/json",
          "OData-Version": "4.0",
          "If-Match": "*",
        },
        body: JSON.stringify(payload),
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      console.error(`[TaskGrid] PATCH failed for ${recordId}:`, response.status, errorText);
      throw new Error(`Save failed (${response.status}): ${errorText}`);
    }

    return response;
  });

  await Promise.all(requests);
}

  private refreshData(context: ComponentFramework.Context<IInputs>): void {
    context.parameters.TaskDataSet.refresh();
  }
  public getOutputs(): IOutputs { return {}; }

  public destroy(): void {
    if (this.container) ReactDOM.unmountComponentAtNode(this.container);
  }
}