import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { TaskGrid } from "./components/TaskGrid";
import { TaskNode } from "./components/buildTree";

interface FlatTask {
  recordId:   string;
  taskName:   string;
  parentGuid: string | null;
  startDate:  string;
  endDate:    string;
  pctDone:    number;
  isSummary:  boolean;
  assignee:   string;
}

function buildTreeByGuid(flat: FlatTask[]): TaskNode[] {
  const map  = new Map<string, TaskNode>();
  const roots: TaskNode[] = [];

  flat.forEach(r => {
    const node: TaskNode = {
      id:           0,
      recordId:     r.recordId,
      taskName:     r.taskName,
      parentTaskId: null,
      parentGuid:   r.parentGuid,
      startDate:    r.startDate,
      endDate:      r.endDate,
      pctDone:      r.pctDone,
      isSummary:    r.isSummary,
      assignee:     r.assignee,
      subRows:      [],
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

  private container: HTMLDivElement | null = null;
  private renderAttempts = 0;

  public init(
    context: ComponentFramework.Context<IInputs>,
    _notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    console.log("[TaskGrid] init fired, container:", container?.tagName ?? "undefined");
    if (container && container.tagName) {
      this.container = container;
    }
    context.parameters.TaskDataSet.paging.setPageSize(5000);
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this.renderAttempts++;
    console.log("[TaskGrid] updateView #" + this.renderAttempts, "container:", this.container?.tagName ?? "undefined");

    // If container wasn't passed in init, try getting it now
    if (!this.container) {
      const el = (context as any).mode?.trackContainerResize
        ? (context as any).factory?.getRoot?.()
        : null;
      console.log("[TaskGrid] factory root:", el);
      if (el) this.container = el;
    }

    if (!this.container) {
      console.warn("[TaskGrid] no container yet, skipping render #" + this.renderAttempts);
      return;
    }

    const dataset = context.parameters.TaskDataSet;

    console.log("[TaskGrid] dataset loading:", dataset.loading,
                "records:", dataset.sortedRecordIds.length);

    if (dataset.loading) return;

    if (dataset.paging.hasNextPage) {
      dataset.paging.loadNextPage();
      return;
    }

    const flat: FlatTask[] = dataset.sortedRecordIds.map((id: string) => {
      const rec = dataset.records[id];
      const parentRef = rec.getValue("parentTaskId") as any;
      const parentGuid: string | null =
        parentRef?.id?.guid ?? parentRef?.id ?? null;

      return {
        recordId:   id,
        taskName:   String(rec.getValue("taskName")  ?? ""),
        parentGuid: parentGuid,
        startDate:  String(rec.getValue("startDate") ?? ""),
        endDate:    String(rec.getValue("endDate")   ?? ""),
        pctDone:    Number(rec.getValue("pctDone")   ?? 0),
        isSummary:  rec.getValue("isSummary")        === true,
        assignee:   "",
      };
    });

    console.log("[TaskGrid] flat records:", flat.length);
    console.log("[TaskGrid] sample record:", flat[0]);

    try {
      ReactDOM.render(
        React.createElement(TaskGrid, {
          data:   buildTreeByGuid(flat),
          onSave: this.saveToDataverse.bind(this),
        }),
        this.container
      );
      console.log("[TaskGrid] render SUCCESS, rows:", flat.length);
    } catch(e) {
      console.error("[TaskGrid] render FAILED:", e);
    }
  }

  private async saveToDataverse(
    changes: Record<string, Partial<TaskNode>>
  ): Promise<void> {
    const entityPluralName = "msdyn_projecttasks";

    const requests = Object.entries(changes).map(([recordId, fields]) => {
      const payload: Record<string, unknown> = {};
      if (fields.taskName  !== undefined) payload["msdyn_subject"]  = fields.taskName;
      if (fields.startDate !== undefined) payload["msdyn_start"]    = fields.startDate;
      if (fields.endDate   !== undefined) payload["msdyn_finish"]   = fields.endDate;
      if (fields.pctDone   !== undefined) payload["msdyn_progress"] = fields.pctDone;

      return fetch(`/api/data/v9.2/${entityPluralName}(${recordId})`, {
        method: "PATCH",
        headers: {
          "Content-Type": "application/json",
          "OData-Version": "4.0",
          "If-Match": "*",
        },
        body: JSON.stringify(payload),
      });
    });

    await Promise.all(requests);
  }

  public getOutputs(): IOutputs { return {}; }

  public destroy(): void {
    if (this.container) ReactDOM.unmountComponentAtNode(this.container);
  }
}