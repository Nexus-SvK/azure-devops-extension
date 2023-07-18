import { WorkRestClient, TeamSettingsIteration } from "azure-devops-extension-api/Work"
import { WorkItem, WorkItemExpand, WorkItemTrackingRestClient } from "azure-devops-extension-api/WorkItemTracking";
import { JsonPatchDocument, Operation } from "azure-devops-extension-api/WebApi";
import { TeamContext } from "azure-devops-extension-api/Core/Core";
import { ParentWorkItem } from "./DataTypeFormats";

export class SprintProcessor {

    public _workHttpClient: WorkRestClient;
    public _witClient: WorkItemTrackingRestClient;
    public _teamContext: TeamContext;
    public _nextSprint: TeamSettingsIteration;
    public readonly systemFields: string[] = [
        "System.IterationId", "System.ExternalLinkCount", "System.HyperLinkCount", "System.AttachedFileCount", "System.NodeName",
        "System.RevisedDate", "System.ChangedDate", "System.Id", "System.AreaId", "System.AuthorizedAs", "System.State", "System.AuthorizedDate", "System.Watermark",
        "System.Rev", "System.ChangedBy", "System.Reason", "System.WorkItemType", "System.CreatedDate", "System.CreatedBy", "System.History", "System.RelatedLinkCount",
        "System.BoardColumn", "System.BoardColumnDone", "System.BoardLane", "System.CommentCount", "System.TeamProject", "System.AreaLevel1", "System.IterationLevel1",
        "System.IterationLevel2", "Microsoft.VSTS.Common.StateChangeDate", "Microsoft.VSTS.Common.ActivatedDate", "Microsoft.VSTS.Common.ActivatedBy", "System.AreaPath",
        "Microsoft.VSTS.Scheduling.CompletedWork", "System.IterationPath", "System.Title", "Microsoft.VSTS.Common.ClosedBy", "Microsoft.VSTS.Common.ClosedDate"];

    constructor(workHttpClient: WorkRestClient, witClient: WorkItemTrackingRestClient, teamContext: TeamContext, nextSprint: TeamSettingsIteration) {
        this._workHttpClient = workHttpClient;
        this._witClient = witClient;
        this._teamContext = teamContext;
        this._nextSprint = nextSprint;
    }

    public async SelectAllWorkItems(iter: TeamSettingsIteration) {
        const items = (await this._workHttpClient.getIterationWorkItems(this._teamContext, iter.id)).workItemRelations;
        const ids = items.map((item) => item.target.id)
        const workItems = await this._witClient.getWorkItems(ids, undefined, undefined, undefined, WorkItemExpand.All);
        const upperRankTypes = ["User Story", "Bug", "Ticket"];
        const parentWIs = workItems.filter(wi => upperRankTypes.includes(wi.fields["System.WorkItemType"]) && wi.fields["System.State"] !== "Closed");
        const lowerRankWorkItems = workItems.filter(wi => wi.fields["System.WorkItemType"] === "Task");
        const parentWorkItems: ParentWorkItem[] = [];
        for (const parentWI of parentWIs) {
            const children: WorkItem[] = lowerRankWorkItems.filter(wI => wI.relations.some(rel => rel.url === parentWI.url));
            parentWorkItems.push(new ParentWorkItem(parentWI, children));
        }
        return parentWorkItems;
    }

    public async ProcessWorkItemsAsync(timeFrame: TeamSettingsIteration) {
        const workItems = await this.SelectAllWorkItems(timeFrame);
        const closeable = [];
        for (const wi of workItems) {
            if (wi.children.every((chWi) => chWi.fields["System.State"] === "New")) { await this.moveParentWorkItemToNextSprint(wi); }
            else if (wi.children.every((chWi) => chWi.fields["System.State"] === "Closed")) { await this.storyResolved(wi); }
            else {
                const copiedParentWI = await this.CopyWIWithParentRelationsAsync(wi.parent);
                const copyPromises = [];

                for (const childrenWI of wi.children) {
                    if (childrenWI.fields["System.State"] !== "Closed") {
                        if (childrenWI.fields["System.State"] === "New") {
                            await this.newTaskToNextSprintStory(copiedParentWI.url, childrenWI.id)
                            wi.children = wi.children.filter((x) => x !== childrenWI)
                        } else {
                            const copyPromise = this.CopyWIWithChildRelationsAsync(childrenWI, copiedParentWI.url);
                            copyPromises.push(copyPromise);
                        }
                    }
                }
                closeable.push(...wi.children, wi.parent);
                await Promise.all(copyPromises);
            }
        }
        await this.CloseWorkItems(closeable)
    }

    public async CopyWIWithChildRelationsAsync(oldWorkItem: WorkItem, parentWIUrl: string): Promise<WorkItem | undefined> {
        const patchDocument: JsonPatchDocument[] = [];
        const systemFields = this.systemFields;
        Object.keys(oldWorkItem.fields).forEach(function (key) {
            if (!systemFields.includes(key)) {
                patchDocument.push({
                    op: Operation.Add,
                    path: `/fields/${key}`,
                    value: oldWorkItem.fields[key]
                });
            }
        });

        patchDocument.push({
            op: Operation.Add,
            path: "/fields/System.Title",
            value: this.ChangeWITitle(oldWorkItem.fields["System.Title"])
        })

        patchDocument.push({
            op: Operation.Add,
            path: "/fields/System.IterationPath",
            value: this._nextSprint.path
        });

        patchDocument.push({
            op: Operation.Add,
            path: "/relations/-",
            value: {
                rel: "System.LinkTypes.Hierarchy-Reverse",
                url: parentWIUrl
            }
        });
        return await this._witClient.createWorkItem(patchDocument, this._teamContext.projectId, oldWorkItem.fields["System.WorkItemType"]);
    }


    public async CopyWIWithParentRelationsAsync(oldWorkItem: WorkItem) {
        let patchDocument: JsonPatchDocument[] = [];
        const systemFields = this.systemFields;
        Object.keys(oldWorkItem.fields).forEach(function (key) {
            if (!systemFields.includes(key)) {
                patchDocument.push({
                    op: Operation.Add,
                    path: "/fields/" + key,
                    value: oldWorkItem.fields[key]
                });
            }
        });
        patchDocument.push({
            op: Operation.Add,
            path: "/fields/System.Title",
            value: this.ChangeWITitle(oldWorkItem.fields["System.Title"])
        });
        patchDocument.push({
            op: Operation.Add,
            path: "/fields/System.IterationPath",
            value: this._nextSprint.path
        });
        patchDocument.push({
            op: Operation.Add,
            path: "/fields/System.State",
            value: "Active"
        });
        if (oldWorkItem.relations) {
            const feature = oldWorkItem.relations.find(relation => relation.rel === "System.LinkTypes.Hierarchy-Reverse");
            if (feature) {
                patchDocument.push({
                    op: Operation.Add,
                    path: "/relations/-",
                    value:
                    {
                        rel: "System.LinkTypes.Hierarchy-Reverse",
                        url: feature.url
                    }
                });
            }
        }
        return await this._witClient.createWorkItem(patchDocument, this._teamContext.projectId, oldWorkItem.fields["System.WorkItemType"]);
    }

    public ChangeWITitle(oldWITitle: string) {
        const match = oldWITitle.match(/\(\d+\)/);
        if (match) {
            const oldIteration = match[0];
            const newIteration = parseInt(
                oldIteration.match(/\d+/)?.[0] ?? '0'
            ) + 1;
            return oldWITitle.replace(oldIteration, '') + `(${newIteration})`;
        } else {
            const sprintUnique = oldWITitle.match(/^.*?(\d+\.\d+)$/);
            if (sprintUnique) {
                const oldIteration = sprintUnique[1];
                const newIteration = this._nextSprint?.name.match(/^.*?(\d+\.\d+)$/);
                const newIter = newIteration ? newIteration[1] : '';
                return oldWITitle.replace(oldIteration, '') + newIter;
            } else {
                return oldWITitle.trimEnd() + ' (1)';
            }
        }
    }

    public async moveParentWorkItemToNextSprint(wI: ParentWorkItem) {
        const sprintUnique = wI.parent.fields["System.Title"].match(/^.*?(\d+\.\d+)$/);
        if (sprintUnique) {
            const newParent = await this.CopyWIWithParentRelationsAsync(wI.parent);
            for (const child of wI.children) {
                await this.newTaskToNextSprintStory(newParent.url, child.id);
            }
        } else {
            for (const workItem of wI.allWorkItems) {
                const patchDocument: JsonPatchDocument[] = [];
                patchDocument.push({
                    op: Operation.Replace,
                    path: "/fields/System.IterationPath",
                    value: this._nextSprint.path
                })
                await this._witClient.updateWorkItem(patchDocument, workItem.id);

            }
        }
    }

    public async newTaskToNextSprintStory(parentUrl: string, childID: number) {
        const patchDocument: JsonPatchDocument[] = [];
        patchDocument.push({
            op: Operation.Replace,
            path: "/fields/System.IterationPath",
            value: this._nextSprint.path
        });
        patchDocument.push({
            op: Operation.Remove,
            path: "/relations/0",
            value: null,
        });
        patchDocument.push({
            op: Operation.Add,
            path: "/relations/0",
            value: {
                rel: "System.LinkTypes.Hierarchy-Reverse",
                url: parentUrl
            }
        });
        await this._witClient.updateWorkItem(patchDocument, childID);
    }

    public async storyResolved(wI: ParentWorkItem) {
        const patchDocument: JsonPatchDocument[] = [];
        patchDocument.push({
            op: Operation.Replace,
            path: "/fields/System.State",
            value: "Resolved"
        });
        await this._witClient.updateWorkItem(patchDocument, wI.parent.id);
    }

    public async CloseWorkItems(items: WorkItem[]): Promise<void> {
        for (const item of items) {
            const patchDocument: JsonPatchDocument[] = [];
            if (item.fields["System.WorkItemType"] !== "Task") {
                patchDocument.push({
                    op: Operation.Replace,
                    path: "/fields/System.State",
                    value: "Resolved"
                });
                await this._witClient.updateWorkItem(patchDocument, item.id);
            } else {
                patchDocument.push({
                    op: Operation.Replace,
                    path: "/fields/System.State",
                    value: "Closed"
                });

                await this._witClient.updateWorkItem(patchDocument, item.id);
            }
        }
    }

}