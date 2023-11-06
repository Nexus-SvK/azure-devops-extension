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
        try {
            // Fetch work items related to the iteration
            const iterationWorkItems = await this._workHttpClient.getIterationWorkItems(this._teamContext, iter.id);
            const workItemRelations = iterationWorkItems.workItemRelations;
            const workItemIds = workItemRelations.map((item) => item.target.id);

            // Fetch details of the work items
            const workItems = await this._witClient.getWorkItems(workItemIds, undefined, undefined, undefined, WorkItemExpand.All);

            // Filter parent and lower rank work items
            const upperRankTypes = ["User Story", "Bug", "Ticket"];
            const parentWorkItems = workItems.filter(wi => upperRankTypes.includes(wi.fields["System.WorkItemType"]) && wi.fields["System.State"] !== "Closed");
            const lowerRankWorkItems = workItems.filter(wi => wi.fields["System.WorkItemType"] === "Task");

            // Group lower rank work items under their respective parent work items
            const parentWorkItemsWithChildren = parentWorkItems.map(parentWI => {
                const children = lowerRankWorkItems.filter(wi => wi.relations.some(rel => rel.url === parentWI.url));
                return new ParentWorkItem(children, parentWI);
            });

            return parentWorkItemsWithChildren;
        } catch (error) {
            // Handle specific errors and provide more informative error messages
            if (error instanceof Error) {
                throw new Error(`"SelectAllWorkItems": Failed to fetch work items: ${error.message}`);
            } else {
                throw new Error(`"SelectAllWorkItems": Failed to fetch work items`);
            }
        }
    }


    public async ProcessWorkItemsAsync(timeFrame: TeamSettingsIteration, callback: Function) {
        const workItems = await this.SelectAllWorkItems(timeFrame);
        const allWorkItems = workItems.reduce((x, i) => x + i.allWorkItems.length, 0);
        let completed = 0;
        const closeable = [];
        for (const wi of workItems) {
            if (wi.children.every((chWi) => chWi.fields["System.State"] === "New" && typeof wi.parent !== 'undefined')) { await this.moveParentWorkItemToNextSprint(wi) }
            else if (wi.children.every((chWi) => chWi.fields["System.State"] === "Closed")) { await this.storyResolved(wi) }
            else {
                if (wi.parent) {
                    const copiedParentWI = await this.CopyWIWithParentRelationsAsync(wi.parent);
                    const copyPromises = [];

                    for (const childrenWI of wi.children) {
                        if (childrenWI.fields["System.State"] !== "Closed") {
                            if (childrenWI.fields["System.State"] === "New") {
                                await this.newTaskToNextSprintStory(copiedParentWI.url, childrenWI.id)
                                wi.allWorkItems = wi.allWorkItems.filter((x) => x.id !== childrenWI.id);
                            } else {
                                const copyPromise = this.CopyWIWithChildRelationsAsync(childrenWI, copiedParentWI.url);
                                copyPromises.push(copyPromise);
                            }
                        }
                    }
                    closeable.push(...wi.allWorkItems);
                    await Promise.all(copyPromises);
                }
            }
            completed += wi.allWorkItems.length;
            callback(Math.trunc((completed / allWorkItems) * 100))
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
        try {
            return await this._witClient.createWorkItem(patchDocument, this._teamContext.projectId, oldWorkItem.fields["System.WorkItemType"]);
        } catch (e) {
            if (e instanceof Error) {
                throw new Error(`CopyWIWithChildRelationsAsync: ${e.message}`);
            } else {
                throw new Error(`Error in CopyWIWithChildRelationsAsync - WorkItem ${oldWorkItem.id}`);
            }
        }
    }


    public async CopyWIWithParentRelationsAsync(oldWorkItem: WorkItem) {
        let updateOldDocument: JsonPatchDocument[] = [];

        updateOldDocument.push({
            op: Operation.Replace,
            path: "/fields/System.Title",
            value: `${oldWorkItem.fields["System.Title"]} ->`
        });

        try {
            await this._witClient.updateWorkItem(updateOldDocument, oldWorkItem.id);
        } catch (e) {
            if (e instanceof Error) {
                throw new Error(`CopyWIWithParentRelationsAsync - Failed to update: ${e.message}`);
            } else {
                throw new Error(`Error in CopyWIWithParentRelationsAsync - Failed to update WorkItem ${oldWorkItem.id}`);
            }
        }

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
        } try {
            return await this._witClient.createWorkItem(patchDocument, this._teamContext.projectId, oldWorkItem.fields["System.WorkItemType"]);
        } catch (e) {
            if (e instanceof Error) {
                throw new Error(`CopyWIWithParentRelationsAsync: ${e.message}`);
            } else {
                throw new Error(`Error in CopyWIWithParentRelationsAsync - WorkItem ${oldWorkItem.id}`);
            }
        }
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
        if (!wI.parent) {
            throw new Error("moveParentWorkItemToNextSprint: No Parent Work Item");
        }
        const sprintUnique = wI.parent?.fields["System.Title"].match(/^.*?(\d+\.\d+)$/);
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
                try {
                    await this._witClient.updateWorkItem(patchDocument, workItem.id);
                } catch (e) {
                    if (e instanceof Error) {
                        throw new Error(`moveParentWorkItemToNextSprint: ${e.message}`);
                    } else {
                        throw new Error('Error in moveParentWorkItemToNextSprint');
                    }
                }

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
        try {
            await this._witClient.updateWorkItem(patchDocument, childID);
        } catch (e) {
            if (e instanceof Error) {
                throw new Error(`newTaskToNextSprintStory: ${e.message}`);
            } else {
                throw new Error(`Error in newTaskToNextSprintStory`);
            }
        }
    }

    public async storyResolved(wI: ParentWorkItem) {
        try {
            const patchDocument: JsonPatchDocument[] = [];
            patchDocument.push({
                op: Operation.Replace,
                path: "/fields/System.State",
                value: "Resolved"
            });
            if (!wI.parent) throw new Error('Parent WorkItem undefined');
            await this._witClient.updateWorkItem(patchDocument, wI.parent?.id);
        } catch (e) {
            if (e instanceof Error) {
                throw new Error(`storyResolved: ${e.message}`)
            } else {
                throw new Error('Error in storyResolved');
            }
        }
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
                try {
                    await this._witClient.updateWorkItem(patchDocument, item.id);
                } catch (e) {
                    if (e instanceof Error) {
                        throw new Error(`CloseWorkItems: ${e.message}`)
                    } else {
                        throw new Error(`Error in CloseWorkItems`)
                    }
                }
            } else {
                patchDocument.push({
                    op: Operation.Replace,
                    path: "/fields/System.State",
                    value: "Closed"
                });

                try {
                    await this._witClient.updateWorkItem(patchDocument, item.id);
                } catch (e) {
                    if (e instanceof Error) {
                        throw new Error(`CloseWorkItems: ${e.message}`)
                    } else {
                        throw new Error(`Error in CloseWorkItems`)
                    }
                }
            }
        }
    }

}