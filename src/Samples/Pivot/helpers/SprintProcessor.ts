import type {
	WorkRestClient,
	TeamSettingsIteration,
} from "azure-devops-extension-api/Work";
import {
	type WorkItem,
	WorkItemExpand,
	type WorkItemTrackingRestClient,
} from "azure-devops-extension-api/WorkItemTracking";
import {
	type JsonPatchDocument,
	Operation,
} from "azure-devops-extension-api/WebApi";
import type { TeamContext } from "azure-devops-extension-api/Core/Core";
import { ParentWorkItem } from "./DataTypeFormats";

export class SprintProcessor {
	public _workHttpClient: WorkRestClient;
	public _witClient: WorkItemTrackingRestClient;
	public _teamContext: TeamContext;
	public _nextSprint: TeamSettingsIteration;
	public readonly systemFields: string[] = [
		"System.IterationId",
		"System.ExternalLinkCount",
		"System.HyperLinkCount",
		"System.AttachedFileCount",
		"System.NodeName",
		"System.RevisedDate",
		"System.ChangedDate",
		"System.Id",
		"System.AreaId",
		"System.AuthorizedAs",
		"System.State",
		"System.AuthorizedDate",
		"System.Watermark",
		"System.Rev",
		"System.ChangedBy",
		"System.Reason",
		"System.WorkItemType",
		"System.CreatedDate",
		"System.CreatedBy",
		"System.History",
		"System.RelatedLinkCount",
		"System.BoardColumn",
		"System.BoardColumnDone",
		"System.BoardLane",
		"System.CommentCount",
		"System.TeamProject",
		"System.AreaLevel1",
		"System.IterationLevel1",
		"System.IterationLevel2",
		"Microsoft.VSTS.Common.StateChangeDate",
		"Microsoft.VSTS.Common.ActivatedDate",
		"Microsoft.VSTS.Common.ActivatedBy",
		"System.AreaPath",
		"Microsoft.VSTS.Scheduling.CompletedWork",
		"System.IterationPath",
		"System.Title",
		"Microsoft.VSTS.Common.ClosedBy",
		"Microsoft.VSTS.Common.ClosedDate",
	];

	constructor(
		workHttpClient: WorkRestClient,
		witClient: WorkItemTrackingRestClient,
		teamContext: TeamContext,
		nextSprint: TeamSettingsIteration,
	) {
		this._workHttpClient = workHttpClient;
		this._witClient = witClient;
		this._teamContext = teamContext;
		this._nextSprint = nextSprint;
		const errors = localStorage.getItem("errors");
		if (!errors) localStorage.setItem("errors", "[]");
	}

	public async SelectAllWorkItems(iter: TeamSettingsIteration) {
		try {
			// Fetch work items related to the iteration
			const iterationWorkItems =
				await this._workHttpClient.getIterationWorkItems(
					this._teamContext,
					iter.id,
				);
			const workItemRelations = iterationWorkItems.workItemRelations;
			const workItemIds = workItemRelations.map((item) => item.target.id);

			// Fetch details of the work items
			const workItems = await this._witClient.getWorkItems(
				workItemIds,
				undefined,
				undefined,
				undefined,
				WorkItemExpand.All,
			);

			// Filter parent and lower rank work items
			const upperRankTypes = ["User Story", "Bug", "Ticket"];
			const parentWorkItems = workItems.filter(
				(wi) =>
					upperRankTypes.includes(wi.fields["System.WorkItemType"]) &&
					wi.fields["System.State"] !== "Closed",
			);
			const lowerRankWorkItems = workItems.filter(
				(wi) => wi.fields["System.WorkItemType"] === "Task",
			);

			// Group lower rank work items under their respective parent work items
			const parentWorkItemsWithChildren = parentWorkItems.map((parentWI) => {
				const children = lowerRankWorkItems.filter((wi) =>
					wi.relations?.some((rel) => rel.url === parentWI.url),
				);
				return new ParentWorkItem(children, parentWI);
			});

			return parentWorkItemsWithChildren;
		} catch (error) {
			throw new Error(
				`SelectAllWorkItems: Failed to fetch work items: ${(error as Error).message}`,
			);
		}
	}

	public async ProcessWorkItemsAsync(
		timeFrame: TeamSettingsIteration,
		callback: (percentage: number) => void,
	) {
		const workItems: ParentWorkItem[] = [];
		try {
			const allWorkItems = await this.SelectAllWorkItems(timeFrame);
			workItems.push(...allWorkItems);
		} catch (e) {
			this.setError({ error: (e as Error).message });
			console.error((e as Error).message);
		}
		const allWorkItems = workItems.reduce(
			(x, i) => x + i.allWorkItems.length,
			0,
		);
		let completed = 0;
		for (const wi of workItems) {
			if (
				wi.children.every(
					(chWi) =>
						chWi.fields["System.State"] === "New" &&
						typeof wi.parent !== "undefined",
				)
			) {
				try {
					await this.moveParentWorkItemToNextSprint(wi);
				} catch (e) {
					this.setError({ error: (e as Error).message });
					console.error((e as Error).message);
				}
				completed += wi.allWorkItems.length;
			} else if (
				wi.children.every((chWi) => chWi.fields["System.State"] === "Closed")
			) {
				try {
					await this.storyResolved(wi);
				} catch (e) {
					this.setError({ error: (e as Error).message });
					console.error((e as Error).message);
				}
				completed += wi.allWorkItems.length;
			} else {
				if (wi.parent) {
					try {
						const copiedParentWI = await this.CopyWIWithParentRelationsAsync(
							wi.parent,
						);

						for (const childrenWI of wi.children) {
							if (childrenWI.fields["System.State"] !== "Closed") {
								if (childrenWI.fields["System.State"] === "New") {
									try {
										await this.newTaskToNextSprintStory(
											copiedParentWI.url,
											childrenWI.id,
										);
									} catch (e) {
										this.setError({ error: (e as Error).message });
										console.error((e as Error).message);
									}
									// wi.allWorkItems = wi.allWorkItems.filter(
									// 	(x) => x.id !== childrenWI.id,
									// );
								} else {
									let failedToCopy = false;
									try {
										await this.CopyWIWithChildRelationsAsync(
											childrenWI,
											copiedParentWI.url,
										);
									} catch (e) {
										failedToCopy = true;
										this.setError({ error: (e as Error).message });
										console.error((e as Error).message);
									}
									try {
										!failedToCopy && (await this.CloseWorkItem(childrenWI));
									} catch (e) {
										this.setError({ error: (e as Error).message });
										console.error((e as Error).message);
									}
									console.log("Closing children");
								}
							}
							completed++;
						}
						try {
							await this.CloseWorkItem(wi.parent);
						} catch (e) {
							this.setError({ error: (e as Error).message });
							console.error((e as Error).message);
						}
						console.log("Closing parent");
						completed++;
					} catch (e) {
						this.setError({ error: (e as Error).message });
						console.error((e as Error).message);
					}
				}
			}
			// completed++;
			callback(Math.trunc((completed / allWorkItems) * 100));
		}
		// await this.CloseWorkItems(closeable);
	}

	public async CopyWIWithChildRelationsAsync(
		oldWorkItem: WorkItem,
		parentWIUrl: string,
	): Promise<WorkItem | undefined> {
		const patchDocument: JsonPatchDocument[] = [];
		const systemFields = this.systemFields;
		Object.keys(oldWorkItem.fields).forEach((key) => {
			if (!systemFields.includes(key)) {
				patchDocument.push({
					op: Operation.Add,
					path: `/fields/${key}`,
					value: oldWorkItem.fields[key],
				});
			}
		});

		patchDocument.push({
			op: Operation.Add,
			path: "/fields/System.Title",
			value: this.ChangeWITitle(oldWorkItem.fields["System.Title"]),
		});

		patchDocument.push({
			op: Operation.Add,
			path: "/fields/System.IterationPath",
			value: this._nextSprint.path,
		});

		patchDocument.push({
			op: Operation.Add,
			path: "/relations/-",
			value: {
				rel: "System.LinkTypes.Hierarchy-Reverse",
				url: parentWIUrl,
			},
		});
		try {
			return await this._witClient.createWorkItem(
				patchDocument,
				this._teamContext.projectId,
				oldWorkItem.fields["System.WorkItemType"],
			);
		} catch (e) {
			throw new Error(`${oldWorkItem.id}: ${(e as Error).message}`);
		}
	}

	public async CopyWIWithParentRelationsAsync(oldWorkItem: WorkItem) {
		const updateOldDocument: JsonPatchDocument[] = [];

		updateOldDocument.push({
			op: Operation.Replace,
			path: "/fields/System.Title",
			value: `${oldWorkItem.fields["System.Title"]} ->`,
		});

		try {
			await this._witClient.updateWorkItem(updateOldDocument, oldWorkItem.id);
		} catch (e) {
			throw new Error(`${oldWorkItem.id}: ${(e as Error).message}`);
		}

		const patchDocument: JsonPatchDocument[] = [];
		const systemFields = this.systemFields;
		Object.keys(oldWorkItem.fields).forEach(function (key) {
			if (!systemFields.includes(key)) {
				patchDocument.push({
					op: Operation.Add,
					path: "/fields/" + key,
					value: oldWorkItem.fields[key],
				});
			}
		});
		patchDocument.push({
			op: Operation.Add,
			path: "/fields/System.Title",
			value: this.ChangeWITitle(oldWorkItem.fields["System.Title"]),
		});
		patchDocument.push({
			op: Operation.Add,
			path: "/fields/System.IterationPath",
			value: this._nextSprint.path,
		});
		patchDocument.push({
			op: Operation.Add,
			path: "/fields/System.State",
			value: "Active",
		});
		if (oldWorkItem.relations) {
			const feature = oldWorkItem.relations.find(
				(relation) => relation.rel === "System.LinkTypes.Hierarchy-Reverse",
			);
			if (feature) {
				patchDocument.push({
					op: Operation.Add,
					path: "/relations/-",
					value: {
						rel: "System.LinkTypes.Hierarchy-Reverse",
						url: feature.url,
					},
				});
			}
		}
		try {
			return await this._witClient.createWorkItem(
				patchDocument,
				this._teamContext.projectId,
				oldWorkItem.fields["System.WorkItemType"],
			);
		} catch (e) {
			throw new Error(`${oldWorkItem.id}: ${(e as Error).message}`);
		}
	}

	public ChangeWITitle(oldWITitle: string) {
		const match = oldWITitle.match(/\(\d+\)/);
		const sprintUnique = oldWITitle.match(/^.*?(\d+\.\d+)$/);
		if (match) {
			const oldIteration = match[0];
			const newIteration =
				Number.parseInt(oldIteration.match(/\d+/)?.[0] ?? "0") + 1;
			return `${oldWITitle.replace(oldIteration, "")}(${newIteration})`;
		} else if (sprintUnique) {
			const oldIteration = sprintUnique[1];
			const newIteration = this._nextSprint?.name.match(/^.*?(\d+\.\d+)$/);
			const newIter = newIteration ? newIteration[1] : "";
			return oldWITitle.replace(oldIteration, "") + newIter;
		} else {
			return oldWITitle.trimEnd() + " (1)";
		}
	}

	public async moveParentWorkItemToNextSprint(wI: ParentWorkItem) {
		if (!wI.parent) {
			throw new Error("moveParentWorkItemToNextSprint: No Parent Work Item");
		}
		const sprintUnique =
			wI.parent?.fields["System.Title"].match(/^.*?(\d+\.\d+)$/);
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
					value: this._nextSprint.path,
				});
				try {
					await this._witClient.updateWorkItem(patchDocument, workItem.id);
				} catch (e) {
					throw new Error(
						`moveParentWorkItemToNextSprint: ${(e as Error).message}`,
					);
				}
			}
		}
	}

	public async newTaskToNextSprintStory(parentUrl: string, childID: number) {
		const patchDocument: JsonPatchDocument[] = [];
		patchDocument.push({
			op: Operation.Replace,
			path: "/fields/System.IterationPath",
			value: this._nextSprint.path,
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
				url: parentUrl,
			},
		});
		try {
			await this._witClient.updateWorkItem(patchDocument, childID);
		} catch (e) {
			throw new Error(`newTaskToNextSprintStory: ${(e as Error).message}`);
		}
	}

	public async storyResolved(wI: ParentWorkItem) {
		try {
			const patchDocument: JsonPatchDocument[] = [];
			patchDocument.push({
				op: Operation.Replace,
				path: "/fields/System.State",
				value: "Resolved",
			});
			if (!wI.parent) throw new Error("storyResolved: No Parent Work Item");
			await this._witClient.updateWorkItem(patchDocument, wI.parent?.id);
		} catch (e) {
			throw new Error(`storyResolved: ${(e as Error).message}`);
		}
	}

	public async CloseWorkItems(items: WorkItem[]): Promise<void> {
		for (const item of items) {
			const patchDocument: JsonPatchDocument[] = [];
			if (item.fields["System.WorkItemType"] !== "Task") {
				patchDocument.push({
					op: Operation.Replace,
					path: "/fields/System.State",
					value: "Resolved",
				});
				try {
					await this._witClient.updateWorkItem(patchDocument, item.id);
				} catch (e) {
					this.setError({ error: (e as Error).message });
					console.error(`CloseWorkItems: ${(e as Error).message}`);
					throw new Error(`CloseWorkItems: ${(e as Error).message}`);
				}
			} else {
				patchDocument.push({
					op: Operation.Replace,
					path: "/fields/System.State",
					value: "Closed",
				});

				try {
					await this._witClient.updateWorkItem(patchDocument, item.id);
				} catch (e) {
					throw new Error(`CloseWorkItems: ${(e as Error).message}`);
				}
			}
		}
	}

	public async CloseWorkItem(item: WorkItem): Promise<void> {
		const patchDocument: JsonPatchDocument[] = [];
		if (item.fields["System.WorkItemType"] !== "Task") {
			patchDocument.push({
				op: Operation.Replace,
				path: "/fields/System.State",
				value: "Resolved",
			});
			try {
				await this._witClient.updateWorkItem(patchDocument, item.id);
			} catch (error) {
				const errorMessage = `CloseWorkItems: Failed to update work item ${item.id}: ${(error as Error).message}`;
				console.error(errorMessage);
				throw new Error(errorMessage);
			}
		} else {
			patchDocument.push({
				op: Operation.Replace,
				path: "/fields/System.State",
				value: "Closed",
			});

			try {
				await this._witClient.updateWorkItem(patchDocument, item.id);
			} catch (error) {
				throw new Error(
					`CloseWorkItem: ${item.id}: ${(error as Error).message}`,
				);
			}
		}
	}

	public setError(e: { workItemId?: number; error: string }) {
		const errors = localStorage.getItem("errors");
		if (!errors) {
			localStorage.setItem("errors", JSON.stringify([e]));
		} else {
			const parsedErrors = JSON.parse(errors);
			parsedErrors.push(e);
			localStorage.setItem("errors", JSON.stringify(parsedErrors));
		}
	}
}
