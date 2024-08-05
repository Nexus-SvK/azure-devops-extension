import {
	type WorkItem,
	WorkItemTrackingRestClient,
} from "azure-devops-extension-api/WorkItemTracking";

export class TeamContext {
	project: string;
	projectId: string;
	team: string;
	teamId: string;

	constructor(
		project: string,
		projectId: string,
		team: string,
		teamId: string,
	) {
		this.project = project;
		this.projectId = projectId;
		this.team = team;
		this.teamId = teamId;
	}
}

export class ParentWorkItem {
	public parent?: WorkItem;
	public children: WorkItem[];
	public allWorkItems: WorkItem[];

	constructor(children: WorkItem[], parent?: WorkItem) {
		this.parent = parent;
		this.children = children;
		this.allWorkItems = parent ? [parent, ...children] : children;
	}
}
