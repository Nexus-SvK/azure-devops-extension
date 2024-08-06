import { showRootComponent } from "../../Common";
import "es6-promise/auto";
import "./Pivot.scss";
import React, { useMemo } from "react";
import { createRef, useEffect, useState } from "react";

import * as SDK from "azure-devops-extension-sdk";
import {
	WorkRestClient,
	TimeFrame,
	type TeamSettingsIteration,
} from "azure-devops-extension-api/Work";
import { WorkItemTrackingRestClient } from "azure-devops-extension-api/WorkItemTracking";
import { ScrollableList, type IListItemDetails, ListSelection, ListItem } from "azure-devops-ui/List";
import { Card } from "azure-devops-ui/Card";
import { CommonServiceIds, type IProjectPageService, getClient } from "azure-devops-extension-api";

import { Button } from "azure-devops-ui/Button";
import { ButtonGroup } from "azure-devops-ui/ButtonGroup";
import { Header, TitleSize } from "azure-devops-ui/Header";
import { Toast } from "azure-devops-ui/Toast";
import { TeamContext } from "./helpers/DataTypeFormats";
import { SprintProcessor } from "./helpers/SprintProcessor";
import { ArrayItemProvider } from "azure-devops-ui/Utilities/Provider";

const PivotContent = () => {
	const [sprintClosed, setSprintClosed] = useState<boolean>(false);
	const [isToastVisible, setIsToastVisible] = useState<boolean>(false);
	const [canceled, setCanceled] = useState<boolean>(false);
	const [teamContext, setTeamContext] = useState<undefined | TeamContext>(
		undefined,
	);
	const [witClient, setWitClient] = useState<
		undefined | WorkItemTrackingRestClient
	>(undefined);
	const [workHttpClient, setWorkHttpClient] = useState<
		undefined | WorkRestClient
	>(undefined);
	const [current, setCurrent] = useState<undefined | TeamSettingsIteration>(
		undefined,
	);
	const [previous, setPrevious] = useState<undefined | TeamSettingsIteration>(
		undefined,
	);
	const [future, setFuture] = useState<undefined | TeamSettingsIteration>(
		undefined,
	);
	const [error, setError] = useState<undefined | Error>(undefined);
	const [percentage, setPercentage] = useState<number>(0);
	const toastRef: React.RefObject<Toast> = createRef<Toast>();

	const init = async () => {
		await SDK.init();
		await SDK.ready();
		const config = SDK.getConfiguration();
		const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService)
		const project = await projectService.getProject()
		if (!project) throw new Error("Project not found");
		const teamContext = new TeamContext(
			project.name,
			project.id,
			config.team.name,
			config.team.id,
		);
		const witClient = getClient(WorkItemTrackingRestClient);
		const workHttpClient = getClient(WorkRestClient);
		if (!teamContext) return;
		const iters = await workHttpClient.getTeamIterations(teamContext);
		const current = iters.find(
			(x) => x.attributes.timeFrame === TimeFrame.Current,
		);
		const previous = iters
			.filter((x) => x.attributes.timeFrame === TimeFrame.Past)
			.pop();
		const future = iters.find(
			(x) => x.attributes.timeFrame === TimeFrame.Future,
		);

		// setWebContext(webContext);
		setTeamContext(teamContext);
		setWitClient(witClient);
		setWorkHttpClient(workHttpClient);
		setCurrent(current);
		setPrevious(previous);
		setFuture(future);
	};
	useEffect(() => {
		init();
	}, []);

	const errors = useMemo(() => {
		const errors = localStorage.getItem("errors");
		if (!errors) return [];
		return JSON.parse(errors) as { error: string, workItemId?: number }[];
	}, [])

	const onButtonClick = () => {
		if (sprintClosed && toastRef.current) {
			// const toastRef = toastRef.current;
			setTimeout(() => {
				toastRef.current?.fadeOut().promise.then(() => {
					setSprintClosed(false);
				});
			}, 2000);
		} else {
			setSprintClosed(true);
		}
	};

	const initializeComponent = async (
		timeFrame: TeamSettingsIteration,
		destination: TeamSettingsIteration,
	) => {
		setIsToastVisible(true);
		if (!(workHttpClient && witClient && teamContext)) return;
		const queryExecutor = new SprintProcessor(
			workHttpClient,
			witClient,
			teamContext,
			destination,
		);
		try {
			await queryExecutor.ProcessWorkItemsAsync(
				timeFrame,
				(percentage: number) => {
					setPercentage(percentage);
				},
			);
		} catch (e) {
			const errors = localStorage.getItem("errors");
			if (!errors) { localStorage.setItem("errors", "[]"); }
			else {
				const parsedErrors = JSON.parse(errors);
				parsedErrors.push({ error: (e as Error).message });
				localStorage.setItem("errors", JSON.stringify(parsedErrors));
			}
			console.error(e);
			setError(e as Error);
		}
		// await SDK.notifyLoadSucceeded();
		setIsToastVisible(false);
		setSprintClosed(typeof error === "undefined");
		onButtonClick();
	};

	return (
		<div
			className="sample-pivot close-sprint"
			style={{
				display: "flex",
				flexDirection: "column",
				justifyContent: "start",
				alignItems: "center",
			}}
		>
			{canceled ? (
				<>
					<Header
						title="If you want to retry action later click button below"
						titleSize={TitleSize.Medium}
						titleAriaLevel={3}
					/>
					<Button
						text="Close Sprint"
						danger={true}
						onClick={() => {
							setCanceled(false);
						}}
					/>
				</>
			) : (
				<>
					<Header
						title="Do you want to proceed in closing a Sprint?"
						titleSize={TitleSize.Medium}
						titleAriaLevel={3}
					/>
					<ButtonGroup className="flex-wrap">
						<Button
							text={`Close Sprint: ${current?.name ?? ""}`}
							disabled={!(current && future)}
							onClick={() => {
								if (!(current && future)) return;
								initializeComponent(current, future);
							}}
						/>
						<Button
							text={`Close Previous Sprint: ${previous?.name ?? ""}`}
							disabled={!(previous && current)}
							onClick={() => {
								if (!(previous && current)) return;
								initializeComponent(previous, current);
							}}
						/>
						<Button
							text="Cancel"
							danger={true}
							onClick={() => {
								setCanceled(true);
							}}
						/>
					</ButtonGroup>
					{errors.length > 0 &&
						<Card>
							<div style={{ display: "flex", height: "300px" }}>
								<ScrollableList
									itemProvider={new ArrayItemProvider(errors)}
									renderRow={Row}
									selection={new ListSelection(false)}
									width="100%"
								/>
							</div>
						</Card>}

					{isToastVisible && (
						<Toast message={`Closing Sprint ${percentage}%`} />
					)}
					{sprintClosed && !error && (
						<Toast ref={toastRef} message={"Sprint successfully closed"} />
					)}
					{error && <Toast message={error.message} />}
				</>
			)}
		</div>
	);
};

const Row = (
	index: number,
	item: { error: string, workItemId?: number },
	details: IListItemDetails<{ error: string, workItemId?: number }>,
	key?: string
): JSX.Element => {
	return (
		<ListItem key={key || `list-item${index}`} index={index} details={details}>
			<div className="list-example-row flex-row h-scroll-hidden">
				<div
					style={{ marginLeft: "10px", padding: "10px 0px" }}
					className="flex-column h-scroll-hidden"
				>
					<span className="wrap-text">{item.error}</span>
					<span className="fontSizeMS font-size-ms secondary-text wrap-text">
						{item.workItemId && `WorkItem: ${item.workItemId}`}
					</span>
				</div>
			</div>
		</ListItem>
	);
};

showRootComponent(<PivotContent />);
