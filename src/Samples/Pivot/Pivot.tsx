import { showRootComponent } from "../../Common";
import "es6-promise/auto";
import "./Pivot.scss";
import React from "react";

import * as SDK from "azure-devops-extension-sdk";
import { WorkRestClient, TimeFrame, TeamSettingsIteration } from "azure-devops-extension-api/Work"
import { WorkItemTrackingRestClient } from "azure-devops-extension-api/WorkItemTracking";
import { getClient } from "azure-devops-extension-api";

import { Button } from "azure-devops-ui/Button";
import { ButtonGroup } from "azure-devops-ui/ButtonGroup";
import { Header, TitleSize } from "azure-devops-ui/Header";
import { Toast } from "azure-devops-ui/Toast";
import { TeamContext } from "./helpers/DataTypeFormats";
import { SprintProcessor } from "./helpers/SprintProcessor";



class PivotContent extends React.Component<{}, {
    sprintClosed: boolean, isToastVisible: boolean,
    canceled: boolean,
    webContext: undefined | any,
    teamContext: undefined | TeamContext,
    witClient: undefined | WorkItemTrackingRestClient,
    workHttpClient: undefined | WorkRestClient,
    current: undefined | TeamSettingsIteration,
    previous: undefined | TeamSettingsIteration,
    future: undefined | TeamSettingsIteration
}> {
    private toastRef: React.RefObject<Toast> = React.createRef<Toast>();
    constructor(props: {}) {
        super(props);
        this.state = {
            isToastVisible: false,
            sprintClosed: false,
            canceled: false,
            webContext: undefined,
            teamContext: undefined,
            witClient: undefined,
            workHttpClient: undefined,
            current: undefined,
            previous: undefined,
            future: undefined
        }

    }

    public async componentDidMount() {
        await SDK.init();
        const webContext = SDK.getWebContext();
        const teamContext = new TeamContext(webContext.project.name, webContext.project.id, webContext.team.name, webContext.team.id);
        const witClient = getClient(WorkItemTrackingRestClient);
        const workHttpClient = getClient(WorkRestClient);
        const iters = await workHttpClient.getTeamIterations(teamContext);
        const current = iters.find((x) => x.attributes.timeFrame === TimeFrame.Current);
        const previous = iters.filter((x) => x.attributes.timeFrame === TimeFrame.Past).pop();
        const future = iters.find((x) => x.attributes.timeFrame === TimeFrame.Future);
        this.setState({ webContext, teamContext, witClient, workHttpClient, current, previous, future });
    }

    private onButtonClick = () => {
        if (this.state.sprintClosed && this.toastRef.current) {
            const toastRef = this.toastRef.current;
            setTimeout(() => {
                toastRef.fadeOut().promise.then(() => {
                    this.setState({ sprintClosed: false });
                });
            }, 2000);
        } else {
            this.setState({ sprintClosed: true });
        }
    };

    private async initializeComponent(timeFrame: TeamSettingsIteration, destination: TeamSettingsIteration) {
        await SDK.ready();
        this.setState({ isToastVisible: true });
        const { workHttpClient, witClient, teamContext } = this.state;
        if (!(workHttpClient && witClient && teamContext)) return;
        const queryExecutor = new SprintProcessor(workHttpClient, witClient, teamContext, destination);
        await queryExecutor.ProcessWorkItemsAsync(timeFrame);
        await SDK.notifyLoadSucceeded();
        this.setState({ isToastVisible: false });
        this.setState({ sprintClosed: true });
        this.onButtonClick()

    }

    public render(): JSX.Element {
        const { isToastVisible, sprintClosed, canceled, current, previous, future } = this.state;
        return (
            <div className="sample-pivot" style={{ display: 'flex', flexDirection: 'column', justifyContent: 'start', alignItems: 'center' }}>
                {canceled ? (<>
                    <Header
                        title="If you want to retry action later click button below"
                        titleSize={TitleSize.Medium}
                        titleAriaLevel={3}
                    />
                    <Button
                        text="Close Sprint"
                        danger={true}
                        onClick={() => { this.setState({ canceled: false }) }}
                    />
                </>)
                    :
                    (<><Header
                        title="Do you want to proceed in closing a Sprint?"
                        titleSize={TitleSize.Medium}
                        titleAriaLevel={3}
                    />
                        <ButtonGroup className="flex-wrap">
                            <Button
                                text={`Close Sprint: ${current?.name ?? ''}`}
                                disabled={!!!(current && future)}
                                onClick={() => {
                                    if (!(current && future)) return;
                                    this.initializeComponent(current, future)
                                }}
                            />
                            <Button
                                text={`Close Previous Sprint: ${previous?.name ?? ''}`}
                                disabled={!!!(previous && current)}
                                onClick={() => {

                                    if (!(previous && current)) return;
                                    this.initializeComponent(previous, current)
                                }}
                            />
                            <Button
                                text="Cancel"
                                danger={true}
                                onClick={() => { this.setState({ canceled: true }) }}
                            />
                        </ButtonGroup>

                        {isToastVisible && (
                            <Toast
                                message="Closing Sprint"
                            />
                        )}
                        {
                            sprintClosed && <Toast
                                ref={this.toastRef}
                                message="Sprint successfully closed"
                            />

                        }
                    </>)
                }
            </div>
        )
    }
}

showRootComponent(<PivotContent />);