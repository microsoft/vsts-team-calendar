import React = require("react");

import { TeamMember } from "azure-devops-extension-api/WebApi/WebApi";


import { getUser } from "azure-devops-extension-sdk";

import { Button } from "azure-devops-ui/Button";
import { ButtonGroup } from "azure-devops-ui/ButtonGroup";
import { CustomDialog } from "azure-devops-ui/Dialog";
import { Dropdown } from "azure-devops-ui/Dropdown";
import { TitleSize } from "azure-devops-ui/Header";
import { IListSelection } from "azure-devops-ui/List";
import { IListBoxItem } from "azure-devops-ui/ListBox";
import { MessageCard, MessageCardSeverity } from "azure-devops-ui/MessageCard";
import { ObservableValue } from "azure-devops-ui/Core/Observable";
import { Observer } from "azure-devops-ui/Observer";
import { PanelHeader, PanelFooter, PanelContent } from "azure-devops-ui/Panel";
import { DropdownSelection } from "azure-devops-ui/Utilities/DropdownSelection";

import { Calendar } from "@fullcalendar/core";

import { ICalendarEvent } from "./Contracts";
import { MessageDialog } from "./MessageDialog";
import { toDate, formatDate } from "./TimeLib";
import { VSOCapacityEventSource, Everyone } from "./VSOCapacityEventSource";
import { TeamSettingsIteration } from "azure-devops-extension-api/Work";

interface IAddEditDaysOffDialogProps {
    /**
     * Calendar api to add event to the Calendar
     */
    calendarApi: Calendar;

    /**
     * End date for event
     */
    end: Date;

    /**
     * Event object if editing an event.
     */
    event?: ICalendarEvent;

    /**
     * Object that stores all event data
     */
    eventSource: VSOCapacityEventSource;

    /**
     * List of members in currently selected team.
     */
    members: TeamMember[];

    /**
     * Callback function on dialog dismiss
     */
    onDismiss: () => void;

    /**
     * Start date for event
     */
    start: Date;
}

/**
 * Dialog that lets user add new days off
 */
export class AddEditDaysOffDialog extends React.Component<IAddEditDaysOffDialogProps> {
    endDate: Date;
    isConfirmationDialogOpen: ObservableValue<boolean>;
    iteration?: TeamSettingsIteration;
    memberSelection: IListSelection;
    message: ObservableValue<string>;
    okButtonEnabled: ObservableValue<boolean>;
    selectedMemberId: string;
    selectedMemberName: string;
    startDate: Date;
    teamMembers: IListBoxItem[];

    constructor(props: IAddEditDaysOffDialogProps) {
        super(props);

        this.okButtonEnabled = new ObservableValue<boolean>(true);
        this.message = new ObservableValue<string>("");
        this.memberSelection = new DropdownSelection();
        this.teamMembers = [];
        this.isConfirmationDialogOpen = new ObservableValue<boolean>(false);

        let selectedIndex = 0;
        if (this.props.event) {
            this.startDate = new Date(this.props.event.startDate);
            this.endDate = new Date(this.props.event.endDate);
            this.teamMembers.push({ id: this.props.event.member!.id, text: this.props.event.member!.displayName });
        } else {
            this.startDate = props.start;
            this.endDate = props.end;
            const userName = getUser().displayName;
            let i = 1;
            this.teamMembers.push({ id: Everyone, text: Everyone });
            this.teamMembers.push(
                ...this.props.members.map(item => {
                    if (userName === item.identity.displayName) {
                        selectedIndex = i;
                    }
                    i++;
                    return { id: item.identity.id, text: item.identity.displayName };
                })
            );
        }

        this.memberSelection.select(selectedIndex);
        this.selectedMemberId = this.teamMembers[selectedIndex].id;
        this.selectedMemberName = this.teamMembers[selectedIndex].text!;

        this.validateSelections();
    }

    public render(): JSX.Element {
        return (
            <>
                <CustomDialog onDismiss={this.props.onDismiss}>
                    <PanelHeader
                        onDismiss={this.props.onDismiss}
                        showCloseButton={false}
                        titleProps={{ size: TitleSize.Small, text: this.props.event ? "Edit days off" : "Add days off" }}
                    />
                    <PanelContent>
                        <div className="flex-grow flex-column event-dialog-content">
                            <Observer message={this.message}>
                                {(props: { message: string }) => {
                                    return props.message !== "" ? (
                                        <MessageCard className="flex-self-stretch" severity={MessageCardSeverity.Info}>
                                            {props.message}
                                        </MessageCard>
                                    ) : null;
                                }}
                            </Observer>
                            <div className="input-row flex-row">
                                <span>Start Date</span>
                                <div className="bolt-textfield column-2">
                                    <input
                                        className="bolt-textfield-input input-date"
                                        defaultValue={formatDate(this.startDate, "YYYY-MM-DD")}
                                        onChange={this.onInputStartDate}
                                        type="date"
                                    />
                                </div>
                            </div>
                            <div className="input-row flex-row">
                                <span>End Date</span>
                                <div className="bolt-textfield column-2">
                                    <input
                                        className="bolt-textfield-input input-date"
                                        defaultValue={formatDate(this.endDate, "YYYY-MM-DD")}
                                        onChange={this.onInputEndDate}
                                        type="date"
                                    />
                                </div>
                            </div>
                            <div className="input-row flex-row">
                                <span>Team Member</span>
                                <Dropdown
                                    className="column-2"
                                    items={this.teamMembers}
                                    onSelect={this.onSelectTeamMember}
                                    selection={this.memberSelection}
                                />
                            </div>
                        </div>
                    </PanelContent>
                    <PanelFooter>
                        <div className="flex-grow flex-row">
                            {this.props.event && <Button onClick={this.onDeleteClick} subtle={true} text="Delete days off" />}
                            <ButtonGroup className="bolt-panel-footer-buttons flex-grow">
                                <Button onClick={this.props.onDismiss} text="Cancel" />
                                <Observer enabled={this.okButtonEnabled}>
                                    {(props: { enabled: boolean }) => {
                                        return <Button disabled={!props.enabled} onClick={this.onOKClick} primary={true} text="Ok" />;
                                    }}
                                </Observer>
                            </ButtonGroup>
                        </div>
                    </PanelFooter>
                </CustomDialog>
                <Observer isDialogOpen={this.isConfirmationDialogOpen}>
                    {(props: { isDialogOpen: boolean }) => {
                        return props.isDialogOpen ? (
                            <MessageDialog
                                message="Are you sure you want to delete the days off?"
                                onConfirm={() => {
                                    this.props.eventSource.deleteEvent(this.props.event!, this.props.event!.iterationId!).then(() => {
                                        this.props.calendarApi.refetchEvents();
                                    });
                                    this.isConfirmationDialogOpen.value = false;
                                    this.props.onDismiss();
                                }}
                                onDismiss={() => {
                                    this.isConfirmationDialogOpen.value = false;
                                }}
                                title="Delete days off"
                            />
                        ) : null;
                    }}
                </Observer>
            </>
        );
    }

    private onDeleteClick = async (): Promise<void> => {
        this.isConfirmationDialogOpen.value = true;
    };

    private onInputEndDate = (e: React.ChangeEvent<HTMLInputElement>): void => {
        let temp = e.target.value;
        if (temp) {
            this.endDate = toDate(temp);
        }
        this.validateSelections();
    };

    private onInputStartDate = (e: React.ChangeEvent<HTMLInputElement>): void => {
        let temp = e.target.value;
        if (temp) {
            this.startDate = toDate(temp);
        }
        this.validateSelections();
    };

    private onOKClick = (): void => {
        let promise;
        if (this.props.event) {
            promise = this.props.eventSource.updateEvent(this.props.event, this.props.event.iterationId!, this.startDate, this.endDate);
        } else {
            promise = this.props.eventSource.addEvent(
                this.iteration!.id,
                this.startDate,
                this.endDate,
                this.selectedMemberName,
                this.selectedMemberId
            );
        }
        promise.then(() => {
            this.props.calendarApi.refetchEvents();
        });
        this.props.onDismiss();
    };

    private onSelectTeamMember = (event: React.SyntheticEvent<HTMLElement>, item: IListBoxItem<{}>) => {
        this.selectedMemberName = item.text!;
        this.selectedMemberId = item.id;
    };

    private validateSelections = () => {
        let valid: boolean = this.startDate <= this.endDate;
        // start date and end date should be in same iteration
        this.iteration = this.props.eventSource.getIterationForDate(this.startDate, this.endDate);
        valid = valid && !!this.iteration;

        if (valid) {
            if (this.message.value !== "") {
                this.message.value = "";
            }
        } else {
            if (this.startDate > this.endDate) {
                this.message.value = "Start date must be same or before the end date.";
            } else {
                this.message.value = "Selected dates are not part of any or same Iteration.";
            }
        }
        this.okButtonEnabled.value = valid;
    };
}
