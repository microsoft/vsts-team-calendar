import React = require("react");
import DatePicker from "react-datepicker";
import "react-datepicker/dist/react-datepicker.css";

import { Button } from "azure-devops-ui/Button";
import { ButtonGroup } from "azure-devops-ui/ButtonGroup";
import { CustomDialog } from "azure-devops-ui/Dialog";
import { EditableDropdown } from "azure-devops-ui/EditableDropdown";
import { Icon } from "azure-devops-ui/Icon";
import { TitleSize } from "azure-devops-ui/Header";
import { IListSelection, ListSelection } from "azure-devops-ui/List";
import { IListBoxItem } from "azure-devops-ui/ListBox";
import { MessageCard, MessageCardSeverity } from "azure-devops-ui/MessageCard";
import { ObservableValue } from "azure-devops-ui/Core/Observable";
import { Observer } from "azure-devops-ui/Observer";
import { PanelHeader, PanelFooter, PanelContent } from "azure-devops-ui/Panel";
import { TextField } from "azure-devops-ui/TextField";

import { Calendar, EventApi } from "@fullcalendar/core";

import { FreeFormEventsSource } from "./FreeFormEventSource";
import { MessageDialog } from "./MessageDialog";

interface IAddEditEventDialogProps {
    /**
     * Calendar api to add event to the Calendar
     */
    calendarApi: Calendar;

    /**
     * Event api to update ui.
     */
    eventApi?: EventApi;

    /**
     * End date for event
     */
    end: Date;

    /**
     * Object that stores all event data
     */
    eventSource: FreeFormEventsSource;

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
 * Dialog that lets user add new event
 */
export class AddEditEventDialog extends React.Component<IAddEditEventDialogProps> {
    startDate: ObservableValue<Date>;
    endDate: ObservableValue<Date>;
    isConfirmationDialogOpen: ObservableValue<boolean>;
    isDatePickerOpen: ObservableValue<boolean>;
    okButtonEnabled: ObservableValue<boolean>;
    title: ObservableValue<string>;
    description: ObservableValue<string>;
    category: string;
    message: ObservableValue<string>;
    catagorySelection: IListSelection;
    startDatePickerRef: React.RefObject<any>;
    endDatePickerRef: React.RefObject<any>;

    constructor(props: IAddEditEventDialogProps) {
        super(props);
        this.startDatePickerRef = React.createRef();
        this.endDatePickerRef = React.createRef();
        this.catagorySelection = new ListSelection();
        if (this.props.eventApi) {
            this.startDate = new ObservableValue<Date>(this.props.eventApi.start!);
            // api end date is +1 day of actual end date
            if (this.props.eventApi.end) {
                const endDate = new Date(this.props.eventApi.end);
                endDate.setDate(this.props.eventApi.end.getDate() - 1);
                this.endDate = new ObservableValue<Date>(endDate);
            } else {
                this.endDate = new ObservableValue<Date>(new Date(this.props.eventApi.start!));
            }
            this.title = new ObservableValue<string>(this.props.eventApi.title);
            this.description = new ObservableValue<string>(this.props.eventApi.extendedProps.description || "");
            this.category = this.props.eventApi.extendedProps.category;
            this.catagorySelection.select(0);
        } else {
            this.startDate = new ObservableValue<Date>(props.start);
            this.endDate = new ObservableValue<Date>(props.end);
            this.title = new ObservableValue<string>("");
            this.description = new ObservableValue<string>("");
            this.category = "";
        }
        this.okButtonEnabled = new ObservableValue<boolean>(false);
        this.isConfirmationDialogOpen = new ObservableValue<boolean>(false);
        this.isDatePickerOpen = new ObservableValue<boolean>(false);
        this.message = new ObservableValue<string>("");
    }

    public render(): JSX.Element {
        return (
            <>
                <CustomDialog onDismiss={this.props.onDismiss}>
                    <PanelHeader
                        onDismiss={this.props.onDismiss}
                        showCloseButton={false}
                        titleProps={{ size: TitleSize.Small, text: this.props.eventApi ? "Edit event" : "Add event" }}
                    />
                    <PanelContent>
                        <Observer isDatePickerOpen={this.isDatePickerOpen}>
                            {(props: { isDatePickerOpen: boolean }) => (
                                <div className={`flex-grow flex-column event-dialog-content ${props.isDatePickerOpen ? 'picker-open' : ''}`}>
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
                                <span>Title</span>
                                <TextField className="column-2" onChange={this.onInputTitle} value={this.title} />
                            </div>
                            <div className="input-row flex-row">
                                <span>Category</span>
                                <EditableDropdown
                                    allowFreeform={true}
                                    className="column-2"
                                    items={Array.from(this.props.eventSource.getCategories())}
                                    onValueChange={this.onCatagorySelectionChange}
                                    placeholder={this.category}
                                />
                            </div>
                            <div className="input-row flex-row">
                                <span>Description</span>
                                <TextField className="column-2" onChange={this.onInputDescription} multiline={true} value={this.description} />
                            </div>
                            <div className="input-row flex-row">
                                <span>Start Date</span>
                                <div className="column-2 date-picker-wrapper">
                                    <Observer startDate={this.startDate}>
                                        {(props: { startDate: Date }) => (
                                            <>
                                                <DatePicker
                                                    ref={this.startDatePickerRef}
                                                    selected={props.startDate}
                                                    onChange={(date: Date | null) => {
                                                        if (date) {
                                                            this.startDate.value = date;
                                                            this.validateSelections();
                                                        }
                                                    }}
                                                    onCalendarOpen={() => this.isDatePickerOpen.value = true}
                                                    onCalendarClose={() => this.isDatePickerOpen.value = false}
                                                    dateFormat="MM/dd/yyyy"
                                                    className="bolt-textfield-input input-date"
                                                    
                                                />
                                                <Icon 
                                                    className="date-picker-icon" 
                                                    iconName="Calendar"
                                                    onClick={() => {
                                                        const input = this.startDatePickerRef.current?.input;
                                                        if (input) {
                                                            input.click();
                                                        }
                                                    }}
                                                />
                                            </>
                                        )}
                                    </Observer>
                                </div>
                            </div>
                            <div className="input-row flex-row">
                                <span>End Date</span>
                                <div className="column-2 date-picker-wrapper">
                                    <Observer endDate={this.endDate}>
                                        {(props: { endDate: Date }) => (
                                            <>
                                                <DatePicker
                                                    ref={this.endDatePickerRef}
                                                    selected={props.endDate}
                                                    onChange={(date: Date | null) => {
                                                        if (date) {
                                                            this.endDate.value = date;
                                                            this.validateSelections();
                                                        }
                                                    }}
                                                    onCalendarOpen={() => this.isDatePickerOpen.value = true}
                                                    onCalendarClose={() => this.isDatePickerOpen.value = false}
                                                    dateFormat="MM/dd/yyyy"
                                                    className="bolt-textfield-input input-date"
                                                />
                                                <Icon 
                                                    className="date-picker-icon" 
                                                    iconName="Calendar"
                                                    onClick={() => {
                                                        const input = this.endDatePickerRef.current?.input;
                                                        if (input) {
                                                            input.click();
                                                        }
                                                    }}
                                                />
                                            </>
                                        )}
                                    </Observer>
                                </div>
                            </div>
                        </div>
                            )}
                        </Observer>
                    </PanelContent>
                    <PanelFooter>
                        <div className="flex-grow flex-row">
                            {this.props.eventApi && <Button onClick={this.onDeleteClick} danger={true} text="Delete event" />}
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
                                message="Are you sure you want to delete the event?"
                                onConfirm={() => {
                                    this.props.eventSource
                                        .deleteEvent(this.props.eventApi!.extendedProps.id, this.props.eventApi!.start!)
                                        .then(() => {
                                            this.props.calendarApi.refetchEvents();
                                        });
                                    this.isConfirmationDialogOpen.value = false;
                                    this.props.onDismiss();
                                }}
                                onDismiss={() => {
                                    this.isConfirmationDialogOpen.value = false;
                                }}
                                title="Delete event"
                            />
                        ) : null;
                    }}
                </Observer>
            </>
        );
    }

    private onCatagorySelectionChange = (value?: IListBoxItem<{}> | undefined): void => {
        this.category = value ? value.text || "Uncategorized" : "Uncategorized";
        this.validateSelections();
    };

    private onDeleteClick = async (): Promise<void> => {
        this.isConfirmationDialogOpen.value = true;
    };

    private onInputDescription = (e: React.ChangeEvent, value: string): void => {
        this.description.value = value;
        this.validateSelections();
    };

    private onInputTitle = (e: React.ChangeEvent, value: string): void => {
        this.title.value = value;
        this.validateSelections();
    };

    private onOKClick = (): void => {
        const excludedEndDate = new Date(this.endDate.value);
        excludedEndDate.setDate(this.endDate.value.getDate() + 1);
        if (this.category === "") {
            this.category = "Uncategorized";
        }
        let promise;
        if (this.props.eventApi) {
            promise = this.props.eventSource.updateEvent(
                this.props.eventApi.extendedProps.id,
                this.title.value,
                this.startDate.value,
                this.endDate.value,
                this.category,
                this.description.value
            );
        } else {
            promise = this.props.eventSource.addEvent(this.title.value, this.startDate.value, this.endDate.value, this.category, this.description.value);
        }
        promise.then(() => {
            this.props.calendarApi.refetchEvents();
        });
        this.props.onDismiss();
    };

    private validateSelections = () => {
        this.okButtonEnabled.value = this.title.value !== "" && this.startDate.value <= this.endDate.value;
        if (this.title.value === "") {
            this.message.value = "Title can not be empty.";
        } else if (this.startDate.value > this.endDate.value) {
            this.message.value = "Start date must be same or before the end date.";
        } else if (this.message.value !== "") {
            this.message.value = "";
        }
    };
}
