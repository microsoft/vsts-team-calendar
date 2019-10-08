import "./Calendar.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { CommonServiceIds, IProjectPageService, getClient } from "azure-devops-extension-api";
import { IExtensionDataService, IExtensionDataManager, ILocationService } from "azure-devops-extension-api/Common";
import { CoreRestClient, WebApiTeam } from "azure-devops-extension-api/Core";
import { TeamMember } from "azure-devops-extension-api/WebApi/WebApi";

import * as SDK from "azure-devops-extension-sdk";

import { Dropdown, DropdownExpandableButton } from "azure-devops-ui/Dropdown";
import { CustomHeader, HeaderTitleArea } from "azure-devops-ui/Header";
import { IHeaderCommandBarItem, HeaderCommandBar } from "azure-devops-ui/HeaderCommandBar";
import { Icon } from "azure-devops-ui/Icon";
import { IListBoxItem } from "azure-devops-ui/ListBox";
import { ContextualMenu } from "azure-devops-ui/Menu";
import { ObservableValue } from "azure-devops-ui/Core/Observable";
import { Observer } from "azure-devops-ui/Observer";
import { Page } from "azure-devops-ui/Page";
import { Location } from "azure-devops-ui/Utilities/Position";

import { View, EventApi, Duration, Calendar } from "@fullcalendar/core";
import dayGridPlugin from "@fullcalendar/daygrid";
import interactionPlugin from "@fullcalendar/interaction";
import FullCalendar from "@fullcalendar/react";
import timeGridPlugin from "@fullcalendar/timegrid";

import { AddEditDaysOffDialog } from "./AddEditDaysOffDialog";
import { AddEditEventDialog } from "./AddEditEventDialog";
import { ICalendarEvent } from "./Contracts";
import { FreeFormId, FreeFormEventsSource } from "./FreeFormEventSource";
import { SummaryComponent } from "./SummaryComponent";
import { MonthAndYear, monthAndYearToString } from "./TimeLib";
import { getQueryVariable, setTeamQueryVariable } from "./UrlLib";
import { DaysOffId, VSOCapacityEventSource, IterationId } from "./VSOCapacityEventSource";

enum Dialogs {
    None,
    NewEventDialog,
    NewDaysOffDialog
}

class HubContent extends React.Component {
    calendarComponentRef = React.createRef<FullCalendar>();
    openDialog: ObservableValue<Dialogs> = new ObservableValue(Dialogs.None);
    anchorElement: ObservableValue<HTMLElement | undefined> = new ObservableValue<HTMLElement | undefined>(undefined);
    showMonthPicker: ObservableValue<boolean> = new ObservableValue<boolean>(false);
    selectedStartDate: Date;
    selectedEndDate: Date;
    currentMonthAndYear: ObservableValue<MonthAndYear>;
    commandBarItems: IHeaderCommandBarItem[];
    eventToEdit?: ICalendarEvent;
    eventApi?: EventApi;
    teams: ObservableValue<WebApiTeam[]>;
    selectedTeamId: ObservableValue<string>;
    projectId: string;
    projectName: string;
    selectedTeamName: string;
    dataManager: IExtensionDataManager | undefined;
    freeFormEventSource: FreeFormEventsSource;
    vsoCapacityEventSource: VSOCapacityEventSource;
    hostUrl: string;
    members: TeamMember[];

    constructor(props: {}) {
        super(props);

        this.currentMonthAndYear = new ObservableValue<MonthAndYear>({
            month: new Date().getMonth(),
            year: new Date().getFullYear()
        });

        this.state = {
            fullScreenMode: false,
            calendarWeekends: true,
            calendarEvents: []
        };

        this.commandBarItems = [
            {
                iconProps: {
                    iconName: "Add"
                },
                id: "newItem",
                important: true,
                onActivate: this.onClickNewItem,
                text: "New Item"
            },
            {
                id: "today",
                important: true,
                onActivate: () => {
                    if (this.calendarComponentRef.current) {
                        this.getCalendarApi().today();
                        this.currentMonthAndYear.value = {
                            month: new Date().getMonth(),
                            year: new Date().getFullYear()
                        };
                    }
                },
                text: "Today"
            },
            {
                iconProps: {
                    iconName: "ChevronLeft"
                },
                important: true,
                id: "prev",
                onActivate: () => {
                    if (this.calendarComponentRef.current) {
                        this.getCalendarApi().prev();
                        this.currentMonthAndYear.value = this.calcMonths(this.currentMonthAndYear.value, -1);
                    }
                },
                text: "Prev"
            },
            {
                iconProps: {
                    iconName: "ChevronRight"
                },
                important: true,
                id: "next",
                onActivate: () => {
                    if (this.calendarComponentRef.current) {
                        this.getCalendarApi().next();
                        this.currentMonthAndYear.value = this.calcMonths(this.currentMonthAndYear.value, 1);
                    }
                },
                text: "Next"
            }
        ];

        this.selectedEndDate = new Date();
        this.selectedStartDate = new Date();
        this.teams = new ObservableValue<WebApiTeam[]>([]);
        this.selectedTeamName = "Select Team";
        this.selectedTeamId = new ObservableValue<string>("");
        this.projectId = "";
        this.projectName = "";
        this.hostUrl = "";
        this.members = [];
        this.freeFormEventSource = new FreeFormEventsSource();
        this.vsoCapacityEventSource = new VSOCapacityEventSource();
    }

    public render(): JSX.Element {
        return (
            <Page className="sample-hub flex-grow flex-row">
                <div className="flex-column scroll-hidden calendar-area">
                    <CustomHeader className="bolt-header-with-commandbar">
                        <HeaderTitleArea className="flex-grow">
                            <div className="flex-grow">
                                <Observer currentMonthAndYear={this.currentMonthAndYear}>
                                    {(props: { currentMonthAndYear: MonthAndYear }) => {
                                        return (
                                            <Dropdown
                                                key={props.currentMonthAndYear.month}
                                                items={this.getMonthPickerOptions()}
                                                placeholder={monthAndYearToString(props.currentMonthAndYear)}
                                                renderExpandable={expandableProps => (
                                                    <DropdownExpandableButton hideDropdownIcon={true} {...expandableProps} />
                                                )}
                                                onSelect={this.onSelectMonthYear}
                                            />
                                        );
                                    }}
                                </Observer>
                                <Icon ariaLabel="Video icon" iconName="ChevronRight" />
                                <Observer teams={this.teams}>
                                    {(props: { teams: WebApiTeam[] }) => {
                                        return props.teams === [] ? null : (
                                            <Dropdown
                                                items={this.getTeamPickerOptions()}
                                                placeholder={this.selectedTeamName}
                                                renderExpandable={expandableProps => <DropdownExpandableButton {...expandableProps} />}
                                                onSelect={this.onSelectTeam}
                                            />
                                        );
                                    }}
                                </Observer>
                            </div>
                        </HeaderTitleArea>
                        <HeaderCommandBar items={this.commandBarItems} />
                    </CustomHeader>
                    <Observer teamId={this.selectedTeamId}>
                        {(props: { teamId: string }) => {
                            return props.teamId === "" ? null : (
                                <div className="calendar-component">
                                    <FullCalendar
                                        defaultView="dayGridMonth"
                                        header={false}
                                        plugins={[dayGridPlugin, timeGridPlugin, interactionPlugin]}
                                        ref={this.calendarComponentRef}
                                        selectable={true}
                                        select={this.handleSelect}
                                        height={this.getCalendarHeight()}
                                        editable={true}
                                        eventClick={this.onEventClick}
                                        eventDrop={this.onEventDrop}
                                        eventRender={this.eventRender}
                                        eventResize={this.onEventResize}
                                        eventSources={[
                                            { events: this.freeFormEventSource.getEvents },
                                            { events: this.vsoCapacityEventSource.getEvents }
                                        ]}
                                    />
                                </div>
                            );
                        }}
                    </Observer>
                </div>
                <SummaryComponent capacityEventSource={this.vsoCapacityEventSource} freeFormEventSource={this.freeFormEventSource} />
                <Observer anchorElement={this.anchorElement}>
                    {(props: { anchorElement: HTMLElement | undefined }) => {
                        return props.anchorElement ? (
                            <ContextualMenu
                                key={this.selectedEndDate!.toString()}
                                onDismiss={() => {
                                    this.anchorElement.value = undefined;
                                }}
                                menuProps={{
                                    id: "foo",
                                    items: [
                                        { id: "event", text: "Add event", iconProps: { iconName: "Add" }, onActivate: this.onClickAddEvent },
                                        { id: "dayOff", text: "Add days off", iconProps: { iconName: "Clock" }, onActivate: this.onClickAddDaysOff }
                                    ]
                                }}
                                anchorElement={props.anchorElement}
                                anchorOrigin={{ horizontal: Location.start, vertical: Location.start }}
                                anchorOffset={{ horizontal: 4, vertical: 4 }}
                            />
                        ) : null;
                    }}
                </Observer>
                <Observer dialog={this.openDialog}>
                    {(props: { dialog: Dialogs }) => {
                        return props.dialog == Dialogs.NewDaysOffDialog ? (
                            <AddEditDaysOffDialog
                                calendarApi={this.getCalendarApi()}
                                end={this.selectedEndDate}
                                event={this.eventToEdit}
                                eventSource={this.vsoCapacityEventSource}
                                onDismiss={this.onDialogDismiss}
                                start={this.selectedStartDate}
                                members={this.members}
                            />
                        ) : props.dialog === Dialogs.NewEventDialog ? (
                            <AddEditEventDialog
                                calendarApi={this.getCalendarApi()}
                                end={this.selectedEndDate}
                                eventApi={this.eventApi}
                                onDismiss={this.onDialogDismiss}
                                start={this.selectedStartDate}
                                eventSource={this.freeFormEventSource}
                            />
                        ) : null;
                    }}
                </Observer>
            </Page>
        );
    }

    componentDidMount() {
        SDK.init();
        this.initialize();
        window.addEventListener("resize", this.updateDimensions);
    }

    private calcMonths(current: MonthAndYear, monthDelta: number): MonthAndYear {
        let month = (current.month + monthDelta) % 12;
        let year = current.year + Math.floor((current.month + monthDelta) / 12);
        if (month < 0) {
            month = 12 + month;
        }
        return { month, year };
    }

    private eventRender = (arg: { isMirror: boolean; isStart: boolean; isEnd: boolean; event: EventApi; el: HTMLElement; view: View }) => {
        if (arg.event.id.startsWith(DaysOffId) && arg.event.start) {
            const capacityEvent = this.vsoCapacityEventSource.getGroupedEventForDate(arg.event.start);
            if (capacityEvent && capacityEvent.icons) {
                capacityEvent.icons.forEach(element => {
                    if (element.src) {
                        var img: HTMLImageElement = document.createElement("img");
                        img.src = element.src;
                        img.className = "event-icon";
                        img.title = element.linkedEvent.title;
                        img.onclick = () => {
                            this.eventToEdit = element.linkedEvent;
                            this.openDialog.value = Dialogs.NewDaysOffDialog;
                        };
                        var content = arg.el.querySelector(".fc-content");
                        if (content) {
                            content.appendChild(img);
                        }
                    }
                });
            }
        } else if (arg.event.id.startsWith(IterationId) && arg.isStart) {
            arg.el.innerText = arg.event.title;
        }
    };

    private getCalendarApi(): Calendar {
        return this.calendarComponentRef.current!.getApi();
    }

    private getCalendarHeight(): number {
        var height = document.getElementById("team-calendar");
        if (height) {
            return height.offsetHeight - 90;
        }
        return 200;
    }

    private getMonthPickerOptions(): IListBoxItem[] {
        const options: IListBoxItem[] = [];
        const listSize = 3;
        for (let i = -listSize; i <= listSize; ++i) {
            const monthAndYear = this.calcMonths(this.currentMonthAndYear.value, i);
            const text = monthAndYearToString(monthAndYear);
            options.push({
                id: text,
                text: text,
                data: monthAndYear
            });
        }
        return options;
    }

    private getTeamPickerOptions(): IListBoxItem[] {
        const options: IListBoxItem[] = [];
        this.teams.value.forEach(function(item) {
            options.push({ id: item.id, text: item.name, data: item });
        });

        return options;
    }

    private handleSelect = (arg: {
        start: Date;
        end: Date;
        startStr: string;
        endStr: string;
        allDay: boolean;
        resource?: any;
        jsEvent: MouseEvent;
        view: View;
    }) => {
        this.selectedEndDate = new Date(arg.end);
        this.selectedEndDate.setDate(arg.end.getDate() - 1);
        this.selectedStartDate = arg.start;
        const dataDate = this.selectedEndDate.toISOString().split("T")[0];
        this.anchorElement.value = document.querySelector("td.fc-day-top[data-date='" + dataDate + "']") as HTMLElement;
    };

    private async initialize() {
        const dataSvc = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        this.dataManager = await dataSvc.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());

        const queryParam = getQueryVariable("team");
        let selectedTeamId;

        if (queryParam) {
            selectedTeamId = queryParam;
        } else {
            // Nothing in URL - check data service
            selectedTeamId = await this.dataManager.getValue<string>("selected-team", { scopeType: "User" });
        }

        const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
        const project = await projectService.getProject();

        const locationService = await SDK.getService<ILocationService>(CommonServiceIds.LocationService);

        if (project) {
            const client = getClient(CoreRestClient);
            const teams = await client.getTeams(project.id, true, 1000);

            this.projectId = project.id;
            this.projectName = project.name;

            teams.sort((a, b) => {
                return a.name.toUpperCase().localeCompare(b.name.toUpperCase());
            });

            if (!selectedTeamId) {
                selectedTeamId = teams[0].id;
            }
            if (!queryParam) {
                setTeamQueryVariable(selectedTeamId);
            }

            this.hostUrl = await locationService.getServiceLocation();
            this.selectedTeamName = (await client.getTeam(project.id, selectedTeamId)).name;
            this.freeFormEventSource.initialize(selectedTeamId, this.dataManager);
            this.vsoCapacityEventSource.initialize(project.id, this.projectName, selectedTeamId, this.selectedTeamName, this.hostUrl);
            this.selectedTeamId.value = selectedTeamId;
            this.dataManager.setValue<string>("selected-team", this.selectedTeamId.value, { scopeType: "User" });
            this.teams.value = teams;
            this.members = await client.getTeamMembersWithExtendedProperties(project.id, selectedTeamId);
        }
    }

    private onClickNewItem = () => {
        this.eventApi = undefined;
        const today = new Date();
        this.selectedStartDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
        this.selectedEndDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
        if (this.calendarComponentRef.current) {
            this.openDialog.value = Dialogs.NewEventDialog;
        }
    };

    private onClickAddEvent = () => {
        this.eventApi = undefined;
        this.openDialog.value = Dialogs.NewEventDialog;
    };

    private onClickAddDaysOff = () => {
        this.eventToEdit = undefined;
        this.openDialog.value = Dialogs.NewDaysOffDialog;
    };

    private onDialogDismiss = () => {
        this.openDialog.value = Dialogs.None;
    };

    private onEventClick = (arg: { el: HTMLElement; event: EventApi; jsEvent: MouseEvent; view: View }) => {
        if (arg.event.id.startsWith(FreeFormId)) {
            this.eventApi = arg.event;
            this.openDialog.value = Dialogs.NewEventDialog;
        }
    };

    private onEventDrop = (arg: {
        el: HTMLElement;
        event: EventApi;
        oldEvent: EventApi;
        delta: Duration;
        revert: () => void;
        jsEvent: Event;
        view: View;
    }) => {
        if (arg.event.id.startsWith(FreeFormId)) {
            let inclusiveEndDate;
            if (arg.event.end) {
                inclusiveEndDate = new Date(arg.event.end);
                inclusiveEndDate.setDate(arg.event.end.getDate() - 1);
            } else {
                inclusiveEndDate = new Date(arg.event.start!);
            }

            this.freeFormEventSource.updateEvent(
                arg.event.extendedProps.id,
                arg.event.title,
                arg.event.start!,
                inclusiveEndDate,
                arg.event.extendedProps.category,
                arg.event.extendedProps.description
            );
        }
    };

    private onEventResize = (arg: {
        el: HTMLElement;
        startDelta: Duration;
        endDelta: Duration;
        prevEvent: EventApi;
        event: EventApi;
        revert: () => void;
        jsEvent: Event;
        view: View;
    }) => {
        if (arg.event.id.startsWith(FreeFormId)) {
            let inclusiveEndDate;
            if (arg.event.end) {
                inclusiveEndDate = new Date(arg.event.end);
                inclusiveEndDate.setDate(arg.event.end.getDate() - 1);
            } else {
                inclusiveEndDate = new Date(arg.event.start!);
            }

            this.freeFormEventSource.updateEvent(
                arg.event.extendedProps.id,
                arg.event.title,
                arg.event.start!,
                inclusiveEndDate,
                arg.event.extendedProps.category,
                arg.event.extendedProps.description
            );
        }
    };

    private onSelectMonthYear = (event: React.SyntheticEvent<HTMLElement, Event>, item: IListBoxItem<{}>) => {
        const date = item.data as MonthAndYear;
        if (this.calendarComponentRef) {
            this.getCalendarApi().gotoDate(new Date(date.year, date.month));
            this.currentMonthAndYear.value = date;
        }
    };

    private onSelectTeam = async (event: React.SyntheticEvent<HTMLElement, Event>, item: IListBoxItem<{}>) => {
        const newTeam = item.data! as WebApiTeam;
        this.selectedTeamName = newTeam.name;
        this.freeFormEventSource.initialize(newTeam.id, this.dataManager!);
        this.vsoCapacityEventSource.initialize(this.projectId, this.projectName, newTeam.id, newTeam.name, this.hostUrl);
        this.getCalendarApi().refetchEvents();
        this.dataManager!.setValue<string>("selected-team", newTeam.id, { scopeType: "User" });
        this.members = await getClient(CoreRestClient).getTeamMembersWithExtendedProperties(this.projectId, newTeam.id);
    };

    private updateDimensions = () => {
        if (this.calendarComponentRef.current) {
            this.calendarComponentRef.current.getApi().setOption("height", this.getCalendarHeight());
        }
    };
}

function showRootComponent(component: React.ReactElement<any>) {
    ReactDOM.render(component, document.getElementById("team-calendar"));
}

showRootComponent(<HubContent />);
