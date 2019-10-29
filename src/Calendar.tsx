import "./Calendar.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { CommonServiceIds, IProjectPageService, getClient } from "azure-devops-extension-api";
import { IExtensionDataService, IExtensionDataManager, ILocationService, IHostNavigationService } from "azure-devops-extension-api/Common";
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

import { localeData } from "moment";

import { AddEditDaysOffDialog } from "./AddEditDaysOffDialog";
import { AddEditEventDialog } from "./AddEditEventDialog";
import { ICalendarEvent } from "./Contracts";
import { FreeFormId, FreeFormEventsSource } from "./FreeFormEventSource";
import { SummaryComponent } from "./SummaryComponent";
import { MonthAndYear, monthAndYearToString, formatDate } from "./TimeLib";
import { DaysOffId, VSOCapacityEventSource, IterationId } from "./VSOCapacityEventSource";

enum Dialogs {
    None,
    NewEventDialog,
    NewDaysOffDialog
}

class ExtensionContent extends React.Component {
    anchorElement: ObservableValue<HTMLElement | undefined> = new ObservableValue<HTMLElement | undefined>(undefined);
    calendarComponentRef = React.createRef<FullCalendar>();
    commandBarItems: IHeaderCommandBarItem[];
    currentMonthAndYear: ObservableValue<MonthAndYear>;
    dataManager: IExtensionDataManager | undefined;
    displayCalendar: ObservableValue<boolean>;
    eventApi?: EventApi;
    eventToEdit?: ICalendarEvent;
    freeFormEventSource: FreeFormEventsSource;
    hostUrl: string;
    members: TeamMember[];
    navigationService: IHostNavigationService | undefined;
    openDialog: ObservableValue<Dialogs> = new ObservableValue(Dialogs.None);
    projectId: string;
    projectName: string;
    selectedEndDate: Date;
    selectedStartDate: Date;
    selectedTeamName: string;
    showMonthPicker: ObservableValue<boolean> = new ObservableValue<boolean>(false);
    teams: ObservableValue<WebApiTeam[]>;
    vsoCapacityEventSource: VSOCapacityEventSource;

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
        this.displayCalendar = new ObservableValue<boolean>(false);
        this.projectId = "";
        this.projectName = "";
        this.hostUrl = "";
        this.members = [];
        this.freeFormEventSource = new FreeFormEventsSource();
        this.vsoCapacityEventSource = new VSOCapacityEventSource();
    }

    public render(): JSX.Element {
        return (
            <Page className="flex-grow flex-row">
                <div className="flex-column scroll-hidden calendar-area">
                    <CustomHeader className="bolt-header-with-commandbar">
                        <HeaderTitleArea className="flex-grow">
                            <div className="flex-grow">
                                <Observer currentMonthAndYear={this.currentMonthAndYear}>
                                    {(props: { currentMonthAndYear: MonthAndYear }) => {
                                        return (
                                            <Dropdown
                                                items={this.getMonthPickerOptions()}
                                                key={props.currentMonthAndYear.month}
                                                onSelect={this.onSelectMonthYear}
                                                placeholder={monthAndYearToString(props.currentMonthAndYear)}
                                                renderExpandable={expandableProps => (
                                                    <DropdownExpandableButton hideDropdownIcon={true} {...expandableProps} />
                                                )}
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
                                                onSelect={this.onSelectTeam}
                                                placeholder={this.selectedTeamName}
                                                renderExpandable={expandableProps => <DropdownExpandableButton {...expandableProps} />}
                                            />
                                        );
                                    }}
                                </Observer>
                            </div>
                        </HeaderTitleArea>
                        <HeaderCommandBar items={this.commandBarItems} />
                    </CustomHeader>
                    <Observer display={this.displayCalendar}>
                        {(props: { display: boolean }) => {
                            return props.display ? (
                                <div className="calendar-component">
                                    <FullCalendar
                                        defaultView="dayGridMonth"
                                        editable={true}
                                        eventClick={this.onEventClick}
                                        eventDrop={this.onEventDrop}
                                        eventRender={this.eventRender}
                                        eventResize={this.onEventResize}
                                        eventSources={[
                                            { events: this.freeFormEventSource.getEvents },
                                            { events: this.vsoCapacityEventSource.getEvents }
                                        ]}
                                        firstDay={localeData(navigator.language).firstDayOfWeek()}
                                        header={false}
                                        height={this.getCalendarHeight()}
                                        plugins={[dayGridPlugin, timeGridPlugin, interactionPlugin]}
                                        ref={this.calendarComponentRef}
                                        select={this.onSelectCalendarDates}
                                        selectable={true}
                                    />
                                </div>
                            ) : null;
                        }}
                    </Observer>
                </div>
                <SummaryComponent capacityEventSource={this.vsoCapacityEventSource} freeFormEventSource={this.freeFormEventSource} />
                <Observer anchorElement={this.anchorElement}>
                    {(props: { anchorElement: HTMLElement | undefined }) => {
                        return props.anchorElement ? (
                            <ContextualMenu
                                anchorElement={props.anchorElement}
                                anchorOffset={{ horizontal: 4, vertical: 4 }}
                                anchorOrigin={{ horizontal: Location.start, vertical: Location.start }}
                                key={this.selectedEndDate!.toString()}
                                menuProps={{
                                    id: "foo",
                                    items: [
                                        { id: "event", text: "Add event", iconProps: { iconName: "Add" }, onActivate: this.onClickAddEvent },
                                        { id: "dayOff", text: "Add days off", iconProps: { iconName: "Clock" }, onActivate: this.onClickAddDaysOff }
                                    ]
                                }}
                                onDismiss={() => {
                                    this.anchorElement.value = undefined;
                                }}
                            />
                        ) : null;
                    }}
                </Observer>
                <Observer dialog={this.openDialog}>
                    {(props: { dialog: Dialogs }) => {
                        return props.dialog === Dialogs.NewDaysOffDialog ? (
                            <AddEditDaysOffDialog
                                calendarApi={this.getCalendarApi()}
                                end={this.selectedEndDate}
                                event={this.eventToEdit}
                                eventSource={this.vsoCapacityEventSource}
                                members={this.members}
                                onDismiss={this.onDialogDismiss}
                                start={this.selectedStartDate}
                            />
                        ) : props.dialog === Dialogs.NewEventDialog ? (
                            <AddEditEventDialog
                                calendarApi={this.getCalendarApi()}
                                end={this.selectedEndDate}
                                eventApi={this.eventApi}
                                eventSource={this.freeFormEventSource}
                                onDismiss={this.onDialogDismiss}
                                start={this.selectedStartDate}
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

    /**
     * Edits the rendered event if required
     */
    private eventRender = (arg: { isMirror: boolean; isStart: boolean; isEnd: boolean; event: EventApi; el: HTMLElement; view: View }) => {
        if (arg.event.id.startsWith(DaysOffId) && arg.event.start) {
            // get grouped event for that date
            const capacityEvent = this.vsoCapacityEventSource.getGroupedEventForDate(arg.event.start);
            if (capacityEvent && capacityEvent.icons) {
                // add all user icons in to event
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
            // iterations are background event, show title for only start
            arg.el.innerText = arg.event.title;
            arg.el.style.color = "black";
        }
    };

    private getCalendarApi(): Calendar {
        return this.calendarComponentRef.current!.getApi();
    }

    /**
     * Manually calculates available vertical space for calendar
     */
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
                data: monthAndYear,
                id: text,
                text: text
            });
        }
        return options;
    }

    private getTeamPickerOptions(): IListBoxItem[] {
        const options: IListBoxItem[] = [];
        this.teams.value.forEach(function(item) {
            options.push({ data: item, id: item.id, text: item.name });
        });

        return options;
    }

    private async initialize() {
        const dataSvc = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
        const project = await projectService.getProject();
        const locationService = await SDK.getService<ILocationService>(CommonServiceIds.LocationService);

        this.dataManager = await dataSvc.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
        this.navigationService = await SDK.getService<IHostNavigationService>(CommonServiceIds.HostNavigationService);

        const queryParam = await this.navigationService.getQueryParams();
        let selectedTeamId;

        // if URL has team id in it, use that
        if (queryParam && queryParam["team"]) {
            selectedTeamId = queryParam["team"];
        }

        if (project) {
            if (!selectedTeamId) {
                // Nothing in URL - check data service
                selectedTeamId = await this.dataManager.getValue<string>("selected-team-" + project.id, { scopeType: "User" });
            }

            const client = getClient(CoreRestClient);
            const teams = await client.getTeams(project.id, false, 1000);

            this.projectId = project.id;
            this.projectName = project.name;

            teams.sort((a, b) => {
                return a.name.toUpperCase().localeCompare(b.name.toUpperCase());
            });

            // if team id wasn't in URL or database use first available team
            if (!selectedTeamId) {
                selectedTeamId = teams[0].id;
            }

            if (!queryParam || !queryParam["team"]) {
                // Add team id to URL
                this.navigationService.setQueryParams({ team: selectedTeamId });
            }

            this.hostUrl = await locationService.getServiceLocation();
            this.selectedTeamName = (await client.getTeam(project.id, selectedTeamId)).name;
            this.freeFormEventSource.initialize(selectedTeamId, this.dataManager);
            this.vsoCapacityEventSource.initialize(project.id, this.projectName, selectedTeamId, this.selectedTeamName, this.hostUrl);
            this.displayCalendar.value = true;
            this.dataManager.setValue<string>("selected-team-" + project.id, selectedTeamId, { scopeType: "User" });
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

    private onSelectCalendarDates = (arg: {
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
        const dataDate = formatDate(this.selectedEndDate, "YYYY-MM-DD");
        this.anchorElement.value = document.querySelector("td.fc-day-top[data-date='" + dataDate + "']") as HTMLElement;
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
        this.dataManager!.setValue<string>("selected-team-" + this.projectId, newTeam.id, { scopeType: "User" });
        this.navigationService!.setQueryParams({ team: newTeam.id });
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

showRootComponent(<ExtensionContent />);
