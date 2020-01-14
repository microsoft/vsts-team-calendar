import * as React from "react";

import * as Calendar from "./Calendar";
import * as Calendar_Contracts from "./Contracts";
import * as Calendar_Dialogs from "./Dialogs";
import { allSettled, realPromise } from "./Utils/Promise";
import * as Controls from "VSS/Controls";
import * as Controls_Dialogs from "VSS/Controls/Dialogs";
import * as Controls_Menus from "VSS/Controls/Menus";
import * as Controls_Navigation from "VSS/Controls/Navigation";
import * as Controls_StatusIndicator from "VSS/Controls/StatusIndicator";
import * as Service from "VSS/Service";
import * as Services_Navigation from "VSS/SDK/Services/Navigation";
import * as Tfs_Core_WebApi from "TFS/Core/RestClient";
import * as Utils_Core from "VSS/Utils/Core";
import * as Utils_Date from "VSS/Utils/Date";
import * as Utils_String from "VSS/Utils/String";
import { Hub } from "vss-ui/Hub";
import { HubHeader } from "vss-ui/HubHeader";
import { PickList, PickListDropdown, IPickListSelection } from "vss-ui/PickList";
import { IHubBreadcrumbItem } from "vss-ui/HubHeader";
import { HubViewState } from "vss-ui/Utilities/HubViewState";
import { IObservableValue, ObservableValue } from "vss-ui/Utilities/Observable";
import * as WebApi_Constants from "VSS/WebApi/Constants";
import * as WebApi_Contracts from "VSS/WebApi/Contracts";
import * as Work_Contracts from "TFS/Work/Contracts";
import * as FullCalendar from "fullcalendar"; // types only

import { BaseComponent, Callout, IBaseProps, IContextualMenuItem, IBreadcrumbItem, SelectionMode } from "office-ui-fabric-react";
import { PivotBarItem } from "vss-ui/PivotBar";
import { WebApiTeam } from "TFS/Core/Contracts";

const months = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
];

// Polyfill Math.trunc...
(Math as any).trunc =
    (Math as any).trunc ||
    function(x) {
        var n = x - x % 1;
        return n === 0 && (x < 0 || (x === 0 && 1 / x !== 1 / 0)) ? -0 : n;
    };

function newElement(tag: string, className?: string, text?: string): JQuery {
    return $("<" + tag + "/>")
        .addClass(className || "")
        .text(text || "");
}

export class EventSourceCollection {
    private static _createPromise: Promise<any>;
    private _collection: Calendar_Contracts.IEventSource[] = [];
    private _map: { [name: string]: Calendar_Contracts.IEventSource } = {};

    constructor(sources: Calendar_Contracts.IEventSource[]) {
        this._collection = sources || [];
        for (const source of this._collection) {
            this._map[source.id] = source;
        }
    }

    public getById(id: string): Calendar_Contracts.IEventSource {
        return this._map[id];
    }

    public getAllSources(): Calendar_Contracts.IEventSource[] {
        return this._collection;
    }

    public updateTeamContext(newTeam: WebApiTeam) {
        for (const eventSource of this._collection) {
            if (typeof eventSource.updateTeamContext === "function") {
                eventSource.updateTeamContext(newTeam);
            }
        }
    }

    public static create(team?: any): PromiseLike<EventSourceCollection> {
        if (!this._createPromise) {
            const extensionContext = VSS.getExtensionContext();
            const eventSourcesTargetId =
                extensionContext.publisherId + "." + extensionContext.extensionId + ".calendar-event-sources";
            this._createPromise = realPromise(VSS.getServiceContributions(eventSourcesTargetId)).then(contributions => {
                const servicePromises = contributions.map(c => c.getInstance(c.id, { team }));
                return allSettled(servicePromises).then(promiseStates => {
                    const services = [];
                    promiseStates.forEach((promiseState, index: number) => {
                        if (promiseState.value) {
                            services.push(promiseState.value);
                        } else {
                            console.log("Failed to get calendar event source instance for: " + contributions[index].id);
                        }
                    });
                    return new EventSourceCollection(services);
                });
            });
        }

        return this._createPromise;
    }
}

export interface CalendarViewOptions extends Calendar.CalendarOptions {
    defaultEvents?: FullCalendar.EventObject[];
    selectedTeam?: WebApiTeam;
}

export interface ICalendarHubProps extends IBaseProps {
    onSelectedTeamChanged: (selectedTeam: WebApiTeam) => void;
    onCalendarViewUpdated: (calendarView: CalendarView) => void;
    selectedTeam?: WebApiTeam;
}

interface MonthAndYear {
    month: number;
    year: number;
}

export interface CalendarState {
    calendar: MonthAndYear;
    showMonthPicker: boolean;
    monthPickerTarget: HTMLElement | null;
}

export class CalendarStateManagerComponent extends BaseComponent<
    {},
    { calendarView?: ObservableValue<CalendarView>; selectedTeam?: WebApiTeam }
> {
    constructor() {
        super();
        this.state = { calendarView: new ObservableValue() };

        this.initialize();
    }

    private async initialize() {
        const navSvc = await VSS.getService<IHostNavigationService>(VSS.ServiceIds.Navigation);
        let selectedTeamId = (await navSvc.getCurrentState()).team;
        if (!selectedTeamId) {
            // Nothing in URL - check data service
            const dataSvc = await VSS.getService<IExtensionDataService>(VSS.ServiceIds.ExtensionData);
            selectedTeamId = await dataSvc.getValue<string>("selected-team-" + VSS.getWebContext().project.id, { scopeType: "User" });
        }
        if (!selectedTeamId) {
            // Nothing in URL or data service - check context
            const contextTeam = VSS.getWebContext().team;
            if (contextTeam && contextTeam.id) {
                selectedTeamId = contextTeam.id;
            }
        }
        if (selectedTeamId) {
            // get the team if we found a way to choose one
            const team = await Tfs_Core_WebApi.getClient().getTeam(VSS.getWebContext().project.id, selectedTeamId);
            if (team) {
                await this.setSelectedTeam(team);

                // Work around a bug where pivot bar wont render properly on initial render.
                setTimeout(() => {
                    this.forceUpdate();
                }, 250);
            }
        }
    }

    private async setSelectedTeam(selectedTeam: WebApiTeam) {
        // Set the URL and data service value to the selected team.
        const navSvc = await VSS.getService<IHostNavigationService>(VSS.ServiceIds.Navigation);
        navSvc.updateHistoryEntry(null, { team: selectedTeam.id }, true, true, null, false);

        const dataSvc = await VSS.getService<IExtensionDataService>(VSS.ServiceIds.ExtensionData);
        dataSvc.setValue<string>("selected-team-" + VSS.getWebContext().project.id, selectedTeam.id, { scopeType: "User" });

        this.setState({ selectedTeam });
    }

    public render(): JSX.Element {
        return (
            <div>
                <div className="calendar-area" id="calendarArea">
                    <CalendarComponent
                        onSelectedTeamChanged={async selectedTeam => {
                            this.setSelectedTeam(selectedTeam);
                        }}
                        onCalendarViewUpdated={calendarView => {
                            this.state.calendarView.value = calendarView;
                            this.forceUpdate();
                        }}
                        selectedTeam={this.state.selectedTeam}
                    />
                </div>
                {this.state.selectedTeam &&
                    this.state.calendarView.value && (
                        <div className="summary-area" id="summaryArea">
                            <SummaryComponent calendarView={this.state.calendarView} />
                        </div>
                    )}
            </div>
        );
    }
}

export class CalendarComponent extends BaseComponent<ICalendarHubProps, CalendarState> {
    private hubViewState: HubViewState;
    private calendarView: CalendarView;
    private calendarDiv: HTMLElement;
    private header: HubHeader;
    private needToEnhance: boolean = false;

    constructor() {
        super();
        this.hubViewState = new HubViewState({ defaultPivot: "calendar" });
        this.state = {
            calendar: { month: new Date().getMonth(), year: new Date().getFullYear() },
            showMonthPicker: false,
            monthPickerTarget: null,
        };
    }

    private setSelectedTeam(team: WebApiTeam) {
        if (team.id !== this.props.selectedTeam.id && typeof this.props.onSelectedTeamChanged === "function") {
            this.props.onSelectedTeamChanged(team);
        }
    }

    private static calcMonths(current: MonthAndYear, monthDelta: number): MonthAndYear {
        let month = (current.month + monthDelta) % 12;
        let year = current.year + Math.floor((current.month + monthDelta) / 12);
        if (month < 0) {
            month = 12 + month;
        }
        return { month, year };
    }

    private onMonthPickerClick(ev?: React.MouseEvent<HTMLElement>, item?: IBreadcrumbItem) {
        this.setState({ showMonthPicker: true, monthPickerTarget: ev.target as HTMLElement });
    }

    /**
     * Get list of current month -5 and +5 months
     */
    private getHeaderItems(): IHubBreadcrumbItem[] {
        return [
            {
                text: `${months[this.state.calendar.month]} ${this.state.calendar.year}`,
                key: "monthpicker",
                leftIconProps: { iconName: "Calendar" },
                onClick: this.onMonthPickerClick.bind(this),
            },
        ];
    }

    private getMonthPickerOptions(): MonthAndYear[] {
        const options: MonthAndYear[] = [];
        const listSize = 3;
        for (let i = -listSize; i <= listSize; ++i) {
            options.push(CalendarComponent.calcMonths(this.state.calendar, i));
        }
        return options;
    }

    private getMonthPickerListItem(item: MonthAndYear) {
        const str = `${months[item.month]} ${item.year}`;
        return {
            name: str,
            key: str,
        };
    }

    private enhanceControls(): void {
        if (this.calendarView) {
            this.calendarView.dispose();
            this.calendarView = null;
        }
        if (this.calendarDiv && this.props.selectedTeam) {
            $(this.calendarDiv).empty();
            this.calendarView = Controls.create(CalendarView, $(this.calendarDiv), {
                selectedTeam: this.props.selectedTeam,
            } as CalendarViewOptions);
            this.calendarView.getCalendar().goTo(new Date(this.state.calendar.year, this.state.calendar.month, 1, 0, 0, 0));
            this.props.onCalendarViewUpdated(this.calendarView);
        }
    }

    public componentDidUpdate(prevProps: ICalendarHubProps, prevState: CalendarState): void {
        if (prevProps.selectedTeam !== this.props.selectedTeam) {
            if (this.needToEnhance) {
                this.enhanceControls();
                this.needToEnhance = false;
            } else {
                this.needToEnhance = true;
            }
        }
    }

    public render(): JSX.Element {
        return (
            <div>
                {this.state.showMonthPicker && (
                    <Callout
                        isBeakVisible={false}
                        target={this.state.monthPickerTarget}
                        onDismiss={() => {
                            this.setState({ showMonthPicker: false });
                        }}>
                        <PickList
                            selectionMode={SelectionMode.single}
                            items={this.getMonthPickerOptions()}
                            getListItem={this.getMonthPickerListItem.bind(this)}
                            onSelectionChanged={selection => {
                                const selectedItem = selection.selectedItems[0] as MonthAndYear;
                                this.setState({ showMonthPicker: false, calendar: selection.selectedItems[0] });
                                this.calendarView.getCalendar().goTo(new Date(selectedItem.year, selectedItem.month, 1, 0, 0, 0));
                            }}
                            selectedItems={[this.state.calendar]}
                        />
                    </Callout>
                )}
                <Hub
                    hideFullScreenToggle={true}
                    hubViewState={this.hubViewState}
                    viewActions={[
                        {
                            key: "move-today",
                            name: "Today",
                            important: true,
                            onClick: this._onCommandClick,
                        },
                        {
                            key: "move-left",
                            name: "Prev",
                            important: true,
                            iconProps: {
                                iconName: "ChevronLeft",
                            },
                            onClick: this._onCommandClick,
                        },
                        {
                            key: "move-right",
                            name: "Next",
                            important: true,
                            iconProps: {
                                iconName: "ChevronRight",
                            },
                            onClick: this._onCommandClick,
                        },
                    ]}
                    commands={[
                        {
                            key: "new-item",
                            name: "New Item",
                            important: true,
                            iconProps: {
                                iconName: "Add",
                            },
                            onClick: this._onCommandClick,
                        },
                        {
                            key: "refresh",
                            name: "Refresh",
                            important: true,
                            iconProps: {
                                iconName: "Refresh",
                            },
                            onClick: this._onCommandClick,
                        },
                    ]}>
                    <HubHeader
                        breadcrumbItems={this.getHeaderItems()}
                        ref={header => {
                            this.header = header;
                        }}
                        headerItemPicker={{
                            isSearchable: true,
                            searchTextPlaceholder: "Filter...",
                            getItems: () => {
                                return new Promise<WebApiTeam[]>((resolve, reject) => {
                                    Tfs_Core_WebApi.getClient()
                                        .getTeams(VSS.getWebContext().project.id, 1000 as any)
                                        .then(result => {
                                            result.sort((a, b) => {
                                                return a.name.toUpperCase().localeCompare(b.name.toUpperCase());
                                            });
                                            resolve(result);
                                        }, reject);
                                });
                            },
                            getListItem: (item: WebApiTeam) => {
                                if (item) {
                                    return {
                                        name: item.name,
                                        key: item.id,
                                    };
                                }
                                return {
                                    name: "Select a team",
                                    key: "selectteammessage",
                                };
                            },
                            selectedItem: this.props.selectedTeam,
                            onSelectedItemChanged: (selectedTeam: WebApiTeam) => {
                                this.setSelectedTeam(selectedTeam);
                            },
                        }}
                    />
                    {this.props.selectedTeam && (
                        <PivotBarItem name="Calendar" itemKey="calendar">
                            <div
                                className="calendar-canvas"
                                id="calendarCanvas"
                                ref={elem => {
                                    if (elem) {
                                        this.calendarDiv = elem;
                                        if (this.needToEnhance) {
                                            this.enhanceControls();
                                            this.needToEnhance = false;
                                        } else {
                                            this.needToEnhance = true;
                                        }
                                    }
                                }}
                            />
                        </PivotBarItem>
                    )}
                </Hub>
            </div>
        );
    }
    private _onCommandClick = (ev: React.MouseEvent<HTMLElement>, item: IContextualMenuItem): void => {
        switch (item.key) {
            case "refresh":
                this.calendarView.getCalendar().refreshEvents();
                break;
            case "new-item":
                this.calendarView.addEventClicked();
                break;
            case "move-left":
                const prevMonth = this.state.calendar.month - 1;
                this.setState({ calendar: CalendarComponent.calcMonths(this.state.calendar, -1) });
                this.calendarView.getCalendar().prev();
                break;
            case "move-right":
                this.setState({ calendar: CalendarComponent.calcMonths(this.state.calendar, 1) });
                this.calendarView.getCalendar().next();
                break;
            case "move-today":
                const date = new Date();
                this.setState({ calendar: { month: date.getMonth(), year: date.getFullYear() } });
                this.calendarView.getCalendar().showToday();
                break;
        }
    };
}

export class CalendarView extends Controls.BaseControl {
    private _eventSources: EventSourceCollection;

    private _calendar: Calendar.Calendar;
    private _popupMenu: Controls_Menus.PopupMenu;
    private _defaultEvents: FullCalendar.EventObject[];
    private _calendarEventSourceMap: { [sourceId: string]: Calendar.CalendarEventSource } = {};
    private _iterations: Work_Contracts.TeamSettingsIteration[];
    private _currentMember: Calendar_Contracts.ICalendarMember;
    private _processedSprints: { [key: string]: boolean } = {};
    private _selectedTeam: WebApiTeam;

    private static calendarEventFields: string[] = [
        "title",
        "__etag",
        "id",
        "category",
        "iterationId",
        "movable",
        "member",
        "description",
        "icons",
        "eventData",
    ];

    constructor(options: CalendarViewOptions) {
        super(options);
        this._eventSources = new EventSourceCollection([]);
        this._defaultEvents = options.defaultEvents || [];
        this._selectedTeam = options.selectedTeam;
    }

    public initialize() {
        const webContext: WebContext = VSS.getWebContext();
        this._currentMember = {
            displayName: webContext.user.name,
            id: webContext.user.id,
            imageUrl: "",
            uniqueName: webContext.user.email,
            url: "",
        };

        // Create calendar control
        this._calendar = Controls.create(
            Calendar.Calendar,
            this._element,
            $.extend(
                {},
                {
                    fullCalendarOptions: {
                        aspectRatio: this._getCalendarAspectRatio(),
                        handleWindowResize: false, // done by manually adjusting aspect ratio when the window is resized.
                    },
                },
                this._options,
            ),
        );

        EventSourceCollection.create(this._selectedTeam).then(eventSources => {
            if (this._disposed) {
                return;
            }
            eventSources.updateTeamContext(this._selectedTeam);
            this._eventSources = eventSources;
            this._calendar.addEvents(this._defaultEvents);

            this._addDefaultEventSources();

            this._calendar.addCallback(Calendar.FullCalendarCallbackType.viewDestroy, this._viewDestroy.bind(this));
            this._calendar.addCallback(Calendar.FullCalendarCallbackType.eventAfterRender, this._eventAfterRender.bind(this));
            this._calendar.addCallback(Calendar.FullCalendarCallbackType.eventClick, this._eventClick.bind(this));
            this._calendar.addCallback(Calendar.FullCalendarCallbackType.eventDrop, this._eventMoved.bind(this));
            this._calendar.addCallback(Calendar.FullCalendarCallbackType.eventResize, this._eventMoved.bind(this));
            this._calendar.addCallback(Calendar.FullCalendarCallbackType.select, this._daysSelected.bind(this));
        });

        const setAspectRatio = Utils_Core.throttledDelegate(this, 300, () => {
            this._calendar.setOption("aspectRatio", this._getCalendarAspectRatio());
        });
        window.addEventListener("resize", () => {
            setAspectRatio();
        });
    }

    public dispose() {
        if (this._calendar) {
            this._calendar.dispose();
            this._calendar = null;
        }
    }

    public getCalendar() {
        return this._calendar;
    }

    private _viewDestroy() {
        this._processedSprints = {};
    }

    private _getCalendarAspectRatio() {
        const extensionFrame = this._element.closest("#extensionFrame");
        const availableHeight =
            extensionFrame.height() -
            extensionFrame.find(".vss-PivotBar--bar-one-line").height() -
            33 /* 23 margin pixels added */;
        const availableWidth = extensionFrame.find(".calendar-area").width();
        return availableWidth / availableHeight;
    }

    public addEventClicked(): void {
        // Find the free form event source
        const eventSource: Calendar_Contracts.IEventSource = $.grep(this._eventSources.getAllSources(), eventSource => {
            return eventSource.id == "freeForm";
        })[0];
        // Setup the event
        const now = new Date();
        const event: Calendar_Contracts.CalendarEvent = {
            title: "",
            startDate: new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0)).toISOString(), // Create date equivalent to UTC midnight on the current Date
        };

        this._addEvent(event, eventSource);
    }

    private _addDefaultEventSources(): void {
        const eventSources = $.map(this._eventSources.getAllSources(), eventSource => {
            const sourceAndOptions = {
                source: eventSource,
                callbacks: {},
            } as Calendar.SourceAndOptions;
            sourceAndOptions.callbacks[Calendar.FullCalendarCallbackType.eventRender] = this._eventRender.bind(this, eventSource);
            if (eventSource.background) {
                sourceAndOptions.options = { rendering: "background" } as any;
            }
            return sourceAndOptions;
        });

        const calendarEventSources = this._calendar.addEventSources(eventSources);
        calendarEventSources.forEach((calendarEventSource, index) => {
            this._calendarEventSourceMap[eventSources[index].source.id] = calendarEventSource;
        });
    }

    private _getCalendarEventSource(eventSourceId: string): Calendar.CalendarEventSource {
        return this._calendarEventSourceMap[eventSourceId];
    }

    private _eventRender(
        eventSource: Calendar_Contracts.IEventSource,
        event: FullCalendar.EventObject,
        element: JQuery,
        view: FullCalendar.ViewObject,
    ) {
        if (event.rendering !== "background") {
            const eventObject = event as Calendar_Contracts.IExtendedCalendarEventObject;
            const calendarEvent = this._eventObjectToCalendarEvent(eventObject);

            if (calendarEvent.icons) {
                for (const icon of calendarEvent.icons) {
                    const $image = $("<img/>")
                        .attr("src", icon.src)
                        .addClass("event-icon")
                        .addClass(icon.cssClass)
                        .prependTo(element.find(".fc-content"));
                    if (icon.title) {
                        $image.attr("title", icon.title);
                    }
                    if (eventSource.getEnhancer) {
                        eventSource.getEnhancer().then(enhancer => {
                            if (icon.action) {
                                const iconEvent = icon.linkedEvent || calendarEvent;
                                $image.bind("click", icon.action.bind(this, iconEvent));
                            }
                            if (icon.linkedEvent) {
                                enhancer.canEdit(calendarEvent, this._currentMember).then((canEdit: boolean) => {
                                    const commands = [
                                        { rank: 5, id: "Edit", text: "Edit", icon: "icon-edit" },
                                        { rank: 10, id: "Delete", text: "Delete", icon: "icon-delete" },
                                    ];
                                    const tempEvent = this._calendarEventToEventObject(icon.linkedEvent, eventSource);
                                    this._buildContextMenu($image, tempEvent, commands);
                                    $image.bind("click", this._editEvent.bind(this, tempEvent));
                                });
                            }
                        });
                    }
                }
            }

            if (eventSource.getEnhancer) {
                eventSource.getEnhancer().then(enhancer => {
                    enhancer.canEdit(calendarEvent, this._currentMember).then((canEdit: boolean) => {
                        if (canEdit) {
                            const commands = [
                                { rank: 5, id: "Edit", text: "Edit", icon: "icon-edit" },
                                { rank: 10, id: "Delete", text: "Delete", icon: "icon-delete" },
                            ];
                            this._buildContextMenu($(element), eventObject, commands);
                        }
                    });
                });
            }
        }
    }

    private _buildContextMenu($element: JQuery, eventObject: Calendar_Contracts.IExtendedCalendarEventObject, commands: any[]) {
        const menuOptions = {
            items: commands,
            executeAction: e => {
                const command = e.get_commandName();

                switch (command) {
                    case "Edit":
                        this._editEvent(eventObject);
                        break;
                    case "Delete":
                        this._deleteEvent(eventObject);
                        break;
                }
            },
        };

        $element.on("contextmenu", (e: JQueryEventObject) => {
            if (this._popupMenu) {
                this._popupMenu.dispose();
                this._popupMenu = null;
            }
            this._popupMenu = Controls.BaseControl.createIn(
                Controls_Menus.PopupMenu,
                this._element,
                $.extend(
                    {
                        align: "left-bottom",
                    },
                    menuOptions,
                    {
                        items: [{ childItems: Controls_Menus.sortMenuItems(commands) }],
                    },
                ),
            ) as Controls_Menus.PopupMenu;
            Utils_Core.delay(this, 10, () => {
                this._popupMenu.popup(this._element, $element);
            });
            e.preventDefault();
        });
    }

    private _eventAfterRender(event: FullCalendar.EventObject, element: JQuery, view: FullCalendar.ViewObject) {
        if (event.rendering === "background") {
            element.addClass("sprint-event");
            element.data("event", event);
            let contentCellIndex = 0;
            const $contentCells = element.closest(".fc-row.fc-widget-content").find(".fc-content-skeleton table tr td");
            element
                .parent()
                .children()
                .each((index: number, child: Element) => {
                    if ($(child).data("event") && $(child).data("event").title === event.title) {
                        return false; // break;
                    }
                    contentCellIndex += parseInt($(child).attr("colspan"));
                });
            if (!this._processedSprints[event.title]) {
                this._processedSprints[event.title] = true;
                $contentCells.eq(contentCellIndex).append(
                    $("<span/>")
                        .addClass("sprint-label")
                        .text(event.title),
                );
            }
        }
    }

    private _daysSelected(startDate: Date, endDate: Date, allDay: boolean, jsEvent: MouseEvent, view: FullCalendar.ViewObject) {
        if (this._eventSources != null && this._eventSources.getAllSources().length > 0) {
            let addEventSources: Calendar_Contracts.IEventSource[];
            addEventSources = $.grep(this._eventSources.getAllSources(), eventSource => {
                return !!eventSource.addEvent;
            });
            const start = new Date(startDate as any).toISOString();
            const end = Utils_Date.addDays(new Date(endDate as any), -1).toISOString();
            let event: Calendar_Contracts.CalendarEvent;

            if (addEventSources.length > 0) {
                event = {
                    title: "",
                    startDate: start,
                    endDate: end,
                };
            }
            const commandsPromise = this._getAddCommands(addEventSources, event, start, end);

            commandsPromise.then(commands => {
                const menuOptions = {
                    items: commands,
                };

                const dataDate = Utils_Date.format(Utils_Date.shiftToUTC(new Date(event.endDate)), "yyyy-MM-dd"); //2015-04-19
                const $element = $("td.fc-day-top[data-date='" + dataDate + "']");
                if (this._popupMenu) {
                    this._popupMenu.dispose();
                    this._popupMenu = null;
                }
                this._popupMenu = Controls.BaseControl.createIn(
                    Controls_Menus.PopupMenu,
                    this._element,
                    $.extend(menuOptions, {
                        align: "left-bottom",
                        items: [{ childItems: Controls_Menus.sortMenuItems(commands) }],
                    }),
                ) as Controls_Menus.PopupMenu;
                Utils_Core.delay(this, 10, () => {
                    this._popupMenu.popup(this._element, $element);
                });
            });
        }
    }

    private _getAddCommands(
        addEventSources: Calendar_Contracts.IEventSource[],
        event: Calendar_Contracts.CalendarEvent,
        start: string,
        end: string,
    ): PromiseLike<any[]> {
        const commandPromises: PromiseLike<any>[] = [];
        let i = 0;
        for (const source of addEventSources) {
            if (source.getEnhancer) {
                commandPromises.push(
                    source.getEnhancer().then(enhancer => {
                        return enhancer.canAdd(event, this._currentMember).then((canAdd: boolean) => {
                            return {
                                rank: source.order || i,
                                id: event.id,
                                text: Utils_String.format("Add {0}", source.name.toLocaleLowerCase()),
                                icon: enhancer.icon || "icon-add",
                                disabled: !canAdd,
                                action: this._addEvent.bind(this, event, source),
                            } as Controls_Menus.IMenuItemSpec;
                        });
                    }),
                );
            } else {
                commandPromises.push(
                    Promise.resolve({
                        rank: source.order || i,
                        id: event.id,
                        text: Utils_String.format("Add {0}", source.name.toLocaleLowerCase()),
                        icon: "icon-add",
                        disabled: false,
                        action: this._addEvent.bind(this, event, source),
                    } as Controls_Menus.IMenuItemSpec),
                );
            }
            i++;
        }
        return Promise.all(commandPromises);
    }

    private _eventClick(event: FullCalendar.EventObject, jsEvent: MouseEvent, view: FullCalendar.ViewObject) {
        const eventObject = event as Calendar_Contracts.IExtendedCalendarEventObject;
        const source = this._getEventSourceFromEvent(eventObject);
        delete eventObject.source;
        if (source.getEnhancer) {
            source.getEnhancer().then(enhancer => {
                enhancer.canEdit(eventObject as any, this._currentMember).then((canEdit: boolean) => {
                    if (canEdit) {
                        this._editEvent(eventObject);
                    }
                });
            });
        }
    }

    private _eventMoved(
        event: Calendar_Contracts.IExtendedCalendarEventObject,
        dayDelta: number,
        minuteDelta: number,
        revertFunc: Function,
        jsEvent: Event,
        ui: any,
        view: FullCalendar.ViewObject,
    ) {
        const calendarEvent = this._eventObjectToCalendarEvent(event);

        const eventSource: Calendar_Contracts.IEventSource = this._getEventSourceFromEvent(event);

        if (eventSource && eventSource.updateEvent) {
            if (eventSource.getEnhancer) {
                eventSource.getEnhancer().then(enhancer => {
                    enhancer.canEdit(calendarEvent, this._currentMember).then((canEdit: boolean) => {
                        if (canEdit) {
                            eventSource
                                .updateEvent(null, calendarEvent)
                                .then((updatedEvent: Calendar_Contracts.CalendarEvent) => {
                                    // Set underlying source to dirty so refresh picks up new changes
                                    const originalEventSource = this._getCalendarEventSource(eventSource.id);
                                    originalEventSource.state.dirty = true;

                                    // Update dates
                                    event.end = Utils_Date.addDays(new Date(updatedEvent.endDate), 1).toISOString();
                                    event.start = updatedEvent.startDate;
                                    event.__etag = updatedEvent.__etag;
                                    this._calendar.updateEvent(event as any);
                                });
                        }
                    });
                });
            }
        }
    }

    private _addEvent(event: Calendar_Contracts.CalendarEvent, eventSource: Calendar_Contracts.IEventSource) {
        const query = this._calendar.getViewQuery();
        event.member = this._currentMember;
        const membersPromise = this._getTeamMembers();

        const dialogOptions: Calendar_Dialogs.IEventDialogOptions = {
            calendarEvent: event,
            resizable: true,
            source: eventSource,
            membersPromise: membersPromise,
            query: query,
            bowtieVersion: 2,
            okCallback: (calendarEvent: Calendar_Contracts.CalendarEvent) => {
                eventSource.addEvent(calendarEvent).then((addedEvent: Calendar_Contracts.CalendarEvent) => {
                    const calendarEventSource = this._getCalendarEventSource(eventSource.id);
                    calendarEventSource.state.dirty = true;
                    if (addedEvent) {
                        this._calendar.renderEvent(addedEvent, eventSource.id);
                    } else {
                        this._calendar.refreshEvents(calendarEventSource);
                    }
                });
            },
        };

        //Calendar_Dialogs.ExternalEventDialog.showDialog(dialogOptions);
        Controls_Dialogs.Dialog.show(Calendar_Dialogs.EditEventDialog, dialogOptions);
    }

    private _getTeamMembers(): PromiseLike<WebApi_Contracts.IdentityRef[]> {
        const webContext = VSS.getWebContext();

        // Hack to temporarily workaround API compat break in M125 which is being reverted
        let coreClient: any = Service.VssConnection.getConnection().getHttpClient(
            Tfs_Core_WebApi.CoreHttpClient2_2,
            WebApi_Constants.ServiceInstanceTypes.TFS,
        );
        if (!coreClient.getTeamMembers) {
            coreClient = Service.VssConnection.getConnection().getHttpClient(
                Tfs_Core_WebApi.CoreHttpClient4,
                WebApi_Constants.ServiceInstanceTypes.TFS,
            );
        }

        return coreClient.getTeamMembers(webContext.project.name, this._selectedTeam.name);
    }

    private _editEvent(event: Calendar_Contracts.IExtendedCalendarEventObject): void {
        const oldEvent = this._eventObjectToCalendarEvent(event);

        const calendarEvent = $.extend({}, oldEvent);

        const eventSource: Calendar_Contracts.IEventSource = this._getEventSourceFromEvent(event);

        if (eventSource && eventSource.updateEvent) {
            const query = this._calendar.getViewQuery();

            const dialogOptions: Calendar_Dialogs.IEventDialogOptions = {
                calendarEvent: calendarEvent,
                source: eventSource,
                resizable: true,
                isEdit: true,
                membersPromise: this._getTeamMembers(),
                query: query,
                bowtieVersion: 2,
                okCallback: (calendarEvent: Calendar_Contracts.CalendarEvent) => {
                    eventSource.updateEvent(oldEvent, calendarEvent).then((updatedEvent: Calendar_Contracts.CalendarEvent) => {
                        // Set underlying source to dirty so refresh picks up new changes
                        const originalEventSource = this._getCalendarEventSource(eventSource.id);
                        originalEventSource.state.dirty = true;

                        if (updatedEvent) {
                            // Update title
                            event.title = updatedEvent.title;

                            // Update category
                            event.category = updatedEvent.category;

                            // Update color
                            event.color = updatedEvent.category.color;

                            // Update description
                            event.description = updatedEvent.description;

                            // Update data
                            event.eventData = updatedEvent.eventData;

                            //Update dates
                            event.end = Utils_Date.addDays(new Date(updatedEvent.endDate), 1).toISOString();
                            event.start = updatedEvent.startDate;
                            event.__etag = updatedEvent.__etag;
                            // Update iteration
                            event.iterationId = updatedEvent.iterationId;
                            this._calendar.updateEvent(event as any);
                        } else {
                            this._calendar.refreshEvents(originalEventSource);
                        }
                    });
                },
            };

            //Calendar_Dialogs.ExternalEventDialog.showDialog(dialogOptions);
            Controls_Dialogs.Dialog.show(Calendar_Dialogs.EditEventDialog, dialogOptions);
        }
    }

    private _deleteEvent(event: Calendar_Contracts.IExtendedCalendarEventObject): void {
        const calendarEvent = this._eventObjectToCalendarEvent(event);

        const eventSource: Calendar_Contracts.IEventSource = this._getEventSourceFromEvent(event);

        if (eventSource && eventSource.removeEvent) {
            if (confirm("Are you sure you want to delete the event?")) {
                eventSource.removeEvent(calendarEvent).then((calendarEvents: Calendar_Contracts.CalendarEvent[]) => {
                    const originalEventSource = this._getCalendarEventSource(eventSource.id);
                    originalEventSource.state.dirty = true;
                    this._calendar.removeEvent(event.id as string);
                    if (!calendarEvents) {
                        this._calendar.refreshEvents(originalEventSource);
                    }
                });
            }
        }
    }

    private _eventObjectToCalendarEvent(
        eventObject: Calendar_Contracts.IExtendedCalendarEventObject,
    ): Calendar_Contracts.CalendarEvent {
        const start = new Date(eventObject.start as any).toISOString();
        const end = eventObject.end ? Utils_Date.addDays(new Date(eventObject.end as any), -1).toISOString() : start;
        const calendarEvent = {
            startDate: start,
            endDate: end,
        };
        CalendarView.calendarEventFields.forEach(prop => {
            calendarEvent[prop] = eventObject[prop];
        });
        return calendarEvent as any;
    }

    private _calendarEventToEventObject(
        calendarEvent: Calendar_Contracts.CalendarEvent,
        source: Calendar_Contracts.IEventSource,
    ): Calendar_Contracts.IExtendedCalendarEventObject {
        const start = calendarEvent.startDate;
        const end = calendarEvent.endDate ? Utils_Date.addDays(new Date(calendarEvent.endDate), 1).toISOString() : start;
        const eventObject = {
            start: start,
            end: end,
            eventType: source.id,
        };
        CalendarView.calendarEventFields.forEach(prop => {
            eventObject[prop] = calendarEvent[prop];
        });
        return eventObject as any;
    }

    private _getEventSourceFromEvent(event: Calendar_Contracts.IExtendedCalendarEventObject): Calendar_Contracts.IEventSource {
        let calendarEventSource: Calendar_Contracts.IEventSource;
        let eventSource;
        if (event.source) {
            calendarEventSource = event.source.func as Calendar_Contracts.IEventSource;
            if (calendarEventSource) {
                eventSource = (calendarEventSource as any).eventSource;
            }
        } else if (event.eventType) {
            eventSource = this._eventSources.getById(event.eventType);
        }
        return eventSource;
    }
}

interface SummaryComponentProps extends IBaseProps {
    calendarView: IObservableValue<CalendarView>;
}

export class SummaryComponent extends BaseComponent<SummaryComponentProps, {}> {
    private summaryDiv: HTMLElement;
    private summaryView: SummaryView;

    private enhanceControls() {
        if (this.summaryDiv) {
            if (this.summaryView) {
                this.summaryView.dispose();
            }
            const $summaryDiv = $(this.summaryDiv);
            $summaryDiv.empty();
            this.summaryView = Controls.create(SummaryView, $summaryDiv, {
                calendar: this.props.calendarView.value.getCalendar(),
            });
        }
    }

    public componentDidMount() {
        this.enhanceControls();
    }

    public componentDidUpdate() {
        this.enhanceControls();
    }

    public render(): JSX.Element {
        return (
            <div
                className=""
                ref={elem => {
                    this.summaryDiv = elem;
                }}
            />
        );
    }
}

export class SummaryView extends Controls.BaseControl {
    private _calendar: Calendar.Calendar;
    private _rendering: boolean;
    private _statusIndicator: Controls_StatusIndicator.StatusIndicator;

    initialize(): void {
        super.initialize();
        this._rendering = false;
        const $statusContainer = this._element.parent().find(".status-indicator-container");
        this._statusIndicator = Controls.BaseControl.createIn(Controls_StatusIndicator.StatusIndicator, $statusContainer, {
            center: true,
            throttleMinTime: 0,
            imageClass: "big-status-progress",
        }) as Controls_StatusIndicator.StatusIndicator;
        this._statusIndicator.start();
        this._calendar = this._options.calendar;

        // Attach to calendar changes to refresh summary view
        this._calendar.addCallback(Calendar.FullCalendarCallbackType.eventAfterAllRender, () => {
            if (!this._disposed) {
                EventSourceCollection.create().then((eventSources: EventSourceCollection) => {
                    if (!this._rendering && !this._disposed) {
                        this._rendering = true;
                        this._loadSections(eventSources);
                    }
                });
            }
        });
    }

    private _loadSections(eventSources: EventSourceCollection): void {
        // Clear DOM elements first
        this.getElement()
            .children()
            .remove();

        // Sort event sources
        const sources = eventSources
            .getAllSources()
            .slice(0)
            .sort((es1: Calendar_Contracts.IEventSource, es2: Calendar_Contracts.IEventSource) => {
                return es1.order - es2.order;
            });

        const categoryPromises = [];

        // Render each section
        for (const source of sources) {
            categoryPromises.push(this._renderSection(source));
        }

        Promise.all(categoryPromises).then(() => {
            this._rendering = false;
            this._statusIndicator.complete();
        });
    }

    private _renderSection(source: Calendar_Contracts.IEventSource): PromiseLike<void> {
        // Form query using the current date of the calendar
        const query = this._calendar.getViewQuery();
        // Get events categorized
        return source.getCategories(query).then((categories: Calendar_Contracts.IEventCategory[]) => {
            //if (categories.length > 0) {
            const $sectionContainer = newElement("div", "category").appendTo(this.getElement());
            const summaryTitle = source.name;
            return source.getTitleUrl(VSS.getWebContext()).then(titleUrl => {
                if (titleUrl) {
                    const $link = newElement("a", "", summaryTitle);
                    $link.on("click", eventObject => {
                        VSS.getService(VSS.ServiceIds.Navigation).then(
                            (navigationService: Services_Navigation.HostNavigationService) => {
                                // Get current hash value from host url
                                navigationService.openNewWindow(titleUrl, "");
                            },
                        );
                    });

                    const $title = $("<h3>");
                    $link.appendTo($title);
                    $title.appendTo($sectionContainer);
                } else {
                    newElement("h3", "", summaryTitle).appendTo($sectionContainer);
                }
                if (categories.length > 0) {
                    for (const category of categories) {
                        const $titleContainer = newElement("div", "category-title").appendTo($sectionContainer);
                        if (category.imageUrl) {
                            newElement("img", "category-icon")
                                .attr("src", category.imageUrl)
                                .appendTo($titleContainer);
                        } else if (category.color) {
                            const $newElem = newElement("div", "category-color")
                                .css("background-color", category.color)
                                .appendTo($titleContainer);
                            if (source.getEnhancer) {
                                source.getEnhancer().then(enhancer => {
                                    $newElem.bind("click", () => {});
                                });
                            }
                        }
                        newElement("span", "category-titletext", category.title).appendTo($titleContainer);
                        newElement(
                            "div",
                            ["category-subtitle", category.color ? "c-color" : "", category.imageUrl ? "c-icon" : ""].join(" "),
                            category.subTitle,
                        ).appendTo($sectionContainer);
                    }
                } else {
                    newElement("span", "", "(none)").appendTo($sectionContainer);
                }
            });
        });
    }
}
