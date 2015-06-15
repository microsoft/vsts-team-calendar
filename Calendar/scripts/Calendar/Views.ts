/// <reference path='../ref/VSS/VSS.d.ts' />
/// <reference path='../ref/fullCalendar/fullCalendar.d.ts' />
/// <reference path='../ref/moment/moment.d.ts' />

import Calendar = require("Calendar/Calendar");
import Calendar_Contracts = require("Calendar/Contracts");
import Calendar_Dialogs = require("Calendar/Dialogs");
import Calendar_Utils_Guid = require("Calendar/Utils/Guid");
import Controls = require("VSS/Controls");
import Controls_Common = require("VSS/Controls/Common");
import Controls_Menus = require("VSS/Controls/Menus");
import Controls_Navigation = require("VSS/Controls/Navigation");
import Q = require("q");
import Service = require("VSS/Service");
import Tfs_Core_WebApi = require("TFS/Core/RestClient");
import TFS_Core_Contracts = require("TFS/Core/Contracts");
import Utils_Core = require("VSS/Utils/Core");
import Utils_Date = require("VSS/Utils/Date");
import WebApi_Constants = require("VSS/WebApi/Constants");
import WebApi_Contracts = require("VSS/WebApi/Contracts");
import Work_Client = require("TFS/Work/RestClient");
import Work_Contracts = require("TFS/Work/Contracts");

function newElement(tag: string, className?: string, text?: string): JQuery {
    return $("<" + tag + "/>")
        .addClass(className || "")
        .text(text || "");
}

export class EventSourceCollection {

    private static _deferred: Q.Deferred<any>;
    private _collection: Calendar_Contracts.IEventSource[] = [];
    private _map: { [name: string]: Calendar_Contracts.IEventSource } = {};

    constructor(sources: Calendar_Contracts.IEventSource[]) {
        this._collection = sources || [];
        $.each(this._collection,(index: number, source: Calendar_Contracts.IEventSource) => {
            this._map[source.id] = source;
        });
    }

    public getById(id: string): Calendar_Contracts.IEventSource {
        return this._map[id];
    }

    public getAllSources(): Calendar_Contracts.IEventSource[] {
        return this._collection;
    }

    public static create(): IPromise<EventSourceCollection> {
        if (!this._deferred) {
            this._deferred = Q.defer();
            VSS.getServiceContributions(VSS.getExtensionContext().namespace + "#eventSources").then((contributions) => {
                var servicePromises = $.map(contributions, contribution => contribution.getInstance(contribution.id));
                Q.allSettled(servicePromises).then((promiseStates) => {
                    var services = [];
                    promiseStates.forEach((promiseState, index: number) => {
                        if (promiseState.value) {
                            services.push(promiseState.value);
                        }
                        else {
                            console.log("Failed to get calendar event source instance for: " + contributions[index].id);
                        }
                    });
                    this._deferred.resolve(new EventSourceCollection(services));
                });
            });
        }

        return this._deferred.promise;
    }
}

export interface CalendarViewOptions extends Calendar.CalendarOptions {
    defaultEvents?: FullCalendar.EventObject[];
}

export class CalendarView extends Controls_Navigation.NavigationView {
    private _eventSources: EventSourceCollection;

    private _calendar: Calendar.Calendar;
    private _popupMenu: Controls_Menus.PopupMenu;
    private _toolbar: Controls_Menus.MenuBar;
    private _defaultEvents: FullCalendar.EventObject[];
    private _calendarEventSourceMap: { [sourceId: string]: Calendar.CalendarEventSource; } = {};
    private _iterations: Work_Contracts.TeamSettingsIteration[];

    constructor(options: CalendarViewOptions) {
        super(options);
        this._eventSources = new EventSourceCollection([]);
        this._defaultEvents = options.defaultEvents || [];
    }

    public initialize() {
        
        this._setupToolbar();

        // Create calendar control
        this._calendar = Controls.create(Calendar.Calendar, this._element.find('.calendar-container'), $.extend({}, {
            fullCalendarOptions: {
                aspectRatio: this._getCalendarAspectRatio(),
                handleWindowResize: false // done by manually adjusting aspect ratio when the window is resized.
            }
        }, this._options));

        EventSourceCollection.create().then((eventSources) => {

            this._eventSources = eventSources;
            this._calendar.addEvents(this._defaultEvents);

            this._addDefaultEventSources();

            this._calendar.addCallback(Calendar.FullCalendarCallbackType.eventAfterRender, this._eventAfterRender.bind(this));
            this._calendar.addCallback(Calendar.FullCalendarCallbackType.eventClick, this._eventClick.bind(this));
            this._calendar.addCallback(Calendar.FullCalendarCallbackType.eventDrop, this._eventMoved.bind(this));
            this._calendar.addCallback(Calendar.FullCalendarCallbackType.eventResize, this._eventMoved.bind(this));
            this._calendar.addCallback(Calendar.FullCalendarCallbackType.select, this._daysSelected.bind(this));
        });

        var setAspectRatio = Utils_Core.throttledDelegate(this, 300, function () {
            this._calendar.setOption("aspectRatio", this._getCalendarAspectRatio());
        });
        window.addEventListener("resize",() => {
            setAspectRatio();
        });

        this._updateTitle();

        this._fetchIterationData().then((iterations: Work_Contracts.TeamSettingsIteration[]) => {
            this._iterations = iterations;
        });
    }

    private _isInIteration(date: Date): boolean {
        var inIteration: boolean = false;
        this._iterations.forEach((iteration: Work_Contracts.TeamSettingsIteration, index: number, array: Work_Contracts.TeamSettingsIteration[]) => {
            if (date >= iteration.attributes.startDate && date <= iteration.attributes.finishDate) {
                inIteration = true;
                return;
            }
        });
        return inIteration;
    }

    private _getCalendarAspectRatio() {
        var leftPane = this._element.closest(".splitter>.leftPane");
        var titleBar = this._element.find("div.calendar-title");
        var toolbar = this._element.find("div.menu-container");
        return leftPane.width() / (leftPane.height() - toolbar.height() - titleBar.height() - 20);
    }

    private _setupToolbar () {
        this._toolbar = <Controls_Menus.MenuBar>Controls.BaseControl.createIn(Controls_Menus.MenuBar, this._element.find('.menu-container'), {
                            items: this._createToolbarItems(),
                            executeAction: Utils_Core.delegate(this, this._onToolbarItemClick)
                        });

        this._element.find('.menu-container').addClass('toolbar');
    }

    private _createToolbarItems() : any {
        return [
            { id: "new-item", text: "New Item", title: "Add event", icon: "icon-add-small", showText: false},
            { separator: true },
            { id: "refresh-items", title: "Refresh", icon: "icon-refresh", showText: false },
            { id: "move-today", text: "Today", title: "Today", noIcon: true, showText: true, cssClass: "right-align"},
            { id: "move-next", text: "Next", icon: "icon-drop-right", title: "Next", noIcon: false, showText: false, cssClass: "right-align"},
            { id: "move-prev", text: "Prev", icon: "icon-drop-left", showText: false , title: "Previous", noIcon: false, cssClass: "right-align"}
        ];
    }

    public _onToolbarItemClick(e?: any): any {
        var command = e ? e.get_commandName() : '';
        var result = false;
        switch (command) {
            case "refresh-items":
                this._calendar.refreshEvents();
                break;
            case "new-item":
                this._addEventClicked();
                break;
            case "move-prev":
                this._calendar.prev();
                this._updateTitle();
                break;
            case "move-next":
                this._calendar.next();
                this._updateTitle();
                break;
            case "move-today":
                this._calendar.showToday();
                this._updateTitle();
                break;
            default:
                result = true;
                break;
        }
        return result;
    }

    private _updateTitle() : void {
        var formattedDate = this._calendar.getFormattedDate('MMMM YYYY');
        $('.calendar-title').text(formattedDate);
    }

    private _addEventClicked() : void {
        // Find the free form event source
        var addEventSources: Calendar_Contracts.IEventSource[] = $.grep(this._eventSources.getAllSources(),(eventSource) => { return !!eventSource.addEvents; });

        // Setup the event
        var event: Calendar_Contracts.CalendarEvent = {
            title: "",
            startDate: Utils_Date.shiftToUTC(new Date()),
            eventId: Calendar_Utils_Guid.newGuid()
        };

        this._addEvent(event, addEventSources[0])		
    }

    private _addDefaultEventSources(): void {

        var eventSources = $.map(this._eventSources.getAllSources(),(eventSource) => {
            var sourceAndOptions = <Calendar.SourceAndOptions>{
                source: eventSource,
                callbacks: {}
            };
            sourceAndOptions.callbacks[Calendar.FullCalendarCallbackType.eventRender] = this._eventRender.bind(this, eventSource);
            if (eventSource.background) {
                sourceAndOptions.options = { rendering: "background" };
            }
            return sourceAndOptions;
        });

        var calendarEventSources = this._calendar.addEventSources(eventSources);
        calendarEventSources.forEach((calendarEventSource, index) => {
            this._calendarEventSourceMap[eventSources[index].source.id] = calendarEventSource;
        });
    }

    private _getCalendarEventSource(eventSourceId: string): Calendar.CalendarEventSource {
        return this._calendarEventSourceMap[eventSourceId];
    }

    private _eventRender(eventSource: Calendar_Contracts.IEventSource, event: FullCalendar.EventObject, element: JQuery, view: FullCalendar.View) {
        if (event.rendering !== 'background') {
            var commands = [];

            if (eventSource.updateEvents) {
                commands.push({ rank: 5, id: "Edit", text: "Edit", icon: "icon-edit" });
            }
            if (eventSource.removeEvents) {
                commands.push({ rank: 10, id: "Delete", text: "Delete", icon: "icon-delete" });
            }

            if (commands.length > 0) {
                var menuOptions = {
                    items: commands,
                    executeAction: (e) => {
                        var command = e.get_commandName();

                        switch (command) {
                            case "Edit":
                                this._editEvent(<Calendar_Contracts.IExtendedCalendarEventObject> event);
                                break;
                            case "Delete":
                                this._deleteEvent(<Calendar_Contracts.IExtendedCalendarEventObject> event);
                                break;
                        }
                    }
                };

                var $element = $(element);
                $element.on("contextmenu",(e: JQueryEventObject) => {
                    if (this._popupMenu) {
                        this._popupMenu.dispose();
                        this._popupMenu = null;
                    }
                    this._popupMenu = <Controls_Menus.PopupMenu>Controls.BaseControl.createIn(Controls_Menus.PopupMenu, this._element, $.extend(
                        {
                            align: "left-bottom"
                        },
                        menuOptions,
                        {
                            items: [{ childItems: Controls_Menus.sortMenuItems(commands) }]
                        }));
                    Utils_Core.delay(this, 10, function () {
                        this._popupMenu.popup(this._element, $element);
                    });
                    e.preventDefault();
                });
            }
        }
    }

    private _eventAfterRender(event: FullCalendar.EventObject, element: JQuery, view: FullCalendar.View) {
        if (event.rendering === "background") {
            element.addClass("sprint-event");
            element.data("event", event);
            var contentCellIndex = 0;
            var $contentCells = element.closest(".fc-row.fc-widget-content").find(".fc-content-skeleton table tr td");
            element.parent().children().each((index: number, child: Element) => {
                if ($(child).data("event") && $(child).data("event").title === event.title) {
                    return false; // break;
                }
                contentCellIndex += parseInt($(child).attr("colspan"));
            });
            if (event["sprintProcessedFor"] === undefined || event["sprintProcessedFor"] !== view["renderId"]) {
                event["sprintProcessedFor"] = view["renderId"];
                $contentCells.eq(contentCellIndex).append($("<span/>").addClass("sprint-label").text(event.title));
            }
        }
    }

    private _daysSelected(startDate: Date, endDate: Date, allDay: boolean, jsEvent: MouseEvent, view: FullCalendar.View) {
        //This should not be hard-coded, we should instead figure out context menu commands
        //and dialog contents dynamically from the source
        var addEventSources: Calendar_Contracts.IEventSource[];
        addEventSources = $.grep(this._eventSources.getAllSources(),(eventSource) => { return !!eventSource.addEvents; });

        if (addEventSources.length > 0) {
            var event: Calendar_Contracts.CalendarEvent = {
                title: "",
                startDate: Utils_Date.shiftToUTC(new Date(startDate.valueOf())),
                endDate: Utils_Date.addDays(Utils_Date.shiftToUTC(new Date(endDate.valueOf())), -1),
                eventId: Calendar_Utils_Guid.newGuid()
            };


            var commands = [];

            commands.push({ rank: 5, id: "addEvent", text: "Add event", icon: "icon-add" });
            commands.push({ rank: 10, id: "addDayOff", text: "Add day off", icon: "icon-tfs-build-reason-schedule", disabled: !this._isInIteration(event.startDate) || !this._isInIteration(event.endDate) });
            var menuOptions = {
                items: commands,
                executeAction: (e) => {
                    var command = e.get_commandName();

                    switch (command) {
                        case "addEvent":
                            this._addEvent(event, addEventSources[0]);
                            break;
                        case "addDayOff":
                            this._addDayOff(event, addEventSources[1]);
                            break;
                    }
                }
            };

            var dataDate = Utils_Date.format(event.endDate, "yyyy-MM-dd"); //2015-04-19
            var $element = $("td.fc-day-number[data-date='" + dataDate + "']");
            if (this._popupMenu) {
                this._popupMenu.dispose();
                this._popupMenu = null;
            }
            this._popupMenu = <Controls_Menus.PopupMenu>Controls.BaseControl.createIn(Controls_Menus.PopupMenu, this._element, $.extend(
                menuOptions,
                {
                    align: "left-bottom",
                    items: [{ childItems: Controls_Menus.sortMenuItems(commands) }]

                }));
            Utils_Core.delay(this, 10, function () {
                this._popupMenu.popup(this._element, $element);
            });
        }
    }

    private _eventClick(event: FullCalendar.EventObject, jsEvent: MouseEvent, view: FullCalendar.View) {
        this._editEvent(<Calendar_Contracts.IExtendedCalendarEventObject> event);
    }

    private _eventMoved(event: Calendar_Contracts.IExtendedCalendarEventObject, dayDelta: number, minuteDelta: number, revertFunc: Function, jsEvent: Event, ui: any, view: FullCalendar.View) {
        var calendarEvent: Calendar_Contracts.CalendarEvent = {
            startDate: new Date((<Date>event.start).valueOf()),
            endDate: new Date((<Date>event.end).valueOf()),
            title: event.title,
            eventId: event.id,
            category: event.category,
            member: event.member
        };
        
        var eventSource: Calendar_Contracts.IEventSource = this._getEventSourceFromEvent(event);

        if (eventSource && eventSource.updateEvents) {
            if (eventSource.id === "freeForm") {
                eventSource.updateEvents([calendarEvent]).then((calendarEvents: Calendar_Contracts.CalendarEvent[]) => {
                    // Set underlying source to dirty so refresh picks up new changes
                    var originalEventSource = this._getCalendarEventSource(eventSource.id);
                    originalEventSource.state.dirty = true;

                    // Update title
                    event.title = calendarEvent.title;

                    // Update category
                    event.category = calendarEvent.category;

                    //Update dates
                    
                    event.end = Utils_Date.addDays(new Date(calendarEvent.endDate.valueOf()), 1);
                    event.start = Utils_Date.addDays(new Date(calendarEvent.startDate.valueOf()), 1);
                    this._calendar.updateEvent(event);
                });
            }
        }
    }

    private _addEvent(event: Calendar_Contracts.CalendarEvent, eventSource: Calendar_Contracts.IEventSource) {
        var query = this._calendar.getViewQuery();
        Controls_Common.Dialog.show(Calendar_Dialogs.EditFreeFormEventDialog, {
            event: event,
            title: "Add Event",
            resizable: false,
            okCallback: (calendarEvent: Calendar_Contracts.CalendarEvent) => {
                eventSource.addEvents([calendarEvent]).then((calendarEvents: Calendar_Contracts.CalendarEvent[]) => {
                    var calendarEventSource = this._getCalendarEventSource(eventSource.id);
                    calendarEventSource.state.dirty = true;
                    this._calendar.renderEvent(calendarEvent, eventSource.id);
                });
            },
            categories: this._eventSources.getById("freeForm").getCategories(query).then(categories => categories.map(category => category.title))
        });
    }

    private _addDayOff(event: Calendar_Contracts.CalendarEvent, eventSource: Calendar_Contracts.IEventSource) {
        var webContext: WebContext = VSS.getWebContext();
        event.member = { displayName: webContext.user.name, id: webContext.user.id, imageUrl: "", uniqueName: "", url: "" };
        Controls_Common.Dialog.show(Calendar_Dialogs.EditCapacityEventDialog, {
            event: event,
            title: "Add Days Off",
            resizable: false,
            okCallback: (calendarEvent: Calendar_Contracts.CalendarEvent) => {
                eventSource.addEvents([calendarEvent]).then((calendarEvents: Calendar_Contracts.CalendarEvent[]) => {
                    var calendarEventSource = this._getCalendarEventSource(eventSource.id);
                    calendarEventSource.state.dirty = true;
                    calendarEvent.category = "DaysOff";
                    this._calendar.renderEvent(calendarEvent, eventSource.id);
                });
            },
            membersPromise: this._getTeamMembers()
        });
    }

    private _getTeamMembers(): IPromise<WebApi_Contracts.IdentityRef[]> {
        var deferred = Q.defer<WebApi_Contracts.IdentityRef[]>();

        var webContext = VSS.getWebContext();
        var workClient: Tfs_Core_WebApi.CoreHttpClient = Service.VssConnection
            .getConnection()
            .getHttpClient(Tfs_Core_WebApi.CoreHttpClient, WebApi_Constants.ServiceInstanceTypes.TFS);

        // fetch the wit events
        workClient.getTeamMembers(webContext.project.name, webContext.team.name).then((members: WebApi_Contracts.IdentityRef[]) => {
            deferred.resolve(members);
        });

        return deferred.promise;
    }


    private _editEvent(event: Calendar_Contracts.IExtendedCalendarEventObject): void {
        var calendarEvent: Calendar_Contracts.CalendarEvent = {
            startDate: Utils_Date.addDays(new Date((<Date>event.start).valueOf()), 1),
            endDate: <Date>event.end,
            title: event.title,
            eventId: event.id,
            category: event.category,
            member: event.member
        };

        var eventSource: Calendar_Contracts.IEventSource = this._getEventSourceFromEvent(event);

        if (eventSource && eventSource.updateEvents) {
            if (eventSource.id === "freeForm") {
                var query = this._calendar.getViewQuery();
                Controls_Common.Dialog.show(Calendar_Dialogs.EditFreeFormEventDialog, {
                    event: calendarEvent,
                    title: "Edit Event",
                    resizable: false,
                    okCallback: (calendarEvent: Calendar_Contracts.CalendarEvent) => {
                        eventSource.updateEvents([calendarEvent]).then((calendarEvents: Calendar_Contracts.CalendarEvent[]) => {
                            // Set underlying source to dirty so refresh picks up new changes
                            var originalEventSource = this._getCalendarEventSource(eventSource.id);
                            originalEventSource.state.dirty = true;

                            // Update title
                            event.title = calendarEvent.title;

                            // Update category
                            event.category = calendarEvent.category;

                            //Update dates
                            var end = Utils_Date.addDays(new Date(calendarEvent.endDate.valueOf()), 1);
                            event.end = end;
                            event.start = calendarEvent.startDate;
                            this._calendar.updateEvent(event);
                        });
                    },
                    categories: this._eventSources.getById("freeForm").getCategories(query).then(categories => categories.map(category => category.title))
                });
            }
            else if (eventSource.id === "daysOff") {
                Controls_Common.Dialog.show(Calendar_Dialogs.EditCapacityEventDialog, {
                    event: calendarEvent,
                    title: "Edit Days Off",
                    resizable: false,
                    isEdit: true,
                    okCallback: (calendarEvent: Calendar_Contracts.CalendarEvent) => {
                        eventSource.updateEvents([calendarEvent]).then((calendarEvents: Calendar_Contracts.CalendarEvent[]) => {
                            // Set underlying source to dirty so refresh picks up new changes
                            var originalEventSource = this._getCalendarEventSource(eventSource.id);
                            originalEventSource.state.dirty = true;


                            //Update dates
                            var end = Utils_Date.addDays(new Date(calendarEvent.endDate.valueOf()), 1);
                            event.end = end;
                            event.start = calendarEvent.startDate;

                            this._calendar.updateEvent(event);
                        });
                    },
                });
            }
        }
    }

    private _deleteEvent(event: Calendar_Contracts.IExtendedCalendarEventObject): void {
        var start = new Date((<Date>event.start).valueOf());
        var calendarEvent: Calendar_Contracts.CalendarEvent = {
            startDate: start,
            title: event.title,
            eventId: event.id,
            category: event.category,
            member: event.member
        };

        var eventSource: Calendar_Contracts.IEventSource = this._getEventSourceFromEvent(event);

        if (eventSource && eventSource.removeEvents) {
            if (confirm("Are you sure you want to delete the event?")) {
                eventSource.removeEvents([calendarEvent]).then((calendarEvents: Calendar_Contracts.CalendarEvent[]) => {
                    var originalEventSource = this._getCalendarEventSource(eventSource.id);
                    originalEventSource.state.dirty = true;
                    this._calendar.removeEvent(<string>event.id);
                });
            }
        }
    }

    private _getEventSourceFromEvent(event: Calendar_Contracts.IExtendedCalendarEventObject): Calendar_Contracts.IEventSource {
        var calendarEventSource: Calendar_Contracts.IEventSource;
        var eventSource;
        if (event.source) {
            calendarEventSource = <Calendar_Contracts.IEventSource>event.source.events;
            if (calendarEventSource) {
                eventSource = (<any>calendarEventSource).eventSource;
            }
        } else if (event.eventType) {
            eventSource = this._eventSources.getById(event.eventType);
        }
        return eventSource;
    }

    private _fetchIterationData(): IPromise<Work_Contracts.TeamSettingsIteration[]> {
        var deferred = Q.defer<Work_Contracts.TeamSettingsIteration[]>();
        var iterationPath: string;
        var iterationPromises: IPromise<Work_Contracts.TeamSettingsIteration>[] = [];
        var result : Work_Contracts.TeamSettingsIteration[] = [];

        var webContext = VSS.getWebContext();
        var teamContext: TFS_Core_Contracts.TeamContext = { projectId: webContext.project.id, teamId: webContext.team.id, project: "", team: "" };
        var workClient: Work_Client.WorkHttpClient = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient, WebApi_Constants.ServiceInstanceTypes.TFS);

        workClient.getTeamIterations(teamContext).then(
            (iterations: Work_Contracts.TeamSettingsIteration[]) => {
                iterations.forEach((iteration: Work_Contracts.TeamSettingsIteration, index: number, array: Work_Contracts.TeamSettingsIteration[]) => {
                    result.push(iteration);
                });


                deferred.resolve(result);

            },
            (e: Error) => {
                deferred.reject(e);
            });
        return deferred.promise;
    }
}

export class SummaryView extends Controls.BaseControl {
    private _calendar: Calendar.Calendar;
    private _rendering: boolean;

    initialize(): void {
        super.initialize();
        this._rendering = false;
        this._calendar = <Calendar.Calendar>Controls.Enhancement.getInstance(Calendar.Calendar, $(".vss-calendar"));

        // Attach to calendar changes to refresh summary view
        this._calendar.addCallback(Calendar.FullCalendarCallbackType.eventAfterAllRender,() => {
            EventSourceCollection.create().then((eventSources: EventSourceCollection) => {
                if (!this._rendering) {
                    this._rendering = true;
                    this._loadSections(eventSources);
                }
            });
        });
    }

    private _loadSections(eventSources: EventSourceCollection): void {
        // Clear DOM elements first
        this.getElement().children().remove();

        // Sort event sources
        var sources = eventSources.getAllSources().slice(0).sort(
            (es1: Calendar_Contracts.IEventSource, es2: Calendar_Contracts.IEventSource) => {
                return es1.order - es2.order;
            });

        var categoryPromises = [];
        // Render each section
        $.each(sources,(index: number, source: Calendar_Contracts.IEventSource) => {
            categoryPromises.push(this._renderSection(source));
        });

        Q.all(categoryPromises).then(() => {
            this._rendering = false;
        });
    }

    private _renderSection(source: Calendar_Contracts.IEventSource): IPromise<Calendar_Contracts.IEventSource> {
        var deferred = Q.defer<Calendar_Contracts.IEventSource>();
        // Form query using the current date of the calendar
        var query = this._calendar.getViewQuery();
        // Get events categorized
        source.getCategories(query).then(
            (categories: Calendar_Contracts.IEventCategory[]) => {
                if (categories.length > 0) {
                    var $sectionContainer = newElement("div", "category").appendTo(this.getElement());
                    newElement("h3", "", source.name).appendTo($sectionContainer);
                    $.each(categories,(index: number, category: Calendar_Contracts.IEventCategory) => {
                        var $titleContainer = newElement("div", "category-title").appendTo($sectionContainer);
                        if (category.imageUrl) {
                            newElement("img", "category-icon").attr("src", category.imageUrl).appendTo($titleContainer);
                        }
                        if (category.color) {
                            newElement("div", "category-color").css("background-color", category.color).appendTo($titleContainer);
                        }
                        newElement("span", "category-titletext", category.title).appendTo($titleContainer);
                        newElement("div", ["category-subtitle", (category.color ? "c-color" : ""), (category.imageUrl ? "c-icon" : "")].join(" "), category.subTitle).appendTo($sectionContainer);
                    });
                }
                deferred.resolve(source);
            });
        return deferred.promise;
    }
}

