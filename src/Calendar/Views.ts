import Calendar = require("./Calendar");
import Calendar_Contracts = require("./Contracts");
import Calendar_Dialogs = require("./Dialogs");
import Calendar_Utils_Guid = require("./Utils/Guid");
import Controls = require("VSS/Controls");
import Controls_Dialogs = require("VSS/Controls/Dialogs");
import Controls_Menus = require("VSS/Controls/Menus");
import Controls_Navigation = require("VSS/Controls/Navigation");
import Controls_StatusIndicator = require("VSS/Controls/StatusIndicator");
import Q = require("q");
import Service = require("VSS/Service");
import Tfs_Core_WebApi = require("TFS/Core/RestClient");
import TFS_Core_Contracts = require("TFS/Core/Contracts");
import Utils_Core = require("VSS/Utils/Core");
import Utils_Date = require("VSS/Utils/Date");
import Utils_String = require("VSS/Utils/String");
import WebApi_Constants = require("VSS/WebApi/Constants");
import WebApi_Contracts = require("VSS/WebApi/Contracts");
import Work_Client = require("TFS/Work/RestClient");
import Work_Contracts = require("TFS/Work/Contracts");

import FullCalendar = require("fullcalendar");

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
            var extensionContext = VSS.getExtensionContext();
            var eventSourcesTargetId = extensionContext.publisherId + "." + extensionContext.extensionId + ".calendar-event-sources";
            VSS.getServiceContributions(eventSourcesTargetId).then((contributions) => {
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
    private _currentMember: Calendar_Contracts.ICalendarMember;
    
    private static calendarEventFields: string[] = ["title", "__etag", "id", "category", "iterationId", "movable", "member", "description", "icons", "eventData"];

    constructor(options: CalendarViewOptions) {
        super(options);
        this._eventSources = new EventSourceCollection([]);
        this._defaultEvents = options.defaultEvents || [];
    }

    public initialize() {
        
        var webContext: WebContext = VSS.getWebContext();
        this._currentMember = { displayName: webContext.user.name, id: webContext.user.id, imageUrl: "", uniqueName: webContext.user.email, url: "" };
        
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
            
            this._toolbar.updateItems(this._createToolbarItems());
        });

        var setAspectRatio = Utils_Core.throttledDelegate(this, 300, () => {
            this._calendar.setOption("aspectRatio", this._getCalendarAspectRatio());
        });
        window.addEventListener("resize",() => {
            setAspectRatio();
        });

        this._updateTitle();
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
                            executeAction: this._onToolbarItemClick.bind(this)
                        });

        this._element.find('.menu-container').addClass('toolbar');
    }

    private _createToolbarItems() : any {
        var addDisabled = this._eventSources == null || this._eventSources.getAllSources().length == 0;
        return [
            { id: "new-item", text: "New Item", title: "Add event", icon: "icon-add-small", showText: false, disabled: addDisabled},
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
        var eventSource: Calendar_Contracts.IEventSource = $.grep(this._eventSources.getAllSources(),(eventSource) => { return eventSource.id == "freeForm"; })[0];
        // Setup the event
        var now = new Date();
        var event: Calendar_Contracts.CalendarEvent = {
            title: "",
            startDate: new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0)).toISOString() // Create date equivalent to UTC midnight on the current Date
        };

        this._addEvent(event, eventSource)		
    }

    private _addDefaultEventSources(): void {

        var eventSources = $.map(this._eventSources.getAllSources(),(eventSource) => {
            var sourceAndOptions = <Calendar.SourceAndOptions>{
                source: eventSource,
                callbacks: {}
            };
            sourceAndOptions.callbacks[Calendar.FullCalendarCallbackType.eventRender] = this._eventRender.bind(this, eventSource);
            if (eventSource.background) {
                sourceAndOptions.options = <any>{ rendering: "background" };
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

    private _eventRender(eventSource: Calendar_Contracts.IEventSource, event: FullCalendar.EventObject, element: JQuery, view: FullCalendar.ViewObject) {
        if (event.rendering !== 'background') {
            var eventObject = <Calendar_Contracts.IExtendedCalendarEventObject>event;
            var calendarEvent = this._eventObjectToCalendarEvent(eventObject);
                                       
            if (calendarEvent.icons) {
                $.each(calendarEvent.icons, (index: number, icon: Calendar_Contracts.IEventIcon) => {
                    var $image = $("<img/>").attr("src", icon.src).addClass("event-icon").addClass(icon.cssClass).prependTo(element.find('.fc-content'));
                    if (icon.title) {
                        $image.attr("title", icon.title);
                    }
                    if (eventSource.getEnhancer) {
                        eventSource.getEnhancer().then((enhancer) => {
                            if (icon.action) {
                                var iconEvent = icon.linkedEvent || calendarEvent
                                $image.bind("click", icon.action.bind(this, iconEvent));
                            }
                            if (icon.linkedEvent) {
                                enhancer.canEdit(calendarEvent, this._currentMember).then((canEdit: boolean) => {
                                        var commands = [
                                            { rank: 5, id: "Edit", text: "Edit", icon: "icon-edit" },
                                            { rank: 10, id: "Delete", text: "Delete", icon: "icon-delete" }
                                        ]
                                        var tempEvent = this._calendarEventToEventObject(icon.linkedEvent, eventSource)
                                        this._buildContextMenu($image, tempEvent, commands);
                                        $image.bind("click", this._editEvent.bind(this, tempEvent));
                                });
                            }
                        });
                    }
                });
            }

            if (eventSource.getEnhancer) {
                eventSource.getEnhancer().then((enhancer) => {
                    enhancer.canEdit(calendarEvent, this._currentMember).then((canEdit: boolean) => {
                        if(canEdit) {
                            var commands = [
                                { rank: 5, id: "Edit", text: "Edit", icon: "icon-edit" },
                                { rank: 10, id: "Delete", text: "Delete", icon: "icon-delete" }];
                            this._buildContextMenu($(element), eventObject, commands);
                        }
                    });
                });
            }
        }
    }
    
    private _buildContextMenu($element: JQuery, eventObject: Calendar_Contracts.IExtendedCalendarEventObject, commands: any[]) {
        var menuOptions = {
            items: commands,
            executeAction: (e) => {
                var command = e.get_commandName();

                switch (command) {
                    case "Edit":
                        this._editEvent(eventObject);
                        break;
                    case "Delete":
                        this._deleteEvent(eventObject);
                        break;
                }
            }
        };
        
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

    private _daysSelected(startDate: Date, endDate: Date, allDay: boolean, jsEvent: MouseEvent, view: FullCalendar.ViewObject) {
        if(this._eventSources != null && this._eventSources.getAllSources().length > 0 ) {
            var addEventSources: Calendar_Contracts.IEventSource[];
            addEventSources = $.grep(this._eventSources.getAllSources(), (eventSource) => { return !!eventSource.addEvent; });
            var start = new Date(<any>startDate).toISOString();
            var end = Utils_Date.addDays(new Date(<any>endDate), -1).toISOString();
            var event: Calendar_Contracts.CalendarEvent;
            
            if (addEventSources.length > 0) {
                event =  {
                    title: "",
                    startDate: start,
                    endDate: end
                };
            }
            var commandsPromise = this._getAddCommands(addEventSources, event, start, end);
            
            commandsPromise.then((commands) => {
                var menuOptions = {
                    items: commands
                };

                var dataDate = Utils_Date.format(Utils_Date.shiftToUTC(new Date(event.endDate)), "yyyy-MM-dd"); //2015-04-19
                var $element = $("td.fc-day-top[data-date='" + dataDate + "']");
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
                Utils_Core.delay(this, 10, () => {
                    this._popupMenu.popup(this._element, $element);
                });
            });
        }
    }
    
    private _getAddCommands(addEventSources: Calendar_Contracts.IEventSource[], event: Calendar_Contracts.CalendarEvent, start: string, end: string): IPromise<any[]>{
        var commandPromises: IPromise<any>[] = [];
        for(var i = 0; i < addEventSources.length; i++) {
            ((source: Calendar_Contracts.IEventSource) => {
                if(source.getEnhancer) {
                    commandPromises.push(source.getEnhancer().then((enhancer) => {
                        return enhancer.canAdd(event, this._currentMember).then((canAdd: boolean) => {
                            return <Controls_Menus.IMenuItemSpec>{
                                rank: source.order || i,
                                id: event.id,
                                text: Utils_String.format("Add {0}", source.name.toLocaleLowerCase()),
                                icon: enhancer.icon || "icon-add",
                                disabled: !canAdd,
                                action: this._addEvent.bind(this, event, source)
                            }                            
                        });
                    }));
                }
                else{
                    commandPromises.push(Q.resolve(<Controls_Menus.IMenuItemSpec>{
                        rank: source.order || i,
                        id: event.id,
                        text: Utils_String.format("Add {0}", source.name.toLocaleLowerCase()),
                        icon: "icon-add",
                        disabled: false, 
                        action: this._addEvent.bind(this, event, source)                       
                    }))
                }
            })(addEventSources[i]);
        }
        return Q.all(commandPromises);   
    }

    private _eventClick(event: FullCalendar.EventObject, jsEvent: MouseEvent, view: FullCalendar.ViewObject) {
        var eventObject = <Calendar_Contracts.IExtendedCalendarEventObject> event;
        var source  = this._getEventSourceFromEvent(eventObject);
        if (source.getEnhancer) {
            source.getEnhancer().then((enhancer) => {
                enhancer.canEdit(<any>eventObject, this._currentMember).then((canEdit: boolean) => {
                    if (canEdit) {
                        this._editEvent(eventObject);
                    }             
                });
            });
        }
    }

    private _eventMoved(event: Calendar_Contracts.IExtendedCalendarEventObject, dayDelta: number, minuteDelta: number, revertFunc: Function, jsEvent: Event, ui: any, view: FullCalendar.ViewObject) {
        var calendarEvent = this._eventObjectToCalendarEvent(event);
        
        var eventSource: Calendar_Contracts.IEventSource = this._getEventSourceFromEvent(event);

        if (eventSource && eventSource.updateEvent) {
            if (eventSource.getEnhancer) {
                eventSource.getEnhancer().then((enhancer) => {
                    enhancer.canEdit(calendarEvent, this._currentMember).then((canEdit: boolean) => {
                        if(canEdit) {
                            eventSource.updateEvent(null, calendarEvent).then((updatedEvent: Calendar_Contracts.CalendarEvent) => {
                                // Set underlying source to dirty so refresh picks up new changes
                                var originalEventSource = this._getCalendarEventSource(eventSource.id);
                                originalEventSource.state.dirty = true;
                                                                
                                // Update dates
                                event.end =  Utils_Date.addDays(new Date(updatedEvent.endDate), 1).toISOString();
                                event.start = updatedEvent.startDate;                                
                                event.__etag = updatedEvent.__etag;
                                this._calendar.updateEvent(<any>event);
                            });
                        }
                    });
                });
            }
        }
    }
    
    private _addEvent(event: Calendar_Contracts.CalendarEvent, eventSource: Calendar_Contracts.IEventSource){
        var query = this._calendar.getViewQuery();
        event.member = this._currentMember;
        var membersPromise = this._getTeamMembers();
        
        var dialogOptions : Calendar_Dialogs.IEventDialogOptions = {
            calendarEvent: event,
            resizable: true,
            source: eventSource,
            membersPromise: membersPromise,
            query: query,
            bowtieVersion: 2,
            okCallback: (calendarEvent: Calendar_Contracts.CalendarEvent) => {
                eventSource.addEvent(calendarEvent).then((addedEvent: Calendar_Contracts.CalendarEvent) => {
                    var calendarEventSource = this._getCalendarEventSource(eventSource.id);
                    calendarEventSource.state.dirty = true;
                    if (addedEvent) {
                        this._calendar.renderEvent(addedEvent, eventSource.id);
                    }
                    else {
                        this._calendar.refreshEvents(calendarEventSource);
                    }
                });
            }
        };
        
        //Calendar_Dialogs.ExternalEventDialog.showDialog(dialogOptions);
        Controls_Dialogs.Dialog.show(Calendar_Dialogs.EditEventDialog, dialogOptions);        
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
        var oldEvent = this._eventObjectToCalendarEvent(event);
        
        var calendarEvent = $.extend({}, oldEvent);

        var eventSource: Calendar_Contracts.IEventSource = this._getEventSourceFromEvent(event);

        if (eventSource && eventSource.updateEvent) {
            var query = this._calendar.getViewQuery();
            
            var dialogOptions : Calendar_Dialogs.IEventDialogOptions = {
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
                        var originalEventSource = this._getCalendarEventSource(eventSource.id);
                        originalEventSource.state.dirty = true;
                        
                        if(updatedEvent) {
                            // Update title
                            event.title = updatedEvent.title;

                            // Update category
                            event.category = updatedEvent.category;
                            
                            // Update color
                            event.color = updatedEvent.category.color;
                            
                            // Update description
                            event.description = updatedEvent.description;
                            
                            // Update data
                            event.eventData = updatedEvent.eventData

                            //Update dates
                            event.end = Utils_Date.addDays(new Date(updatedEvent.endDate), 1).toISOString();
                            event.start = updatedEvent.startDate;
                            event.__etag = updatedEvent.__etag; 
                            // Update iteration
                            event.iterationId = updatedEvent.iterationId;
                            this._calendar.updateEvent(<any>event);
                        }
                        else {
                            this._calendar.refreshEvents(originalEventSource);
                        }
                    });
                }
            };
            
            //Calendar_Dialogs.ExternalEventDialog.showDialog(dialogOptions);
            Controls_Dialogs.Dialog.show(Calendar_Dialogs.EditEventDialog, dialogOptions);
        }
    }

    private _deleteEvent(event: Calendar_Contracts.IExtendedCalendarEventObject): void {        
        var calendarEvent = this._eventObjectToCalendarEvent(event)

        var eventSource: Calendar_Contracts.IEventSource = this._getEventSourceFromEvent(event);

        if (eventSource && eventSource.removeEvent) {
            if (confirm("Are you sure you want to delete the event?")) {
                eventSource.removeEvent(calendarEvent).then((calendarEvents: Calendar_Contracts.CalendarEvent[]) => {
                    var originalEventSource = this._getCalendarEventSource(eventSource.id);
                    originalEventSource.state.dirty = true;
                    this._calendar.removeEvent(<string>event.id);
                    if (!calendarEvents) {
                        this._calendar.refreshEvents(originalEventSource);
                    }
                });
            }
        }
    }
    
    private _eventObjectToCalendarEvent(eventObject: Calendar_Contracts.IExtendedCalendarEventObject): Calendar_Contracts.CalendarEvent {
        var start = new Date(<any>eventObject.start).toISOString();
        var end = eventObject.end ? Utils_Date.addDays(new Date(<any>eventObject.end), -1).toISOString() : start;
        var calendarEvent = {
            startDate: start,
            endDate: end
        }
        CalendarView.calendarEventFields.forEach(prop => {
            calendarEvent[prop] = eventObject[prop];
        });
        
        return <any>calendarEvent;
    }
    
    private _calendarEventToEventObject(calendarEvent: Calendar_Contracts.CalendarEvent, source: Calendar_Contracts.IEventSource): Calendar_Contracts.IExtendedCalendarEventObject {
        var start = calendarEvent.startDate;
        var end = calendarEvent.endDate ? Utils_Date.addDays(new Date(calendarEvent.endDate), 1).toISOString(): start;
        var eventObject = {
            start: start,
            end: end,
            eventType: source.id
        }
        CalendarView.calendarEventFields.forEach(prop => {
            eventObject[prop] = calendarEvent[prop];
        });
        return <any>eventObject;
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
}

export class SummaryView extends Controls.BaseControl {
    private _calendar: Calendar.Calendar;
    private _rendering: boolean;
    private _statusIndicator: Controls_StatusIndicator.StatusIndicator;

    initialize(): void {
        super.initialize();
        this._rendering = false;
        var $statusContainer = this._element.parent().find(".status-indicator-container");
        this._statusIndicator = <Controls_StatusIndicator.StatusIndicator>Controls.BaseControl.createIn(Controls_StatusIndicator.StatusIndicator, $statusContainer, {
            center: true,
            throttleMinTime: 0,
            imageClass: "big-status-progress"
        });
        this._statusIndicator.start();
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
            this._statusIndicator.complete();
        });
    }

    private _renderSection(source: Calendar_Contracts.IEventSource): IPromise<Calendar_Contracts.IEventSource> {
        var deferred = Q.defer<Calendar_Contracts.IEventSource>();
        // Form query using the current date of the calendar
        var query = this._calendar.getViewQuery();
        // Get events categorized
        source.getCategories(query).then(
            (categories: Calendar_Contracts.IEventCategory[]) => {
                //if (categories.length > 0) {
                    var $sectionContainer = newElement("div", "category").appendTo(this.getElement());
                    var summaryTitle = source.name;
                    source.getTitleUrl(VSS.getWebContext()).then((titleUrl) => {
                        if (titleUrl) {
                            $("<h3>").html("<a target='_blank' href='" + titleUrl + "'>" + summaryTitle + "</a>").appendTo($sectionContainer);  
                        }
                        else {
                            newElement("h3", "", summaryTitle).appendTo($sectionContainer);                            
                        }
                        if (categories.length > 0) {
                            $.each(categories,(index: number, category: Calendar_Contracts.IEventCategory) => {
                                var $titleContainer = newElement("div", "category-title").appendTo($sectionContainer);
                                if (category.imageUrl) {
                                    newElement("img", "category-icon").attr("src", category.imageUrl).appendTo($titleContainer);
                                }
                                else if (category.color) {
                                    var $newElem = newElement("div", "category-color").css("background-color", category.color).appendTo($titleContainer);
                                    if(source.getEnhancer) {
                                        source.getEnhancer().then((enhancer) => {
                                            $newElem.bind("click", () => {});
                                        });
                                    }
                                }
                                newElement("span", "category-titletext", category.title).appendTo($titleContainer);
                                newElement("div", ["category-subtitle", (category.color ? "c-color" : ""), (category.imageUrl ? "c-icon" : "")].join(" "), category.subTitle).appendTo($sectionContainer);
                            });
                        }
                        else {
                          newElement("span", "", "(none)").appendTo($sectionContainer);
                        }  
                    });
                //}
                deferred.resolve(source);
            });
        return deferred.promise;
    }
}

