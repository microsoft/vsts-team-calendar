import * as Calendar_ColorUtils from "./Utils/Color";
import * as Calendar_Contracts from "./Contracts";
import * as Calendar_Utils_Guid from "./Utils/Guid";
import { timeout } from "./Utils/Promise";
import * as Controls from "VSS/Controls";
import * as Culture from "VSS/Utils/Culture";
import * as Utils_Date from "VSS/Utils/Date";
import * as Utils_String from "VSS/Utils/String";
import * as FullCalendar from "fullCalendar";

export interface CalendarOptions {
    fullCalendarOptions: IDictionaryStringTo<any>;
    viewRender: any;
}

export interface SourceAndOptions {
    source: Calendar_Contracts.IEventSource;
    options?: FullCalendar.Options;
    callbacks?: { [callbackType: number]: Function };
}

export interface CalendarEventSource extends FullCalendar.EventSource {
    (source: Calendar_Contracts.IEventSource, options: FullCalendar.Options): void;

    eventSource: Calendar_Contracts.IEventSource;
    state?: CalendarEventSourceState;
}

export interface CalendarEventSourceState {
    dirty?: boolean;
    cachedEvents?: FullCalendar.EventObject[];
}

export enum FullCalendarCallbackType {
    // General display events
    viewRender,
    viewDestroy,
    dayRender,
    windowResize,

    // Clicking, hovering, selecting
    dayClick,
    eventClick,
    eventMouseover,
    eventMouseout,
    select,
    unselect,

    // Event rendering events
    eventRender,
    eventAfterRender,
    eventDestroy,
    eventAfterAllRender,

    // Dragging, dropping
    eventDragStart,
    eventDragStop,
    eventDrop,
    eventResizeStart,
    eventResizeStop,
    eventResize,
    drop,
    eventReceive,
}

export enum FullCalendarEventRenderingCallbackType {}

export class Calendar extends Controls.Control<CalendarOptions> {
    private _callbacks: { [callbackType: number]: Function[] };
    private _calendarSources: CalendarEventSource[];

    constructor(options: CalendarOptions) {
        super($.extend({ cssClass: "vss-calendar" }, options));
        this._callbacks = {};
        this._calendarSources = [];
    }

    public initialize() {
        super.initialize();

        //  Determine optimal aspect ratio
        const aspectRatio = parseFloat(($(".leftPane").width() / ($(".leftPane").height() - 85)).toFixed(1));
        const firstDay = Culture.getDateTimeFormat().FirstDayOfWeek;

        this._element.fullCalendar(
            $.extend(
                {
                    eventRender: this._getComposedCallback(FullCalendarCallbackType.eventRender),
                    eventAfterRender: this._getComposedCallback(FullCalendarCallbackType.eventAfterRender),
                    eventAfterAllRender: this._getComposedCallback(FullCalendarCallbackType.eventAfterAllRender),
                    eventDestroy: this._getComposedCallback(FullCalendarCallbackType.eventDestroy),
                    viewRender: (view: FullCalendar.ViewObject, element) => this._viewRender(view, element),
                    viewDestroy: this._getComposedCallback(FullCalendarCallbackType.viewDestroy),
                    dayRender: this._getComposedCallback(FullCalendarCallbackType.dayRender),
                    windowResize: this._getComposedCallback(FullCalendarCallbackType.windowResize),
                    dayClick: this._getComposedCallback(FullCalendarCallbackType.dayClick),
                    eventClick: this._getComposedCallback(FullCalendarCallbackType.eventClick),
                    eventMouseover: this._getComposedCallback(FullCalendarCallbackType.eventMouseover),
                    eventMouseout: this._getComposedCallback(FullCalendarCallbackType.eventMouseout),
                    select: this._getComposedCallback(FullCalendarCallbackType.select),
                    unselect: this._getComposedCallback(FullCalendarCallbackType.unselect),
                    eventDragStart: this._getComposedCallback(FullCalendarCallbackType.eventDragStart),
                    eventDragStop: this._getComposedCallback(FullCalendarCallbackType.eventDragStop),
                    eventDrop: this._getComposedCallback(FullCalendarCallbackType.eventDrop),
                    eventResizeStart: this._getComposedCallback(FullCalendarCallbackType.eventResizeStart),
                    eventResizeStop: this._getComposedCallback(FullCalendarCallbackType.eventResizeStop),
                    eventResize: this._getComposedCallback(FullCalendarCallbackType.eventResize),
                    drop: this._getComposedCallback(FullCalendarCallbackType.drop),
                    eventReceive: this._getComposedCallback(FullCalendarCallbackType.eventReceive),
                    header: false,
                    aspectRatio: aspectRatio,
                    columnFormat: "dddd",
                    selectable: true,
                    firstDay: firstDay,
                },
                this._options.fullCalendarOptions,
            ),
        );
    }

    public addEventSource(
        source: Calendar_Contracts.IEventSource,
        options?: FullCalendar.Options,
        callbacks?: { [callbackType: number]: Function },
    ) {
        this.addEventSources([{ source: source, options: options, callbacks: callbacks }]);
    }

    public addEventSources(sources: SourceAndOptions[]): CalendarEventSource[] {
        if (this._disposed) {
            return [];
        }
        sources.forEach(source => {
            const calendarSource = this._createEventSource(source.source, source.options || {});
            this._calendarSources.push(calendarSource);
            this._element.fullCalendar("addEventSource", calendarSource);

            if (source.callbacks) {
                const callbackTypes = Object.keys(source.callbacks);
                for (let i = 0; i < callbackTypes.length; ++i) {
                    const callbackType: FullCalendarCallbackType = parseInt(callbackTypes[i]);
                    if (callbackType === FullCalendarCallbackType.eventAfterAllRender) {
                        // This callback doesn't make sense for individual events.
                        continue;
                    }
                    this.addCallback(
                        callbackType,
                        this._createFilteredCallback(
                            source.callbacks[callbackTypes[i]],
                            event => event["eventType"] === source.source.id,
                        ),
                    );
                }
            }
        });

        return this._calendarSources;
    }

    public getViewQuery(): Calendar_Contracts.IEventQuery {
        const view = this._element.fullCalendar("getView");

        return {
            startDate: Utils_Date.shiftToUTC(new Date(view.start.valueOf() as string)),
            endDate: Utils_Date.shiftToUTC(new Date(view.end.valueOf() as string)),
        };
    }

    public next(): void {
        this._element.fullCalendar("next");
    }

    public prev(): void {
        this._element.fullCalendar("prev");
    }

    public goTo(date: Date) {
        this._element.fullCalendar("gotoDate", date);
    }

    public showToday(): void {
        this._element.fullCalendar("today");
    }

    public getFormattedDate(format: string): string {
        const currentDate: any = this._element.fullCalendar("getDate");
        return currentDate.format(format);
    }

    public renderEvent(event: Calendar_Contracts.CalendarEvent, eventType: string) {
        const end = Utils_Date.addDays(new Date(<any>event.endDate), 1).toISOString();
        const calEvent: any = {
            id: event.id,
            title: event.title,
            description: event.description,
            allDay: true,
            start: event.startDate,
            end: end,
            eventType: eventType,
            iterationId: event.iterationId,
            category: event.category,
            editable: event.movable,
            icons: event.icons,
            eventData: event.eventData,
        };

        if (event.__etag) {
            calEvent.__etag = event.__etag;
        }

        if (event.member) {
            calEvent.member = event.member;
        }

        calEvent.color = event.category.color;
        calEvent.backgroundColor = calEvent.color;
        calEvent.borderColor = calEvent.color;
        calEvent.textColor = calEvent.category.textColor || "#FFFFFF";

        this._element.fullCalendar("renderEvent", calEvent, false);
    }

    public updateEvent(event: FullCalendar.EventObject) {
        event.color = (<any>event).category.color;
        event.backgroundColor = event.color;
        event.borderColor = event.color;

        this._element.fullCalendar("updateEvent", event);
    }

    public setOption(key: string, value: any) {
        this._element.fullCalendar("option", key, value);
    }

    public refreshEvents(eventSource?: FullCalendar.EventSource) {
        if (!eventSource) {
            $(".sprint-label").remove();
            this._element.fullCalendar("refetchEvents");
        } else {
            this._element.fullCalendar("removeEventSource", eventSource);
            this._element.fullCalendar("addEventSource", eventSource);
        }
    }

    public removeEvent(id: string) {
        if (id) {
            this._element.fullCalendar("removeEvents", id);
        }
    }
    private _viewRender(view: FullCalendar.ViewObject, element: JQuery) {
        view["renderId"] = Math.random();
    }

    private _createFilteredCallback(
        original: (event: FullCalendar.EventObject, element: JQuery, view: FullCalendar.ViewObject) => any,
        filter: (event: FullCalendar.EventObject) => boolean,
    ): Function {
        return (event: FullCalendar.EventObject, element: JQuery, view: FullCalendar.ViewObject) => {
            if (filter(event)) {
                return original(event, element, view);
            }
        };
    }

    private _createEventSource(source: Calendar_Contracts.IEventSource, options: FullCalendar.Options): CalendarEventSource {
        const state: CalendarEventSourceState = {};

        const getEventsMethod = (
            start: Date,
            end: Date,
            timezone: string | boolean,
            callback: (events: FullCalendar.EventObject[]) => void,
        ) => {
            if (!state.dirty && state.cachedEvents) {
                callback(state.cachedEvents);
                return;
            }
            const loadSourcePromise = <PromiseLike<Calendar_Contracts.CalendarEvent[]>>source.load();
            timeout(loadSourcePromise, 5000, "Could not load event source " + source.name + ". Request timed out.").then(
                results => {
                    const calendarEvents = results.map((value, index) => {
                        const start = value.startDate;
                        const end = value.endDate ? Utils_Date.addDays(new Date(value.endDate), 1).toISOString() : start;
                        const event: any = {
                            id: value.id || Calendar_Utils_Guid.newGuid(),
                            title: value.title,
                            description: value.description,
                            allDay: true,
                            start: start,
                            end: end,
                            eventType: source.id,
                            rendering: (<any>options).rendering || "",
                            category: value.category,
                            iterationId: value.iterationId,
                            member: value.member,
                            editable: value.movable,
                            icons: value.icons,
                            eventData: value.eventData,
                        };

                        if (value.__etag) {
                            event.__etag = value.__etag;
                        }

                        if ($.isFunction(source.addEvent) && value.category) {
                            const color =
                                <any>value.category.color ||
                                Calendar_ColorUtils.generateColor(
                                    (<string>value.category.title || "uncategorized").toLowerCase(),
                                );
                            event.backgroundColor = color;
                            event.borderColor = color;
                            event.textColor = value.category.textColor || "#FFFFFF";
                        }

                        if ((<any>options).rendering === "background" && value.category) {
                            const color =
                                <any>value.category.color ||
                                Calendar_ColorUtils.generateBackgroundColor(
                                    (<string>event.category.title || "uncategorized").toLowerCase(),
                                );
                            event.backgroundColor = color;
                            event.borderColor = color;
                            event.textColor = value.category.textColor || "#FFFFFF";
                        }

                        return event;
                    });

                    state.dirty = false;
                    state.cachedEvents = calendarEvents;
                    callback(calendarEvents);
                },
                reason => {
                    console.error(
                        Utils_String.format("Error getting event data.\nEvent source: {0}\nReason: {1}", source.name, reason),
                    );
                    callback([]);
                },
            );
        };

        const calendarEventSource: CalendarEventSource = <any>getEventsMethod;
        calendarEventSource.eventSource = source;
        calendarEventSource.state = state;

        return calendarEventSource;
    }

    public addCallback(callbackType: FullCalendarCallbackType, callback: Function) {
        if (!this._callbacks[callbackType]) {
            this._callbacks[callbackType] = [];
        }
        this._callbacks[callbackType].push(callback);
    }

    private _getComposedCallback(callbackType: FullCalendarCallbackType) {
        const args = arguments;
        return (event: FullCalendar.EventObject, element: JQuery, view: FullCalendar.ViewObject): JQuery | boolean => {
            const fns = this._callbacks[callbackType];
            if (!fns) {
                return undefined;
            }
            let broken = false;
            let updatedElement = element;
            for (let i = 0; i < fns.length; ++i) {
                const fn = fns[i];

                const result = fn(event, updatedElement, view);
                if (callbackType === FullCalendarCallbackType.eventRender && result === false) {
                    broken = true;
                    break;
                }
                if (callbackType === FullCalendarCallbackType.eventRender && result instanceof jQuery) {
                    updatedElement = result;
                }
            }
            if (broken) {
                return false;
            }
            return updatedElement;
        };
    }

    /**
     * Adds a one-off event directly to the calendar. Consider adding directly to an event 
     * source rather than calling this method.
     * @param FullCalendar.EventObject[] Array of events to add.
     */
    public addEvents(events: FullCalendar.EventObject[]) {
        events.forEach(event => {
            this._element.fullCalendar("renderEvent", event, true);
        });
    }

    /**
     * Remove events from the calendar
     * @param Array of event IDs to remove, or a filter function, accepting an event, returing true if it is to be removed, false if it is to be kept. Leave null to remove all events.
     * @return EventObject[] - events that were removed.
     */
    public removeEvents(filter: any[] | ((event: FullCalendar.EventObject) => boolean)): FullCalendar.EventObject[] {
        const clientEvents = this._element.fullCalendar("clientEvents", filter);
        this._element.fullCalendar("removeEvents", filter);
        return clientEvents;
    }

    /**
     * Gets the current date of the calendar
     @ return Date
     */
    public getDate(): Date {
        return this._element.fullCalendar("getDate");
    }
}
