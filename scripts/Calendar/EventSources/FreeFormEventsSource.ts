/// <reference path='../../../typings/VSS.d.ts' />
/// <reference path='../../../typings/q.d.ts' />
/// <reference path='../../../typings/jquery.d.ts' />

import Calendar_Contracts = require("Calendar/Contracts");
import Calendar_DateUtils = require("Calendar/Utils/Date");
import Calendar_ColorUtils = require("Calendar/Utils/Color");
import Contracts_Platform = require("VSS/Common/Contracts/Platform");
import Contributions_Contracts = require("VSS/Contributions/Contracts");
import ExtensionManagement_RestClient = require("VSS/ExtensionManagement/RestClient");
import Services_ExtensionData = require("VSS/SDK/Services/ExtensionData");
import Q = require("q");
import Service = require("VSS/Service");
import Utils_Core = require("VSS/Utils/Core");
import Utils_Date = require("VSS/Utils/Date");
import Utils_String = require("VSS/Utils/String");
import WebApi_Constants = require("VSS/WebApi/Constants");

export class FreeFormEventsSource implements Calendar_Contracts.IEventSource {

    public id = "freeForm";
    public name = "Events";
    public order = 10;

    private _teamId: string;
    private _events: Calendar_Contracts.CalendarEvent[];

    constructor() {
        var webContext = VSS.getWebContext();
        this._teamId = webContext.team.id;
    }
    
    public load(): IPromise<Calendar_Contracts.CalendarEvent[]> {
        return this.getEvents().then((events: Calendar_Contracts.CalendarEvent[]) => {
            var updatedEvents: Calendar_Contracts.CalendarEvent[] = events.slice();
            $.each(events, (index: number, event: Calendar_Contracts.CalendarEvent) => {
                var start = Utils_Date.parseDateString(event.startDate);
                var end = Utils_Date.parseDateString(event.endDate);
                // For now, skip events with date strngs we can't parse.
                if(!start || !end) {
                    var eventInArray: Calendar_Contracts.CalendarEvent = $.grep(updatedEvents, function (e: Calendar_Contracts.CalendarEvent) { return e.id === event.id; })[0];
                    var index = updatedEvents.indexOf(eventInArray);
                    if (index > -1) {
                        updatedEvents.splice(index, 1);
                    }                    
                }
                else {
                    start = Utils_Date.shiftToUTC(start);
                    end = Utils_Date.shiftToUTC(end);
                    if(start.getHours() !== 0) {
                        // Set dates back to midnight                    
                        start.setHours(0);
                        end.setHours(0);
                        // update the event in the list
                        var updatedEvent = $.extend({}, event);
                        updatedEvent.startDate = Utils_Date.shiftToLocal(start).toISOString();
                        updatedEvent.endDate = Utils_Date.shiftToLocal(end).toISOString();
                        var eventInArray: Calendar_Contracts.CalendarEvent = $.grep(updatedEvents, function (e: Calendar_Contracts.CalendarEvent) { return e.id === updatedEvent.id; })[0];
                        var index = updatedEvents.indexOf(eventInArray);
                        if (index > -1) {
                            updatedEvents.splice(index, 1);
                        }
                        updatedEvents.push(updatedEvent);
                        this.updateEvents([updatedEvent]);
                    }
                }
            });
            return updatedEvents;
        });
    }

    public getEvents(query?: Calendar_Contracts.IEventQuery): IPromise<Calendar_Contracts.CalendarEvent[]> {
        var deferred = Q.defer<Calendar_Contracts.CalendarEvent[]>();
        VSS.getService("ms.vss-web.data-service").then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            extensionDataService.queryCollectionNames([this._teamId]).then(
                (collections: Contributions_Contracts.ExtensionDataCollection[]) => {
                    if (collections[0] && collections[0].documents) {
                        this._events = collections[0].documents;
                    }
                    else {
                        this._events = [];
                    }
                    
                    deferred.resolve(this._events);
                },
                (e: Error) => {
                    this._events = [];
                    deferred.resolve(this._events);
                });
        });

        return deferred.promise;
    }

    public getCategories(query?: Calendar_Contracts.IEventQuery): IPromise<Calendar_Contracts.IEventCategory[]> {
        return this.getEvents().then(
            (events: Calendar_Contracts.CalendarEvent[]) => {
                return this._categorizeEvents(events, query);
            });
    }

    public addEvents(events: Calendar_Contracts.CalendarEvent[]): IPromise<Calendar_Contracts.CalendarEvent> {
        var deferred = Q.defer();
        VSS.getService("ms.vss-web.data-service").then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            extensionDataService.createDocument(this._teamId, events[0]).then(
                (addedEvent: Calendar_Contracts.CalendarEvent) => {
                    this._events.push(addedEvent);
                    deferred.resolve(addedEvent);
                },
                (e: Error) => {
                    deferred.reject(e);
                });
        });
        return deferred.promise;
    }

    public removeEvents(events: Calendar_Contracts.CalendarEvent[]): IPromise<Calendar_Contracts.CalendarEvent[]> {
        var deferred = Q.defer();
        VSS.getService("ms.vss-web.data-service").then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            extensionDataService.deleteDocument(this._teamId, events[0].id).then(
                () => {
                    var eventInArray: Calendar_Contracts.CalendarEvent = $.grep(this._events, function (e: Calendar_Contracts.CalendarEvent) { return e.id === events[0].id; })[0]; //better check here
                    var index = this._events.indexOf(eventInArray);
                    if (index > -1) {
                        this._events.splice(index, 1);
                    }
                    deferred.resolve(this._events);
                },
                (e: Error) => {
                    //Handle event has already been deleted
                    deferred.reject(e);
                });
        });
        return deferred.promise;
    }

    public updateEvents(events: Calendar_Contracts.CalendarEvent[]): IPromise<Calendar_Contracts.CalendarEvent[]> {
        var deferred = Q.defer();
        VSS.getService("ms.vss-web.data-service").then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            extensionDataService.updateDocument(this._teamId, events[0]).then(
                (updatedEvent: Calendar_Contracts.CalendarEvent) => {
                    var eventInArray: Calendar_Contracts.CalendarEvent = $.grep(this._events, function (e: Calendar_Contracts.CalendarEvent) { return e.id === updatedEvent.id; })[0]; //better check here
                    var index = this._events.indexOf(eventInArray);
                    if (index > -1) {
                        this._events.splice(index, 1);
                    }
                    this._events.push(updatedEvent);
                    deferred.resolve(this._events);
                },
                (e: Error) => {
                    //Handle concurrency issue
                    deferred.reject(e);
                });
        },
        (e: Error) => {
            //Handle concurrency issue
            deferred.reject(e);
        });
        return deferred.promise;
    }
    
    public getTitleUrl(webContext: WebContext): IPromise<string> {
        var deferred = Q.defer();
        deferred.resolve("");
        return deferred.promise;
    }

    private _categorizeEvents(events: Calendar_Contracts.CalendarEvent[], query?: Calendar_Contracts.IEventQuery): Calendar_Contracts.IEventCategory[] {
        var categories: Calendar_Contracts.IEventCategory[] = [];
        var categoryMap: { [name: string]: Calendar_Contracts.IEventCategory } = {};
        var countMap: { [name: string]: number } = {};
        $.each(events || [],(index: number, event: Calendar_Contracts.CalendarEvent) => {
            var name = (event.category || "uncategorized").toLocaleLowerCase();
            if (Calendar_DateUtils.eventIn(event, query)) {
                var count = 0;
                if (!categoryMap[name]) {
                    categoryMap[name] = {
                        title: name,
                        subTitle: "",
                        color: Calendar_ColorUtils.generateColor(name)
                    };

                    categories.push(categoryMap[name]);
                    countMap[name] = 0;
                }

                // Update sub title with the count
                count = countMap[name] + 1;
                if (count === 1) {
                    categoryMap[name].subTitle = event.title;
                }
                else {
                    categoryMap[name].subTitle = Utils_String.format("{0} event{1}", count, count > 1 ? "s" : "");
                }
                countMap[name] = count;
            }
        });

        return categories;
    }
}