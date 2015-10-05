/// <reference path='../../../typings/vss/VSS.d.ts' />

import Calendar_Contracts = require("Calendar/Contracts");
import Calendar_ColorUtils = require("Calendar/Utils/Color");
import Calendar_DateUtils = require("Calendar/Utils/Date");
import Q = require("q");
import Service = require("VSS/Service");
import TFS_Core_Contracts = require("TFS/Core/Contracts");
import Utils_Core = require("VSS/Utils/Core");
import Utils_Date = require("VSS/Utils/Date");
import Utils_String = require("VSS/Utils/String");
import WebApi_Constants = require("VSS/WebApi/Constants");
import Work_Client = require("TFS/Work/RestClient");
import Work_Contracts = require("TFS/Work/Contracts");

export class VSOIterationEventSource implements Calendar_Contracts.IEventSource {

    public id = "iterations";
    public name = "Iterations";
    public order = 20;
    public background = true;
    private _events: Calendar_Contracts.CalendarEvent[];

    public getEvents(query?: Calendar_Contracts.IEventQuery): IPromise<Calendar_Contracts.CalendarEvent[]> {

        var result: Calendar_Contracts.CalendarEvent[] = [];
        var deferred = Q.defer<Calendar_Contracts.CalendarEvent[]>();
        this._events = null;

        var webContext = VSS.getWebContext();
        var teamContext: TFS_Core_Contracts.TeamContext = { projectId: webContext.project.id, teamId: webContext.team.id, project: "", team: "" };
        var workClient: Work_Client.WorkHttpClient = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient, WebApi_Constants.ServiceInstanceTypes.TFS);

        // fetch the wit events
        workClient.getTeamIterations(teamContext).then(
            (iterations: Work_Contracts.TeamSettingsIteration[]) => {
                iterations.forEach((iteration: Work_Contracts.TeamSettingsIteration, index: number, array: Work_Contracts.TeamSettingsIteration[]) => {
                    if (iteration && iteration.attributes && iteration.attributes.startDate) {
                        var event: any = {};
                        event.startDate = Utils_Date.shiftToUTC(iteration.attributes.startDate);
                        if (iteration.attributes.finishDate) {
                            event.endDate = Utils_Date.shiftToUTC(iteration.attributes.finishDate);
                        }
                        event.title = iteration.name;
                        if (this._isCurrentIteration(event)) {
                            event.category = iteration.name;
                        }

                        result.push(event);
                    }
            });

            result.sort((a, b) => { return a.startDate.valueOf() - b.startDate.valueOf(); });
            this._events = result;
            deferred.resolve(result);

        },
        (e: Error) => {
            deferred.reject(e);
        });

        return deferred.promise;
    }

    public getCategories(query: Calendar_Contracts.IEventQuery): IPromise<Calendar_Contracts.IEventCategory[]> {
        var deferred = Q.defer<any>();
        if (this._events) {
            deferred.resolve(this._getCategoryData(this._events.slice(0), query));
            
        }
        else {
            this.getEvents().then(
                (events: Calendar_Contracts.CalendarEvent[]) => {
                    deferred.resolve(this._getCategoryData(events, query));
                });
        }

        return deferred.promise;
    }

    private _getCategoryData(events: Calendar_Contracts.CalendarEvent[], query: Calendar_Contracts.IEventQuery): Calendar_Contracts.IEventCategory[]{
        var categories: Calendar_Contracts.IEventCategory[] = [];

        $.each(events.splice(0).sort((e1: Calendar_Contracts.CalendarEvent, e2: Calendar_Contracts.CalendarEvent) => {
            if (!e1.startDate || !e2.endDate) {
                return 0;
            }

            return e1.startDate.getTime() - e2.startDate.getTime();
        }),
            (index: number, event: Calendar_Contracts.CalendarEvent) => {
                if (Calendar_DateUtils.eventIn(event, query)) {
                    var category: Calendar_Contracts.IEventCategory = {
                        title: event.title,
                        subTitle: Utils_String.format("{0} - {1}",
                            Utils_Date.format(event.startDate, "M"),
                            Utils_Date.format(event.endDate, "M")),
                    };
                    if (event.category) {
                        category.color = Calendar_ColorUtils.generateBackgroundColor(event.title)
                    }
                    else {
                        category.color = "#FFFFFF";
                    }
                    categories.push(category);
                }
            });

        return categories;
    }

    private _isCurrentIteration(event: Calendar_Contracts.CalendarEvent): boolean {
        if (event.startDate && event.endDate) {
            var today: Date = Utils_Date.shiftToUTC(new Date());
            return today >= event.startDate && today <= event.endDate;
        }
        return false;
    }
}