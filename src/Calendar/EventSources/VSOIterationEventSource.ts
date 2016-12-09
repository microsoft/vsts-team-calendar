import Calendar_Contracts = require("../Contracts");
import Calendar_ColorUtils = require("../Utils/Color");
import Calendar_DateUtils = require("../Utils/Date");
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
    private _categories: Calendar_Contracts.IEventCategory[];

    public load(): IPromise<Calendar_Contracts.CalendarEvent[]> {
        return this.getEvents();
    }

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
                        event.startDate = (iteration.attributes.startDate).toISOString();
                        if (iteration.attributes.finishDate) {
                            event.endDate = (iteration.attributes.finishDate).toISOString();
                        }
                        
                        event.title = iteration.name;
                        var start = new Date(event.startDate);
                        var end = new Date(event.endDate);
                        var startAsUtc = new Date(start.getUTCFullYear(), start.getUTCMonth(), start.getUTCDate(), start.getUTCHours(), start.getUTCMinutes(), start.getUTCSeconds());
                        var endAsUtc = new Date(end.getUTCFullYear(), end.getUTCMonth(), end.getUTCDate(), end.getUTCHours(), end.getUTCMinutes(), end.getUTCSeconds());
                        
                        event.category = <Calendar_Contracts.IEventCategory> {
                            id: this.id + "." + iteration.name,
                            title: iteration.name,
                            subTitle: Utils_String.format("{0} - {1}",
                                Utils_Date.format(startAsUtc, "M"),
                                Utils_Date.format(endAsUtc, "M")),                          
                        }
                        if (this._isCurrentIteration(event)) {
                            event.category.color = Calendar_ColorUtils.generateBackgroundColor(event.title)
                        }
                        else {
                            event.category.color = "#FFFFFF";
                        }

                        result.push(event);
                    }
            });

            result.sort((a, b) => { return new Date(a.startDate).valueOf() - new Date(b.startDate).valueOf(); });
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
    
    public getTitleUrl(webContext: WebContext): IPromise<string> {
        var deferred = Q.defer();
        deferred.resolve(webContext.host.uri + webContext.project.name + "/_admin/_iterations");
        return deferred.promise;
    }

    private _getCategoryData(events: Calendar_Contracts.CalendarEvent[], query: Calendar_Contracts.IEventQuery): Calendar_Contracts.IEventCategory[]{
        var categories: Calendar_Contracts.IEventCategory[] = [];

        $.each(events.splice(0).sort((e1: Calendar_Contracts.CalendarEvent, e2: Calendar_Contracts.CalendarEvent) => {
            if (!e1.startDate || !e2.endDate) {
                return 0;
            }

            return new Date(e1.startDate).getTime() - new Date(e2.startDate).getTime();
        }),
            (index: number, event: Calendar_Contracts.CalendarEvent) => {
                if (Calendar_DateUtils.eventIn(event, query)) {
                    categories.push(event.category);
                }
            });

        return categories;
    }

    private _isCurrentIteration(event: Calendar_Contracts.CalendarEvent): boolean {
        if (event.startDate && event.endDate) {
            var now = new Date();
            var today: number =  new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0)).valueOf();
            return today >= new Date(event.startDate).valueOf() && today <= new Date(event.endDate).valueOf();
        }
        return false;
    }
}