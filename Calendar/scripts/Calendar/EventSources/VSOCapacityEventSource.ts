/// <reference path='../../VSS/References/VSS-Common.d.ts' />
/// <reference path='../../VSS/VSS.SDK.d.ts' />

import Calendar_Contracts = require("Calendar/Contracts");
import Calendar_DateUtils = require("Calendar/Utils/Date");
import Q = require("q");
import Service = require("VSS/Service");
import TFS_Core_Contracts = require("TFS/Core/Contracts");
import Utils_Core = require("VSS/Utils/Core");
import WebApi_Constants = require("VSS/WebApi/Constants");
import Work_Client = require("TFS/Work/RestClient");
import Work_Contracts = require("TFS/Work/Contracts");

export class VSOCapacityEventSource implements Calendar_Contracts.IEventSource {

    public id = "daysOff";
    public name = "Days off";
    public order = 30;
    private _events: Calendar_Contracts.CalendarEvent[];


    public getEvents(query?: Calendar_Contracts.IEventQuery): IPromise<Calendar_Contracts.CalendarEvent[]> {

        var result: Calendar_Contracts.CalendarEvent[] = [];
        var deferred = Q.defer<Calendar_Contracts.CalendarEvent[]>();
        var capacityPromises: IPromise<Work_Contracts.Capacities>[] = [];
        var iterationTeamDaysOffPromises: IPromise<Work_Contracts.TeamSettingsDaysOff>[] = [];
        this._events = null;

        var webContext = VSS.getWebContext();
        var teamContext: TFS_Core_Contracts.TeamContext = { projectId: webContext.project.id, teamId: webContext.team.id, project: "", team: "" };
        var workClient: Work_Client.WorkHttpClient = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient, WebApi_Constants.ServiceInstanceTypes.TFS);

        workClient.getTeamIterations(teamContext).then(
            (iterations: Work_Contracts.TeamSettingsIterations) => {
                iterations.values.forEach((iteration: Work_Contracts.TeamSettingsIteration, index: number, array: Work_Contracts.TeamSettingsIteration[]) => {
                    iterationTeamDaysOffPromises.push(workClient.getTeamDaysOff(teamContext, iteration.id));
                    iterationTeamDaysOffPromises[iterationTeamDaysOffPromises.length - 1].then(
                        (teamDaysOff: Work_Contracts.TeamSettingsDaysOff) => {
                            if (teamDaysOff && teamDaysOff.daysOff && teamDaysOff.daysOff.length) {
                                teamDaysOff.daysOff.forEach((daysOffRange: Work_Contracts.DateRange, i: number, array: Work_Contracts.DateRange[]) => {
                                    var event: any = {};
                                    event.startDate = Utils_Core.DateUtils.shiftToUTC(new Date(daysOffRange.start.valueOf()));
                                    event.endDate = Utils_Core.DateUtils.shiftToUTC(new Date(daysOffRange.end.valueOf()));
                                    event.title = "Team Day Off";
                                    event.member = {
                                        displayName: webContext.team.name,
                                        id: webContext.team.id,
                                        imageUrl: this._buildTeamImageUrl(webContext.host.uri, webContext.team.id)
                                    };
                                    event.category = "DaysOff";

                                    result.push(event);
                                });
                            }
                        });

                    capacityPromises.push(workClient.getCapacities(teamContext, iteration.id));
                    capacityPromises[capacityPromises.length - 1].then(
                        (capacities: Work_Contracts.Capacities) => {
                            if (capacities && capacities.values && capacities.values.length) {
                                for (var i = 0, l = capacities.values.length; i < l; i++) {
                                    var capacity = capacities.values[i];
                                    capacity.daysOff.forEach((daysOffRange: Work_Contracts.DateRange, i: number, array: Work_Contracts.DateRange[]) => {
                                        var event: any = {};
                                        event.startDate = Utils_Core.DateUtils.shiftToUTC(new Date(daysOffRange.start.valueOf()));
                                        event.endDate = Utils_Core.DateUtils.shiftToUTC(new Date(daysOffRange.end.valueOf()));
                                        event.title = capacity.teamMember.displayName + " Day Off";
                                        event.member = capacity.teamMember;
                                        event.category = "DaysOff";

                                        result.push(event);
                                    });
                                }
                            }

                            return result;
                        });

                    Q.all(iterationTeamDaysOffPromises).then(
                        () => {
                            Q.all(capacityPromises).then(
                                () => {
                                    this._events = result;
                                    deferred.resolve(result);
                                });
                        },
                        (e: Error) => {
                            deferred.reject(e);
                        });
                });
            },
            (e: Error) => {
                deferred.reject(e);
            });

        return deferred.promise;
    }

    public getCategories(query: Calendar_Contracts.IEventCategory): IPromise<Calendar_Contracts.IEventCategory[]> {
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
    private _getCategoryData(events: Calendar_Contracts.CalendarEvent[], query: Calendar_Contracts.IEventQuery): Calendar_Contracts.IEventCategory[] {
        var memberMap: { [id: string]: boolean } = {};
        var categories: Calendar_Contracts.IEventCategory[] = [];
        $.each(events,(index: number, event: Calendar_Contracts.CalendarEvent) => {
            if (Calendar_DateUtils.eventIn(event, query)) {
                var member = <Work_Contracts.Member>(<any>event).member;
                if (!memberMap[member.id]) {
                    memberMap[member.id] = true;
                    categories.push({
                        title: member.displayName,
                        imageUrl: member.imageUrl
                    });
                }

                // TODO calculate the days off
            }
        });

        return categories;
    }

    public addEvents(events: Calendar_Contracts.CalendarEvent[]): IPromise<Calendar_Contracts.CalendarEvent[]> {
        this._events = null;
        var deferred = Q.defer();
        var dayOffStart = events[0].startDate;
        var dayOffEnd = events[0].endDate;
        var isTeam: boolean = events[0].member.displayName === "Everyone";
        var memberId: string = events[0].member.id;
        var webContext = VSS.getWebContext();
        var teamContext: TFS_Core_Contracts.TeamContext = { projectId: webContext.project.id, teamId: webContext.team.id, project: "", team: "" };
        var workClient: Work_Client.WorkHttpClient = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient, WebApi_Constants.ServiceInstanceTypes.TFS);
        Calendar_DateUtils.getIterationId(dayOffStart).then((iterationId: string) => {
            if (isTeam) {
                this._getTeamDaysOff(workClient, teamContext, iterationId).then((teamDaysOff: Work_Contracts.TeamSettingsDaysOff) => {
                    var teamDaysOffPatch: Work_Contracts.TeamSettingsDaysOffPatch = { daysOff: teamDaysOff.daysOff };
                    teamDaysOffPatch.daysOff.push({ start: dayOffStart, end: dayOffEnd });
                    workClient.updateTeamDaysOff(teamDaysOffPatch, teamContext, iterationId).then((value: Work_Contracts.TeamSettingsDaysOff) => {
                        deferred.resolve(events[0]);
                    });
                });
            }
            else {
                this._getCapacity(workClient, teamContext, iterationId, memberId).then((capacity: Work_Contracts.TeamMemberCapacity) => {
                    var capacityPatch: Work_Contracts.CapacityPatch = { activities: capacity.activities, daysOff: capacity.daysOff };
                    capacityPatch.daysOff.push({ start: dayOffStart, end: dayOffEnd });
                    workClient.updateCapacity(capacityPatch, teamContext, iterationId, memberId).then((value: Work_Contracts.TeamMemberCapacity) => {
                        deferred.resolve(events[0]);
                    });
                });
            }
        });
        return deferred.promise;
    }

    public removeEvents(events: Calendar_Contracts.CalendarEvent[]): IPromise<Calendar_Contracts.CalendarEvent[]> {
        this._events = null;
        var deferred = Q.defer();
        var dayOffStart = Utils_Core.DateUtils.shiftToUTC(events[0].startDate);
        var memberId = events[0].member.id;
        var isTeam: boolean = events[0].member.uniqueName === undefined;
        var webContext = VSS.getWebContext();
        var teamContext: TFS_Core_Contracts.TeamContext = { projectId: webContext.project.id, teamId: webContext.team.id, project: "", team: "" };
        var workClient: Work_Client.WorkHttpClient = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient, WebApi_Constants.ServiceInstanceTypes.TFS);
        Calendar_DateUtils.getIterationId(dayOffStart).then((iterationId: string) => {
            if (isTeam) {
                this._getTeamDaysOff(workClient, teamContext, iterationId).then((teamDaysOff: Work_Contracts.TeamSettingsDaysOff) => {
                    var teamDaysOffPatch: Work_Contracts.TeamSettingsDaysOffPatch = { daysOff: teamDaysOff.daysOff };
                    teamDaysOffPatch.daysOff.some((dateRange: Work_Contracts.DateRange, index: number, array: Work_Contracts.DateRange[]) => {
                        if (dateRange.start.valueOf() === dayOffStart.valueOf()) {
                            teamDaysOffPatch.daysOff.splice(index, 1);
                            return true;
                        }
                        return false;
                    });
                    workClient.updateTeamDaysOff(teamDaysOffPatch, teamContext, iterationId).then((value: Work_Contracts.TeamSettingsDaysOff) => {
                        deferred.resolve(events[0]);
                    });
                });
            }
            else {
                this._getCapacity(workClient, teamContext, iterationId, memberId).then((capacity: Work_Contracts.TeamMemberCapacity) => {
                    var capacityPatch: Work_Contracts.CapacityPatch = { activities: capacity.activities, daysOff: capacity.daysOff };
                    capacityPatch.daysOff.some((dateRange: Work_Contracts.DateRange, index: number, array: Work_Contracts.DateRange[]) => {
                        if (dateRange.start.valueOf() === dayOffStart.valueOf()) {
                            capacityPatch.daysOff.splice(index, 1);
                            return true;
                        }
                        return false;
                    });
                    workClient.updateCapacity(capacityPatch, teamContext, iterationId, memberId).then((value: Work_Contracts.TeamMemberCapacity) => {
                        deferred.resolve(events[0]);
                    });
                });
            }
        });
        return deferred.promise;
    }

    public updateEvents(events: Calendar_Contracts.CalendarEvent[]): IPromise<Calendar_Contracts.CalendarEvent[]> {
        this._events = null;
        var deferred = Q.defer();
        var dayOffStart = events[0].startDate;
        var dayOffEnd = events[0].endDate;
        var memberId = events[0].member.id;
        var isTeam: boolean = events[0].member.uniqueName === undefined;
        var webContext = VSS.getWebContext();
        var teamContext: TFS_Core_Contracts.TeamContext = { projectId: webContext.project.id, teamId: webContext.team.id, project: "", team: "" };
        var workClient: Work_Client.WorkHttpClient = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient, WebApi_Constants.ServiceInstanceTypes.TFS);
        Calendar_DateUtils.getIterationId(dayOffStart).then((iterationId: string) => {
            if (isTeam) {
                this._getTeamDaysOff(workClient, teamContext, iterationId).then((teamDaysOff: Work_Contracts.TeamSettingsDaysOff) => {
                    var teamDaysOffPatch: Work_Contracts.TeamSettingsDaysOffPatch = { daysOff: teamDaysOff.daysOff };
                    var updated : boolean = teamDaysOffPatch.daysOff.some((dateRange: Work_Contracts.DateRange, index: number, array: Work_Contracts.DateRange[]) => {
                        if (dateRange.start.valueOf() === dayOffStart.valueOf()) {
                            teamDaysOffPatch.daysOff[index].end = dayOffEnd;
                            return true;
                        }
                        if (dateRange.end.valueOf() === dayOffEnd.valueOf()) {
                            teamDaysOffPatch.daysOff[index].start = dayOffStart;
                            return true;
                        }
                        return false;
                    });
                    workClient.updateTeamDaysOff(teamDaysOffPatch, teamContext, iterationId).then((value: Work_Contracts.TeamSettingsDaysOff) => {
                        deferred.resolve(events[0]);
                    });
                });
            }
            else {
                this._getCapacity(workClient, teamContext, iterationId, memberId).then((capacity: Work_Contracts.TeamMemberCapacity) => {
                    var capacityPatch: Work_Contracts.CapacityPatch = { activities: capacity.activities, daysOff: capacity.daysOff };
                    capacityPatch.daysOff.some((dateRange: Work_Contracts.DateRange, index: number, array: Work_Contracts.DateRange[]) => {
                        if (dateRange.start.valueOf() === dayOffStart.valueOf()) {
                            capacityPatch.daysOff[index].end = dayOffEnd;
                            return true;
                        }
                        if (dateRange.end.valueOf() === dayOffEnd.valueOf()) {
                            capacityPatch.daysOff[index].start = dayOffStart;
                            return true;
                        }
                        return false;
                    });
                    workClient.updateCapacity(capacityPatch, teamContext, iterationId, memberId).then((value: Work_Contracts.TeamMemberCapacity) => {
                        deferred.resolve(events[0]);
                    });
                });
            }
        });
        return deferred.promise;
    }

    private _getTeamDaysOff(workClient: Work_Client.WorkHttpClient, teamContext: TFS_Core_Contracts.TeamContext, iterationId): IPromise<Work_Contracts.TeamSettingsDaysOff> {
        var deferred = Q.defer();
        workClient.getTeamDaysOff(teamContext, iterationId).then((value: Work_Contracts.TeamSettingsDaysOff) => {
            deferred.resolve(value);
        });

        return deferred.promise;
    }

    private _getCapacity(workClient: Work_Client.WorkHttpClient, teamContext: TFS_Core_Contracts.TeamContext, iterationId, memberId: string): IPromise<Work_Contracts.TeamMemberCapacity> {
        var deferred = Q.defer();
        workClient.getCapacity(teamContext, iterationId, memberId).then((value: Work_Contracts.TeamMemberCapacity) => {
            deferred.resolve(value);
        });

        return deferred.promise;
    }

    private _buildTeamImageUrl(hostUri: string, id: string): string {
        return Utils_Core.StringUtils.format("{0}_api/_common/IdentityImage?id={1}", hostUri, id);
    }
}