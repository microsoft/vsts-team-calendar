/// <reference path='../../../typings/VSS.d.ts' />
/// <reference path='../../../typings/TFS.d.ts' />
/// <reference path='../../../typings/q.d.ts' />
/// <reference path='../../../typings/jquery.d.ts' />

import Calendar_Contracts = require("Calendar/Contracts");
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

export class VSOCapacityEventSource implements Calendar_Contracts.IEventSource {

    public id = "daysOff";
    public name = "Days off";
    public order = 30;
    private _events: Calendar_Contracts.CalendarEvent[];


    public getEvents(query?: Calendar_Contracts.IEventQuery): IPromise<Calendar_Contracts.CalendarEvent[]> {

        var result: Calendar_Contracts.CalendarEvent[] = [];
        var deferred = Q.defer<Calendar_Contracts.CalendarEvent[]>();
        var capacityPromises: IPromise<Work_Contracts.TeamMemberCapacity[]>[] = [];
        var iterationTeamDaysOffPromises: IPromise<Work_Contracts.TeamSettingsDaysOff>[] = [];
        this._events = null;

        var webContext = VSS.getWebContext();
        var teamContext: TFS_Core_Contracts.TeamContext = { projectId: webContext.project.id, teamId: webContext.team.id, project: "", team: "" };
        var workClient: Work_Client.WorkHttpClient = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient, WebApi_Constants.ServiceInstanceTypes.TFS);

        workClient.getTeamIterations(teamContext).then(
            (iterations: Work_Contracts.TeamSettingsIteration[]) => {
                if (!iterations || iterations.length === 0) {
                    this._events = result;
                    deferred.resolve(result);
                }
                iterations.forEach((iteration: Work_Contracts.TeamSettingsIteration, index: number, array: Work_Contracts.TeamSettingsIteration[]) => {
                    iterationTeamDaysOffPromises.push(workClient.getTeamDaysOff(teamContext, iteration.id));
                    iterationTeamDaysOffPromises[iterationTeamDaysOffPromises.length - 1].then(
                        (teamDaysOff: Work_Contracts.TeamSettingsDaysOff) => {
                            if (teamDaysOff && teamDaysOff.daysOff && teamDaysOff.daysOff.length) {
                                teamDaysOff.daysOff.forEach((daysOffRange: Work_Contracts.DateRange, i: number, array: Work_Contracts.DateRange[]) => {
                                    var event: any = {};
                                    event.startDate = Utils_Date.shiftToUTC(new Date(daysOffRange.start.valueOf()));
                                    event.endDate = Utils_Date.shiftToUTC(new Date(daysOffRange.end.valueOf()));
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
                        (capacities: Work_Contracts.TeamMemberCapacity[]) => {
                            if (capacities && capacities.length) {
                                for (var i = 0, l = capacities.length; i < l; i++) {
                                    var capacity = capacities[i];
                                    capacity.daysOff.forEach((daysOffRange: Work_Contracts.DateRange, i: number, array: Work_Contracts.DateRange[]) => {
                                        var event: any = {};
                                        event.startDate = Utils_Date.shiftToUTC(new Date(daysOffRange.start.valueOf()));
                                        event.endDate = Utils_Date.shiftToUTC(new Date(daysOffRange.end.valueOf()));
                                        event.title = IdentityHelper.parseUniquefiedIdentityName(capacity.teamMember.displayName) + " Day Off";
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

    public addEvents(events: Calendar_Contracts.CalendarEvent[]): IPromise<Calendar_Contracts.CalendarEvent> {
        this._events = null;
        var deferred = Q.defer();
        var dayOffStart = Utils_Date.shiftToUTC(new Date(events[0].startDate));
        var dayOffEnd = Utils_Date.shiftToUTC(new Date(events[0].endDate));
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
                    var capacityPatch: Work_Contracts.CapacityPatch = {
                        activities: [
                        {
                            "capacityPerDay": 0,
                            "name": null
                        }],
                        daysOff: capacity.daysOff
                    };
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
        var dayOffStart = Utils_Date.shiftToUTC(new Date(events[0].startDate));
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
                        if (Utils_Date.shiftToUTC(dateRange.start).valueOf() === dayOffStart.valueOf()) {
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
                        if (Utils_Date.shiftToUTC(dateRange.start).valueOf() === dayOffStart.valueOf()) {
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
        var dayOffStart = new Date(events[0].startDate);
        var dayOffEnd = new Date(events[0].endDate);
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
                        if (Utils_Date.shiftToUTC(dateRange.start).valueOf() === dayOffStart.valueOf()) {
                            capacityPatch.daysOff[index].end = dayOffEnd;
                            return true;
                        }
                        if (Utils_Date.shiftToUTC(dateRange.end).valueOf() === dayOffEnd.valueOf()) {
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
    
    public getTitleUrl(webContext: WebContext): IPromise<string> {
        var deferred = Q.defer();
        var workClient: Work_Client.WorkHttpClient = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient, WebApi_Constants.ServiceInstanceTypes.TFS);
            var teamContext: TFS_Core_Contracts.TeamContext = { projectId: webContext.project.id, teamId: webContext.team.id, project: "", team: "" };
            workClient.getTeamIterations(teamContext, "current").then((iterations: Work_Contracts.TeamSettingsIteration[]) => {
                if (iterations.length > 0) {
                    var iterationPath = iterations[0].path.substr(iterations[0].path.indexOf('\\') + 1);
                    deferred.resolve(webContext.host.uri + webContext.project.name + "/" + webContext.team.name + "/_backlogs/capacity/" + iterationPath);
                }
                else {
                    deferred.resolve(webContext.host.uri + webContext.project.name + "/" + webContext.team.name + "/_admin/_iterations");         
                }
            }, (error) => {
                deferred.resolve(webContext.host.uri + webContext.project.name + "/" + webContext.team.name + "/_admin/_iterations");       
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
        workClient.getCapacities(teamContext, iterationId).then((capacities: Work_Contracts.TeamMemberCapacity[]) => {
            var foundCapacity = capacities.some((value: Work_Contracts.TeamMemberCapacity, index, array) => {
                if (value.teamMember.id === memberId) {
                    deferred.resolve(value);
                    return true;
                }
                return false;          
            });
            if(!foundCapacity) {
                var value = {
                    daysOff: [],
                    activities: []
                };
                deferred.resolve(value);
            }
        });

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
                        title: IdentityHelper.parseUniquefiedIdentityName(member.displayName),
                        imageUrl: member.imageUrl
                    });
                }

                // TODO calculate the days off
            }
        });

        return categories;
    }

    private _buildTeamImageUrl(hostUri: string, id: string): string {
        return Utils_String.format("{0}_api/_common/IdentityImage?id={1}", hostUri, id);
    }

}

export class IdentityHelper {
    public static IDENTITY_UNIQUEFIEDNAME_SEPERATOR_START = "<";
    public static IDENTITY_UNIQUEFIEDNAME_SEPERATOR_END = ">";
    public static AAD_IDENTITY_UNIQUEFIEDNAME_SEPERATOR_START = "<<";
    public static AAD_IDENTITY_UNIQUEFIEDNAME_SEPERATOR_END = ">>";
    public static AAD_IDENTITY_USER_PREFIX = "user:";
    public static AAD_IDENTITY_GROUP_PREFIX = "group:";
    public static IDENTITY_UNIQUENAME_SEPARATOR = "\\";
    /**
     * Parse a distinct display name string into an identity reference object
     * 
     * @param name A distinct display name for an identity
     */
    public static parseUniquefiedIdentityName(name: string): string {
        if (!name) { return null; }

        var i = name.lastIndexOf(IdentityHelper.AAD_IDENTITY_UNIQUEFIEDNAME_SEPERATOR_START);
        var j = name.lastIndexOf(IdentityHelper.AAD_IDENTITY_UNIQUEFIEDNAME_SEPERATOR_END);
        var isContainer: boolean = false;
        var isAad: boolean = false;
        if (i >= 0 && j > i) {
            isAad = true;
        }

        // replace "<<" with "<" and ">>" with ">" in case of an AAD identity string representation to make further processing easier
        name = name.replace(IdentityHelper.AAD_IDENTITY_UNIQUEFIEDNAME_SEPERATOR_START, IdentityHelper.IDENTITY_UNIQUEFIEDNAME_SEPERATOR_START)
            .replace(IdentityHelper.AAD_IDENTITY_UNIQUEFIEDNAME_SEPERATOR_END, IdentityHelper.IDENTITY_UNIQUEFIEDNAME_SEPERATOR_END);

        i = name.lastIndexOf(IdentityHelper.IDENTITY_UNIQUEFIEDNAME_SEPERATOR_START);
        j = name.lastIndexOf(IdentityHelper.IDENTITY_UNIQUEFIEDNAME_SEPERATOR_END);
        var displayName = name;
        var alias = "";
        var id = "";
        var localScopeId = "";
        if (i >= 0 && j > i) {
            displayName = $.trim(name.substr(0, i));
            if (isAad) {
                // if its an AAD identity, the string would be in format - name <<object id>>
                id = $.trim(name.substr(i + 1, j - i - 1));  // this would be in format objectid\email

                if (id.indexOf(IdentityHelper.AAD_IDENTITY_USER_PREFIX) === 0) {
                    id = id.substr(IdentityHelper.AAD_IDENTITY_USER_PREFIX.length);
                }
                else if (id.indexOf(IdentityHelper.AAD_IDENTITY_GROUP_PREFIX) === 0) {
                    isContainer = true;
                    id = id.substr(IdentityHelper.AAD_IDENTITY_GROUP_PREFIX.length);
                }

                var ii = id.lastIndexOf("\\");
                if (ii > 0) {
                    alias = $.trim(id.substr(ii + 1));
                    id = $.trim(id.substr(0, ii));
                }
            }
            else {
                alias = $.trim(name.substr(i + 1, j - i - 1));
                // If the alias component is just a guid then this is not a uniqueName
                // but the localScopeId which is used only for TFS/AAD groups
                if (Utils_String.isGuid(alias)) {
                    localScopeId = alias;
                    alias = "";
                }
            }
        }
        return displayName;
    }
}
