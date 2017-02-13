import Calendar_Contracts = require("../Contracts");
import Calendar_DateUtils = require("../Utils/Date");
import Calendar_ColorUtils = require("../Utils/Color");
import Capacity_Enhancer = require("../Enhancers/VSOCapacityEnhancer");
import Contributions_Contracts = require("VSS/Contributions/Contracts");
import Culture = require("VSS/Utils/Culture")
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
    private _enhancer: Capacity_Enhancer.VSOCapacityEnhancer;
    private _events: Calendar_Contracts.CalendarEvent[];
    private _renderedEvents: Calendar_Contracts.CalendarEvent[];
    private _categoryColor: string = "transparent";
    
    public load(): IPromise<Calendar_Contracts.CalendarEvent[]> {
        return this.getEvents().then((events: Calendar_Contracts.CalendarEvent[]) => {
            $.each(events, (index: number, event: Calendar_Contracts.CalendarEvent) => {
                var start = Utils_Date.shiftToUTC(new Date(event.startDate));
                var end = Utils_Date.shiftToUTC(new Date(event.endDate))
                if(start.getHours() !== 0) {
                    // Set dates back to midnight                    
                    start.setHours(0);
                    end.setHours(0);
                    // update the event in the list
                    var newEvent = $.extend({}, event);
                    newEvent.startDate = Utils_Date.shiftToLocal(start).toISOString();
                    newEvent.endDate = Utils_Date.shiftToLocal(end).toISOString();
                    var eventInArray: Calendar_Contracts.CalendarEvent = $.grep(events, function (e: Calendar_Contracts.CalendarEvent) { return e.id === newEvent.id; })[0];
                    var index = events.indexOf(eventInArray);
                    if (index > -1) {
                        events.splice(index, 1);
                    }
                    events.push(newEvent);
                    
                    // Update event
                    this.updateEvent(event, newEvent);
                }
            });
            
            return events;
        });
    }
       
    public getEnhancer(): IPromise<Calendar_Contracts.IEventEnhancer> {
        if(!this._enhancer){
            this._enhancer = new Capacity_Enhancer.VSOCapacityEnhancer();
        }
        return Q.resolve(this._enhancer);
    }

    public getEvents(query?: Calendar_Contracts.IEventQuery): IPromise<Calendar_Contracts.CalendarEvent[]> {
        
        var deferred = Q.defer<Calendar_Contracts.CalendarEvent[]>();
        var capacityPromises: IPromise<Work_Contracts.TeamMemberCapacity[]>[] = [];
        var iterationTeamDaysOffPromises: IPromise<Work_Contracts.TeamSettingsDaysOff>[] = [];
        var eventMap: { [dateString: string]: Calendar_Contracts.CalendarEvent} = {};
        this._events = null;
        this._renderedEvents = null;        
        var events: Calendar_Contracts.CalendarEvent[] = [];
        var renderedEvents: Calendar_Contracts.CalendarEvent[] = [];

        var webContext = VSS.getWebContext();
        var teamContext: TFS_Core_Contracts.TeamContext = { projectId: webContext.project.id, teamId: webContext.team.id, project: "", team: "" };
        var workClient: Work_Client.WorkHttpClient = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient, WebApi_Constants.ServiceInstanceTypes.TFS);
        
        this.getIterations().then(
            (iterations: Work_Contracts.TeamSettingsIteration[]) => {
                if (!iterations || iterations.length === 0) {
                    this._events = events;
                    this._renderedEvents = renderedEvents
                    deferred.resolve(renderedEvents);
                }
                iterations.forEach((iteration: Work_Contracts.TeamSettingsIteration, index: number, array: Work_Contracts.TeamSettingsIteration[]) => {
                    iterationTeamDaysOffPromises.push(workClient.getTeamDaysOff(teamContext, iteration.id));
                    iterationTeamDaysOffPromises[iterationTeamDaysOffPromises.length - 1].then(
                        (teamDaysOff: Work_Contracts.TeamSettingsDaysOff) => {
                            if (teamDaysOff && teamDaysOff.daysOff && teamDaysOff.daysOff.length) {
                                teamDaysOff.daysOff.forEach((daysOffRange: Work_Contracts.DateRange, i: number, array: Work_Contracts.DateRange[]) => {
                                    var event: any = {};
                                    event.startDate = new Date(daysOffRange.start.valueOf()).toISOString();
                                    event.endDate = new Date(daysOffRange.end.valueOf()).toISOString();
                                    event.title = "Team Day Off";
                                    event.member = {
                                        displayName: webContext.team.name,
                                        id: webContext.team.id,
                                        imageUrl: this._buildTeamImageUrl(webContext.host.uri, webContext.team.id)
                                    };
                                    event.category = <Calendar_Contracts.IEventCategory>{
                                        id: this.id + "." + "Everyone",
                                        title: IdentityHelper.parseUniquefiedIdentityName(event.member.displayName),
                                        imageUrl: this._buildTeamImageUrl(webContext.host.uri, teamContext.teamId),
                                        color: this._categoryColor
                                    };
                                    event.id = this._buildCapacityEventId(event);
                                    event.iterationId = iteration.id;
                                    event.icons = [{
                                        src: event.category.imageUrl,
                                        title: event.title,
                                        linkedEvent: event
                                        }];
                                    
                                    events.push(event);
                                    
                                    // add personal day off event to calendar day off events
                                    var dates = Calendar_DateUtils.getDatesInRange(daysOffRange.start, daysOffRange.end);
                                    for (var i = 0; i < dates.length; i++) {
                                        var date = dates[i].toISOString();
                                        if(!eventMap[date]) {
                                            var regroupedEvent: Calendar_Contracts.CalendarEvent = {
                                                startDate: date,
                                                endDate: date,
                                                member: event.member,
                                                title: "",
                                                id: this.id + "." + date,
                                                category: <Calendar_Contracts.IEventCategory> {
                                                    id: "",
                                                    title: "Grouped Event", 
                                                    color: this._categoryColor,                                                  
                                                },
                                                icons: []
                                            }
                                            eventMap[date] = regroupedEvent;
                                            renderedEvents.push(regroupedEvent);
                                        }
                                        eventMap[date].icons.push(event.icons[0]);
                                    }
                                });
                            }
                            return renderedEvents
                        });

                    capacityPromises.push(workClient.getCapacities(teamContext, iteration.id));
                    capacityPromises[capacityPromises.length - 1].then(
                        (capacities: Work_Contracts.TeamMemberCapacity[]) => {
                            if (capacities && capacities.length) {
                                for (var i = 0, l = capacities.length; i < l; i++) {
                                    var capacity = capacities[i];
                                    capacity.daysOff.forEach((daysOffRange: Work_Contracts.DateRange, i: number, array: Work_Contracts.DateRange[]) => {
                                        var event: any = {};
                                        event.startDate = new Date(daysOffRange.start.valueOf()).toISOString();
                                        event.endDate = new Date(daysOffRange.end.valueOf()).toISOString();
                                        event.title = IdentityHelper.parseUniquefiedIdentityName(capacity.teamMember.displayName) + " Day Off";
                                        event.member = capacity.teamMember;
                                        event.category = <Calendar_Contracts.IEventCategory>{
                                            id: this.id + "." + capacity.teamMember.uniqueName,
                                            title: IdentityHelper.parseUniquefiedIdentityName(event.member.displayName),
                                            imageUrl: event.member.imageUrl,
                                            color: this._categoryColor
                                        };
                                        event.id = this._buildCapacityEventId(event);
                                        event.iterationId = iteration.id;
                                        event.icons = [{
                                            src: event.category.imageUrl,
                                            title: event.title,
                                            linkedEvent: event
                                            }];
                                        events.push(event)
                                        
                                        // add personal day off event to calendar day off events
                                        var dates = Calendar_DateUtils.getDatesInRange(daysOffRange.start, daysOffRange.end);
                                        for (var i = 0; i < dates.length; i++) {
                                            var date = dates[i].toISOString();
                                            if(!eventMap[date]) {
                                                var regroupedEvent: Calendar_Contracts.CalendarEvent = {
                                                    startDate: date,
                                                    endDate: date,
                                                    member: event.member,
                                                    title: "",
                                                    id: this.id + "." + date,
                                                    category: <Calendar_Contracts.IEventCategory> {
                                                        id: "",
                                                        title: "Grouped Event", 
                                                        color: this._categoryColor                                               
                                                    },
                                                    icons: []
                                                }
                                                eventMap[date] = regroupedEvent;
                                                renderedEvents.push(regroupedEvent);
                                            }
                                            eventMap[date].icons.push(event.icons[0]);
                                        }
                                    });
                                }
                            }

                            return renderedEvents;
                        });
                });

                Q.all(iterationTeamDaysOffPromises).then(
                    () => {
                        Q.all(capacityPromises).then(
                            () => {
                                this._events = events;
                                this._renderedEvents = renderedEvents;
                                deferred.resolve(renderedEvents);
                            });
                    },
                    (e: Error) => {
                        deferred.reject(e);
                    });
            },
            (e: Error) => {
                deferred.reject(e);
            });

        return deferred.promise;
    }
    
    public getIterations(): IPromise<Work_Contracts.TeamSettingsIteration[]> {
        var webContext = VSS.getWebContext();
        var teamContext: TFS_Core_Contracts.TeamContext = { projectId: webContext.project.id, teamId: webContext.team.id, project: "", team: "" };
        var workClient: Work_Client.WorkHttpClient = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient, WebApi_Constants.ServiceInstanceTypes.TFS);

        return workClient.getTeamIterations(teamContext)        
    }

    public getCategories(query: Calendar_Contracts.IEventQuery): IPromise<Calendar_Contracts.IEventCategory[]> {        
        var deferred = Q.defer<any>();
        if (this._events) {
            deferred.resolve(this._getCategoryData(this._events.slice(0), query));

        }
        else {
            this.getEvents().then(
                (events: Calendar_Contracts.CalendarEvent[]) => {
                    deferred.resolve(this._getCategoryData(this._events, query));
                });
        }

        return deferred.promise;
    }

    public addEvent(event: Calendar_Contracts.CalendarEvent): IPromise<Calendar_Contracts.CalendarEvent> {
        var deferred = Q.defer();
        var dayOffStart = new Date(event.startDate);
        var dayOffEnd = new Date(event.endDate);
        var isTeam: boolean = event.member.displayName === "Everyone";
        var memberId: string = event.member.id;
        var iterationId: string = event.iterationId;
        var webContext = VSS.getWebContext();
        var teamContext: TFS_Core_Contracts.TeamContext = { projectId: webContext.project.id, teamId: webContext.team.id, project: "", team: "" };
        var workClient: Work_Client.WorkHttpClient = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient, WebApi_Constants.ServiceInstanceTypes.TFS);
                    
        if (isTeam) {
            this._getTeamDaysOff(workClient, teamContext, iterationId).then((teamDaysOff: Work_Contracts.TeamSettingsDaysOff) => {
                var teamDaysOffPatch: Work_Contracts.TeamSettingsDaysOffPatch = { daysOff: teamDaysOff.daysOff };
                teamDaysOffPatch.daysOff.push({ start: dayOffStart, end: dayOffEnd });
                workClient.updateTeamDaysOff(teamDaysOffPatch, teamContext, iterationId).then((value: Work_Contracts.TeamSettingsDaysOff) => {
                    // Resolve null to tell views.js to reload the entire event source instead of re-rendering the updated event
                    deferred.resolve(null);
                });
            });
        }
        else {
            this._getCapacity(workClient, teamContext, iterationId, memberId).then((capacity: Work_Contracts.TeamMemberCapacity) => {                
                var capacityPatch: Work_Contracts.CapacityPatch = { activities: capacity.activities, daysOff: capacity.daysOff };
                capacityPatch.daysOff.push({ start: dayOffStart, end: dayOffEnd });
                workClient.updateCapacity(capacityPatch, teamContext, iterationId, memberId).then((value: Work_Contracts.TeamMemberCapacity) => {
                    // Resolve null to tell views.js to reload the entire event source instead of re-rendering the updated event
                    deferred.resolve(null);
                });
            });
        }
        return deferred.promise;
    }
    
    public removeEvent(event: Calendar_Contracts.CalendarEvent): IPromise<Calendar_Contracts.CalendarEvent[]> {
        var deferred = Q.defer();
        var dayOffStart = new Date(event.startDate);
        var memberId = event.member.id;
        var isTeam: boolean = event.member.uniqueName === undefined;
        var iterationId: string = event.iterationId;
        var webContext = VSS.getWebContext();
        var teamContext: TFS_Core_Contracts.TeamContext = { projectId: webContext.project.id, teamId: webContext.team.id, project: "", team: "" };
        var workClient: Work_Client.WorkHttpClient = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient, WebApi_Constants.ServiceInstanceTypes.TFS);
                    
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
                    // Resolve null to tell views.js to reload the entire event source instead removing one event
                    deferred.resolve(null);
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
                    // Resolve null to tell views.js to reload the entire event source instead removing one event
                    deferred.resolve(null);
                });
            });
        }
        return deferred.promise;
    }

    public updateEvent(oldEvent: Calendar_Contracts.CalendarEvent, newEvent: Calendar_Contracts.CalendarEvent): IPromise<Calendar_Contracts.CalendarEvent> {
        var deferred = Q.defer();
        var dayOffStart = new Date(oldEvent.startDate);
        var memberId = oldEvent.member.id;
        var iterationId = oldEvent.iterationId;
        var isTeam: boolean = oldEvent.member.uniqueName === undefined;
        var webContext = VSS.getWebContext();
        var teamContext: TFS_Core_Contracts.TeamContext = { projectId: webContext.project.id, teamId: webContext.team.id, project: "", team: "" };
        var workClient: Work_Client.WorkHttpClient = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient, WebApi_Constants.ServiceInstanceTypes.TFS);
                    
        if (isTeam) {
            this._getTeamDaysOff(workClient, teamContext, iterationId).then((teamDaysOff: Work_Contracts.TeamSettingsDaysOff) => {
                var teamDaysOffPatch: Work_Contracts.TeamSettingsDaysOffPatch = { daysOff: teamDaysOff.daysOff };
                var updated : boolean = teamDaysOffPatch.daysOff.some((dateRange: Work_Contracts.DateRange, index: number, array: Work_Contracts.DateRange[]) => {
                    if (dateRange.start.valueOf() === dayOffStart.valueOf()) {
                        teamDaysOffPatch.daysOff[index].start = new Date(newEvent.startDate);
                        teamDaysOffPatch.daysOff[index].end = new Date(newEvent.endDate);
                        return true;
                    }
                    return false;
                });
                workClient.updateTeamDaysOff(teamDaysOffPatch, teamContext, iterationId).then((value: Work_Contracts.TeamSettingsDaysOff) => {
                    deferred.resolve(null);
                });
            });
        }
        else {
            this._getCapacity(workClient, teamContext, iterationId, memberId).then((capacity: Work_Contracts.TeamMemberCapacity) => {
                var capacityPatch: Work_Contracts.CapacityPatch = { activities: capacity.activities, daysOff: capacity.daysOff };
                capacityPatch.daysOff.some((dateRange: Work_Contracts.DateRange, index: number, array: Work_Contracts.DateRange[]) => {
                    if (dateRange.start.valueOf() === dayOffStart.valueOf()) {
                        capacityPatch.daysOff[index].start =new Date(newEvent.startDate);
                        capacityPatch.daysOff[index].end = new Date(newEvent.endDate);
                        return true;
                    }
                    return false;
                });
                workClient.updateCapacity(capacityPatch, teamContext, iterationId, memberId).then((value: Work_Contracts.TeamMemberCapacity) => {
                    deferred.resolve(null);
                });
            });
        }
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
                    activities: [{
                        "capacityPerDay": 0,
                        "name": null
                    }],
                    daysOff: []
                };
                deferred.resolve(value);
            }
        });

        return deferred.promise;
    }    
    
    private _getCategoryData(events: Calendar_Contracts.CalendarEvent[], query: Calendar_Contracts.IEventQuery): Calendar_Contracts.IEventCategory[] {
        var memberMap: { [id: string]: Calendar_Contracts.IEventCategory } = {};
        var categories: Calendar_Contracts.IEventCategory[] = [];
        $.each(events,(index: number, event: Calendar_Contracts.CalendarEvent) => {
            if (Calendar_DateUtils.eventIn(event, query)) {
                var member = <Work_Contracts.Member>(<any>event).member;
                if (!memberMap[member.id]) {
                    event.category.events = [event.id];
                    event.category.subTitle = this._getCategorySubTitle(event.category, events, query);
                    memberMap[member.id] = event.category;
                    categories.push(event.category);
                }
                else {
                    var category = memberMap[member.id];
                    category.events.push(event.id);
                    category.subTitle = this._getCategorySubTitle(category, events, query);
                }                
            }
        });

        return categories;
    }
    
    private _getCategorySubTitle(category: Calendar_Contracts.IEventCategory, events: Calendar_Contracts.CalendarEvent[], query: Calendar_Contracts.IEventQuery): string {
        // add up days off per person
        var daysOffInRange: Date[] = [];
        var queryStartInUtc = new Date(query.startDate.getUTCFullYear(), query.startDate.getUTCMonth(), query.startDate.getUTCDate(), query.startDate.getUTCHours(), query.startDate.getUTCMinutes(), query.startDate.getUTCSeconds());
        var queryEndInUtc = new Date(query.endDate.getUTCFullYear(), query.endDate.getUTCMonth(), query.endDate.getUTCDate(), query.endDate.getUTCHours(), query.endDate.getUTCMinutes(), query.endDate.getUTCSeconds());
        category.events.forEach(e => {
            var event = events.filter(event => event.id === e)[0];
            var datesInRange = Calendar_DateUtils.getDatesInRange(Utils_Date.shiftToUTC(new Date(event.startDate)), Utils_Date.shiftToUTC(new Date(event.endDate)));
            datesInRange.forEach((dateToCheck: Date, index: number, array: Date[]) => {
                var dateToCheckInUtc = new Date(dateToCheck.getUTCFullYear(), dateToCheck.getUTCMonth(), dateToCheck.getUTCDate(), dateToCheck.getUTCHours(), dateToCheck.getUTCMinutes(), dateToCheck.getUTCSeconds());
               if (Calendar_DateUtils.isBetween(dateToCheckInUtc, queryStartInUtc, queryEndInUtc)) {
                   daysOffInRange.push(dateToCheck);
               } 
            });
        });
        
        // if user has only one day off, return that date
        if(daysOffInRange.length === 1) {
             return Utils_Date.localeFormat(daysOffInRange[0], Culture.getDateTimeFormat().ShortDatePattern, true);
        }
        // else return total number of days off
        return Utils_String.format("{0} days off", daysOffInRange.length)
    }

    private _buildTeamImageUrl(hostUri: string, id: string): string {
        return Utils_String.format("{0}_api/_common/IdentityImage?id={1}", hostUri, id);
    }
    
    private _buildCapacityEventId(event: Calendar_Contracts.CalendarEvent): string {
        return Utils_String.format("{0}.{1}.{2}", this.id, event.title, event.startDate);
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
