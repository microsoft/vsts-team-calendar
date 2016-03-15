/// <reference path='../../../typings/VSS.d.ts' />
/// <reference path='../../../typings/TFS.d.ts' />
/// <reference path='../../../typings/q.d.ts' />
/// <reference path='../../../typings/jquery.d.ts' />

import Calendar_Contracts = require("Calendar/Contracts");
import Calendar_DateUtils = require("Calendar/Utils/Date");
import Calendar_ColorUtils = require("Calendar/Utils/Color");
import Capacity_Enhancer = require("Calendar/Enhancers/VSOCapacityEnhancer");
import Contributions_Contracts = require("VSS/Contributions/Contracts");
import Q = require("q");
import Service = require("VSS/Service");
import Services_ExtensionData = require("VSS/SDK/Services/ExtensionData");
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
    private _categories: Calendar_Contracts.IEventCategory[];
    
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
    
    public load2(): IPromise<boolean> {
        return this.getEvents().then((events: Calendar_Contracts.CalendarEvent[]) => {
            return this._initializeCategories().then((categories: Calendar_Contracts.IEventCategory[]) => { return true; });
        });
    }
       
    public getEnhancer(): IPromise<Calendar_Contracts.IEventEnhancer> {
        if(!this._enhancer){
            this._enhancer = new Capacity_Enhancer.VSOCapacityEnhancer();
        }
        return Q.resolve(this._enhancer);
    }

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
        
        this.getCategories(query).then((categories: Calendar_Contracts.IEventCategory[]) => {
            this._categories = categories;
            this.getIterations().then(
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
                                        // move to local
                                        event.startDate =new Date(daysOffRange.start.valueOf()).toISOString();
                                        event.endDate = new Date(daysOffRange.end.valueOf()).toISOString();
                                        event.title = "Team Day Off";
                                        event.id = this.id + "." + "Everyone" +"." + new Date(Utils_Date.shiftToUTC(daysOffRange.start).valueOf());
                                        event.member = {
                                            displayName: webContext.team.name,
                                            id: webContext.team.id,
                                            imageUrl: this._buildTeamImageUrl(webContext.host.uri, webContext.team.id)
                                        };
                                        event.category = <Calendar_Contracts.IEventCategory>{
                                                id: this.id + "." + "Everyone"
                                            };
                                        event.iterationId = iteration.id;

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
                                            // move to local
                                            event.startDate = new Date(daysOffRange.start.valueOf()).toISOString();
                                            event.endDate = new Date(daysOffRange.end.valueOf()).toISOString();
                                            event.title = IdentityHelper.parseUniquefiedIdentityName(capacity.teamMember.displayName) + " Day Off";
                                            event.id = this.id + "." + capacity.teamMember.uniqueName +"." + new Date(Utils_Date.shiftToUTC(daysOffRange.start).valueOf());
                                            event.member = capacity.teamMember;
                                            event.category = <Calendar_Contracts.IEventCategory>{
                                                id: this.id + "." + capacity.teamMember.uniqueName
                                            };
                                            event.iterationId = iteration.id;

                                            result.push(event);
                                        });
                                    }
                                }

                                return result;
                            });
                    });

                    Q.all(iterationTeamDaysOffPromises).then(
                        () => {
                            Q.all(capacityPromises).then(
                                () => {
                                    this._events = result;
                                    this._updateCategoryForEvents(this._events);
                                    deferred.resolve(result);
                                });
                        },
                        (e: Error) => {
                            deferred.reject(e);
                        });
                },
                (e: Error) => {
                    deferred.reject(e);
                });
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
        var deferred = Q.defer();
        var webContext = VSS.getWebContext();
        VSS.getService("ms.vss-web.data-service").then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            extensionDataService.queryCollectionNames([webContext.team.id]).then((collections: Contributions_Contracts.ExtensionDataCollection[]) => {
                if(collections[0] && collections[0].documents) {
                    this._categories = collections[0].documents.filter((document: any) => { return (!document.startDate  && document.id.split(".")[0] === this.id); });
                }
                else {
                    this._categories = [];
                }
                deferred.resolve(this._categories);
            }, (e: Error) => {
                this._categories = [];
                deferred.resolve(this._categories);
            });
        });
        return deferred.promise;
    }
    
    private _initializeCategories(): IPromise<Calendar_Contracts.IEventCategory[]> {
        var updatedCategories = [];        
        $.each(this._categories, (index: number, category: Calendar_Contracts.IEventCategory) => {
            var deletedEvents = category.events.length == 0;
            $.each(category.events, (index: number, eventId: string) => {
                if(this._events.filter(e => e.id === eventId).length === 0)
                {
                    category.events.splice(category.events.indexOf(eventId));
                    deletedEvents = true;
                }
            });
            // Update or delete category if events have changes
            if(deletedEvents) {
                updatedCategories.push(category);
            }     
        }); 
        return this.updateCategories(updatedCategories);
    }

    public addEvent(event: Calendar_Contracts.CalendarEvent): IPromise<Calendar_Contracts.CalendarEvent> {
        this._events = null;
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
            // Update Team Days Off category
            event.category = <Calendar_Contracts.IEventCategory> {
                id: this.id + "." + "Everyone"
            };
            this._updateCategoryForEvents([event]);
            event.title = "Team Day Off";
            event.iterationId = iterationId;
            this._getTeamDaysOff(workClient, teamContext, iterationId).then((teamDaysOff: Work_Contracts.TeamSettingsDaysOff) => {
                var teamDaysOffPatch: Work_Contracts.TeamSettingsDaysOffPatch = { daysOff: teamDaysOff.daysOff };
                teamDaysOffPatch.daysOff.push({ start: dayOffStart, end: dayOffEnd });
                workClient.updateTeamDaysOff(teamDaysOffPatch, teamContext, iterationId).then((value: Work_Contracts.TeamSettingsDaysOff) => {
                    deferred.resolve(event);
                });
            });
        }
        else {
            // Update member Days Off category
            event.category = <Calendar_Contracts.IEventCategory> {
                id: this.id + "." + event.member.uniqueName
            };
            this._updateCategoryForEvents([event]);            
            event.title = IdentityHelper.parseUniquefiedIdentityName(event.member.displayName) + " Day Off";
            event.iterationId = iterationId;
            
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
                    deferred.resolve(event);
                });
            });
        }
        return deferred.promise;
    }
    
    public addCategory(category: Calendar_Contracts.IEventCategory): IPromise<Calendar_Contracts.IEventCategory> {
        var deferred = Q.defer();
        var webContext = VSS.getWebContext();
        VSS.getService("ms.vss-web.data-service").then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            extensionDataService.createDocument(webContext.team.id, category).then((addedCategory: Calendar_Contracts.IEventCategory) => {
                this._categories.push(addedCategory);
                deferred.resolve(addedCategory);
            }, (e: Error) => {
                deferred.reject(e);
            });
        });
        return deferred.promise;
    }

    public removeEvent(event: Calendar_Contracts.CalendarEvent): IPromise<Calendar_Contracts.CalendarEvent[]> {
        this._events = null;
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
            
        // Update categories
        event.category = null;
        this._updateCategoryForEvents([event]);
        
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
                    deferred.resolve([event]);
                });
            });
        }
        else {
            this._getCapacity(workClient, teamContext, iterationId, memberId).then((capacity: Work_Contracts.TeamMemberCapacity) => {
                var capacityPatch = { daysOff: capacity.daysOff };
                capacityPatch.daysOff.some((dateRange: Work_Contracts.DateRange, index: number, array: Work_Contracts.DateRange[]) => {
                    if (dateRange.start.valueOf() === dayOffStart.valueOf()) {
                        capacityPatch.daysOff.splice(index, 1);
                        return true;
                    }
                    return false;
                });
                workClient.updateCapacity(capacityPatch, teamContext, iterationId, memberId).then((value: Work_Contracts.TeamMemberCapacity) => {
                    deferred.resolve([event]);
                });
            });
        }
        return deferred.promise;
    }
    
    public removeCategory(category: Calendar_Contracts.IEventCategory): IPromise<Calendar_Contracts.IEventCategory[]> {
        var deferred = Q.defer();
        var webContext = VSS.getWebContext();
        VSS.getService("ms.vss-web.data-service").then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
           extensionDataService.deleteDocument(webContext.team.id, category.id).then(() => {
               var categoryInArray: Calendar_Contracts.IEventCategory = $.grep(this._categories, function (cat: Calendar_Contracts.IEventCategory) { return cat.id === category.id})[0];
               var index = this._categories.indexOf(categoryInArray);
               if(index > -1) {
                   this._categories.splice(index, 1);
               }
               deferred.resolve(this._categories);
           }, (e: Error) => {
               deferred.reject(e);
           }); 
        });
        return deferred.promise;
    }

    public updateEvent(oldEvent: Calendar_Contracts.CalendarEvent, newEvent: Calendar_Contracts.CalendarEvent): IPromise<Calendar_Contracts.CalendarEvent> {
        this._events = null;
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
            // Update Team Day Off category
            newEvent.category.id = this.id + "." + "Everyone";
            this._updateCategoryForEvents([newEvent]);
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
                    deferred.resolve(newEvent);
                });
            });
        }
        else {
            // Update member Day Off category
            newEvent.category.id = this.id + "." + newEvent.member.uniqueName;
            this._updateCategoryForEvents([newEvent]);
            this._getCapacity(workClient, teamContext, iterationId, memberId).then((capacity: Work_Contracts.TeamMemberCapacity) => {
                var capacityPatch = { daysOff: capacity.daysOff };
                capacityPatch.daysOff.some((dateRange: Work_Contracts.DateRange, index: number, array: Work_Contracts.DateRange[]) => {
                    if (dateRange.start.valueOf() === dayOffStart.valueOf()) {
                        capacityPatch.daysOff[index].start =new Date(newEvent.startDate);
                        capacityPatch.daysOff[index].end = new Date(newEvent.endDate);
                        return true;
                    }
                    return false;
                });
                workClient.updateCapacity(<any>capacityPatch, teamContext, iterationId, memberId).then((value: Work_Contracts.TeamMemberCapacity) => {
                    deferred.resolve(newEvent);
                });
            });
        }
        return deferred.promise;
    }   
    
    public updateCategories(categories: Calendar_Contracts.IEventCategory[]): IPromise<Calendar_Contracts.IEventCategory[]> {
        var webContext = VSS.getWebContext();
        return VSS.getService("ms.vss-web.data-service").then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            var updatedCategoriesPromises: IPromise<Calendar_Contracts.IEventCategory>[] = [];
            $.each(categories, (index: number, category: Calendar_Contracts.IEventCategory) => {
                // add new category
                if ($.grep(this._categories, (cat: Calendar_Contracts.IEventCategory) => { return cat.id === category.id}).length === 0) {
                    updatedCategoriesPromises.push(this.addCategory(categories[index]).then((updated) => { return null; }));                    
                }
                // remove empty category
                else if (categories[index].events.length == 0) {
                    updatedCategoriesPromises.push(this.removeCategory(categories[index]).then((updated) => { return null; }));
                }
                // updated categories
                else {
                    updatedCategoriesPromises.push(extensionDataService.updateDocument(webContext.team.id, categories[index]).then((updatedCategory: Calendar_Contracts.IEventCategory) => {
                        var categoryInArray: Calendar_Contracts.IEventCategory = $.grep(this._categories, (cat: Calendar_Contracts.IEventCategory) => { return cat.id === category.id})[0];
                        var index = this._categories.indexOf(categoryInArray);
                        if(index > -1) {
                            this._categories.splice(index, 1);
                        }
                        this._categories.push(updatedCategory);
                        return updatedCategory;
                    }));
                }
            });
            return Q.all(updatedCategoriesPromises);
        });
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
    
    private _updateCategoryForEvents(events: Calendar_Contracts.CalendarEvent[]): IPromise<Calendar_Contracts.IEventCategory[]> {
        var webContext = VSS.getWebContext();
        var updatedCategories = [];
        
        for(var i = 0; i < events.length; i++) {
            var event = events[i]
            var isTeam: boolean = event.category && event.category.id === this.id + "." + "Everyone";
            
            // remove event from current category
            var categoryForEvent = $.grep(this._categories, (cat: Calendar_Contracts.IEventCategory) => {
            return cat.events.indexOf(event.id) > -1;
            })[0];
            if(categoryForEvent){
                // Do nothing if category hasn't changed
                if(event.category && (event.category.id === categoryForEvent.id)) {
                    event.category = categoryForEvent;
                    continue;
                }
                var index = categoryForEvent.events.indexOf(event.id);
                categoryForEvent.events.splice(index, 1);
                if(updatedCategories.indexOf(categoryForEvent) < 0) {
                    updatedCategories.push(categoryForEvent);
                }
            }
            // add event to new category
            if(event.category){
                // cate gory already exists
                var newCategory = $.grep(this._categories, (cat: Calendar_Contracts.IEventCategory) => {
                    return cat.id === event.category.id;
                })[0];
                if (!newCategory){
                    // category has been created locally but not stored yet
                    newCategory = $.grep(updatedCategories, (cat: Calendar_Contracts.IEventCategory) => {
                        return cat.id === event.category.id;
                    })[0];
                }
                if(newCategory){
                // category already exists
                    newCategory.events.push(event.id);
                    event.category = newCategory;
                    if(updatedCategories.indexOf(newCategory) < 0) {
                        updatedCategories.push(newCategory);
                    }
                }
                else {
                // category doesn't exist yet
                    var newCategory = event.category;
                    newCategory.title = isTeam ? "Team Days Off" : IdentityHelper.parseUniquefiedIdentityName(event.member.displayName)
                    newCategory.events = [event.id];
                    newCategory.color = Calendar_ColorUtils.generateColor("daysoff");
                    newCategory.imageUrl = isTeam ? this._buildTeamImageUrl(webContext.host.uri, webContext.team.id) : event.member.imageUrl;
                    updatedCategories.push(newCategory);
                }
            }
        }
        // update categories
        return this.updateCategories(updatedCategories);        
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
                    daysOff: []
                };
                deferred.resolve(value);
            }
        });

        return deferred.promise;
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
