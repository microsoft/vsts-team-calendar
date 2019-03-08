import { WebApiTeam } from 'TFS/Core/Contracts';
import * as Calendar_Contracts from '../Contracts';
import * as Calendar_DateUtils from '../Utils/Date';
import { realPromise } from "../Utils/Promise";
import * as Capacity_Enhancer from '../Enhancers/VSOCapacityEnhancer';
import * as Culture from 'VSS/Utils/Culture';
import * as Service from 'VSS/Service';
import * as TFS_Core_Contracts from 'TFS/Core/Contracts';
import * as Utils_Date from 'VSS/Utils/Date';
import * as Utils_String from 'VSS/Utils/String';
import * as WebApi_Constants from 'VSS/WebApi/Constants';
import * as Work_Client from 'TFS/Work/RestClient';
import * as Work_Contracts from 'TFS/Work/Contracts';

export class VSOCapacityEventSource implements Calendar_Contracts.IEventSource {
    public id = "daysOff";
    public name = "Days off";
    public order = 30;
    private _enhancer: Capacity_Enhancer.VSOCapacityEnhancer;
    private _events: Calendar_Contracts.CalendarEvent[];
    private _renderedEvents: Calendar_Contracts.CalendarEvent[];
    private _categoryColor: string = "transparent";
    private _team: WebApiTeam;

    constructor(context?: any) {
        this.updateTeamContext(context.team);
    }

    public updateTeamContext(newTeam: WebApiTeam) {
        this._team = newTeam;
    }

    public load(): PromiseLike<Calendar_Contracts.CalendarEvent[]> {
        return this.getEvents().then((events: Calendar_Contracts.CalendarEvent[]) => {
            for (const event of events) {
                const start = Utils_Date.shiftToUTC(new Date(event.startDate));
                const end = Utils_Date.shiftToUTC(new Date(event.endDate));
                if (start.getHours() !== 0) {
                    // Set dates back to midnight
                    start.setHours(0);
                    end.setHours(0);
                    // update the event in the list
                    const newEvent = $.extend({}, event);
                    newEvent.startDate = Utils_Date.shiftToLocal(start).toISOString();
                    newEvent.endDate = Utils_Date.shiftToLocal(end).toISOString();
                    const eventInArray: Calendar_Contracts.CalendarEvent = $.grep(events, function(
                        e: Calendar_Contracts.CalendarEvent,
                    ) {
                        return e.id === newEvent.id;
                    })[0];
                    const index = events.indexOf(eventInArray);
                    if (index > -1) {
                        events.splice(index, 1);
                    }
                    events.push(newEvent);

                    // Update event
                    this.updateEvent(event, newEvent);
                }
            }

            return events;
        });
    }

    public getEnhancer(): PromiseLike<Calendar_Contracts.IEventEnhancer> {
        if (!this._enhancer) {
            this._enhancer = new Capacity_Enhancer.VSOCapacityEnhancer();
        }
        return Promise.resolve(this._enhancer);
    }

    public getEvents(query?: Calendar_Contracts.IEventQuery): PromiseLike<Calendar_Contracts.CalendarEvent[]> {
        const capacityPromises: PromiseLike<Work_Contracts.TeamMemberCapacity[]>[] = [];
        const iterationTeamDaysOffPromises: PromiseLike<Work_Contracts.TeamSettingsDaysOff>[] = [];
        const eventMap: { [dateString: string]: Calendar_Contracts.CalendarEvent } = {};
        this._events = null;
        this._renderedEvents = null;
        const events: Calendar_Contracts.CalendarEvent[] = [];
        const renderedEvents: Calendar_Contracts.CalendarEvent[] = [];

        const webContext = VSS.getWebContext();
        const teamContext: TFS_Core_Contracts.TeamContext = {
            projectId: webContext.project.id,
            teamId: this._team.id,
            project: "",
            team: "",
        };
        const workClient: Work_Client.WorkHttpClient5 = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient5, WebApi_Constants.ServiceInstanceTypes.TFS);

        return realPromise(this.getIterations()).then(iterations => {
            if (!iterations || iterations.length === 0) {
                this._events = events;
                this._renderedEvents = renderedEvents;
                return renderedEvents;
            }

            for (const iteration of iterations) {
                iterationTeamDaysOffPromises.push(workClient.getTeamDaysOff(teamContext, iteration.id));
                iterationTeamDaysOffPromises[
                    iterationTeamDaysOffPromises.length - 1
                ].then((teamDaysOff: Work_Contracts.TeamSettingsDaysOff) => {
                    if (teamDaysOff && teamDaysOff.daysOff && teamDaysOff.daysOff.length) {
                        for (const daysOffRange of teamDaysOff.daysOff) {
                            const event: any = {};
                            event.startDate = new Date(daysOffRange.start.valueOf()).toISOString();
                            event.endDate = new Date(daysOffRange.end.valueOf()).toISOString();
                            event.title = "Team Day Off";
                            event.member = {
                                displayName: this._team.name,
                                id: this._team.id,
                                imageUrl: this._buildTeamImageUrl(webContext.host.uri, this._team.id),
                            };
                            event.category = <Calendar_Contracts.IEventCategory>{
                                id: this.id + "." + "Everyone",
                                title: IdentityHelper.parseUniquefiedIdentityName(event.member.displayName),
                                imageUrl: this._buildTeamImageUrl(webContext.host.uri, teamContext.teamId),
                                color: this._categoryColor,
                            };
                            event.id = this._buildCapacityEventId(event);
                            event.iterationId = iteration.id;
                            event.icons = [
                                {
                                    src: event.category.imageUrl,
                                    title: event.title,
                                    linkedEvent: event,
                                },
                            ];

                            events.push(event);

                            // add personal day off event to calendar day off events
                            const dates = Calendar_DateUtils.getDatesInRange(daysOffRange.start, daysOffRange.end);
                            for (const dateObj of dates) {
                                const date = dateObj.toISOString();
                                if (!eventMap[date]) {
                                    const regroupedEvent: Calendar_Contracts.CalendarEvent = {
                                        startDate: date,
                                        endDate: date,
                                        member: event.member,
                                        title: "",
                                        id: this.id + "." + date,
                                        category: <Calendar_Contracts.IEventCategory>{
                                            id: "",
                                            title: "Grouped Event",
                                            color: this._categoryColor,
                                        },
                                        icons: [],
                                    };
                                    eventMap[date] = regroupedEvent;
                                    renderedEvents.push(regroupedEvent);
                                }
                                eventMap[date].icons.push(event.icons[0]);
                            }
                        }
                    }
                    return renderedEvents;
                });

                capacityPromises.push(workClient.getCapacities(teamContext, iteration.id));
                capacityPromises[capacityPromises.length - 1].then((capacities: Work_Contracts.TeamMemberCapacity[]) => {
                    if (capacities && capacities.length) {
                        for (const capacity of capacities) {
                            for (const daysOffRange of capacity.daysOff) {
                                const event: any = {};
                                event.startDate = new Date(daysOffRange.start.valueOf()).toISOString();
                                event.endDate = new Date(daysOffRange.end.valueOf()).toISOString();
                                event.title =
                                    IdentityHelper.parseUniquefiedIdentityName(capacity.teamMember.displayName) + " Day Off";
                                event.member = capacity.teamMember;
                                event.category = <Calendar_Contracts.IEventCategory>{
                                    id: this.id + "." + capacity.teamMember.uniqueName,
                                    title: IdentityHelper.parseUniquefiedIdentityName(event.member.displayName),
                                    imageUrl: event.member.imageUrl,
                                    color: this._categoryColor,
                                };
                                event.id = this._buildCapacityEventId(event);
                                event.iterationId = iteration.id;
                                event.icons = [
                                    {
                                        src: event.category.imageUrl,
                                        title: event.title,
                                        linkedEvent: event,
                                    },
                                ];
                                events.push(event);

                                // add personal day off event to calendar day off events
                                const dates = Calendar_DateUtils.getDatesInRange(daysOffRange.start, daysOffRange.end);
                                for (const dateObj of dates) {
                                    const date = dateObj.toISOString();
                                    if (!eventMap[date]) {
                                        const regroupedEvent: Calendar_Contracts.CalendarEvent = {
                                            startDate: date,
                                            endDate: date,
                                            member: event.member,
                                            title: "",
                                            id: this.id + "." + date,
                                            category: <Calendar_Contracts.IEventCategory>{
                                                id: "",
                                                title: "Grouped Event",
                                                color: this._categoryColor,
                                            },
                                            icons: [],
                                        };
                                        eventMap[date] = regroupedEvent;
                                        renderedEvents.push(regroupedEvent);
                                    }
                                    eventMap[date].icons.push(event.icons[0]);
                                }
                            }
                        }
                    }

                    return renderedEvents;
                });
            }

            return Promise.all(iterationTeamDaysOffPromises).then(() => {
                return Promise.all(capacityPromises).then(() => {
                    this._events = events;
                    this._renderedEvents = renderedEvents;
                    return renderedEvents;
                });
            });
        });
    }

    public getIterations(): PromiseLike<Work_Contracts.TeamSettingsIteration[]> {
        const webContext = VSS.getWebContext();
        const teamContext: TFS_Core_Contracts.TeamContext = {
            projectId: webContext.project.id,
            teamId: this._team.id,
            project: "",
            team: "",
        };
        const workClient: Work_Client.WorkHttpClient5 = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient5, WebApi_Constants.ServiceInstanceTypes.TFS);

        return workClient.getTeamIterations(teamContext);
    }

    public getCategories(query: Calendar_Contracts.IEventQuery): PromiseLike<Calendar_Contracts.IEventCategory[]> {
        if (this._events) {
            return Promise.resolve(this._getCategoryData(this._events.slice(0), query));
        } else {
            return this.getEvents().then((events: Calendar_Contracts.CalendarEvent[]) => {
                return this._getCategoryData(this._events, query);
            });
        }
    }

    public addEvent(event: Calendar_Contracts.CalendarEvent): PromiseLike<Calendar_Contracts.CalendarEvent> {
        const dayOffStart = new Date(event.startDate);
        const dayOffEnd = new Date(event.endDate);
        const isTeam: boolean = event.member.displayName === "Everyone";
        const memberId: string = event.member.id;
        const iterationId: string = event.iterationId;
        const webContext = VSS.getWebContext();
        const teamContext: TFS_Core_Contracts.TeamContext = {
            projectId: webContext.project.id,
            teamId: this._team.id,
            project: "",
            team: "",
        };
        const workClient: Work_Client.WorkHttpClient5 = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient5, WebApi_Constants.ServiceInstanceTypes.TFS);

        if (isTeam) {
            return realPromise(
                this._getTeamDaysOff(workClient, teamContext, iterationId),
            ).then((teamDaysOff: Work_Contracts.TeamSettingsDaysOff) => {
                const teamDaysOffPatch: Work_Contracts.TeamSettingsDaysOffPatch = { daysOff: teamDaysOff.daysOff };
                teamDaysOffPatch.daysOff.push({ start: dayOffStart, end: dayOffEnd });
                return workClient
                    .updateTeamDaysOff(teamDaysOffPatch, teamContext, iterationId)
                    .then((value: Work_Contracts.TeamSettingsDaysOff) => {
                        // Resolve null to tell views.js to reload the entire event source instead of re-rendering the updated event
                        return null;
                    });
            });
        } else {
            return realPromise(
                this._getCapacity(workClient, teamContext, iterationId, memberId),
            ).then((capacity: Work_Contracts.TeamMemberCapacity) => {
                const capacityPatch: Work_Contracts.CapacityPatch = {
                    activities: capacity.activities,
                    daysOff: capacity.daysOff,
                };
                capacityPatch.daysOff.push({ start: dayOffStart, end: dayOffEnd });
                return workClient
                    .updateCapacity(capacityPatch, teamContext, iterationId, memberId)
                    .then((value: Work_Contracts.TeamMemberCapacity) => {
                        // Resolve null to tell views.js to reload the entire event source instead of re-rendering the updated event
                        return null;
                    });
            });
        }
    }

    public removeEvent(event: Calendar_Contracts.CalendarEvent): PromiseLike<Calendar_Contracts.CalendarEvent[]> {
        const dayOffStart = new Date(event.startDate);
        const memberId = event.member.id;
        const isTeam: boolean = event.member.uniqueName === undefined;
        const iterationId: string = event.iterationId;
        const webContext = VSS.getWebContext();
        const teamContext: TFS_Core_Contracts.TeamContext = {
            projectId: webContext.project.id,
            teamId: this._team.id,
            project: "",
            team: "",
        };
        const workClient: Work_Client.WorkHttpClient5 = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient5, WebApi_Constants.ServiceInstanceTypes.TFS);

        if (isTeam) {
            return realPromise(
                this._getTeamDaysOff(workClient, teamContext, iterationId),
            ).then((teamDaysOff: Work_Contracts.TeamSettingsDaysOff) => {
                const teamDaysOffPatch: Work_Contracts.TeamSettingsDaysOffPatch = { daysOff: teamDaysOff.daysOff };
                teamDaysOffPatch.daysOff.some(
                    (dateRange: Work_Contracts.DateRange, index: number, array: Work_Contracts.DateRange[]) => {
                        if (dateRange.start.valueOf() === dayOffStart.valueOf()) {
                            teamDaysOffPatch.daysOff.splice(index, 1);
                            return true;
                        }
                        return false;
                    },
                );
                return workClient
                    .updateTeamDaysOff(teamDaysOffPatch, teamContext, iterationId)
                    .then((value: Work_Contracts.TeamSettingsDaysOff) => {
                        // Resolve null to tell views.js to reload the entire event source instead removing one event
                        return null;
                    });
            });
        } else {
            return realPromise(
                this._getCapacity(workClient, teamContext, iterationId, memberId),
            ).then((capacity: Work_Contracts.TeamMemberCapacity) => {
                const capacityPatch: Work_Contracts.CapacityPatch = {
                    activities: capacity.activities,
                    daysOff: capacity.daysOff,
                };
                capacityPatch.daysOff.some(
                    (dateRange: Work_Contracts.DateRange, index: number, array: Work_Contracts.DateRange[]) => {
                        if (dateRange.start.valueOf() === dayOffStart.valueOf()) {
                            capacityPatch.daysOff.splice(index, 1);
                            return true;
                        }
                        return false;
                    },
                );
                return workClient
                    .updateCapacity(capacityPatch, teamContext, iterationId, memberId)
                    .then((value: Work_Contracts.TeamMemberCapacity) => {
                        // Resolve null to tell views.js to reload the entire event source instead removing one event
                        return null;
                    });
            });
        }
    }

    public updateEvent(
        oldEvent: Calendar_Contracts.CalendarEvent,
        newEvent: Calendar_Contracts.CalendarEvent,
    ): PromiseLike<Calendar_Contracts.CalendarEvent> {
        const dayOffStart = new Date(oldEvent.startDate);
        const memberId = oldEvent.member.id;
        const iterationId = oldEvent.iterationId;
        const isTeam: boolean = oldEvent.member.uniqueName === undefined;
        const webContext = VSS.getWebContext();
        const teamContext: TFS_Core_Contracts.TeamContext = {
            projectId: webContext.project.id,
            teamId: this._team.id,
            project: "",
            team: "",
        };
        const workClient: Work_Client.WorkHttpClient5 = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient5, WebApi_Constants.ServiceInstanceTypes.TFS);

        if (isTeam) {
            return realPromise(
                this._getTeamDaysOff(workClient, teamContext, iterationId),
            ).then((teamDaysOff: Work_Contracts.TeamSettingsDaysOff) => {
                const teamDaysOffPatch: Work_Contracts.TeamSettingsDaysOffPatch = { daysOff: teamDaysOff.daysOff };
                const updated: boolean = teamDaysOffPatch.daysOff.some(
                    (dateRange: Work_Contracts.DateRange, index: number, array: Work_Contracts.DateRange[]) => {
                        if (dateRange.start.valueOf() === dayOffStart.valueOf()) {
                            teamDaysOffPatch.daysOff[index].start = new Date(newEvent.startDate);
                            teamDaysOffPatch.daysOff[index].end = new Date(newEvent.endDate);
                            return true;
                        }
                        return false;
                    },
                );
                return workClient
                    .updateTeamDaysOff(teamDaysOffPatch, teamContext, iterationId)
                    .then((value: Work_Contracts.TeamSettingsDaysOff) => {
                        return null;
                    });
            });
        } else {
            return realPromise(
                this._getCapacity(workClient, teamContext, iterationId, memberId),
            ).then((capacity: Work_Contracts.TeamMemberCapacity) => {
                const capacityPatch: Work_Contracts.CapacityPatch = {
                    activities: capacity.activities,
                    daysOff: capacity.daysOff,
                };
                capacityPatch.daysOff.some(
                    (dateRange: Work_Contracts.DateRange, index: number, array: Work_Contracts.DateRange[]) => {
                        if (dateRange.start.valueOf() === dayOffStart.valueOf()) {
                            capacityPatch.daysOff[index].start = new Date(newEvent.startDate);
                            capacityPatch.daysOff[index].end = new Date(newEvent.endDate);
                            return true;
                        }
                        return false;
                    },
                );
                return workClient
                    .updateCapacity(capacityPatch, teamContext, iterationId, memberId)
                    .then((value: Work_Contracts.TeamMemberCapacity) => {
                        return null;
                    });
            });
        }
    }

    public getTitleUrl(webContext: WebContext): PromiseLike<string> {
        const workClient: Work_Client.WorkHttpClient5 = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient5, WebApi_Constants.ServiceInstanceTypes.TFS);
        const teamContext: TFS_Core_Contracts.TeamContext = {
            projectId: webContext.project.id,
            teamId: this._team.id,
            project: "",
            team: "",
        };
        return realPromise(workClient.getTeamIterations(teamContext, "current")).then(
            (iterations: Work_Contracts.TeamSettingsIteration[]) => {
                if (iterations.length > 0) {
                    const iterationPath = iterations[0].path.substr(iterations[0].path.indexOf("\\") + 1);
                    return (
                        webContext.host.uri +
                        webContext.project.name +
                        "/" +
                        this._team.name +
                        "/_backlogs/capacity/" +
                        iterationPath
                    );
                } else {
                    return webContext.host.uri + webContext.project.name + "/" + this._team.name + "/_admin/_iterations";
                }
            },
            error => {
                return webContext.host.uri + webContext.project.name + "/" + this._team.name + "/_admin/_iterations";
            },
        );
    }

    private _getTeamDaysOff(
        workClient: Work_Client.WorkHttpClient5,
        teamContext: TFS_Core_Contracts.TeamContext,
        iterationId,
    ): PromiseLike<Work_Contracts.TeamSettingsDaysOff> {
        return workClient.getTeamDaysOff(teamContext, iterationId).then((value: Work_Contracts.TeamSettingsDaysOff) => {
            return value;
        });
    }

    private _getCapacity(
        workClient: Work_Client.WorkHttpClient5,
        teamContext: TFS_Core_Contracts.TeamContext,
        iterationId,
        memberId: string,
    ): PromiseLike<Work_Contracts.TeamMemberCapacity> {
        return workClient.getCapacities(teamContext, iterationId).then((capacities: Work_Contracts.TeamMemberCapacity[]) => {
            const foundCapacities = capacities.filter(value => value.teamMember.id === memberId);
            if (foundCapacities.length > 0) {
                return foundCapacities[0];
            }

            const value = {
                activities: [
                    {
                        capacityPerDay: 0,
                        name: null,
                    },
                ],
                daysOff: [],
            } as Work_Contracts.TeamMemberCapacity;
            return value;
        });
    }

    private _getCategoryData(
        events: Calendar_Contracts.CalendarEvent[],
        query: Calendar_Contracts.IEventQuery,
    ): Calendar_Contracts.IEventCategory[] {
        const memberMap: { [id: string]: Calendar_Contracts.IEventCategory } = {};
        const categories: Calendar_Contracts.IEventCategory[] = [];
        for (const event of events) {
            if (Calendar_DateUtils.eventIn(event, query)) {
                const member = <Work_Contracts.Member>(<any>event).member;
                if (!memberMap[member.id]) {
                    event.category.events = [event.id];
                    event.category.subTitle = this._getCategorySubTitle(event.category, events, query);
                    memberMap[member.id] = event.category;
                    categories.push(event.category);
                } else {
                    const category = memberMap[member.id];
                    category.events.push(event.id);
                    category.subTitle = this._getCategorySubTitle(category, events, query);
                }
            }
        }

        return categories;
    }

    private _getCategorySubTitle(
        category: Calendar_Contracts.IEventCategory,
        events: Calendar_Contracts.CalendarEvent[],
        query: Calendar_Contracts.IEventQuery,
    ): string {
        // add up days off per person
        const daysOffInRange: Date[] = [];
        const queryStartInUtc = new Date(
            query.startDate.getUTCFullYear(),
            query.startDate.getUTCMonth(),
            query.startDate.getUTCDate(),
            query.startDate.getUTCHours(),
            query.startDate.getUTCMinutes(),
            query.startDate.getUTCSeconds(),
        );
        const queryEndInUtc = new Date(
            query.endDate.getUTCFullYear(),
            query.endDate.getUTCMonth(),
            query.endDate.getUTCDate(),
            query.endDate.getUTCHours(),
            query.endDate.getUTCMinutes(),
            query.endDate.getUTCSeconds(),
        );
        category.events.forEach(e => {
            const event = events.filter(event => event.id === e)[0];
            const datesInRange = Calendar_DateUtils.getDatesInRange(
                Utils_Date.shiftToUTC(new Date(event.startDate)),
                Utils_Date.shiftToUTC(new Date(event.endDate)),
            );
            datesInRange.forEach((dateToCheck: Date, index: number, array: Date[]) => {
                const dateToCheckInUtc = new Date(
                    dateToCheck.getUTCFullYear(),
                    dateToCheck.getUTCMonth(),
                    dateToCheck.getUTCDate(),
                    dateToCheck.getUTCHours(),
                    dateToCheck.getUTCMinutes(),
                    dateToCheck.getUTCSeconds(),
                );
                if (Calendar_DateUtils.isBetween(dateToCheckInUtc, queryStartInUtc, queryEndInUtc)) {
                    daysOffInRange.push(dateToCheck);
                }
            });
        });

        // if user has only one day off, return that date
        if (daysOffInRange.length === 1) {
            return Utils_Date.localeFormat(daysOffInRange[0], Culture.getDateTimeFormat().ShortDatePattern, true);
        }
        // else return total number of days off
        return Utils_String.format("{0} days off", daysOffInRange.length);
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
        if (!name) {
            return null;
        }

        let i = name.lastIndexOf(IdentityHelper.AAD_IDENTITY_UNIQUEFIEDNAME_SEPERATOR_START);
        let j = name.lastIndexOf(IdentityHelper.AAD_IDENTITY_UNIQUEFIEDNAME_SEPERATOR_END);
        let isContainer: boolean = false;
        let isAad: boolean = false;
        if (i >= 0 && j > i) {
            isAad = true;
        }

        // replace "<<" with "<" and ">>" with ">" in case of an AAD identity string representation to make further processing easier
        name = name
            .replace(
                IdentityHelper.AAD_IDENTITY_UNIQUEFIEDNAME_SEPERATOR_START,
                IdentityHelper.IDENTITY_UNIQUEFIEDNAME_SEPERATOR_START,
            )
            .replace(
                IdentityHelper.AAD_IDENTITY_UNIQUEFIEDNAME_SEPERATOR_END,
                IdentityHelper.IDENTITY_UNIQUEFIEDNAME_SEPERATOR_END,
            );

        i = name.lastIndexOf(IdentityHelper.IDENTITY_UNIQUEFIEDNAME_SEPERATOR_START);
        j = name.lastIndexOf(IdentityHelper.IDENTITY_UNIQUEFIEDNAME_SEPERATOR_END);
        let displayName = name;
        let alias = "";
        let id = "";
        let localScopeId = "";
        if (i >= 0 && j > i) {
            displayName = name.substr(0, i).trim();
            if (isAad) {
                // if its an AAD identity, the string would be in format - name <<object id>>
                id = name.substr(i + 1, j - i - 1).trim(); // this would be in format objectid\email

                if (id.indexOf(IdentityHelper.AAD_IDENTITY_USER_PREFIX) === 0) {
                    id = id.substr(IdentityHelper.AAD_IDENTITY_USER_PREFIX.length);
                } else if (id.indexOf(IdentityHelper.AAD_IDENTITY_GROUP_PREFIX) === 0) {
                    isContainer = true;
                    id = id.substr(IdentityHelper.AAD_IDENTITY_GROUP_PREFIX.length);
                }

                const ii = id.lastIndexOf("\\");
                if (ii > 0) {
                    alias = id.substr(ii + 1).trim();
                    id = id.substr(0, ii).trim();
                }
            } else {
                alias = name.substr(i + 1, j - i - 1).trim();
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
