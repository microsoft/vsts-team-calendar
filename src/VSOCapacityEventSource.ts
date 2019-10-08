import { getClient } from "azure-devops-extension-api";
import { TeamContext } from "azure-devops-extension-api/Core";
import {
    CapacityPatch,
    TeamMemberCapacity,
    TeamSettingsDaysOff,
    TeamSettingsDaysOffPatch,
    TeamSettingsIteration,
    WorkRestClient,
    TeamMemberCapacityIdentityRef
} from "azure-devops-extension-api/work";

import { ObservableValue, ObservableArray } from "azure-devops-ui/Core/Observable";

import { EventInput } from "@fullcalendar/core";
import { EventSourceError } from "@fullcalendar/core/structs/event-source";

import { generateColor } from "./Color";
import { ICalendarEvent, IEventIcon, IEventCategory } from "./Contracts";
import { formatDate, getDatesInRange, shiftToLocal, shiftToUTC, toDate } from "./TimeLib";

export const DaysOffId = "daysOff";
export const Everyone = "Everyone";
export const IterationId = "iteration";

export class VSOCapacityEventSource {
    private capacityMap: { [iterationId: string]: { [memberId: string]: TeamMemberCapacityIdentityRef } } = {};
    private capacitySummaryData: ObservableArray<IEventCategory> = new ObservableArray<IEventCategory>([]);
    private capacityUrl: ObservableValue<string> = new ObservableValue("");
    private groupedEventMap: { [dateString: string]: ICalendarEvent } = {};
    private hostUrl: string = "";
    private iterations: TeamSettingsIteration[] = [];
    private iterationSummaryData: ObservableArray<IEventCategory> = new ObservableArray<IEventCategory>([]);
    private iterationUrl: ObservableValue<string> = new ObservableValue("");
    private teamContext: TeamContext = { projectId: "", teamId: "", project: "", team: "" };
    private teamDayOffMap: { [iterationId: string]: TeamSettingsDaysOff } = {};
    private workClient: WorkRestClient = getClient(WorkRestClient, {});

    /**
     * Add new day off for a member or a team
     */
    public addEvent = (iterationId: string, startDate: Date, endDate: Date, memberName: string, memberId: string) => {
        const isTeam = memberName === Everyone;
        startDate = shiftToLocal(startDate);
        endDate = shiftToLocal(endDate);
        if (isTeam) {
            const teamDaysOff = this.teamDayOffMap[iterationId];
            // delete from cached copy
            delete this.teamDayOffMap[iterationId];
            const teamDaysOffPatch: TeamSettingsDaysOffPatch = { daysOff: teamDaysOff.daysOff };
            teamDaysOffPatch.daysOff.push({ start: startDate, end: endDate });
            return this.workClient.updateTeamDaysOff(teamDaysOffPatch, this.teamContext, iterationId);
        } else {
            const capacity = this.capacityMap[iterationId][memberId];
            delete this.capacityMap[iterationId];
            const capacityPatch: CapacityPatch = { activities: capacity.activities, daysOff: capacity.daysOff };
            capacityPatch.daysOff.push({ start: startDate, end: endDate });
            return this.workClient.updateCapacityWithIdentityRef(capacityPatch, this.teamContext, iterationId, memberId);
        }
    };

    public deleteEvent = (event: ICalendarEvent, iterationId: string) => {
        const isTeam = event.member!.displayName === Everyone;
        const startDate = shiftToLocal(new Date(event.startDate));
        if (isTeam) {
            const teamDaysOff = this.teamDayOffMap[iterationId];
            delete this.teamDayOffMap[iterationId];
            var i;
            for (i = 0; i < teamDaysOff.daysOff.length; i++) {
                if (teamDaysOff.daysOff[i].start.valueOf() === startDate.valueOf()) {
                    teamDaysOff.daysOff.splice(i, 1);
                    break;
                }
            }
            const teamDaysOffPatch: TeamSettingsDaysOffPatch = { daysOff: teamDaysOff.daysOff };
            return this.workClient.updateTeamDaysOff(teamDaysOffPatch, this.teamContext, iterationId);
        } else {
            const capacity = this.capacityMap[iterationId][event.member!.id];
            delete this.capacityMap[iterationId];
            var i;
            for (i = 0; i < capacity.daysOff.length; i++) {
                if (capacity.daysOff[i].start.valueOf() === startDate.valueOf()) {
                    capacity.daysOff.splice(i, 1);
                    break;
                }
            }
            const capacityPatch: CapacityPatch = { activities: capacity.activities, daysOff: capacity.daysOff };
            return this.workClient.updateCapacityWithIdentityRef(capacityPatch, this.teamContext, iterationId, event.member!.id);
        }
    };

    public getCapacitySummaryData = (): ObservableArray<IEventCategory> => {
        return this.capacitySummaryData;
    };

    public getCapacityUrl = (): ObservableValue<string> => {
        return this.capacityUrl;
    };

    public getEvents = (
        arg: {
            start: Date;
            end: Date;
            timeZone: string;
        },
        successCallback: (events: EventInput[]) => void,
        failureCallback: (error: EventSourceError) => void
    ): void | PromiseLike<EventInput[]> => {
        const capacityPromises: PromiseLike<TeamMemberCapacity[]>[] = [];
        const teamDaysOffPromises: PromiseLike<TeamSettingsDaysOff>[] = [];
        const renderedEvents: EventInput[] = [];
        const capacityCatagoryMap: { [id: string]: IEventCategory } = {};
        const currentIterations: IEventCategory[] = [];

        this.groupedEventMap = {};

        this.fetchIterations().then(iterations => {
            if (!iterations) {
                iterations = [];
            }
            this.iterations = iterations;

            // convert end date to inclusive end date
            const calendarStart = arg.start;
            const calendarEnd = new Date(arg.end);
            calendarEnd.setDate(arg.end.getDate() - 1);

            for (const iteration of iterations) {
                const start = shiftToUTC(iteration.attributes.startDate);
                const end = shiftToUTC(iteration.attributes.finishDate);

                if (
                    (calendarStart <= start && start <= calendarEnd) ||
                    (calendarStart <= end && end <= calendarEnd) ||
                    (start <= calendarStart && end >= calendarEnd)
                ) {
                    const now = new Date();
                    let color = generateColor("otherIteration");
                    if (iteration.attributes.startDate <= now && now <= iteration.attributes.finishDate) {
                        color = generateColor("currentIteration");
                    }
                    const exclusiveEndDate = new Date(end);
                    exclusiveEndDate.setDate(end.getDate() + 1);

                    renderedEvents.push({
                        id: IterationId + iteration.name,
                        allDay: true,
                        start: start,
                        end: exclusiveEndDate,
                        title: iteration.name,
                        textColor: "#FFFFFF",
                        backgroundColor: color,
                        rendering: "background"
                    });

                    currentIterations.push({
                        color: color,
                        subTitle: formatDate(start, "MONTH-DD") + " - " + formatDate(end, "MONTH-DD"),
                        title: iteration.name,
                        eventCount: 1
                    });

                    const teamsDayOffPromise = this.fetchTeamDaysOff(iteration.id);
                    teamDaysOffPromises.push(teamsDayOffPromise);
                    teamsDayOffPromise.then((teamDaysOff: TeamSettingsDaysOff) => {
                        this.processTeamDaysOff(teamDaysOff, iteration.id, capacityCatagoryMap, calendarStart, calendarEnd);
                    });

                    const capacityPromise = this.fetchCapacities(iteration.id);
                    capacityPromises.push(capacityPromise);
                    capacityPromise.then((capacities: TeamMemberCapacityIdentityRef[]) => {
                        this.processCapacity(capacities, iteration.id, capacityCatagoryMap, calendarStart, calendarEnd);
                    });
                }
            }

            Promise.all(teamDaysOffPromises).then(() => {
                Promise.all(capacityPromises).then(() => {
                    Object.keys(this.groupedEventMap).forEach(id => {
                        const event = this.groupedEventMap[id];
                        // skip events with date strings we can't parse.
                        const start = new Date(event.startDate);
                        const end = new Date(event.endDate);
                        if ((calendarStart <= start && start <= calendarEnd) || (calendarStart <= end && end <= calendarEnd)) {
                            renderedEvents.push({
                                allDay: true,
                                color: "transparent",
                                editable: false,
                                end: end,
                                id: event.id,
                                start: start,
                                title: ""
                            });
                        }
                    });
                    successCallback(renderedEvents);
                    this.iterationSummaryData.splice(0, this.iterationSummaryData.length, ...currentIterations);
                    this.capacitySummaryData.splice(
                        0,
                        this.capacitySummaryData.length,
                        ...Object.keys(capacityCatagoryMap).map(key => {
                            const catagory = capacityCatagoryMap[key];
                            if (catagory.eventCount > 1) {
                                catagory.subTitle = catagory.eventCount + " days off";
                            }
                            return catagory;
                        })
                    );
                });
            });
        });
    };

    public getGroupedEventForDate = (date: Date): ICalendarEvent => {
        const dateString = date.toISOString();
        return this.groupedEventMap[dateString];
    };

    public getIterationForDate = (startDate: Date, endDate: Date): TeamSettingsIteration | undefined => {
        let iteration = undefined;
        startDate = shiftToLocal(startDate);
        endDate = shiftToLocal(endDate);
        this.iterations.forEach(item => {
            if (
                item.attributes.startDate <= startDate &&
                startDate <= item.attributes.finishDate &&
                item.attributes.startDate <= endDate &&
                endDate <= item.attributes.finishDate
            ) {
                iteration = item;
            }
        });

        return iteration;
    };

    public getIterationSummaryData = (): ObservableArray<IEventCategory> => {
        return this.iterationSummaryData;
    };

    public getIterationUrl = (): ObservableValue<string> => {
        return this.iterationUrl;
    };

    public initialize(projectId: string, projectName: string, teamId: string, teamName: string, hostUrl: string) {
        this.hostUrl = hostUrl;
        this.teamContext = {
            projectId: projectId,
            teamId: teamId,
            project: projectName,
            team: teamName
        };
        this.teamDayOffMap = {};
        this.capacityMap = {};
        this.iterations = [];
        this.updateUrls();
    }

    public updateEvent = (oldEvent: ICalendarEvent, iterationId: string, startDate: Date, endDate: Date) => {
        const isTeam = oldEvent.member!.displayName === Everyone;
        const orignalStartDate = shiftToLocal(new Date(oldEvent.startDate));
        startDate = shiftToLocal(startDate);
        endDate = shiftToLocal(endDate);
        if (isTeam) {
            const teamDaysOff = this.teamDayOffMap[iterationId];
            delete this.teamDayOffMap[iterationId];
            var i;
            for (i = 0; i < teamDaysOff.daysOff.length; i++) {
                if (teamDaysOff.daysOff[i].start.valueOf() === orignalStartDate.valueOf()) {
                    teamDaysOff.daysOff[i].start = startDate;
                    teamDaysOff.daysOff[i].end = endDate;
                    break;
                }
            }
            const teamDaysOffPatch: TeamSettingsDaysOffPatch = { daysOff: teamDaysOff.daysOff };
            return this.workClient.updateTeamDaysOff(teamDaysOffPatch, this.teamContext, iterationId);
        } else {
            const capacity = this.capacityMap[iterationId][oldEvent.member!.id];
            delete this.capacityMap[iterationId];
            var i;
            for (i = 0; i < capacity.daysOff.length; i++) {
                if (capacity.daysOff[i].start.valueOf() === orignalStartDate.valueOf()) {
                    capacity.daysOff[i].start = startDate;
                    capacity.daysOff[i].end = endDate;
                    break;
                }
            }
            const capacityPatch: CapacityPatch = { activities: capacity.activities, daysOff: capacity.daysOff };
            return this.workClient.updateCapacityWithIdentityRef(capacityPatch, this.teamContext, iterationId, oldEvent.member!.id);
        }
    };

    private buildTeamImageUrl(id: string): string {
        return this.hostUrl + "_api/_common/IdentityImage?id=" + id;
    }

    private fetchCapacities = (iterationId: string): Promise<TeamMemberCapacityIdentityRef[]> => {
        if (this.capacityMap[iterationId]) {
            const capacities = [];
            for (var key in this.capacityMap[iterationId]) {
                capacities.push(this.capacityMap[iterationId][key]);
            }
            return Promise.resolve(capacities);
        }
        return this.workClient.getCapacitiesWithIdentityRef(this.teamContext, iterationId);
    };

    private fetchIterations = (): Promise<TeamSettingsIteration[]> => {
        if (this.iterations.length > 0) {
            return Promise.resolve(this.iterations);
        }
        return this.workClient.getTeamIterations(this.teamContext);
    };

    private fetchTeamDaysOff = (iterationId: string): Promise<TeamSettingsDaysOff> => {
        if (this.teamDayOffMap[iterationId]) {
            return Promise.resolve(this.teamDayOffMap[iterationId]);
        }
        return this.workClient.getTeamDaysOff(this.teamContext, iterationId);
    };

    private processCapacity = (
        capacities: TeamMemberCapacityIdentityRef[],
        iterationId: string,
        capacityCatagoryMap: { [id: string]: IEventCategory },
        calendarStart: Date,
        calendarEnd: Date
    ) => {
        if (capacities && capacities.length) {
            for (const capacity of capacities) {
                if (this.capacityMap[iterationId]) {
                    this.capacityMap[iterationId][capacity.teamMember.id] = capacity;
                } else {
                    const temp: { [memberId: string]: TeamMemberCapacityIdentityRef } = {};
                    temp[capacity.teamMember.id] = capacity;
                    this.capacityMap[iterationId] = temp;
                }

                for (const daysOffRange of capacity.daysOff) {
                    const start = shiftToUTC(daysOffRange.start);
                    const end = shiftToUTC(daysOffRange.end);
                    const title = capacity.teamMember.displayName + " Day Off";

                    const event: ICalendarEvent = {
                        title: title,
                        startDate: start.toISOString(),
                        endDate: end.toISOString(),
                        member: capacity.teamMember,
                        category: title,
                        iterationId: iterationId
                    };

                    const icon: IEventIcon = {
                        src: capacity.teamMember.imageUrl,
                        linkedEvent: event
                    };

                    // add personal day off event to calendar day off events
                    const dates = getDatesInRange(start, end);
                    for (const dateObj of dates) {
                        if (calendarStart <= dateObj && dateObj <= calendarEnd) {
                            if (capacityCatagoryMap[capacity.teamMember.id]) {
                                capacityCatagoryMap[capacity.teamMember.id].eventCount++;
                            } else {
                                capacityCatagoryMap[capacity.teamMember.id] = {
                                    eventCount: 1,
                                    imageUrl: capacity.teamMember.imageUrl,
                                    subTitle: formatDate(dateObj, "MM-DD-YYYY"),
                                    title: capacity.teamMember.displayName
                                };
                            }

                            const date = dateObj.toISOString();
                            if (!this.groupedEventMap[date]) {
                                const regroupedEvent: ICalendarEvent = {
                                    startDate: date,
                                    endDate: date,
                                    member: event.member,
                                    title: "Grouped Event",
                                    id: DaysOffId + "." + date,
                                    category: "Grouped Event",
                                    icons: []
                                };
                                this.groupedEventMap[date] = regroupedEvent;
                            }
                            this.groupedEventMap[date].icons!.push(icon);
                        }
                    }
                }
            }
        }
    };

    private processTeamDaysOff = (
        teamDaysOff: TeamSettingsDaysOff,
        iterationId: string,
        capacityCatagoryMap: { [id: string]: IEventCategory },
        calendarStart: Date,
        calendarEnd: Date
    ) => {
        if (teamDaysOff && teamDaysOff.daysOff) {
            this.teamDayOffMap[iterationId] = teamDaysOff;
            for (const daysOffRange of teamDaysOff.daysOff) {
                const teamImage = this.buildTeamImageUrl(this.teamContext.teamId);
                const start = shiftToUTC(daysOffRange.start);
                const end = shiftToUTC(daysOffRange.end);

                const event: ICalendarEvent = {
                    title: "Team Day Off",
                    startDate: start.toISOString(),
                    endDate: end.toISOString(),
                    member: {
                        displayName: Everyone,
                        id: this.teamContext.teamId
                    },
                    category: this.teamContext.team,
                    iterationId: iterationId
                };

                const icon: IEventIcon = {
                    src: teamImage,
                    linkedEvent: event
                };

                // add personal day off event to calendar day off events
                const dates = getDatesInRange(start, end);
                for (const dateObj of dates) {
                    if (calendarStart <= dateObj && dateObj <= calendarEnd) {
                        if (capacityCatagoryMap[this.teamContext.team]) {
                            capacityCatagoryMap[this.teamContext.team].eventCount++;
                        } else {
                            capacityCatagoryMap[this.teamContext.team] = {
                                eventCount: 1,
                                imageUrl: teamImage,
                                subTitle: formatDate(dateObj, "MM-DD-YYYY"),
                                title: this.teamContext.team
                            };
                        }

                        const date = dateObj.toISOString();
                        if (!this.groupedEventMap[date]) {
                            const regroupedEvent: ICalendarEvent = {
                                startDate: date,
                                endDate: date,
                                member: event.member,
                                title: "Grouped Event",
                                id: DaysOffId + "." + date,
                                category: "Grouped Event",
                                icons: []
                            };
                            this.groupedEventMap[date] = regroupedEvent;
                        }
                        this.groupedEventMap[date].icons!.push(icon);
                    }
                }
            }
        }
    };

    private updateUrls = () => {
        this.iterationUrl.value = this.hostUrl + this.teamContext.project + "/" + this.teamContext.team + "/_admin/_iterations";

        this.workClient.getTeamIterations(this.teamContext, "current").then(
            iterations => {
                if (iterations.length > 0) {
                    const iterationPath = iterations[0].path.substr(iterations[0].path.indexOf("\\") + 1);
                    this.capacityUrl.value =
                        this.hostUrl + this.teamContext.project + "/" + this.teamContext.team + "/_backlogs/capacity/" + iterationPath;
                } else {
                    this.capacityUrl.value = this.hostUrl + this.teamContext.project + "/" + this.teamContext.team + "/_admin/_iterations";
                }
            },
            error => {
                this.capacityUrl.value = this.hostUrl + this.teamContext.project + "/" + this.teamContext.team + "/_admin/_iterations";
            }
        );
    };
}
