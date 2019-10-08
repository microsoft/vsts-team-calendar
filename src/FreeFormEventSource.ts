import { EventInput } from "@fullcalendar/core";
import { EventSourceError } from "@fullcalendar/core/structs/event-source";

import { IExtensionDataManager, ExtensionDataCollection } from "azure-devops-extension-api";

import { generateColor } from "./Color";
import { ICalendarEvent, IEventCategory } from "./Contracts";
import { shiftToUTC, shiftToLocal, getMonthYearInRange, toMonthYear } from "./TimeLib";
import { ObservableValue, ObservableArray, Observable } from "azure-devops-ui/Core/Observable";

export const FreeFormId = "FreeForm";

export class FreeFormEventsSource {
    categories: Set<string> = new Set<string>();
    dataManager?: IExtensionDataManager;
    eventMap: { [id: string]: ICalendarEvent } = {};
    fetchedCollections: Set<string> = new Set<string>();
    selectedTeamId: string = "";
    summaryData: ObservableArray<IEventCategory> = new ObservableArray<IEventCategory>([]);

    public addEvent = (title: string, startDate: Date, endDate: Date, category: string, description: string): PromiseLike<ICalendarEvent> => {
        const start = shiftToLocal(startDate);
        const end = shiftToLocal(endDate);

        const event: ICalendarEvent = {
            category: category,
            description: description,
            endDate: end.toISOString(),
            startDate: start.toISOString(),
            title: title
        };

        if (typeof event.category !== "string") {
            event.category = event.category.title;
        }
        if (event.category !== "Uncategorized") {
            this.categories.add(event.category);
        }

        return this.dataManager!.createDocument(this.selectedTeamId! + "." + toMonthYear(startDate), event).then((addedEvent: ICalendarEvent) => {
            // add event
            this.eventMap[addedEvent.id!] = addedEvent;
            const start = shiftToUTC(new Date(addedEvent.startDate));
            const end = shiftToUTC(new Date(addedEvent.endDate));

            addedEvent.startDate = start.toISOString();
            addedEvent.endDate = end.toISOString();

            return addedEvent;
        });
    };

    public deleteEvent = (eventId: string, startDate: Date) => {
        delete this.eventMap[eventId];
        return this.dataManager!.deleteDocument(this.selectedTeamId! + "." + toMonthYear(startDate), eventId);
    };

    public getCategories = (): Set<string> => {
        return this.categories;
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
        // convert end date to inclusive end date
        const calendarStart = arg.start;
        const calendarEnd = new Date(arg.end);
        calendarEnd.setDate(arg.end.getDate() - 1);

        this.fetchEvents(calendarStart, calendarEnd).then(() => {
            const inputs: EventInput[] = [];
            const catagoryMap: { [id: string]: IEventCategory } = {};
            Object.keys(this.eventMap).forEach(id => {
                const event = this.eventMap[id];
                // skip events with date strings we can't parse.
                if (Date.parse(event.startDate) && event.endDate && Date.parse(event.endDate)) {
                    if (event.category && typeof event.category !== "string") {
                        event.category = event.category.title;
                    }
                    if (event.category !== "Uncategorized") {
                        this.categories.add(event.category);
                    }

                    const start = shiftToUTC(new Date(event.startDate));
                    const end = shiftToUTC(new Date(event.endDate));

                    // check if event should be shown
                    if ((calendarStart <= start && start <= calendarEnd) || (calendarStart <= end && end <= calendarEnd)) {
                        const excludedEndDate = new Date(end);
                        excludedEndDate.setDate(end.getDate() + 1);

                        const eventColor = generateColor(event.category);

                        inputs.push({
                            id: FreeFormId + "." + event.id,
                            allDay: true,
                            editable: true,
                            start: start,
                            end: excludedEndDate,
                            title: event.title,
                            color: eventColor,
                            extendedProps: {
                                category: event.category,
                                description: event.description,
                                id: event.id
                            }
                        });

                        if (catagoryMap[event.category]) {
                            catagoryMap[event.category].eventCount++;
                        } else {
                            catagoryMap[event.category] = {
                                color: eventColor,
                                eventCount: 1,
                                subTitle: event.title,
                                title: event.category
                            };
                        }
                    }
                }
            });
            successCallback(inputs);
            this.summaryData.splice(
                0,
                this.summaryData.length,
                ...Object.keys(catagoryMap).map(key => {
                    const catagory = catagoryMap[key];
                    if (catagory.eventCount > 1) {
                        catagory.subTitle = catagory.eventCount + " events";
                    }
                    return catagory;
                })
            );
        });
    };

    getSummaryData = (): ObservableArray<IEventCategory> => {
        return this.summaryData;
    };

    public initialize(teamId: string, manager: IExtensionDataManager) {
        this.selectedTeamId = teamId;
        this.dataManager = manager;
        this.eventMap = {};
        this.categories.clear();
        this.fetchedCollections.clear();
    }

    public updateEvent = (
        id: string,
        title: string,
        startDate: Date,
        endDate: Date,
        category: string,
        description: string
    ): PromiseLike<ICalendarEvent> => {
        const oldEvent = this.eventMap[id];
        const oldStartDate = shiftToLocal(new Date(oldEvent.startDate));
        const start = shiftToLocal(startDate);
        const end = shiftToLocal(endDate);

        oldEvent.category = category;
        oldEvent.description = description;
        oldEvent.endDate = end.toISOString();
        oldEvent.startDate = start.toISOString();
        oldEvent.title = title;

        if (toMonthYear(oldStartDate) == toMonthYear(start)) {
            return this.dataManager!.updateDocument(this.selectedTeamId! + "." + toMonthYear(startDate), oldEvent).then(
                (updatedEvent: ICalendarEvent) => {
                    // add event
                    this.eventMap[updatedEvent.id!] = updatedEvent;
                    return updatedEvent;
                }
            );
        } else {
            // move data to new month's collection
            return this.dataManager!.deleteDocument(this.selectedTeamId! + "." + toMonthYear(oldStartDate), oldEvent.id!).then(() => {
                return this.dataManager!.createDocument(this.selectedTeamId! + "." + toMonthYear(startDate), oldEvent).then(
                    (updatedEvent: ICalendarEvent) => {
                        // add event
                        this.eventMap[updatedEvent.id!] = updatedEvent;
                        return updatedEvent;
                    }
                );
            });
        }
    };

    private fetchEvents = (start: Date, end: Date): Promise<{ [id: string]: ICalendarEvent }> => {
        const collectionNames = getMonthYearInRange(start, end).map(item => {
            return this.selectedTeamId! + "." + item;
        });

        const collectionsToFetch: string[] = [];
        collectionNames.forEach(collection => {
            if (!this.fetchedCollections.has(collection)) {
                collectionsToFetch.push(collection);
                this.fetchedCollections.add(collection);
            }
        });

        return this.dataManager!.queryCollectionsByName(collectionsToFetch).then((collections: ExtensionDataCollection[]) => {
            collections.forEach(collection => {
                if (collection && collection.documents) {
                    collection.documents.forEach(doc => {
                        this.eventMap[doc.id] = doc;
                    });
                }
            });

            // if there is old data get it and convert it
            if (!this.fetchedCollections.has(this.selectedTeamId!)) {
                return this.dataManager!.queryCollectionsByName([this.selectedTeamId!]).then((collections: ExtensionDataCollection[]) => {
                    this.fetchedCollections.add(this.selectedTeamId!);
                    if (collections && collections[0] && collections[0].documents) {
                        let lastPromise = Promise.resolve();
                        collections[0].documents.forEach(doc => {
                            this.eventMap[doc.id!] = doc;
                            this.dataManager!.createDocument(this.selectedTeamId! + "." + toMonthYear(new Date(doc.startDate)), doc).then(() => {
                                lastPromise = this.dataManager!.deleteDocument(this.selectedTeamId!, doc.id!);
                            });
                        });
                    }
                    return this.eventMap;
                });
            }

            return Promise.resolve(this.eventMap);
        });
    };
}
