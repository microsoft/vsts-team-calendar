import { IExtensionDataManager, ExtensionDataCollection } from "azure-devops-extension-api";

import { ObservableArray } from "azure-devops-ui/Core/Observable";

import { EventInput } from "@fullcalendar/core";
import { EventSourceError } from "@fullcalendar/core/structs/event-source";

import { generateColor } from "./Color";
import { ICalendarEvent, IEventCategory } from "./Contracts";
import { shiftToLocal, shiftToUTC, getMonthYearInRange, formatDate } from "./TimeLib";

export const FreeFormId = "FreeForm";

export class FreeFormEventsSource {
    categories: Set<string> = new Set<string>();
    dataManager?: IExtensionDataManager;
    eventMap: { [id: string]: ICalendarEvent } = {};
    fetchedCollections: Set<string> = new Set<string>();
    selectedTeamId: string = "";
    summaryData: ObservableArray<IEventCategory> = new ObservableArray<IEventCategory>([]);

    public addEvent = (title: string, startDate: Date, endDate: Date, category: string, description: string): PromiseLike<ICalendarEvent> => {
        const start = shiftToUTC(startDate);
        const end = shiftToUTC(endDate);

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

        return this.dataManager!.createDocument(this.selectedTeamId! + "." + formatDate(startDate, "MM-YYYY"), event).then(
            (addedEvent: ICalendarEvent) => {
                // add event to cache
                // use times from current zone
                this.eventMap[addedEvent.id!] = addedEvent;
                addedEvent.startDate = start.toISOString();
                addedEvent.endDate = end.toISOString();
                return addedEvent;
            }
        ).catch((error: any) => {
            // If document already exists, provide a more user-friendly error
            if (error.status === 400 && error.responseText && error.responseText.indexOf('DocumentExistsException') > -1) {
                throw new Error(`Event "${title}" already exists or there was a conflict creating it.`);
            }
            throw error;
        });
    };

    public deleteEvent = (eventId: string, startDate: Date) => {
        delete this.eventMap[eventId];
        return this.dataManager!.deleteDocument(this.selectedTeamId! + "." + formatDate(startDate, "MM-YYYY"), eventId).catch((error: any) => {
            // If document/collection doesn't exist or any 404 error, treat as successful deletion
            if (error.status === 404) {
                return; // Treat as success
            }
            throw error;
        });
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

        return this.fetchEvents(calendarStart, calendarEnd).then(() => {
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
                    if (!event.category) {
                        event.category = "Uncategorized";
                    }

                    const start = shiftToLocal(new Date(event.startDate));
                    const end = shiftToLocal(new Date(event.endDate));

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
            this.summaryData.value = Object.keys(catagoryMap).map(key => {
                const catagory = catagoryMap[key];
                if (catagory.eventCount > 1) {
                    catagory.subTitle = catagory.eventCount + " events";
                }
                return catagory;
            });
        }).catch((error: any) => {
            console.error("[FreeFormEventSource] Error fetching events:", error);
            failureCallback({ message: error.message || "Failed to load custom events" });
        });
    };

    public getSummaryData = (): ObservableArray<IEventCategory> => {
        return this.summaryData;
    };

    public initialize(teamId: string, manager: IExtensionDataManager) {
        this.selectedTeamId = teamId;
        this.dataManager = manager;
        this.eventMap = {};
        this.categories.clear();
        this.fetchedCollections.clear();
    }

    public preloadCurrentMonthEvents(): Promise<void> {
        if (!this.dataManager || !this.selectedTeamId) {
            return Promise.resolve();
        }

        const now = new Date();
        const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
        const endOfMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0);
        
        // First, check for any legacy data that might not have been migrated
        return this.ensureLegacyDataMigration().then(() => {
            // Then fetch current month events
            return this.fetchEvents(startOfMonth, endOfMonth).then(() => {
                // Preload complete
            });
        });
    }

    private ensureLegacyDataMigration(): Promise<void> {
        if (!this.dataManager || !this.selectedTeamId) {
            return Promise.resolve();
        }

        // Check the main team collection for any remaining legacy events
        return this.dataManager.queryCollectionsByName([this.selectedTeamId]).then((collections: ExtensionDataCollection[]) => {
            if (collections && collections[0] && collections[0].documents && collections[0].documents.length > 0) {
                const oldData: ICalendarEvent[] = [];
                collections[0].documents.forEach((doc: ICalendarEvent) => {
                    // Add to current cache immediately
                    this.eventMap[doc.id!] = doc;
                    oldData.push(doc);
                });
                
                // Convert/migrate the data
                this.convertData(oldData);
                
                // Mark the legacy collection as processed
                this.fetchedCollections.add(this.selectedTeamId);
            } else {
                // Still mark as processed to avoid future checks
                this.fetchedCollections.add(this.selectedTeamId);
            }
        }).catch(error => {
            // Mark as processed even if there's an error to avoid infinite retries
            this.fetchedCollections.add(this.selectedTeamId);
        });
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
        const oldStartDate = new Date(oldEvent.startDate);

        oldEvent.category = category;
        oldEvent.description = description;
        oldEvent.endDate = shiftToUTC(endDate).toISOString();
        oldEvent.startDate = shiftToUTC(startDate).toISOString();
        oldEvent.title = title;

        const collectionNameOld = this.selectedTeamId! + "." + formatDate(oldStartDate, "MM-YYYY");
        const collectionNameNew = this.selectedTeamId! + "." + formatDate(startDate, "MM-YYYY");

        if (collectionNameOld == collectionNameNew) {
            return this.dataManager!.updateDocument(collectionNameNew, oldEvent).then((updatedEvent: ICalendarEvent) => {
                // add event
                this.eventMap[updatedEvent.id!] = updatedEvent;
                return updatedEvent;
            }).catch((error: any) => {
                throw error;
            });
        } else {
            // move data to new month's collection
            return this.dataManager!.deleteDocument(collectionNameOld, oldEvent.id!).then(() => {
                return this.dataManager!.createDocument(collectionNameNew, oldEvent).then((updatedEvent: ICalendarEvent) => {
                    // add event
                    this.eventMap[updatedEvent.id!] = updatedEvent;
                    return updatedEvent;
                }).catch((createError: any) => {
                    // If creation fails due to document already existing, treat as successful
                    if (createError.status === 400 && createError.responseText && createError.responseText.indexOf('DocumentExistsException') > -1) {
                        return oldEvent; // Return the original event as if it was successful
                    }
                    throw createError;
                });
            }).catch((deleteError: any) => {
                // If delete fails with 404, try to create in new location anyway
                if (deleteError.status === 404) {
                    return this.dataManager!.createDocument(collectionNameNew, oldEvent).then((updatedEvent: ICalendarEvent) => {
                        this.eventMap[updatedEvent.id!] = updatedEvent;
                        return updatedEvent;
                    }).catch((createError: any) => {
                        if (createError.status === 400 && createError.responseText && createError.responseText.indexOf('DocumentExistsException') > -1) {
                            return oldEvent;
                        }
                        throw createError;
                    });
                }
                throw deleteError;
            });
        }
    };

    private convertData = (oldData: ICalendarEvent[]) => {
        // chain all actions in to max 10 queues
        let queue: Promise<void>[] = [];
        const maxSize = oldData.length < 10 ? oldData.length : 10;

        let index: number;
        for (index = 0; index < maxSize; index++) {
            queue[index] = Promise.resolve();
        }

        // create new event and delete old one
        oldData.forEach(doc => {
            if (index === maxSize) {
                index = 0;
            }
            queue[index] = queue[index].then(() => {
                const collectionName = this.selectedTeamId! + "." + formatDate(new Date(doc.startDate), "MM-YYYY");
                return this.dataManager!.createDocument(collectionName, doc).catch((error: any) => {
                    // If document already exists, skip creation
                    if (error.status === 400 && error.responseText && error.responseText.indexOf('DocumentExistsException') > -1) {
                        return;
                    }
                    throw error;
                });
            });
            queue[index] = queue[index].then(() => {
                return this.dataManager!.deleteDocument(this.selectedTeamId!, doc.id!).catch((error: any) => {
                    // If any 404 error, treat as successful deletion
                    if (error.status === 404) {
                        return;
                    }
                    throw error;
                });
            });
            index++;
        });

        // delete categories data if there is any
        this.dataManager!.queryCollectionsByName([this.selectedTeamId! + "-categories"]).then((collections: ExtensionDataCollection[]) => {
            if (collections && collections[0] && collections[0].documents) {
                collections[0].documents.forEach(doc => {
                    if (index === maxSize) {
                        index = 0;
                    }
                    queue[index] = queue[index].then(() => {
                        return this.dataManager!.deleteDocument(this.selectedTeamId! + "-categories", doc.id!).catch((error: any) => {
                            // If any 404 error, treat as successful deletion
                            if (error.status === 404) {
                                return;
                            }
                            throw error;
                        });
                    });
                    index++;
                });
            }
        }).catch((error: any) => {
            // This is not critical, so we can continue
        });
    };

    private generatePotentialLegacyCollections(start: Date, end: Date): string[] {
        const legacyCollections: string[] = [];
        
        // Check for potential variations in date formatting that might have been used
        const monthsInRange = getMonthYearInRange(start, end);
        
        monthsInRange.forEach(monthYear => {
            // Current format: teamId.M.YYYY or teamId.MM.YYYY
            const teamPrefix = this.selectedTeamId + ".";
            
            // Try different potential formats that might have been used historically
            const [month, year] = monthYear.split('.');
            const monthNum = parseInt(month);
            
            // Variations to check:
            // 1. Zero-padded month: teamId.09.2025
            const paddedMonth = monthNum < 10 ? "0" + monthNum : monthNum.toString();
            legacyCollections.push(teamPrefix + paddedMonth + "." + year);
            
            // 2. Different separator: teamId-M-YYYY  
            legacyCollections.push(this.selectedTeamId + "-" + month + "-" + year);
            legacyCollections.push(this.selectedTeamId + "-" + paddedMonth + "-" + year);
            
            // 3. Different order: teamId.YYYY.M
            legacyCollections.push(teamPrefix + year + "." + month);
            legacyCollections.push(teamPrefix + year + "." + paddedMonth);
        });
        
        // Also check for the plain team ID collection (legacy format)
        legacyCollections.push(this.selectedTeamId);
        
        return legacyCollections.filter((item, index, arr) => arr.indexOf(item) === index); // Remove duplicates
    }

    private fetchEvents = (start: Date, end: Date): Promise<{ [id: string]: ICalendarEvent }> => {
        const collectionNames = getMonthYearInRange(start, end).map(item => {
            return this.selectedTeamId! + "." + item;
        });

        const collectionsToFetch: string[] = [];
        collectionNames.forEach(collection => {
            if (!this.fetchedCollections.has(collection)) {
                collectionsToFetch.push(collection);
                // DON'T mark as fetched yet - wait until after successful fetch
            }
        });

        // Also check for potential legacy collection variations for backward compatibility
        const legacyCollections = this.generatePotentialLegacyCollections(start, end);
        legacyCollections.forEach(collection => {
            if (!this.fetchedCollections.has(collection) && collectionsToFetch.indexOf(collection) === -1) {
                collectionsToFetch.push(collection);
            }
        });

        return this.dataManager!.queryCollectionsByName(collectionsToFetch).then((collections: ExtensionDataCollection[]) => {
            // Mark collections as fetched AFTER successful retrieval
            collectionsToFetch.forEach(collectionName => {
                this.fetchedCollections.add(collectionName);
            });
            
            collections.forEach(collection => {
                if (collection && collection.documents && collection.documents.length > 0) {
                    // Check if this is a legacy collection that needs migration
                    const isLegacyCollection = collection.collectionName === this.selectedTeamId || 
                                             collection.collectionName.indexOf('-') > -1 || // Old separator format
                                             !collection.collectionName.match(/\.\d{1,2}\.\d{4}$/); // Doesn't match current format
                    
                    if (isLegacyCollection && collection.documents.length > 0) {
                        const eventsToMigrate: ICalendarEvent[] = [];
                        
                        collection.documents.forEach(doc => {
                            // Add to current cache immediately for availability
                            this.eventMap[doc.id] = doc;
                            eventsToMigrate.push(doc);
                        });
                        
                        // Migrate the events to proper monthly collections (async, don't wait)
                        if (eventsToMigrate.length > 0) {
                            setTimeout(() => {
                                this.convertData(eventsToMigrate);
                            }, 100); // Small delay to ensure current operation completes first
                        }
                    } else {
                        // Regular collection processing
                        collection.documents.forEach(doc => {
                            this.eventMap[doc.id] = doc;
                        });
                    }
                }
            });

            // if there is old data get it and convert it
            if (!this.fetchedCollections.has(this.selectedTeamId!)) {
                return this.dataManager!.queryCollectionsByName([this.selectedTeamId!]).then((collections: ExtensionDataCollection[]) => {
                    this.fetchedCollections.add(this.selectedTeamId!);
                    if (collections && collections[0] && collections[0].documents) {
                        const oldData: ICalendarEvent[] = [];
                        collections[0].documents.forEach((doc: ICalendarEvent) => {
                            this.eventMap[doc.id!] = doc;
                            oldData.push(doc);
                        });
                        // Migrate the data (async, don't wait)
                        setTimeout(() => {
                            this.convertData(oldData);
                        }, 100);
                    }
                    return this.eventMap;
                }).catch((error: any) => {
                    // Mark as fetched even on error to prevent retries
                    this.fetchedCollections.add(this.selectedTeamId!);
                    return this.eventMap;
                });
            }
            return Promise.resolve(this.eventMap);
        }).catch((error: any) => {
            // Mark collections as fetched even on error to prevent infinite retries
            collectionsToFetch.forEach(collectionName => {
                this.fetchedCollections.add(collectionName);
            });
            return this.eventMap;
        });
    };
}
