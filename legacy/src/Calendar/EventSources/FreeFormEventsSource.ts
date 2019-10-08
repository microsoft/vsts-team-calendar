import { WebApiTeam } from 'TFS/Core/Contracts';
import * as Calendar_Contracts from '../Contracts';
import * as Calendar_DateUtils from '../Utils/Date';
import * as Calendar_ColorUtils from '../Utils/Color';
import * as Contributions_Contracts from 'VSS/Contributions/Contracts';
import * as FreeForm_Enhancer from '../Enhancers/FreeFormEnhancer';
import * as Services_ExtensionData from 'VSS/SDK/Services/ExtensionData';
import * as Utils_Date from 'VSS/Utils/Date';
import * as Utils_String from 'VSS/Utils/String';

export class FreeFormEventsSource implements Calendar_Contracts.IEventSource {
    public id = "freeForm";
    public name = "Event";
    public order = 10;
    private _enhancer: FreeForm_Enhancer.FreeFormEnhancer;

    private _teamId: string;
    private _categoryId: string;
    private _events: Calendar_Contracts.CalendarEvent[];
    private _categories: Calendar_Contracts.IEventCategory[];

    constructor(context?: any) {
        this.updateTeamContext(context.team);
    }

    public updateTeamContext(newTeam: WebApiTeam) {
        this._teamId = newTeam.id;
        this._categoryId = Utils_String.format("{0}-categories", this._teamId);
    }

    public load(): PromiseLike<Calendar_Contracts.CalendarEvent[]> {
        return this.getCategories().then((categories: Calendar_Contracts.IEventCategory[]) => {
            return this.getEvents().then((events: Calendar_Contracts.CalendarEvent[]) => {
                const updatedEvents: Calendar_Contracts.CalendarEvent[] = [];
                for (const event of events) {
                    // For now, skip events with date strngs we can't parse.
                    if (Date.parse(event.startDate) && Date.parse(event.endDate)) {
                        // update legacy events to match new contract
                        event.movable = true;
                        const category = event.category;
                        if (!category || typeof category === "string") {
                            event.category = <Calendar_Contracts.IEventCategory>{
                                title: category || "Uncategorized",
                                id: this.id + "." + category || "Uncategorized",
                            };
                            this._updateCategoryForEvents([event]);
                        }
                        // fix times
                        const start = Utils_Date.shiftToUTC(new Date(event.startDate));
                        const end = Utils_Date.shiftToUTC(new Date(event.endDate));
                        if (start.getHours() !== 0) {
                            // Set dates back to midnight
                            start.setHours(0);
                            end.setHours(0);
                            // update the event in the list
                            event.startDate = Utils_Date.shiftToLocal(start).toISOString();
                            event.endDate = Utils_Date.shiftToLocal(end).toISOString();
                            this.updateEvent(null, event);
                        }
                        updatedEvents.push(event);
                    }
                }
                return updatedEvents;
            });
        });
    }

    public getEnhancer(): PromiseLike<Calendar_Contracts.IEventEnhancer> {
        if (!this._enhancer) {
            this._enhancer = new FreeForm_Enhancer.FreeFormEnhancer();
        }
        return Promise.resolve(this._enhancer);
    }

    public getEvents(query?: Calendar_Contracts.IEventQuery): PromiseLike<Calendar_Contracts.CalendarEvent[]> {
        return VSS.getService(
            "ms.vss-web.data-service",
        ).then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            return extensionDataService.queryCollectionNames([this._teamId]).then(
                (collections: Contributions_Contracts.ExtensionDataCollection[]) => {
                    if (collections[0] && collections[0].documents) {
                        this._events = collections[0].documents;
                    } else {
                        this._events = [];
                    }
                    return this._events;
                },
                (e: Error) => {
                    this._events = [];
                    return this._events;
                },
            );
        });
    }

    public getCategories(query?: Calendar_Contracts.IEventQuery): PromiseLike<Calendar_Contracts.IEventCategory[]> {
        return VSS.getService(
            "ms.vss-web.data-service",
        ).then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            return extensionDataService.queryCollectionNames([this._categoryId]).then(
                (collections: Contributions_Contracts.ExtensionDataCollection[]) => {
                    this._categories = [];
                    if (collections[0] && collections[0].documents) {
                        this._categories = collections[0].documents;
                    }
                    return this._filterCategories(query);
                },
                (e: Error) => {
                    return [];
                },
            );
        });
    }

    public addEvent(event: Calendar_Contracts.CalendarEvent): PromiseLike<Calendar_Contracts.CalendarEvent> {
        return VSS.getService(
            "ms.vss-web.data-service",
        ).then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            return extensionDataService
                .createDocument(this._teamId, event)
                .then((addedEvent: Calendar_Contracts.CalendarEvent) => {
                    // update category for event
                    addedEvent.category.id = this.id + "." + addedEvent.category.title;
                    this._updateCategoryForEvents([addedEvent]);
                    // add event
                    this._events.push(addedEvent);
                    return addedEvent;
                });
        });
    }

    public addCategory(category: Calendar_Contracts.IEventCategory): PromiseLike<Calendar_Contracts.IEventCategory> {
        return VSS.getService(
            "ms.vss-web.data-service",
        ).then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            return extensionDataService
                .createDocument(this._categoryId, category)
                .then((addedCategory: Calendar_Contracts.IEventCategory) => {
                    this._categories.push(addedCategory);
                    return addedCategory;
                });
        });
    }

    public removeEvent(event: Calendar_Contracts.CalendarEvent): PromiseLike<Calendar_Contracts.CalendarEvent[]> {
        return VSS.getService(
            "ms.vss-web.data-service",
        ).then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            return extensionDataService.deleteDocument(this._teamId, event.id).then(() => {
                // update category for event
                event.category = null;
                this._updateCategoryForEvents([event]);
                // remove event
                const eventInArray: Calendar_Contracts.CalendarEvent = $.grep(
                    this._events,
                    (e: Calendar_Contracts.CalendarEvent) => {
                        return e.id === event.id;
                    },
                )[0]; //better check here
                const index = this._events.indexOf(eventInArray);
                if (index > -1) {
                    this._events.splice(index, 1);
                }
                return this._events;
            });
        });
    }

    public removeCategory(category: Calendar_Contracts.IEventCategory): PromiseLike<Calendar_Contracts.IEventCategory[]> {
        return VSS.getService(
            "ms.vss-web.data-service",
        ).then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            return extensionDataService.deleteDocument(this._categoryId, category.id).then(() => {
                const categoryInArray: Calendar_Contracts.IEventCategory = $.grep(
                    this._categories,
                    (cat: Calendar_Contracts.IEventCategory) => {
                        return cat.id === category.id;
                    },
                )[0];
                const index = this._categories.indexOf(categoryInArray);
                if (index > -1) {
                    this._categories.splice(index, 1);
                }
                return this._categories;
            });
        });
    }

    public updateEvent(
        oldEvent: Calendar_Contracts.CalendarEvent,
        newEvent: Calendar_Contracts.CalendarEvent,
    ): PromiseLike<Calendar_Contracts.CalendarEvent> {
        return VSS.getService<Services_ExtensionData.ExtensionDataService>(
            "ms.vss-web.data-service",
        ).then(extensionDataService => {
            newEvent.category.id = `${this.id}.${newEvent.category.title}`;
            return extensionDataService.updateDocument(this._teamId, newEvent).then(updatedEvent => {
                const eventInArray = this._events.filter(e => e.id === updatedEvent.id)[0];
                const index = this._events.indexOf(eventInArray);
                if (index >= 0) {
                    this._events.splice(index, 1);
                }
                if (oldEvent && newEvent.category.id !== oldEvent.category.id) {
                    return this._updateCategoryForEvents([newEvent]).then(categories => {
                        this._events.push(updatedEvent);
                        return updatedEvent;
                    });
                } else {
                    this._events.push(updatedEvent);
                    return updatedEvent;
                }
            });
        });
    }

    public updateCategories(categories: Calendar_Contracts.IEventCategory[]): PromiseLike<Calendar_Contracts.IEventCategory[]> {
        return VSS.getService(
            "ms.vss-web.data-service",
        ).then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            const updatedCategoriesPromises: PromiseLike<Calendar_Contracts.IEventCategory>[] = [];
            let index = 0;
            for (const category of categories) {
                if (category.events.length === 0) {
                    updatedCategoriesPromises.push(this.removeCategory(category)[0]);
                } else if (this._categories.filter(cat => cat.id === category.id).length === 0) {
                    updatedCategoriesPromises.push(this.addCategory(category));
                } else {
                    updatedCategoriesPromises.push(
                        extensionDataService
                            .updateDocument(this._categoryId, categories[index++])
                            .then((updatedCategory: Calendar_Contracts.IEventCategory) => {
                                const categoryInArray: Calendar_Contracts.IEventCategory = $.grep(
                                    this._categories,
                                    (cat: Calendar_Contracts.IEventCategory) => {
                                        return cat.id === category.id;
                                    },
                                )[0];
                                const index = this._categories.indexOf(categoryInArray);
                                if (index > -1) {
                                    this._categories.splice(index, 1);
                                }
                                this._categories.push(updatedCategory);
                                return updatedCategory;
                            }),
                    );
                }
            }
            return Promise.all(updatedCategoriesPromises);
        });
    }

    public getTitleUrl(webContext: WebContext): PromiseLike<string> {
        return Promise.resolve("");
    }

    private _updateCategoryForEvents(
        events: Calendar_Contracts.CalendarEvent[],
    ): PromiseLike<Calendar_Contracts.IEventCategory[]> {
        const categoryMap: { [id: string]: boolean } = {};
        const updatedCategories = [];

        // remove event from current category
        for (const event of events) {
            const categoryForEvent = $.grep(this._categories, (cat: Calendar_Contracts.IEventCategory) => {
                return cat.events.indexOf(event.id) > -1;
            })[0];
            if (categoryForEvent) {
                // Do nothing if category hasn't changed
                if (event.category && event.category.title === categoryForEvent.title) {
                    event.category = categoryForEvent;
                    return;
                }
                const index = categoryForEvent.events.indexOf(event.id);
                categoryForEvent.events.splice(index, 1);
                const count = categoryForEvent.events.length;
                categoryForEvent.subTitle = Utils_String.format("{0} event{1}", count, count > 1 ? "s" : "");
                if (!categoryMap[categoryForEvent.id]) {
                    categoryMap[categoryForEvent.id] = true;
                    updatedCategories.push(categoryForEvent);
                }
            }
            // add event to new category
            if (event.category) {
                let newCategory = $.grep(this._categories, (cat: Calendar_Contracts.IEventCategory) => {
                    return cat.id === event.category.id;
                })[0];
                if (newCategory) {
                    // category already exists
                    newCategory.events.push(event.id);
                    const count = newCategory.events.length;
                    newCategory.subTitle = Utils_String.format("{0} event{1}", count, count > 1 ? "s" : "");
                    event.category = newCategory;
                } else {
                    // category doesn't exist yet
                    newCategory = event.category;
                    newCategory.events = [event.id];
                    newCategory.subTitle = event.title;
                    newCategory.color = Calendar_ColorUtils.generateColor(event.category.title);
                }
                if (!categoryMap[newCategory.id]) {
                    categoryMap[newCategory.id] = true;
                    updatedCategories.push(newCategory);
                }
            }
        }
        // update categories
        return this.updateCategories(updatedCategories);
    }

    private _filterCategories(query?: Calendar_Contracts.IEventQuery): Calendar_Contracts.IEventCategory[] {
        if (!this._events) {
            return [];
        }
        if (!query) {
            return this._categories;
        }

        return $.grep(this._categories, (category: Calendar_Contracts.IEventCategory) => {
            const categoryAdded: boolean = false;
            return category.events.some((event: string, eventIndex: number) => {
                const eventInList: Calendar_Contracts.CalendarEvent = $.grep(
                    this._events,
                    (e: Calendar_Contracts.CalendarEvent) => {
                        return e.id === event;
                    },
                )[0];
                return eventInList && Calendar_DateUtils.eventIn(eventInList, query);
            });
        });
    }
}
