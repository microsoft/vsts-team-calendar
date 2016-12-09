import Calendar_Contracts = require("../Contracts");
import Calendar_DateUtils = require("../Utils/Date");
import Calendar_ColorUtils = require("../Utils/Color");
import Contracts_Platform = require("VSS/Common/Contracts/Platform");
import Contributions_Contracts = require("VSS/Contributions/Contracts");
import ExtensionManagement_RestClient = require("VSS/ExtensionManagement/RestClient");
import FreeForm_Enhancer = require("../Enhancers/FreeFormEnhancer");
import Services_ExtensionData = require("VSS/SDK/Services/ExtensionData");
import Q = require("q");
import Service = require("VSS/Service");
import Utils_Core = require("VSS/Utils/Core");
import Utils_Date = require("VSS/Utils/Date");
import Utils_String = require("VSS/Utils/String");
import WebApi_Constants = require("VSS/WebApi/Constants");

export class FreeFormEventsSource implements Calendar_Contracts.IEventSource {

    public id = "freeForm";
    public name = "Event";
    public order = 10;
    private _enhancer: FreeForm_Enhancer.FreeFormEnhancer;

    private _teamId: string;
    private _categoryId: string;
    private _events: Calendar_Contracts.CalendarEvent[];
    private _categories: Calendar_Contracts.IEventCategory[];

    constructor() {
        var webContext = VSS.getWebContext();
        this._teamId = webContext.team.id;
        this._categoryId = Utils_String.format("{0}-categories", this._teamId);
    }
    
    public load(): IPromise<Calendar_Contracts.CalendarEvent[]> {

        return this.getCategories().then((categories: Calendar_Contracts.IEventCategory[]) => {
             return this.getEvents().then((events: Calendar_Contracts.CalendarEvent[]) => {
                var updatedEvents: Calendar_Contracts.CalendarEvent[] = [];
                $.each(events, (index: number, event: Calendar_Contracts.CalendarEvent) => {
                    // For now, skip events with date strngs we can't parse.
                    if(Date.parse(event.startDate) && Date.parse(event.endDate)) {
                        
                        // update legacy events to match new contract
                        event.movable = true;
                        var category = event.category
                        if(!category || typeof(category) === 'string') {
                            event.category = <Calendar_Contracts.IEventCategory> {
                                title: category || "Uncategorized",
                                id: (this.id + "." + category) || "Uncategorized"
                            }
                            this._updateCategoryForEvents([event]);
                        }
                        // fix times
                        var start = Utils_Date.shiftToUTC(new Date(event.startDate));
                        var end = Utils_Date.shiftToUTC(new Date(event.endDate));
                        if(start.getHours() !== 0) {
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
                });
                return updatedEvents;
            });
        });
    }
    
    public getEnhancer(): IPromise<Calendar_Contracts.IEventEnhancer> {
        if(!this._enhancer){
            this._enhancer = new FreeForm_Enhancer.FreeFormEnhancer();
        }
        return Q.resolve(this._enhancer);
    }

    public getEvents(query?: Calendar_Contracts.IEventQuery): IPromise<Calendar_Contracts.CalendarEvent[]> {
        var deferred = Q.defer<Calendar_Contracts.CalendarEvent[]>();
        VSS.getService("ms.vss-web.data-service").then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            extensionDataService.queryCollectionNames([this._teamId]).then(
                (collections: Contributions_Contracts.ExtensionDataCollection[]) => {
                    if (collections[0] && collections[0].documents) {
                        this._events = collections[0].documents;
                    }
                    else {
                        this._events = [];
                    }
                    deferred.resolve(this._events);
                },
                (e: Error) => {
                    this._events = [];
                    deferred.resolve(this._events);
                });
        });

        return deferred.promise;
    }

    public getCategories(query?: Calendar_Contracts.IEventQuery): IPromise<Calendar_Contracts.IEventCategory[]> {
        var deferred = Q.defer();
        VSS.getService("ms.vss-web.data-service").then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
           extensionDataService.queryCollectionNames([this._categoryId]).then((collections: Contributions_Contracts.ExtensionDataCollection[]) => {
               this._categories = [];
               if(collections[0] && collections[0].documents) {
                   this._categories = collections[0].documents;
               }
               deferred.resolve(this._filterCategories(query));
           }, (e: Error) => {
               deferred.resolve([]);
           });
        });
        return deferred.promise;
    }

    public addEvent(event: Calendar_Contracts.CalendarEvent): IPromise<Calendar_Contracts.CalendarEvent> {
        var deferred = Q.defer();
        VSS.getService("ms.vss-web.data-service").then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            
            extensionDataService.createDocument(this._teamId, event).then(
                (addedEvent: Calendar_Contracts.CalendarEvent) => {
                    // update category for event
                    addedEvent.category.id = this.id + "." + addedEvent.category.title;
                    this._updateCategoryForEvents([addedEvent]);
                    // add event
                    this._events.push(addedEvent);
                    deferred.resolve(addedEvent);
                },
                (e: Error) => {
                    deferred.reject(e);
                });
        });
        return deferred.promise;
    }
    
    public addCategory(category: Calendar_Contracts.IEventCategory): IPromise<Calendar_Contracts.IEventCategory> {
        var deferred = Q.defer();
        VSS.getService("ms.vss-web.data-service").then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            extensionDataService.createDocument(this._categoryId, category).then((addedCategory: Calendar_Contracts.IEventCategory) => {
                this._categories.push(addedCategory);
                deferred.resolve(addedCategory);
            }, (e: Error) => {
                deferred.reject(e);
            });
        });
        return deferred.promise;
    }
    
    public removeEvent(event: Calendar_Contracts.CalendarEvent): IPromise<Calendar_Contracts.CalendarEvent[]> {
        var deferred = Q.defer();
        VSS.getService("ms.vss-web.data-service").then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            extensionDataService.deleteDocument(this._teamId, event.id).then(
                () => {
                    // update category for event
                    event.category = null;
                    this._updateCategoryForEvents([event]);
                    // remove event
                    var eventInArray: Calendar_Contracts.CalendarEvent = $.grep(this._events, (e: Calendar_Contracts.CalendarEvent) => { return e.id === event.id; })[0]; //better check here
                    var index = this._events.indexOf(eventInArray);
                    if (index > -1) {
                        this._events.splice(index, 1);
                    }
                    deferred.resolve(this._events);
                },
                (e: Error) => {
                    //Handle event has already been deleted
                    deferred.reject(e);
                });
        });
        return deferred.promise;
    }
    
    public removeCategory(category: Calendar_Contracts.IEventCategory): IPromise<Calendar_Contracts.IEventCategory[]> {
        var deferred = Q.defer();
        VSS.getService("ms.vss-web.data-service").then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
           extensionDataService.deleteDocument(this._categoryId, category.id).then(() => {
               var categoryInArray: Calendar_Contracts.IEventCategory = $.grep(this._categories, (cat: Calendar_Contracts.IEventCategory) => { return cat.id === category.id})[0];
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
        var deferred = Q.defer();
        return VSS.getService("ms.vss-web.data-service").then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            // update category for event
            newEvent.category.id = this.id + "." + newEvent.category.title;
            
            extensionDataService.updateDocument(this._teamId, newEvent).then(
                (updatedEvent: Calendar_Contracts.CalendarEvent) => {
                    var eventInArray: Calendar_Contracts.CalendarEvent = $.grep(this._events, (e: Calendar_Contracts.CalendarEvent) => { return e.id === updatedEvent.id; })[0]; //better check here
                    var index = this._events.indexOf(eventInArray);
                    if (index > -1) {
                        this._events.splice(index, 1);
                    }
                    if (oldEvent && newEvent.category.id !== oldEvent.category.id) {                        
                        this._updateCategoryForEvents([newEvent]).then((categories: Calendar_Contracts.IEventCategory[]) => {                            
                            this._events.push(updatedEvent);
                            deferred.resolve(updatedEvent);
                        });
                    }
                    else {
                        this._events.push(updatedEvent);
                        deferred.resolve(updatedEvent);
                    }
                },
                (e: Error) => {
                    //Handle concurrency issue
                    return Q.reject(e);
                });
            return deferred.promise;
        },
        (e: Error) => {
            //Handle concurrency issue
            return Q.reject(e);
        });
    }
    
    public updateCategories(categories: Calendar_Contracts.IEventCategory[]): IPromise<Calendar_Contracts.IEventCategory[]> {
        return VSS.getService("ms.vss-web.data-service").then((extensionDataService: Services_ExtensionData.ExtensionDataService) => {
            var updatedCategoriesPromises: IPromise<Calendar_Contracts.IEventCategory>[] = [];
            $.each(categories, (index: number, category: Calendar_Contracts.IEventCategory) => {
                if (category.events.length === 0) {
                    updatedCategoriesPromises.push(this.removeCategory(category)[0]);
                }
                else if (this._categories.filter(cat => cat.id === category.id).length ===0) {
                    updatedCategoriesPromises.push(this.addCategory(category));
                }
                else {
                    updatedCategoriesPromises.push(extensionDataService.updateDocument(this._categoryId, categories[index]).then((updatedCategory: Calendar_Contracts.IEventCategory) => {
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
        deferred.resolve("");
        return deferred.promise;
    }
    
    private _updateCategoryForEvents(events: Calendar_Contracts.CalendarEvent[]): IPromise<Calendar_Contracts.IEventCategory[]> {
        var categoryMap: { [id: string]: boolean } = {};
        var updatedCategories = [];
                
        // remove event from current category        
        for(var i = 0; i < events.length; i++) {
            var event = events[i];
            var categoryForEvent = $.grep(this._categories, (cat: Calendar_Contracts.IEventCategory) => {
                return cat.events.indexOf(event.id) > -1;
            })[0];
            if(categoryForEvent){
                // Do nothing if category hasn't changed
                if(event.category && (event.category.title === categoryForEvent.title)) {
                    event.category = categoryForEvent;
                    return;
                }
                var index = categoryForEvent.events.indexOf(event.id);
                categoryForEvent.events.splice(index, 1);
                var count = categoryForEvent.events.length
                categoryForEvent.subTitle = Utils_String.format("{0} event{1}", count, count > 1 ? "s" : "");
                if(!categoryMap[categoryForEvent.id]) {
                    categoryMap[categoryForEvent.id] = true;
                    updatedCategories.push(categoryForEvent);
                }
            }
            // add event to new category
            if(event.category){
                var newCategory = $.grep(this._categories, (cat: Calendar_Contracts.IEventCategory) => {
                    return cat.id === event.category.id;
                })[0];
                if(newCategory){
                // category already exists
                    newCategory.events.push(event.id);
                    var count = newCategory.events.length
                    newCategory.subTitle = Utils_String.format("{0} event{1}", count, count > 1 ? "s" : "");
                    event.category = newCategory;
                }
                else {
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
        if(!this._events) {
            return [];
        }
        if (!query) {
           return this._categories;
        }
        
        return $.grep(this._categories, (category: Calendar_Contracts.IEventCategory) => {
            var categoryAdded: boolean = false;
            return category.events.some((event: string, eventIndex: number) => {
                var eventInList: Calendar_Contracts.CalendarEvent = $.grep(this._events, (e: Calendar_Contracts.CalendarEvent) => { return e.id === event; })[0];    
                return (eventInList && Calendar_DateUtils.eventIn(eventInList, query));
            });
        });
    }
}