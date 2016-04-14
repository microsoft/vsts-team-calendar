/// <reference path='../../../typings/VSS.d.ts' />
/// <reference path='../../../typings/q.d.ts' />
/// <reference path='../../../typings/jquery.d.ts' />

import Calendar_Contracts = require("Calendar/Contracts");
import Q = require("q");

export class FreeFormEnhancer implements Calendar_Contracts.IEventEnhancer {
    
    public id: string = "freeform";    
    public addDialogId: string;
        
    constructor(){
            var extensionContext = VSS.getExtensionContext();
            this.addDialogId = extensionContext.publisherId + "." + extensionContext.extensionId + ".add-freeform-control";        
    }
        
    public canEdit(event: Calendar_Contracts.CalendarEvent, member: Calendar_Contracts.ICalendarMember): IPromise<boolean> {
        return Q.resolve(true);
    }
    
    public canAdd(event: Calendar_Contracts.CalendarEvent, member: Calendar_Contracts.ICalendarMember): IPromise<boolean>{
        return Q.resolve(true);
    }
}