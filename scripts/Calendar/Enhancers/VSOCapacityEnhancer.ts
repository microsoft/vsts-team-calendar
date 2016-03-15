/// <reference path='../../../typings/VSS.d.ts' />
/// <reference path='../../../typings/q.d.ts' />
/// <reference path='../../../typings/jquery.d.ts' />

import Calendar_Contracts = require("Calendar/Contracts");
import Calendar_Dialogs = require("Calendar/Dialogs");
import Controls = require("VSS/Controls");
import Controls_Menus = require("VSS/Controls/Menus");
import Q = require("q");

export class VSOCapacityEnhancer implements Calendar_Contracts.IEventEnhancer {
    private static _deferred: Q.Deferred<any>;
    public id: string = "daysOff";
    public addDialogId: string;  
            
    public icon = "icon-tfs-build-reason-schedule";
            
    constructor() {
            var extensionContext = VSS.getExtensionContext();
            this.addDialogId = extensionContext.publisherId + "." + extensionContext.extensionId + ".add-daysoff-control";  
    }
    
    public getTitle(isEdit: boolean): string{
        return isEdit ? "Edit Days Off" : "Add Days Off";
    }
    
    public canEdit(event: Calendar_Contracts.CalendarEvent, member: Calendar_Contracts.ICalendarMember): IPromise<boolean> {
        return Q.resolve(member === event.member);
    }

    public canAdd(event: Calendar_Contracts.CalendarEvent, member: Calendar_Contracts.ICalendarMember): IPromise<boolean> {
        return Q.resolve(true);
    }
}