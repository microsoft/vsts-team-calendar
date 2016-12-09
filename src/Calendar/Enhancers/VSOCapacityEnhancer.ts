import Calendar_Contracts = require("../Contracts");
import Q = require("q");

export class VSOCapacityEnhancer implements Calendar_Contracts.IEventEnhancer {
    
    public id: string = "daysOff";
    public addDialogId: string;  
            
    public icon = "icon-tfs-build-reason-schedule";
            
    constructor() {
            var extensionContext = VSS.getExtensionContext();
            this.addDialogId = extensionContext.publisherId + "." + extensionContext.extensionId + ".add-daysoff-control";  
    }
    
    public canEdit(event: Calendar_Contracts.CalendarEvent, member: Calendar_Contracts.ICalendarMember): IPromise<boolean> {
        return Q.resolve(event.category.title !== "Grouped Event");
    }

    public canAdd(event: Calendar_Contracts.CalendarEvent, member: Calendar_Contracts.ICalendarMember): IPromise<boolean> {
        return Q.resolve(true);
    }
}
