import * as Calendar_Contracts from "../Contracts";

export class VSOCapacityEnhancer implements Calendar_Contracts.IEventEnhancer {
    public id: string = "daysOff";
    public addDialogId: string;

    public icon = "icon-tfs-build-reason-schedule";

    constructor() {
        const extensionContext = VSS.getExtensionContext();
        this.addDialogId = extensionContext.publisherId + "." + extensionContext.extensionId + ".add-daysoff-control";
    }

    public canEdit(event: Calendar_Contracts.CalendarEvent, member: Calendar_Contracts.ICalendarMember): PromiseLike<boolean> {
        return Promise.resolve(event.category.title !== "Grouped Event");
    }

    public canAdd(event: Calendar_Contracts.CalendarEvent, member: Calendar_Contracts.ICalendarMember): PromiseLike<boolean> {
        return Promise.resolve(true);
    }
}
