import * as Calendar_Contracts from "../Contracts";

export class FreeFormEnhancer implements Calendar_Contracts.IEventEnhancer {
    public id: string = "freeForm";
    public addDialogId: string;

    constructor() {
        const extensionContext = VSS.getExtensionContext();
        this.addDialogId = extensionContext.publisherId + "." + extensionContext.extensionId + ".add-freeform-control";
    }

    public canEdit(event: Calendar_Contracts.CalendarEvent, member: Calendar_Contracts.ICalendarMember): Promise<boolean> {
        return Promise.resolve(true);
    }

    public canAdd(event: Calendar_Contracts.CalendarEvent, member: Calendar_Contracts.ICalendarMember): Promise<boolean> {
        return Promise.resolve(true);
    }
}
