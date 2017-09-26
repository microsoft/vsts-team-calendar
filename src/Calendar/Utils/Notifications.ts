import Calendar_Contracts = require("../Contracts");
import Locations = require("VSS/Locations");
import VSS_Common_Contracts = require("VSS/WebApi/Contracts");
import Notifications_Extensions = require("Notifications/Extensions");

export function publishCalendarEvent(calendarEvent: Calendar_Contracts.CalendarEvent, action: string) {
    // Publish an event to the notification system
    const notificationEvent: VSS_Common_Contracts.VssNotificationEvent = {
        actors: [
            {
                id: VSS.getWebContext().user.id,
                role: "initiator"
            },
            {
                id: VSS.getWebContext().team.id,
                role: "team"
            }
        ],
        artifactUris: [],
        data: {
            calendarEvent: calendarEvent,
            action: action,
            links: {
                eventPage: Locations.urlHelper.getMvcUrl({ 
                    controller: "apps",
                    action: "hub",
                    parameters: [ VSS.getContribution().id ]
                })
            }
        },
        eventType: "ms-devlabs.team-calendar.calendar-notification-event",
        scopes: [
            {
                id: VSS.getWebContext().collection.id,
                type: "collection"
            },
            {
                id: VSS.getWebContext().project.id,
                type: "project"
            }
        ]
    };

    Notifications_Extensions.publishEvent(notificationEvent);
}