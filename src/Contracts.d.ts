/**
 * Represents a single calendar event
 */
export interface ICalendarEvent {
    /**
     * Title of the event
     */
    title: string;

    /**
     * Used by collection to
     */
    __etag?: number;

    /**
     * Event start date
     */
    startDate: string;

    /**
     * Event end date
     */
    endDate: string;

    /**
     * Unique id for the event
     */
    id?: string;

    /**
     * Id of the iteration to which the event is linked
     * (previous version of calendear allowed days off to be in wrong iteration)
     */
    iterationId?: string;

    /**
     * Category of the service
     */
    category: string | IEventCategory;

    /**
     * The member associated with this event
     */
    member?: ICalendarMember;

    /**
     * A description of the event
     */
    description?: string;

    /**
     * Icons to be displayed on the event
     */
    icons?: IEventIcon[];
}

export interface ICalendarMember {
    /**
     * Display name of the member
     */
    displayName: string;

    /**
     * Unique ID for the member
     */
    id: string;
}

/**
 * Summary item for events
 */
export interface IEventCategory {
    /**
     * Title of the event category
     */
    title: string;

    /**
     * Sub title of the event category
     */
    subTitle?: string;

    /**
     * Image url of the category
     */
    imageUrl?: string;

    /**
     * Color of the category
     */
    color?: string;

    /**
     * Number of event under this Category
     */
    eventCount: number;
}

/**
 * An icon displayed on the calendar representing an event
 */
export interface IEventIcon {
    /**
     * src url for the icon
     */
    src: string;

    /**
     * The event to edit or delete when the icon is selected
     */
    linkedEvent: ICalendarEvent;
}
