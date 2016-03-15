/**
 * Interface for Add Event dialog content
 */
export interface IDialogContent{
    /**
     * Returns the updated calendar item
     */
    onOkClick: () => IPromise<CalendarEvent>; 
    
    /**
     * Returns the current title of the dialog
     */
    getTitle?: () => IPromise<string>;
}

/**
* Interface for a calendar event provider view
*/
export interface IEventEnhancer {
    
    /**
    * Unique id of the enhancer
    */
    id: string;

    /**
    * Unique id of the add dialog content control
    */
    addDialogId?: string;
    
    /**
    * Unique id of the side panel control
    */
    sidePanelId?: string;
    
    /**
    * optional customizable add-icon for context menu
    */
    icon?: string;
    
    /**
    * Determines whether an event is editable in the current context
    *
    * @param event
    */
    canEdit: (event: CalendarEvent, member: ICalendarMember) => IPromise<boolean>;

    /**
     * Determines whether an event can be added
     */
    canAdd: (event: CalendarEvent, member: ICalendarMember) => IPromise<boolean>;
}

/**
* Interface for a calendar event provider
*/
export interface IEventSource {

    /**
    * Unique id of the event source
    */
    id: string;

    /**
    * Friendly display name of the event source
    */
    name: string;

    /**
    * Order used in sorting the event sources
    */
    order: number;

    /**
     * Returns the UI enhancer for the event source
     */
    getEnhancer?: () => IPromise<IEventEnhancer>;

    /**
    * Set to true if events from this source should be rendered in the background.
    */
    background?: boolean;
    
    /**
     * Returns true when the event source is loaded
     */
    load: () => IPromise<CalendarEvent[]>;

    /**
    * Get the events that match a certain criteria
    *
    * @param query Events query
    */
    getEvents: (query?: IEventQuery) => IPromise<CalendarEvent[]>;

    /**
     * Get the event categories that match a certain criteria
     */
    getCategories(query?: IEventQuery): IPromise<IEventCategory[]>;

    /**
    * Optional method to add events to a given source
    */
    addEvent?: (events: CalendarEvent) => IPromise<CalendarEvent>;

    /**
    * Optional method to remove events from this event source
    */
    removeEvent?: (events: CalendarEvent) => IPromise<CalendarEvent[]>;

    /**
    * Optional method to update an event in this event source
    */
    updateEvent?: (oldEvent: CalendarEvent, newEvent: CalendarEvent) => IPromise<CalendarEvent>;
    
    /**
    * Forms the url which is linked to the title of the summary section for the source
    */
    getTitleUrl(webContext: WebContext): IPromise<string>;
}

/**
 * Summary item for events
 */
export interface IEventCategory {
    /**
     * Unique id of the category
     * {source id}.{category title}
     */
    id: string;
    /**
     * Title of the event category
     */
    title: string;

    /**
     * Sub title of the event category
     */
    subTitle?: string;
    
    /**
     * Ids of the events in the category
     */
    events?: string[];

    /**
     * Image url of this category
     */
    imageUrl?: string;

    /**
     * Color of this category
     */
    color?: string;
}

/**
* Query criteria for events
*/
export interface IEventQuery {

    /**
    * If specified, only include events on or after the given date
    */
    startDate?: Date;
    
    /**
    * If specified, only include events on or before the given date
    */
    endDate?: Date;
}

/**
* Represents a single calendar event
*/
export interface CalendarEvent {

    /**
    * Title of the event
    */
    title: string;

    __etag?: number;

    /**
    * Event start date
    */
    startDate: string;

    /**
    * Event end date
    */
    endDate?: string;

    /**
    * Unique id for the event
    */
    id?: string;
    
    /**
     * Category of the service
     */
    category?: IEventCategory;
    
    /**
     * Id of the iteration to which the event is linked
     */
    iterationId?: string;
    
    /**
     * Whether the event is movable on the calendar
     */
    movable?: boolean;

    /**
     * The member associated with this event
     */
    member?: ICalendarMember;
    
    /**
     * A description of the event
     */
    description?: string;
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

    /**
    * URL to the identity image for the member
    */
    imageUrl: string;

    /**
    * Unique name for the member
    */
    uniqueName: string;

    /**
    * URL for the member
    */
    url: string;
}

/**
* Represents a single calendar event
*/
export interface IExtendedCalendarEventObject {
    color?: string;
    backgroundColor?: string;
    borderColor?: string;
    textColor?: string;
    className?: string|string[];
    editable?: boolean;
    startEditable?: boolean;
    durationEditable?: boolean;
    rendering?: string;
    overlap?: boolean;
    constraint?: string;
    id?: string;
    __etag?: number;
    title: string;
    description?: string;
    allDay?: boolean;
    start: Date|string;
    end?: Date|string;
    url?: string;
    source?: any | IExtendedCalendarEventSource;
    member?: ICalendarMember;
    category?: IEventCategory;
    iterationId?: string;
    eventType?: string;
}

/**
* Represents a single calendar event
*/
export interface IExtendedCalendarEventSource {
    events?: IExtendedCalendarEventObject[] | IEventSource;
}


