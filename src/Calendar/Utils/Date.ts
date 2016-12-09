import Calendar_Contracts = require("../Contracts");
import Q = require("q");
import Service = require("VSS/Service");
import TFS_Core_Contracts = require("TFS/Core/Contracts");
import Utils_Date = require("VSS/Utils/Date");
import WebApi_Constants = require("VSS/WebApi/Constants");
import Work_Client = require("TFS/Work/RestClient");
import Work_Contracts = require("TFS/Work/Contracts");

function ensureDate(date: string | Date): Date {
    if (typeof date === "string") {
        return new Date(<string>date);
    }

    return <Date>date;
}

/**
 * Checks whether the specified date is between 2 dates
 * @param date Date to check
 * @param startDate Start date
 * @param endDate End date
 * @return True if date is between 2 dates, otherwise false
 */
export function isBetween(date: Date, startDate: Date, endDate: Date): boolean {
    var ticks = date.getTime();
    return ticks >= startDate.getTime() && ticks <= endDate.getTime();
}

/**
 * Checks whether the specified event is within the dates of specified query
 * @param event Event to query
 * @param query Start date and end dates to check
 * @return True if event satisfies query, otherwise false
 */
export function eventIn(event: Calendar_Contracts.CalendarEvent, query: Calendar_Contracts.IEventQuery): boolean {
    if (!query || !query.startDate || !query.endDate) {
        return false;
    }

    if (isBetween(ensureDate(event.startDate), query.startDate, query.endDate)) {
        return true;
    }

    if (isBetween(ensureDate(event.endDate), query.startDate, query.endDate)) {
        return true;
    }

    return false;
}

/**
 * Turns a start date and end date into a list of dates within the range, inclusive
 * @param startDate Start date
 * @param endDate End Date
 * @return Date[] containing each date in the range
 */
export function getDatesInRange(startDate:Date, endDate: Date): Date[] {
    var dates = [];
    var current: Date = startDate;
    while (current.getTime() <= endDate.getTime()) {
        dates.push(new Date(<any>current));
        current = Utils_Date.addDays(current, 1);
    }
    return dates;
}