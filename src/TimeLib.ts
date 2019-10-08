const minute = 1000 * 60;

export const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
export interface MonthAndYear {
    month: number;
    year: number;
}

export function formatDate(date: Date, format?: string): string {
    if (format === "YYYY-MM-DD" || format === "MM-DD-YYYY") {
        let month = "" + (date.getMonth() + 1);
        let day = "" + date.getDate();
        const year = date.getFullYear();

        if (month.length < 2) month = "0" + month;
        if (day.length < 2) day = "0" + day;

        return format === "YYYY-MM-DD" ? [year, month, day].join("-") : [month, day, year].join("-");
    } else if (format === "MONTH-DD") {
        return months[date.getMonth()] + " " + date.getDate();
    }
    return date.toISOString();
}

export function monthAndYearToString(monthAndYear: MonthAndYear): string {
    return months[monthAndYear.month] + " " + monthAndYear.year;
}

/**
 * Turns a start date and end date into a list of dates within the range, inclusive
 * @param startDate Start date
 * @param endDate End Date
 * @return Date[] containing each date in the range
 */
export function getDatesInRange(startDate: Date, endDate: Date): Date[] {
    const dates = [];
    let current: Date = startDate;
    while (current.getTime() <= endDate.getTime()) {
        dates.push(new Date(current));
        current.setDate(current.getDate() + 1);
    }
    return dates;
}

/**
 * Turns a start date and end date into a list of MonthYear within the range and previous month, inclusive
 * @param startDate Start date
 * @param endDate End Date
 * @return string[] containing all months and years in "MM.YYYY" format
 */
export function getMonthYearInRange(startDate: Date, endDate: Date): string[] {
    const monthYear = [];
    let current: Date = new Date(startDate);
    current.setMonth(current.getMonth() - 1);
    while (current.getTime() <= endDate.getTime()) {
        monthYear.push(toMonthYear(current));
        current.setMonth(current.getMonth() + 1);
    }
    return monthYear;
}

export function shiftToLocal(date: Date): Date {
    return new Date(date.getTime() - date.getTimezoneOffset() * minute);
}

export function shiftToUTC(date: Date): Date {
    return new Date(date.getTime() + date.getTimezoneOffset() * minute);
}

export function toDate(dateString: string): Date {
    const [year, month, day] = dateString.split("-").map(Number);
    const date = new Date(year, month - 1, day);
    return date;
}

export function toMonthYear(current: Date): string {
    return current.getMonth() + 1 + "." + current.getFullYear();
}
