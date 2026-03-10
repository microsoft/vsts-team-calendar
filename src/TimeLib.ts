const minute = 1000 * 60;

export const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
export interface MonthAndYear {
    month: number;
    year: number;
}

export function formatDate(date: Date, format?: string): string {
    if (format === "YYYY-MM-DD") {
        let month = "" + (date.getMonth() + 1);
        let day = "" + date.getDate();
        const year = date.getFullYear();

        if (month.length < 2) month = "0" + month;
        if (day.length < 2) day = "0" + day;

        return [year, month, day].join("-");
    } else if (format === "MM/DD/YYYY") {
        let month = "" + (date.getMonth() + 1);
        let day = "" + date.getDate();
        const year = date.getFullYear();

        if (month.length < 2) month = "0" + month;
        if (day.length < 2) day = "0" + day;

        return [month, day, year].join("/");
    } else if (format === "MONTH-DD") {
        return months[date.getMonth()] + " " + date.getDate();
    } else if (format === "MM-YYYY") {
        return date.getMonth() + 1 + "." + date.getFullYear();
    }
    return date.toISOString();
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
        monthYear.push(formatDate(current, "MM-YYYY"));
        current.setMonth(current.getMonth() + 1);
    }
    return monthYear;
}

export function monthAndYearToString(monthAndYear: MonthAndYear): string {
    return months[monthAndYear.month] + " " + monthAndYear.year;
}

/**
 * Create date which has same time as of UTC time in Local timezone
 * @param date to be converted
 */
export function shiftToLocal(date: Date): Date {
    return new Date(date.getTime() + date.getTimezoneOffset() * minute);
}

/**
 * Create date which has same time in UTC timezone
 * @param date to be converted
 */
export function shiftToUTC(date: Date): Date {
    return new Date(date.getTime() - date.getTimezoneOffset() * minute);
}

/**
 * Converts "YYYY-MM-DD" formated string to date
 * @param dateString date string to be converted
 */
export function toDate(dateString: string): Date {
    const [year, month, day] = dateString.split("-").map(Number);
    const date = new Date(year, month - 1, day);
    return date;
}

/**
 * Gets the user's locale from browser settings
 * @returns User's locale string (e.g., "en-US", "en-GB", "de-DE")
 */
export function getUserLocale(): string {
    return navigator.language || 'en-US';
}

/**
 * Formats a date using the user's locale settings
 * @param date Date to format
 * @param options Optional Intl.DateTimeFormat options
 * @returns Localized date string
 */
export function formatDateLocalized(date: Date, options?: Intl.DateTimeFormatOptions): string {
    const locale = getUserLocale();
    const defaultOptions: Intl.DateTimeFormatOptions = {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        ...options
    };
    
    try {
        return new Intl.DateTimeFormat(locale, defaultOptions).format(date);
    } catch (error) {
        // Fallback to ISO format if Intl is not available
        console.warn('Intl.DateTimeFormat not available, using fallback', error);
        return formatDate(date, "YYYY-MM-DD");
    }
}

/**
 * Gets the date format string for react-datepicker based on user's locale
 * @returns Format string compatible with react-datepicker
 */
export function getLocalizedDateFormat(): string {
    const locale = getUserLocale();
    
    // Common locale patterns for react-datepicker
    const localeFormats: { [key: string]: string } = {
        'en-US': 'MM/dd/yyyy',
        'en-GB': 'dd/MM/yyyy',
        'en-CA': 'dd/MM/yyyy',
        'en-AU': 'dd/MM/yyyy',
        'en-NZ': 'dd/MM/yyyy',
        'de': 'dd.MM.yyyy',
        'de-DE': 'dd.MM.yyyy',
        'de-AT': 'dd.MM.yyyy',
        'de-CH': 'dd.MM.yyyy',
        'fr': 'dd/MM/yyyy',
        'fr-FR': 'dd/MM/yyyy',
        'fr-CA': 'yyyy-MM-dd',
        'ja': 'yyyy/MM/dd',
        'ja-JP': 'yyyy/MM/dd',
        'zh': 'yyyy/MM/dd',
        'zh-CN': 'yyyy/MM/dd',
        'ko': 'yyyy.MM.dd',
        'ko-KR': 'yyyy.MM.dd',
        'es': 'dd/MM/yyyy',
        'es-ES': 'dd/MM/yyyy',
        'it': 'dd/MM/yyyy',
        'it-IT': 'dd/MM/yyyy',
        'pt': 'dd/MM/yyyy',
        'pt-BR': 'dd/MM/yyyy',
        'nl': 'dd-MM-yyyy',
        'nl-NL': 'dd-MM-yyyy',
        'sv': 'yyyy-MM-dd',
        'sv-SE': 'yyyy-MM-dd',
        'da': 'dd-MM-yyyy',
        'da-DK': 'dd-MM-yyyy',
        'nb': 'dd.MM.yyyy',
        'nb-NO': 'dd.MM.yyyy',
        'fi': 'dd.MM.yyyy',
        'fi-FI': 'dd.MM.yyyy',
        'pl': 'dd.MM.yyyy',
        'pl-PL': 'dd.MM.yyyy',
        'ru': 'dd.MM.yyyy',
        'ru-RU': 'dd.MM.yyyy'
    };
    
    // Try exact match first
    if (localeFormats[locale]) {
        return localeFormats[locale];
    }
    
    // Try language code without region
    const languageCode = locale.split('-')[0];
    if (localeFormats[languageCode]) {
        return localeFormats[languageCode];
    }
    
    // Default to MM/dd/yyyy for unknown locales
    return 'MM/dd/yyyy';
}
