/**
 * @jest-environment jsdom
 */
import {
    formatDate,
    getDatesInRange,
    getMonthYearInRange,
    monthAndYearToString,
    shiftToLocal,
    shiftToUTC,
    toDate,
    getUserLocale,
    formatDateLocalized,
    getLocalizedDateFormat
} from "../src/TimeLib";

describe("formatDate", () => {
    const date = new Date(2024, 2, 5); // March 5, 2024 (month is 0-indexed)

    it("formats as YYYY-MM-DD with zero-padded month and day", () => {
        expect(formatDate(date, "YYYY-MM-DD")).toBe("2024-03-05");
    });

    it("formats as MM/DD/YYYY with zero-padded month and day", () => {
        expect(formatDate(date, "MM/DD/YYYY")).toBe("03/05/2024");
    });

    it("formats as full month name and day", () => {
        expect(formatDate(date, "MONTH-DD")).toBe("March 5");
    });

    it("formats as MM-YYYY without zero-padding", () => {
        expect(formatDate(date, "MM-YYYY")).toBe("3.2024");
    });

    it("falls back to ISO string when no known format is given", () => {
        expect(formatDate(date)).toBe(date.toISOString());
    });
});

describe("getDatesInRange", () => {
    it("returns a single date when start and end are the same day", () => {
        const day = new Date(2024, 0, 1);
        const result = getDatesInRange(day, new Date(2024, 0, 1));
        expect(result).toHaveLength(1);
        expect(formatDate(result[0], "YYYY-MM-DD")).toBe("2024-01-01");
    });

    it("returns every date inclusive of start and end", () => {
        const start = new Date(2024, 0, 1);
        const end = new Date(2024, 0, 4);
        const result = getDatesInRange(start, end);
        expect(result.map((d) => formatDate(d, "YYYY-MM-DD"))).toEqual([
            "2024-01-01",
            "2024-01-02",
            "2024-01-03",
            "2024-01-04"
        ]);
    });

    it("returns an empty array when the start date is after the end date", () => {
        const start = new Date(2024, 0, 5);
        const end = new Date(2024, 0, 1);
        expect(getDatesInRange(start, end)).toEqual([]);
    });

    it("correctly spans a month boundary", () => {
        const start = new Date(2024, 0, 30);
        const end = new Date(2024, 1, 2);
        const result = getDatesInRange(start, end);
        expect(result.map((d) => formatDate(d, "YYYY-MM-DD"))).toEqual([
            "2024-01-30",
            "2024-01-31",
            "2024-02-01",
            "2024-02-02"
        ]);
    });
});

describe("getMonthYearInRange", () => {
    it("includes the month prior to the start date through the end month", () => {
        const start = new Date(2024, 2, 1); // March 2024
        const end = new Date(2024, 4, 31); // May 2024
        expect(getMonthYearInRange(start, end)).toEqual(["2.2024", "3.2024", "4.2024", "5.2024"]);
    });

    it("handles a year boundary", () => {
        const start = new Date(2024, 0, 1); // January 2024
        const end = new Date(2024, 1, 29); // February 2024 (leap year)
        expect(getMonthYearInRange(start, end)).toEqual(["12.2023", "1.2024", "2.2024"]);
    });
});

describe("monthAndYearToString", () => {
    it("combines the month name and year", () => {
        expect(monthAndYearToString({ month: 0, year: 2024 })).toBe("January 2024");
        expect(monthAndYearToString({ month: 11, year: 1999 })).toBe("December 1999");
    });
});

describe("shiftToLocal / shiftToUTC", () => {
    const original = new Date(2024, 5, 15, 10, 30, 0);
    const offsetMs = original.getTimezoneOffset() * 60 * 1000;

    it("shiftToUTC adds the timezone offset", () => {
        expect(shiftToUTC(original).getTime()).toBe(original.getTime() - offsetMs);
    });

    it("shiftToLocal subtracts the timezone offset", () => {
        expect(shiftToLocal(original).getTime()).toBe(original.getTime() + offsetMs);
    });

    it("round-trips a date through UTC and back to local without changing its value", () => {
        const roundTripped = shiftToLocal(shiftToUTC(original));
        expect(roundTripped.getTime()).toBe(original.getTime());
    });
});

describe("toDate", () => {
    it("parses a YYYY-MM-DD string into a local Date at midnight", () => {
        const result = toDate("2024-03-05");
        expect(result.getFullYear()).toBe(2024);
        expect(result.getMonth()).toBe(2); // 0-indexed
        expect(result.getDate()).toBe(5);
        expect(result.getHours()).toBe(0);
    });
});

describe("getUserLocale", () => {
    const originalLanguage = Object.getOwnPropertyDescriptor(window.navigator, "language");

    afterEach(() => {
        if (originalLanguage) {
            Object.defineProperty(window.navigator, "language", originalLanguage);
        }
    });

    it("returns the browser-reported locale", () => {
        Object.defineProperty(window.navigator, "language", { value: "de-DE", configurable: true });
        expect(getUserLocale()).toBe("de-DE");
    });

    it("falls back to en-US when navigator.language is empty", () => {
        Object.defineProperty(window.navigator, "language", { value: "", configurable: true });
        expect(getUserLocale()).toBe("en-US");
    });
});

describe("formatDateLocalized", () => {
    const originalLanguage = Object.getOwnPropertyDescriptor(window.navigator, "language");

    afterEach(() => {
        if (originalLanguage) {
            Object.defineProperty(window.navigator, "language", originalLanguage);
        }
        jest.restoreAllMocks();
    });

    it("formats the date according to the mocked locale", () => {
        Object.defineProperty(window.navigator, "language", { value: "en-US", configurable: true });
        const date = new Date(2024, 2, 5);
        expect(formatDateLocalized(date)).toBe("03/05/2024");
    });

    it("falls back to YYYY-MM-DD when Intl.DateTimeFormat throws", () => {
        Object.defineProperty(window.navigator, "language", { value: "en-US", configurable: true });
        const spy = jest.spyOn(Intl, "DateTimeFormat").mockImplementation(() => {
            throw new Error("Intl not supported");
        });
        const warnSpy = jest.spyOn(console, "warn").mockImplementation(() => undefined);

        const date = new Date(2024, 2, 5);
        expect(formatDateLocalized(date)).toBe("2024-03-05");
        expect(warnSpy).toHaveBeenCalled();

        spy.mockRestore();
        warnSpy.mockRestore();
    });
});

describe("getLocalizedDateFormat", () => {
    const originalLanguage = Object.getOwnPropertyDescriptor(window.navigator, "language");

    afterEach(() => {
        if (originalLanguage) {
            Object.defineProperty(window.navigator, "language", originalLanguage);
        }
    });

    it("returns the exact match format for a known locale", () => {
        Object.defineProperty(window.navigator, "language", { value: "de-DE", configurable: true });
        expect(getLocalizedDateFormat()).toBe("dd.MM.yyyy");
    });

    it("falls back to the language-only format when region is unknown", () => {
        Object.defineProperty(window.navigator, "language", { value: "fr-XX", configurable: true });
        expect(getLocalizedDateFormat()).toBe("dd/MM/yyyy");
    });

    it("defaults to MM/dd/yyyy for a completely unknown locale", () => {
        Object.defineProperty(window.navigator, "language", { value: "xx-XX", configurable: true });
        expect(getLocalizedDateFormat()).toBe("MM/dd/yyyy");
    });
});
