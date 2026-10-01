import { shiftToUTC } from "../src/TimeLib";
import type { IExtensionDataManager } from "azure-devops-extension-api";

// FreeFormEventSource pulls in azure-devops-ui and @fullcalendar packages purely for
// type declarations and unrelated (non-persistence) functionality. Those packages ship
// as ES modules that Jest/ts-jest cannot parse out of the box, so we stub them here -
// none of the mocked members are exercised by the addEvent/deleteEvent tests below.
jest.mock("azure-devops-ui/Core/Observable", () => ({
    ObservableArray: class ObservableArray<T> {
        constructor(_initial?: T[]) {}
    }
}));

import { FreeFormEventsSource } from "../src/FreeFormEventSource";
import type { ICalendarEvent } from "../src/Contracts";

/**
 * Builds a minimal, strongly-typed mock of IExtensionDataManager, only implementing the
 * methods FreeFormEventsSource actually depends on for the scenarios below. Using
 * jest.Mocked/MockedFunction (rather than an untyped `jest.fn()` bag cast with
 * `as unknown as`) ensures mock call signatures and return types stay checked against
 * the real interface.
 */
type MockDataManager = Pick<jest.Mocked<IExtensionDataManager>, "createDocument" | "deleteDocument">;

function createMockDataManager(overrides: Partial<MockDataManager> = {}): IExtensionDataManager {
    const manager: MockDataManager = {
        createDocument: jest.fn(),
        deleteDocument: jest.fn(),
        ...overrides
    };
    return manager as unknown as IExtensionDataManager;
}

describe("FreeFormEventsSource.addEvent", () => {
    const startDate = new Date(2024, 2, 1);
    const endDate = new Date(2024, 2, 2);

    function createSource(dataManager: IExtensionDataManager): FreeFormEventsSource {
        const source = new FreeFormEventsSource();
        source.dataManager = dataManager;
        source.selectedTeamId = "team-1";
        return source;
    }

    it("saves the event with UTC-shifted dates, caches it, and normalizes dates from the server response", async () => {
        // The server intentionally returns different date values than what was sent, so this
        // test actually exercises the post-response normalization logic (overwriting
        // addedEvent.startDate/endDate with the locally computed UTC values) instead of trivially
        // echoing back values that were already correct.
        const createDocument: jest.MockedFunction<IExtensionDataManager["createDocument"]> = jest
            .fn()
            .mockImplementation((_collection: string, event: ICalendarEvent) =>
                Promise.resolve({ ...event, id: "event-1", startDate: "2000-01-01T00:00:00.000Z", endDate: "2000-01-02T00:00:00.000Z" })
            );
        const source = createSource(createMockDataManager({ createDocument }));

        const result = await source.addEvent("Launch", startDate, endDate, "Release", "Launch day");

        expect(createDocument).toHaveBeenCalledWith(
            "team-1.3.2024",
            expect.objectContaining({
                title: "Launch",
                category: "Release",
                description: "Launch day",
                startDate: shiftToUTC(startDate).toISOString(),
                endDate: shiftToUTC(endDate).toISOString()
            })
        );
        expect(result.id).toBe("event-1");
        // The server's (stale/incorrect) dates must be overwritten with the locally-shifted ones.
        expect(result.startDate).toBe(shiftToUTC(startDate).toISOString());
        expect(result.endDate).toBe(shiftToUTC(endDate).toISOString());
        expect(source.getCategories().has("Release")).toBe(true);
        // The added event must be cached under its returned id for later lookup/deletion.
        expect(source.eventMap["event-1"]).toBe(result);
    });

    it("does not track 'Uncategorized' as a selectable category", async () => {
        const createDocument: jest.MockedFunction<IExtensionDataManager["createDocument"]> = jest
            .fn()
            .mockImplementation((_collection: string, event: ICalendarEvent) => Promise.resolve({ ...event, id: "event-2" }));
        const source = createSource(createMockDataManager({ createDocument }));

        await source.addEvent("Misc", startDate, endDate, "Uncategorized", "");

        expect(source.getCategories().has("Uncategorized")).toBe(false);
    });

    it("raises a friendly error when the event already exists", async () => {
        const createDocument: jest.MockedFunction<IExtensionDataManager["createDocument"]> = jest.fn().mockRejectedValue({
            status: 400,
            responseText: "DocumentExistsException: conflict"
        });
        const source = createSource(createMockDataManager({ createDocument }));

        await expect(source.addEvent("Launch", startDate, endDate, "Release", "")).rejects.toThrow(
            'Event "Launch" already exists or there was a conflict creating it.'
        );
    });

    it("propagates a 400 error unchanged when it is not a DocumentExistsException", async () => {
        // Guards the compound condition in the source: status === 400 alone is not enough to
        // trigger the friendly message - responseText must also mention DocumentExistsException.
        const failure = { status: 400, responseText: "ValidationException: bad payload" };
        const createDocument: jest.MockedFunction<IExtensionDataManager["createDocument"]> = jest.fn().mockRejectedValue(failure);
        const source = createSource(createMockDataManager({ createDocument }));

        await expect(source.addEvent("Launch", startDate, endDate, "Release", "")).rejects.toBe(failure);
    });

    it("propagates unexpected errors from the data manager unchanged", async () => {
        const failure = { status: 500, responseText: "server error" };
        const createDocument: jest.MockedFunction<IExtensionDataManager["createDocument"]> = jest.fn().mockRejectedValue(failure);
        const source = createSource(createMockDataManager({ createDocument }));

        await expect(source.addEvent("Launch", startDate, endDate, "Release", "")).rejects.toBe(failure);
    });
});

describe("FreeFormEventsSource.deleteEvent", () => {
    const startDate = new Date(2024, 2, 1);

    function createSource(dataManager: IExtensionDataManager): FreeFormEventsSource {
        const source = new FreeFormEventsSource();
        source.dataManager = dataManager;
        source.selectedTeamId = "team-1";
        source.eventMap["event-1"] = { id: "event-1" } as ICalendarEvent;
        return source;
    }

    it("removes the event from the local cache and deletes it remotely", async () => {
        const deleteDocument: jest.MockedFunction<IExtensionDataManager["deleteDocument"]> = jest.fn().mockResolvedValue(undefined);
        const source = createSource(createMockDataManager({ deleteDocument }));

        await source.deleteEvent("event-1", startDate);

        expect(source.eventMap["event-1"]).toBeUndefined();
        expect(deleteDocument).toHaveBeenCalledWith("team-1.3.2024", "event-1");
    });

    it("treats a 404 from the server as a successful deletion", async () => {
        const deleteDocument: jest.MockedFunction<IExtensionDataManager["deleteDocument"]> = jest
            .fn()
            .mockRejectedValue({ status: 404 });
        const source = createSource(createMockDataManager({ deleteDocument }));

        await expect(source.deleteEvent("event-1", startDate)).resolves.toBeUndefined();
    });

    it("propagates non-404 errors from the data manager", async () => {
        const failure = { status: 500 };
        const deleteDocument: jest.MockedFunction<IExtensionDataManager["deleteDocument"]> = jest.fn().mockRejectedValue(failure);
        const source = createSource(createMockDataManager({ deleteDocument }));

        await expect(source.deleteEvent("event-1", startDate)).rejects.toBe(failure);
    });
});

