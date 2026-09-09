import { generateColor } from "../src/Color";

describe("generateColor", () => {
    it("returns the same color for the same name every time (deterministic)", () => {
        const first = generateColor("Alice");
        const second = generateColor("Alice");
        expect(first).toBe(second);
    });

    it("returns a valid hex color for a regular name", () => {
        expect(generateColor("Alice")).toMatch(/^#[0-9A-F]{6}$/);
    });

    it("is case-insensitive when generating colors for names", () => {
        expect(generateColor("Alice")).toBe(generateColor("alice"));
        expect(generateColor("BOB")).toBe(generateColor("bob"));
    });

    it("returns the reserved color for the current iteration", () => {
        expect(generateColor("currentIteration")).toBe("rgba(193, 230, 255, 0.7)");
    });

    it("deterministically selects an 'other iteration' color based on the iteration name", () => {
        const first = generateColor("otherIteration", "Sprint 1");
        const second = generateColor("otherIteration", "Sprint 1");
        expect(first).toBe(second);
        expect(["rgba(255, 218, 193, 0.7)", "rgba(230, 255, 193, 0.3)", "rgba(255, 193, 230, 0.3)"]).toContain(first);
    });

    it("falls back to the first 'other iteration' color when no iteration name is given", () => {
        expect(generateColor("otherIteration")).toBe("rgba(255, 218, 193, 0.7)");
    });

    it("selects a different color when the iteration name actually changes the hash", () => {
        // Verifies iterationName genuinely drives the selection (not just accepted and ignored).
        // Hash sums: "Sprint 1" -> 2931 (% 3 === 0 -> index 0), "Sprint 2" -> 2939 (% 3 === 2 -> index 2).
        const colorA = generateColor("otherIteration", "Sprint 1");
        const colorB = generateColor("otherIteration", "Sprint 2");
        expect(colorA).toBe("rgba(255, 218, 193, 0.7)");
        expect(colorB).toBe("rgba(255, 193, 230, 0.3)");
        expect(colorA).not.toBe(colorB);
    });
});
