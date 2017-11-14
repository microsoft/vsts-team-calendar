import * as Utils_String from "VSS/Utils/String";

const allColors = [
    "0072C6",
    "4617B4",
    "8C0095",
    "008A17",
    "D24726",
    "008299",
    "AC193D",
    "DC4FAD",
    "FF8F32",
    "82BA00",
    "03B3B2",
    "5DB2FF",
];

const allColorsOld = [
    "D6252E" /*red*/,
    "4AB63F" /*green*/,
    "40BD95" /*teal*/,
    "859A52" /*olive*/,
    "3267B8" /*blue*/,
    "613DB4" /*purple*/,
    "A34E78" /*maroon*/,
    "C4CCDD" /*steal*/,
    "8C9CBD" /*dark steal*/,
    "AF1E25" /*dark red*/,
    "B14F0D" /*dark orange*/,
    "AB7B05" /*dark peach*/,
    "999400" /*dark yellow*/,
    "35792B" /*dark green*/,
    "2E7D64" /*dark teal*/,
    "5F6C3A" /*dark olive*/,
    "2A5191" /*dark blue*/,
    "50328F" /*dark purple*/,
    "82375F" /*dark maroon*/,
];

const daysOffColor = "#F06C15"; /*orange*/
const nonWorkingDayColor = "#F5F5F5";
const currentIterationColor = "#C1E6FF"; /*dark gray*/

/**
 * Generates a color from the specified name
 * @param name String value used to generate a color
 * @return RGB color in the form of #RRGGBB
 */
export function generateColor(name: string): string {
    const id = name.slice().toLowerCase();
    if (id === "daysoff") {
        return daysOffColor;
    }
    let value = 0;
    for (let i = 0; i < (id || "").length; i++) {
        value += id.charCodeAt(i) * (i + 1);
    }

    return Utils_String.format("#{0}", allColors[value % allColors.length]);
}

/**
 * Generates a background color from the specified name
 * @param name String value used to generate a color
 * @return RGB color in the form of #RRGGBB
 */
export function generateBackgroundColor(name: string): string {
    return currentIterationColor;
}
