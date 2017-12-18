import * as FreeFormEventsSource from "./EventSources/FreeFormEventsSource";
import * as VSOCapacityEventSource from "./EventSources/VSOCapacityEventSource";
import * as VSOIterationEventSource from "./EventSources/VSOIterationEventSource";

// Register the eventSource contributions using the fully qualified contributionId
var context = VSS.getExtensionContext();
VSS.register(context.publisherId + "." + context.extensionId + "." + "freeForm", function (context) {
    return new FreeFormEventsSource.FreeFormEventsSource(context);
});
VSS.register(context.publisherId + "." + context.extensionId + "." + "daysOff", function (context) {
    return new VSOCapacityEventSource.VSOCapacityEventSource(context);
});
VSS.register(context.publisherId + "." + context.extensionId + "." + "iterations", function (context) {
    return new VSOIterationEventSource.VSOIterationEventSource(context);
});

// Deprecated form of contribution registration using the short Id for back-compat.
// DO NOT register this form any longer, use the fully qualified contributionId
//  unless you have previous shipped dependencies on the short Id's
VSS.register("freeForm", function (context) {
    return new FreeFormEventsSource.FreeFormEventsSource(context);
});
VSS.register("daysOff", function (context) {
    return new VSOCapacityEventSource.VSOCapacityEventSource(context);
});
VSS.register("iterations", function (context) {
    return new VSOIterationEventSource.VSOIterationEventSource(context);
});