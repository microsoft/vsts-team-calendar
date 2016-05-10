
 ({
     appDir: "built/debug",
     baseUrl: "./",
     dir: "built/min",
     paths: {
         "VSS": "empty:",
         "q": "empty:",
         "jQuery": "empty:",
         "TFS": "empty:"
     },
     modules: [
         {
             name: "Calendar/Extension"
         },
         {
             name: "Calendar/Dialogs"
         },
        //  {
        //      name: "Calendar/EventSources",
        //      include: [
        //          "Calendar/EventSources/FreeFormEventsSource",
        //          "Calendar/EventSources/VSOCapacityEventSource",
        //          "Calendar/EventSources/VSOIterationEventSource"
        //      ]
        //  }
         {
             name: "Calendar/EventSources/FreeFormEventsSource"
         },
         {
             name: "Calendar/EventSources/VSOCapacityEventSource"
         },
         {
             name: "Calendar/EventSources/VSOIterationEventSource"
         },
     ],
 })