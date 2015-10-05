/// <reference path='../../../typings/vss/VSS.d.ts' />

import Calendar_Contracts = require("Calendar/Contracts");
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

var _iterations: Work_Contracts.TeamSettingsIteration[];
var _iterationsDeferred: Q.Deferred<any> = Q.defer<any>();
var _iterationsLoaded = _iterationsDeferred.promise;

export function getIterationId(dayOff: Date): IPromise<string> {
    var deferred = Q.defer<string>();
    if (!_iterations) {
        loadIterations();
    }
    _iterationsLoaded.then(() => {
        _iterations.some((value: Work_Contracts.TeamSettingsIteration, index: number, array: Work_Contracts.TeamSettingsIteration[]) => {
            if (value && value.attributes && value.attributes.startDate && value.attributes.finishDate) {
                if (dayOff >= Utils_Date.shiftToUTC(value.attributes.startDate) && dayOff <= Utils_Date.shiftToUTC(value.attributes.finishDate)) {
                    deferred.resolve(value.id);
                    return true;
                }
            }
            return false;
        });
    });
    return deferred.promise;
}

function loadIterations(): void {
    _iterations = [];
    var webContext = VSS.getWebContext();
    var teamContext: TFS_Core_Contracts.TeamContext = {projectId: webContext.project.id, teamId: webContext.team.id, project: "", team: ""};
    var workClient: Work_Client.WorkHttpClient = Service.VssConnection
        .getConnection()
        .getHttpClient(Work_Client.WorkHttpClient, WebApi_Constants.ServiceInstanceTypes.TFS);
    
    workClient.getTeamIterations(teamContext).then(
        (iterations: Work_Contracts.TeamSettingsIteration[]) => {
            iterations.forEach((iteration: Work_Contracts.TeamSettingsIteration, index: number, array: Work_Contracts.TeamSettingsIteration[]) => {
                _iterations.push(iteration);
            });

            _iterationsDeferred.resolve([]);

        },
        (e: Error) => {
            _iterationsDeferred.resolve([]);
        });
}