import { WebApiTeam } from 'TFS/Core/Contracts';
import * as Calendar_Contracts from '../Contracts';
import * as Calendar_ColorUtils from '../Utils/Color';
import * as Calendar_DateUtils from '../Utils/Date';
import * as Service from 'VSS/Service';
import * as TFS_Core_Contracts from 'TFS/Core/Contracts';
import * as Utils_Date from 'VSS/Utils/Date';
import * as Utils_String from 'VSS/Utils/String';
import * as WebApi_Constants from 'VSS/WebApi/Constants';
import * as Work_Client from 'TFS/Work/RestClient';
import * as Work_Contracts from 'TFS/Work/Contracts';

export class VSOIterationEventSource implements Calendar_Contracts.IEventSource {
    public id = "iterations";
    public name = "Iterations";
    public order = 20;
    public background = true;
    private _events: Calendar_Contracts.CalendarEvent[];
    private _categories: Calendar_Contracts.IEventCategory[];
    private _teamId: string;

    constructor(context?: any) {
        this.updateTeamContext(context.team);
    }

    public updateTeamContext(newTeam: WebApiTeam) {
        this._teamId = newTeam.id;
    }

    public load(): PromiseLike<Calendar_Contracts.CalendarEvent[]> {
        return this.getEvents();
    }

    public getEvents(query?: Calendar_Contracts.IEventQuery): PromiseLike<Calendar_Contracts.CalendarEvent[]> {
        const result: Calendar_Contracts.CalendarEvent[] = [];
        this._events = null;

        const webContext = VSS.getWebContext();
        const teamContext: TFS_Core_Contracts.TeamContext = {
            projectId: webContext.project.id,
            teamId: this._teamId,
            project: "",
            team: "",
        };
        const workClient: Work_Client.WorkHttpClient2_1 = Service.VssConnection
            .getConnection()
            .getHttpClient(Work_Client.WorkHttpClient2_1, WebApi_Constants.ServiceInstanceTypes.TFS);

        // fetch the wit events
        return workClient.getTeamIterations(teamContext).then(iterations => {
            for (const iteration of iterations) {
                if (iteration && iteration.attributes && iteration.attributes.startDate) {
                    const event: any = {};
                    event.startDate = iteration.attributes.startDate.toISOString();
                    if (iteration.attributes.finishDate) {
                        event.endDate = iteration.attributes.finishDate.toISOString();
                    }

                    event.title = iteration.name;
                    const start = new Date(event.startDate);
                    const end = new Date(event.endDate);
                    const startAsUtc = new Date(
                        start.getUTCFullYear(),
                        start.getUTCMonth(),
                        start.getUTCDate(),
                        start.getUTCHours(),
                        start.getUTCMinutes(),
                        start.getUTCSeconds(),
                    );
                    const endAsUtc = new Date(
                        end.getUTCFullYear(),
                        end.getUTCMonth(),
                        end.getUTCDate(),
                        end.getUTCHours(),
                        end.getUTCMinutes(),
                        end.getUTCSeconds(),
                    );

                    event.category = <Calendar_Contracts.IEventCategory>{
                        id: this.id + "." + iteration.name,
                        title: iteration.name,
                        subTitle: Utils_String.format(
                            "{0} - {1}",
                            Utils_Date.format(startAsUtc, "M"),
                            Utils_Date.format(endAsUtc, "M"),
                        ),
                    };
                    if (this._isCurrentIteration(event)) {
                        event.category.color = Calendar_ColorUtils.generateBackgroundColor(event.title);
                    } else {
                        event.category.color = "#FFFFFF";
                    }

                    result.push(event);
                }
            }

            result.sort((a, b) => {
                return new Date(a.startDate).valueOf() - new Date(b.startDate).valueOf();
            });
            this._events = result;
            return result;
        });
    }

    public getCategories(query: Calendar_Contracts.IEventQuery): PromiseLike<Calendar_Contracts.IEventCategory[]> {
        if (this._events) {
            return Promise.resolve(this._getCategoryData(this._events.slice(0), query));
        } else {
            return this.getEvents().then((events: Calendar_Contracts.CalendarEvent[]) => {
                return this._getCategoryData(events, query);
            });
        }
    }

    public getTitleUrl(webContext: WebContext): PromiseLike<string> {
        return Promise.resolve(webContext.host.uri + webContext.project.name + "/_admin/_iterations");
    }

    private _getCategoryData(
        events: Calendar_Contracts.CalendarEvent[],
        query: Calendar_Contracts.IEventQuery,
    ): Calendar_Contracts.IEventCategory[] {
        return events
            .filter(e => Calendar_DateUtils.eventIn(e, query))
            .sort((a, b) => new Date(a.startDate || (0 as any)).getTime() - new Date(b.startDate || (0 as any)).getTime())
            .map(e => e.category);
    }

    private _isCurrentIteration(event: Calendar_Contracts.CalendarEvent): boolean {
        if (event.startDate && event.endDate) {
            const now = new Date();
            const today: number = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0)).valueOf();
            return today >= new Date(event.startDate).valueOf() && today <= new Date(event.endDate).valueOf();
        }
        return false;
    }
}
