import React = require("react");

import { Button } from "azure-devops-ui/Button";
import { Card } from "azure-devops-ui/Card";
import { Link } from "azure-devops-ui/Link";
import { IListItemDetails, ListItem, ScrollableList } from "azure-devops-ui/List";
import { Observer } from "azure-devops-ui/Observer";
import { Surface, SurfaceBackground } from "azure-devops-ui/Surface";

import { IEventCategory } from "./Contracts";
import { FreeFormEventsSource } from "./FreeFormEventSource";
import { VSOCapacityEventSource } from "./VSOCapacityEventSource";

interface ISummaryComponentProps {
    /**
     * Object that stores all event data
     */
    capacityEventSource: VSOCapacityEventSource;

    /**
     * Object that stores all event data
     */
    freeFormEventSource: FreeFormEventsSource;

    /**
     * Callback to open edit dialog for days off
     */
    onEditDaysOff?: (event: import("./Contracts").ICalendarEvent) => void;

    /**
     * Callback to open edit dialog for event
     */
    onEditEvent?: (eventId: string) => void;

    /**
     * Callback to toggle pane visibility
     */
    onTogglePane?: () => void;
}

export class SummaryComponent extends React.Component<ISummaryComponentProps> {
    constructor(props: ISummaryComponentProps) {
        super(props);
    }

    public render(): JSX.Element {
        return (
            <div className="summary-area">
                <Surface background={SurfaceBackground.neutral}>
                    <div className="summary-header">
                    <div className="summary-title">Calendar Summary</div>
                    <Button
                        iconProps={{ iconName: "DoubleChevronRight" }}
                        onClick={this.props.onTogglePane}
                        tooltipProps={{ text: "Close pane" }}
                        ariaLabel="Close pane"
                        subtle
                    />
                </div>
                <div className="summary-content">
                    <Card className="category-card">
                        <div className="category-section">
                            <Observer url={this.props.capacityEventSource.getIterationUrl()}>
                                {(props: { url: string }) => {
                                    return (
                                        <Link className="category-heading" href={props.url} key={props.url} target="_blank">
                                            Iterations
                                        </Link>
                                    );
                                }}
                            </Observer>
                            <Observer iterationSummaryData={this.props.capacityEventSource.getIterationSummaryData()}>
                                {(props: { iterationSummaryData: IEventCategory[] }) => {
                                    return props.iterationSummaryData.length === 0 ? (
                                        <div className="empty-message">No iterations</div>
                                    ) : (
                                        <ScrollableList
                                            itemProvider={this.props.capacityEventSource.getIterationSummaryData()}
                                            renderRow={this.renderRow}
                                            width="100%"
                                        />
                                    );
                                }}
                            </Observer>
                        </div>
                    </Card>
                    <Card className="category-card">
                        <div className="category-section">
                            <Observer url={this.props.capacityEventSource.getCapacityUrl()}>
                                {(props: { url: string }) => {
                                    return (
                                        <Link className="category-heading" href={props.url} key={props.url} target="_blank">
                                            Days off
                                        </Link>
                                    );
                                }}
                            </Observer>
                            <Observer capacitySummaryData={this.props.capacityEventSource.getCapacitySummaryData()}>
                                {(props: { capacitySummaryData: IEventCategory[] }) => {
                                    return props.capacitySummaryData.length === 0 ? (
                                        <div className="empty-message">No days off</div>
                                    ) : (
                                        <ScrollableList
                                            itemProvider={this.props.capacityEventSource.getCapacitySummaryData()}
                                            renderRow={this.renderRow}
                                            width="100%"
                                        />
                                    );
                                }}
                            </Observer>
                        </div>
                    </Card>
                    <Card className="category-card">
                        <div className="category-section">
                            <div className="category-heading category-heading-static">Events</div>
                            <Observer eventSummaryData={this.props.freeFormEventSource.getSummaryData()}>
                                {(props: { eventSummaryData: IEventCategory[] }) => {
                                    return props.eventSummaryData.length === 0 ? (
                                        <div className="empty-message">No events</div>
                                    ) : (
                                        <ScrollableList itemProvider={this.props.freeFormEventSource.getSummaryData()} renderRow={this.renderRow} width="100%" />
                                    );
                                }}
                            </Observer>
                        </div>
                    </Card>
                </div>
                </Surface>
            </div>
        );
    }

    private renderRow = (index: number, item: IEventCategory, details: IListItemDetails<IEventCategory>, key?: string): JSX.Element => {
        const handleClick = () => {
            if (item.url) {
                // For iterations - open in new tab
                window.open(item.url, "_blank");
            } else if (item.linkedEvent && item.linkedEvent.id && this.props.onEditEvent) {
                // For free-form events - open event edit dialog (events have an id property)
                this.props.onEditEvent(item.linkedEvent.id);
            } else if (item.linkedEvent && this.props.onEditDaysOff) {
                // For days off - open days off edit dialog (days off events typically don't have id)
                this.props.onEditDaysOff(item.linkedEvent);
            }
        };

        const isClickable = item.url || item.linkedEvent;

        return (
            <ListItem key={key || "list-item" + index} index={index} details={details}>
                <div 
                    className="catagory-summary-row flex-row h-scroll-hidden" 
                    style={{ cursor: isClickable ? "pointer" : "default" }}
                    onClick={isClickable ? handleClick : undefined}
                >
                    {item.imageUrl && <img alt="" className="category-icon" src={item.imageUrl} />}
                    {!item.imageUrl && item.color && <div className="category-color" style={{ backgroundColor: item.color }} />}
                    <div className="flex-column h-scroll-hidden catagory-data">
                        <div className="category-titletext">{item.title}</div>
                        <div className="category-subtitle">{item.subTitle}</div>
                    </div>
                </div>
            </ListItem>
        );
    };
}
