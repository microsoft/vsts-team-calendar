import React = require("react");

import { Button } from "azure-devops-ui/Button";
import { Card } from "azure-devops-ui/Card";
import { Icon } from "azure-devops-ui/Icon";
import { Link } from "azure-devops-ui/Link";
import { Observer } from "azure-devops-ui/Observer";
import { Surface, SurfaceBackground } from "azure-devops-ui/Surface";

import { IEventCategory } from "./Contracts";
import { FreeFormEventsSource } from "./FreeFormEventSource";
import { formatDateLocalized } from "./TimeLib";
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
     * Custom color map for event categories
     */
    eventColorMap?: Map<string, string>;

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

interface ISummaryComponentState {
    showAllIterations: boolean;
    showAllDaysOff: boolean;
    showAllEvents: boolean;
    expandedCategories: Set<string>;
}

export class SummaryComponent extends React.Component<ISummaryComponentProps, ISummaryComponentState> {
    private contentRef = React.createRef<HTMLDivElement>();
    private resizeObserver: ResizeObserver | null = null;

    constructor(props: ISummaryComponentProps) {
        super(props);
        this.state = {
            showAllIterations: false,
            showAllDaysOff: false,
            showAllEvents: false,
            expandedCategories: new Set<string>()
        };
    }

    componentWillUnmount() {
        if (this.resizeObserver && this.contentRef.current) {
            this.resizeObserver.unobserve(this.contentRef.current);
            this.resizeObserver.disconnect();
        }
    }

    public render(): JSX.Element {
        return (
            <div className="summary-area">
                <Surface background={SurfaceBackground.neutral}>
                    <div className="summary-header">
                        <div className="summary-title">Calendar Events</div>
                        {this.props.onTogglePane && (
                            <Button
                                ariaLabel="Close panel"
                                iconProps={{ iconName: "Cancel" }}
                                onClick={this.props.onTogglePane}
                                subtle
                                tooltipProps={{ text: "Close panel" }}
                            />
                        )}
                    </div>
                    <div className="summary-content" ref={this.contentRef}>
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
                                    if (props.iterationSummaryData.length === 0) {
                                        return <div className="empty-message">No iterations</div>;
                                    }

                                    // Sort by start date (earliest first)
                                    const displayData = [...props.iterationSummaryData].sort((a, b) => {
                                        const dateA = a.linkedEvent?.startDate ? new Date(a.linkedEvent.startDate).getTime() : 0;
                                        const dateB = b.linkedEvent?.startDate ? new Date(b.linkedEvent.startDate).getTime() : 0;
                                        return dateA - dateB;
                                    });

                                    return (
                                        <>
                                            {displayData.map((item, index) => this.renderSimpleRow(item, index))}
                                        </>
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
                                    if (props.capacitySummaryData.length === 0) {
                                        return <div className="empty-message">No days off</div>;
                                    }
                                    // Sort by start date (earliest first)
                                    const displayData = [...props.capacitySummaryData].sort((a, b) => {
                                        const dateA = a.linkedEvent?.startDate ? new Date(a.linkedEvent.startDate).getTime() : 0;
                                        const dateB = b.linkedEvent?.startDate ? new Date(b.linkedEvent.startDate).getTime() : 0;
                                        return dateA - dateB;
                                    });
                                    return (
                                        <>
                                            {displayData.map((item, index) => this.renderSimpleRow(item, index))}
                                        </>
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
                                    if (props.eventSummaryData.length === 0) {
                                        return <div className="empty-message">No events</div>;
                                    }
                                    // Sort by start date (earliest first)
                                    const displayData = [...props.eventSummaryData].sort((a, b) => {
                                        const dateA = a.linkedEvent?.startDate ? new Date(a.linkedEvent.startDate).getTime() : 0;
                                        const dateB = b.linkedEvent?.startDate ? new Date(b.linkedEvent.startDate).getTime() : 0;
                                        return dateA - dateB;
                                    });
                                    return (
                                        <>
                                            {displayData.map((item, index) => this.renderSimpleRow(item, index))}
                                        </>
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

    private getDisplayColor = (item: IEventCategory): string | undefined => {
        // Check if there's a custom color for this category
        if (this.props.eventColorMap && item.title) {
            const customColor = this.props.eventColorMap.get(item.title);
            if (customColor) {
                return customColor;
            }
        }
        // Fall back to the default color
        return item.color;
    };

    private renderSimpleRow = (item: IEventCategory, index: number): JSX.Element => {
        const hasMultipleEvents = item.eventCount > 1 && item.linkedEvents && item.linkedEvents.length > 1;
        const isExpanded = this.state.expandedCategories.has(item.title);

        const handleClick = () => {
            if (hasMultipleEvents) {
                // Toggle expansion
                const newExpanded = new Set(this.state.expandedCategories);
                if (isExpanded) {
                    newExpanded.delete(item.title);
                } else {
                    newExpanded.add(item.title);
                }
                this.setState({ expandedCategories: newExpanded });
            } else if (item.url) {
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

        const displayColor = this.getDisplayColor(item);

        return (
            <>
                <div
                    key={index}
                    className="catagory-summary-row flex-row h-scroll-hidden"
                    style={{ cursor: "pointer", padding: "2px 8px 2px 8px", alignItems: "center" }}
                    onClick={handleClick}
                >
                    {item.imageUrl && <img alt="" className="category-icon" src={item.imageUrl} />}
                    {!item.imageUrl && displayColor && <div className="category-color" style={{ backgroundColor: displayColor }} />}
                    <div className="flex-column h-scroll-hidden catagory-data" style={{ flex: 1 }}>
                        <div className="category-titletext">{item.title}</div>
                        <div className="category-subtitle">{item.subTitle}</div>
                    </div>
                    {hasMultipleEvents && (
                        <div style={{ marginLeft: "8px", fontSize: "10px", lineHeight: "1" }}>
                            {isExpanded ? "▲" : "▼"}
                        </div>
                    )}
                </div>
                {isExpanded && hasMultipleEvents && [...item.linkedEvents!].sort((a, b) => {
                    const dateA = new Date(a.startDate).getTime();
                    const dateB = new Date(b.startDate).getTime();
                    return dateA - dateB;
                }).map((event, eventIndex) => (
                    <div
                        key={`${index}-event-${eventIndex}`}
                        className="catagory-summary-row flex-row h-scroll-hidden"
                        style={{
                            cursor: "pointer",
                            padding: "2px 8px 2px 40px",
                            backgroundColor: "rgba(0, 0, 0, 0.04)",
                            alignItems: "center"
                        }}
                        onClick={() => {
                            if (event.id && this.props.onEditEvent) {
                                // Free-form event with id
                                this.props.onEditEvent(event.id);
                            } else if (!event.id && this.props.onEditDaysOff) {
                                // Days off event without id
                                this.props.onEditDaysOff(event);
                            }
                        }}
                    >
                        <div style={{ 
                            fontSize: "16px", 
                            marginRight: "8px", 
                            opacity: 0.5,
                            lineHeight: "1",
                            alignSelf: "center"
                        }}>•</div>
                        <div className="flex-column h-scroll-hidden catagory-data">
                            <div className="category-titletext" style={{ fontSize: "13px", opacity: 1 }}>{event.title}</div>
                            <div className="category-subtitle" style={{ fontSize: "11px", opacity: 0.9 }}>
                                {formatDateLocalized(new Date(event.startDate))} - {formatDateLocalized(new Date(event.endDate))}
                            </div>
                        </div>
                    </div>
                ))}
            </>
        );
    };
}
