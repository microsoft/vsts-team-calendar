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

interface ISummaryComponentState {
    showAllIterations: boolean;
    showAllDaysOff: boolean;
    showAllEvents: boolean;
    itemsToShow: number;
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
            itemsToShow: 5 // Default fallback
        };
    }

    componentDidMount() {
        this.calculateItemsToShow();
        if (this.contentRef.current) {
            this.resizeObserver = new ResizeObserver(() => {
                this.calculateItemsToShow();
            });
            this.resizeObserver.observe(this.contentRef.current);
        }
    }

    componentWillUnmount() {
        if (this.resizeObserver && this.contentRef.current) {
            this.resizeObserver.unobserve(this.contentRef.current);
            this.resizeObserver.disconnect();
        }
    }

    private calculateItemsToShow = () => {
        if (!this.contentRef.current) return;
        
        const containerHeight = this.contentRef.current.clientHeight;
        // Simplified calculation: 50px per card overhead, 50px per item
        const cardOverhead = 50;
        const itemHeight = 50;
        const totalOverhead = cardOverhead * 3; // 3 cards
        const availableHeight = containerHeight - totalOverhead;
        
        // Items per section
        const itemsPerSection = Math.floor(availableHeight / itemHeight / 3);
        const itemsToShow = Math.max(5, Math.min(itemsPerSection, 25)); // Between 5 and 25 items per section
        
        if (itemsToShow !== this.state.itemsToShow) {
            this.setState({ itemsToShow });
        }
    };

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
                                    const hasMore = props.iterationSummaryData.length > this.state.itemsToShow;
                                    const displayData = this.state.showAllIterations 
                                        ? props.iterationSummaryData 
                                        : props.iterationSummaryData.slice(0, this.state.itemsToShow);
                                    return (
                                        <>
                                            {displayData.map((item, index) => this.renderSimpleRow(item, index))}
                                            {hasMore && !this.state.showAllIterations && (
                                                <Button
                                                    text={`Show ${props.iterationSummaryData.length - this.state.itemsToShow} more`}
                                                    onClick={() => this.setState({ showAllIterations: true })}
                                                    subtle
                                                    className="show-more-button"
                                                />
                                            )}
                                            {hasMore && this.state.showAllIterations && (
                                                <Button
                                                    text="Hide"
                                                    onClick={() => this.setState({ showAllIterations: false })}
                                                    subtle
                                                    className="show-more-button"
                                                />
                                            )}
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
                                    const hasMore = props.capacitySummaryData.length > this.state.itemsToShow;
                                    const displayData = this.state.showAllDaysOff 
                                        ? props.capacitySummaryData 
                                        : props.capacitySummaryData.slice(0, this.state.itemsToShow);
                                    return (
                                        <>
                                            {displayData.map((item, index) => this.renderSimpleRow(item, index))}
                                            {hasMore && !this.state.showAllDaysOff && (
                                                <Button
                                                    text={`Show ${props.capacitySummaryData.length - this.state.itemsToShow} more`}
                                                    onClick={() => this.setState({ showAllDaysOff: true })}
                                                    subtle
                                                    className="show-more-button"
                                                />
                                            )}
                                            {hasMore && this.state.showAllDaysOff && (
                                                <Button
                                                    text="Hide"
                                                    onClick={() => this.setState({ showAllDaysOff: false })}
                                                    subtle
                                                    className="show-more-button"
                                                />
                                            )}
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
                                    const hasMore = props.eventSummaryData.length > this.state.itemsToShow;
                                    const displayData = this.state.showAllEvents 
                                        ? props.eventSummaryData 
                                        : props.eventSummaryData.slice(0, this.state.itemsToShow);
                                    return (
                                        <>
                                            {displayData.map((item, index) => this.renderSimpleRow(item, index))}
                                            {hasMore && !this.state.showAllEvents && (
                                                <Button
                                                    text={`Show ${props.eventSummaryData.length - this.state.itemsToShow} more`}
                                                    onClick={() => this.setState({ showAllEvents: true })}
                                                    subtle
                                                    className="show-more-button"
                                                />
                                            )}
                                            {hasMore && this.state.showAllEvents && (
                                                <Button
                                                    text="Hide"
                                                    onClick={() => this.setState({ showAllEvents: false })}
                                                    subtle
                                                    className="show-more-button"
                                                />
                                            )}
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

    private renderSimpleRow = (item: IEventCategory, index: number): JSX.Element => {
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
            <div 
                key={index}
                className="catagory-summary-row flex-row h-scroll-hidden" 
                style={{ cursor: isClickable ? "pointer" : "default", padding: "8px 16px" }}
                onClick={isClickable ? handleClick : undefined}
            >
                {item.imageUrl && <img alt="" className="category-icon" src={item.imageUrl} />}
                {!item.imageUrl && item.color && <div className="category-color" style={{ backgroundColor: item.color }} />}
                <div className="flex-column h-scroll-hidden catagory-data">
                    <div className="category-titletext">{item.title}</div>
                    <div className="category-subtitle">{item.subTitle}</div>
                </div>
            </div>
        );
    };
}
