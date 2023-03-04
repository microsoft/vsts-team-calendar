import React = require("react");

import { Link } from "azure-devops-ui/Link";
import { IListItemDetails, ListItem, ScrollableList } from "azure-devops-ui/List";
import { Observer } from "azure-devops-ui/Observer";

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
}

export class SummaryComponent extends React.Component<ISummaryComponentProps> {
    constructor(props: ISummaryComponentProps) {
        super(props);
    }

    public render(): JSX.Element {
        return (
            <div className="summary-area">
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
                            <div className="empty">(None)</div>
                        ) : (
                            <ScrollableList
                                itemProvider={this.props.capacityEventSource.getIterationSummaryData()}
                                renderRow={this.renderRow}
                                width="100%"
                            />
                        );
                    }}
                </Observer>
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
                            <div className="empty">(None)</div>
                        ) : (
                            <ScrollableList
                                itemProvider={this.props.capacityEventSource.getCapacitySummaryData()}
                                renderRow={this.renderRow}
                                width="100%"
                            />
                        );
                    }}
                </Observer>
                <a className="category-heading">Event</a>
                <Observer eventSummaryData={this.props.freeFormEventSource.getSummaryData()}>
                    {(props: { eventSummaryData: IEventCategory[] }) => {
                        return props.eventSummaryData.length === 0 ? (
                            <div className="empty">(None)</div>
                        ) : (
                            <ScrollableList itemProvider={this.props.freeFormEventSource.getSummaryData()} renderRow={this.renderRow} width="100%" />
                        );
                    }}
                </Observer>
            </div>
        );
    }

    private renderRow = (index: number, item: IEventCategory, details: IListItemDetails<IEventCategory>, key?: string): JSX.Element => {
        return (
            <ListItem key={key || "list-item" + index} index={index} details={details}>
                <div className="catagory-summary-row flex-row h-scroll-hidden">
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
