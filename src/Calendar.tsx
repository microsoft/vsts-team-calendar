import "./Calendar.scss";

import * as React from "react";
import * as ReactDOM from "react-dom";

import { CommonServiceIds, IProjectPageService, getClient } from "azure-devops-extension-api";
import { IExtensionDataService, IExtensionDataManager, ILocationService, IHostNavigationService } from "azure-devops-extension-api/Common";
import { CoreRestClient, WebApiTeam } from "azure-devops-extension-api/Core";
import { TeamMember } from "azure-devops-extension-api/WebApi/WebApi";

import * as SDK from "azure-devops-extension-sdk";

import { Button } from "azure-devops-ui/Button";
import { Checkbox } from "azure-devops-ui/Checkbox";
import { Dropdown, DropdownExpandableButton } from "azure-devops-ui/Dropdown";
import { Toggle } from "azure-devops-ui/Toggle";
import { CustomHeader, HeaderTitleArea } from "azure-devops-ui/Header";
import { IHeaderCommandBarItem } from "azure-devops-ui/HeaderCommandBar";
import { IListBoxItem } from "azure-devops-ui/ListBox";
import { ContextualMenu, MenuItemType } from "azure-devops-ui/Menu";
import { ObservableValue } from "azure-devops-ui/Core/Observable";
import { Observer } from "azure-devops-ui/Observer";
import { Page } from "azure-devops-ui/Page";
import { Panel } from "azure-devops-ui/Panel";
import { Spinner, SpinnerSize } from "azure-devops-ui/Spinner";
import { TextField } from "azure-devops-ui/TextField";
import { Location } from "azure-devops-ui/Utilities/Position";

import { View, EventApi, Duration, Calendar } from "@fullcalendar/core";
import dayGridPlugin from "@fullcalendar/daygrid";
import interactionPlugin from "@fullcalendar/interaction";
import FullCalendar from "@fullcalendar/react";
import timeGridPlugin from "@fullcalendar/timegrid";

import { localeData } from "moment";

import { AddEditDaysOffDialog } from "./AddEditDaysOffDialog";
import { AddEditEventDialog } from "./AddEditEventDialog";
import { generateColor } from "./Color";
import { ICalendarEvent } from "./Contracts";
import { FreeFormId, FreeFormEventsSource } from "./FreeFormEventSource";
import { SummaryComponent } from "./SummaryComponent";
import { MonthAndYear, monthAndYearToString, formatDate } from "./TimeLib";
import { DaysOffId, VSOCapacityEventSource, IterationId } from "./VSOCapacityEventSource";

enum Dialogs {
    None,
    NewEventDialog,
    NewDaysOffDialog
}

class ExtensionContent extends React.Component {
    anchorElement: ObservableValue<HTMLElement | undefined> = new ObservableValue<HTMLElement | undefined>(undefined);
    calendarComponentRef = React.createRef<FullCalendar>();
    commandBarItems: IHeaderCommandBarItem[];
    currentMonthAndYear: ObservableValue<MonthAndYear>;
    dataManager: IExtensionDataManager | undefined;
    displayCalendar: ObservableValue<boolean>;
    eventApi?: EventApi;
    eventToEdit?: ICalendarEvent;
    freeFormEventSource: FreeFormEventsSource;
    hostUrl: string;
    locationService: ILocationService | undefined;
    members: TeamMember[];
    navigationService: IHostNavigationService | undefined;
    openDialog: ObservableValue<Dialogs> = new ObservableValue(Dialogs.None);
    isPaneOpen: ObservableValue<boolean> = new ObservableValue<boolean>(true);
    isColorPanelOpen: ObservableValue<boolean> = new ObservableValue<boolean>(false);
    isExpandedView: ObservableValue<boolean> = new ObservableValue<boolean>(true);
    eventColorMap: ObservableValue<Map<string, string>> = new ObservableValue<Map<string, string>>(new Map());
    tempColorMap: Map<string, string> = new Map();
    projectId: string;
    projectName: string;
    selectedEndDate: Date;
    selectedStartDate: Date;
    selectedTeamName: string;
    showMonthPicker: ObservableValue<boolean> = new ObservableValue<boolean>(false);
    sidePanelAnchorElement: ObservableValue<HTMLElement | undefined> = new ObservableValue<HTMLElement | undefined>(undefined);
    teams: ObservableValue<WebApiTeam[]>;
    vsoCapacityEventSource: VSOCapacityEventSource;

    constructor(props: {}) {
        super(props);

        this.currentMonthAndYear = new ObservableValue<MonthAndYear>({
            month: new Date().getMonth(),
            year: new Date().getFullYear()
        });

        this.state = {
            fullScreenMode: false,
            calendarWeekends: true,
            calendarEvents: []
        };

        // Navigation items for lower command bar
        this.commandBarItems = [
            {
                id: "today",
                important: true,
                onActivate: () => {
                    if (this.calendarComponentRef.current) {
                        this.getCalendarApi().today();
                        this.currentMonthAndYear.value = {
                            month: new Date().getMonth(),
                            year: new Date().getFullYear()
                        };
                    }
                },
                text: "Today"
            },
            {
                iconProps: {
                    iconName: "ChevronLeft"
                },
                important: true,
                id: "prev",
                onActivate: () => {
                    if (this.calendarComponentRef.current) {
                        this.getCalendarApi().prev();
                        this.currentMonthAndYear.value = this.calcMonths(this.currentMonthAndYear.value, -1);
                    }
                },
                text: "Prev"
            },
            {
                iconProps: {
                    iconName: "ChevronRight"
                },
                important: true,
                id: "next",
                onActivate: () => {
                    if (this.calendarComponentRef.current) {
                        this.getCalendarApi().next();
                        this.currentMonthAndYear.value = this.calcMonths(this.currentMonthAndYear.value, 1);
                    }
                },
                text: "Next"
            }
        ];

        this.selectedEndDate = new Date();
        this.selectedStartDate = new Date();
        this.teams = new ObservableValue<WebApiTeam[]>([]);
        this.selectedTeamName = "Select Team";
        this.displayCalendar = new ObservableValue<boolean>(false);
        this.projectId = "";
        this.projectName = "";
        this.hostUrl = "";
        this.members = [];
        this.freeFormEventSource = new FreeFormEventsSource();
        this.vsoCapacityEventSource = new VSOCapacityEventSource();
    }

    public render(): JSX.Element {
        return (
            <Page className="flex-grow flex-column bolt-page-grey">
                <Observer displayCalendar={this.displayCalendar}>
                    {(props: { displayCalendar: boolean }) => {
                        if (!props.displayCalendar) {
                            return (
                                <div style={{ 
                                    display: "flex", 
                                    justifyContent: "center", 
                                    alignItems: "center", 
                                    height: "100vh",
                                    width: "100%"
                                }}>
                                    <Spinner size={SpinnerSize.large} label="Loading Calendar..." />
                                </div>
                            );
                        }
                        return (
                            <>
                {/* Single header: All controls on one row - spans full width above calendar and panel */}
                <CustomHeader className="bolt-header-with-commandbar full-width-header">
                    <HeaderTitleArea>
                        <div style={{ display: "flex", alignItems: "center", width: "100%", justifyContent: "space-between" }}>
                            <div style={{ display: "flex", alignItems: "center", gap: "16px" }}>
                                <Observer teams={this.teams}>
                                    {(teamProps: { teams: WebApiTeam[] }) => {
                                        return teamProps.teams.length === 0 ? null : (
                                            <Dropdown
                                                items={this.getTeamPickerOptions()}
                                                onSelect={this.onSelectTeam}
                                                placeholder={this.selectedTeamName}
                                                renderExpandable={expandableProps => <DropdownExpandableButton {...expandableProps} />}
                                                showFilterBox={true}
                                                filterPlaceholderText="Filter teams"
                                                width={280}
                                            />
                                        );
                                    }}
                                </Observer>
                                <Observer currentMonthAndYear={this.currentMonthAndYear}>
                                    {(dateProps: { currentMonthAndYear: MonthAndYear }) => {
                                        return (
                                            <Dropdown
                                                items={this.getMonthPickerOptions()}
                                                key={dateProps.currentMonthAndYear.month}
                                                onSelect={this.onSelectMonthYear}
                                                placeholder={monthAndYearToString(dateProps.currentMonthAndYear)}
                                                renderExpandable={expandableProps => (
                                                    <DropdownExpandableButton {...expandableProps} />
                                                )}
                                                width={180}
                                            />
                                        );
                                    }}
                                </Observer>
                                <Button
                                    text="Today"
                                    onClick={() => {
                                        if (this.calendarComponentRef.current) {
                                            this.getCalendarApi().today();
                                            this.currentMonthAndYear.value = {
                                                month: new Date().getMonth(),
                                                year: new Date().getFullYear()
                                            };
                                        }
                                    }}
                                />
                                <div className="nav-button-group">
                                    <Button
                                        iconProps={{ iconName: "ChevronLeft" }}
                                        ariaLabel="Previous month"
                                        subtle
                                        tooltipProps={{ text: "Previous month" }}
                                        onClick={() => {
                                            if (this.calendarComponentRef.current) {
                                                this.getCalendarApi().prev();
                                                this.currentMonthAndYear.value = this.calcMonths(this.currentMonthAndYear.value, -1);
                                            }
                                        }}
                                    />
                                    <Button
                                        iconProps={{ iconName: "ChevronRight" }}
                                        ariaLabel="Next month"
                                        subtle
                                        tooltipProps={{ text: "Next month" }}
                                        onClick={() => {
                                            if (this.calendarComponentRef.current) {
                                                this.getCalendarApi().next();
                                                this.currentMonthAndYear.value = this.calcMonths(this.currentMonthAndYear.value, 1);
                                            }
                                        }}
                                    />
                                </div>
                            </div>
                            <div style={{ display: "flex", alignItems: "center", gap: "8px", marginRight: "8px" }}>
                                <Button
                                    text="New Item"
                                    iconProps={{ iconName: "Add" }}
                                    primary={true}
                                    onClick={this.onClickNewItem}
                                />
                            </div>
                        </div>
                    </HeaderTitleArea>
                </CustomHeader>
                {/* Side panel options row */}
                <div className="side-panel-options-row">
                    <Observer isPaneOpen={this.isPaneOpen} isExpandedView={this.isExpandedView}>
                        {(props: { isPaneOpen: boolean; isExpandedView: boolean }) => {
                            return (
                                <>
                                    <Button
                                        onClick={() => {
                                            this.isColorPanelOpen.value = true;
                                        }}
                                        iconProps={{ iconName: "Color" }}
                                        ariaLabel="Color settings"
                                        tooltipProps={{ text: "Color settings" }}
                                        subtle
                                    />
                                    <Button
                                        onClick={(e) => {
                                            const currentElement = e.currentTarget as HTMLElement;

                                            if (this.sidePanelAnchorElement.value === currentElement) {
                                                this.sidePanelAnchorElement.value = undefined;
                                            } else {
                                                this.sidePanelAnchorElement.value = currentElement;
                                            }
                                        }}
                                        iconProps={{ iconName: "Equalizer" }}
                                        ariaLabel="Side pane options"
                                        tooltipProps={{ text: "View options" }}
                                        subtle
                                    />
                                </>
                            );
                        }}
                    </Observer>
                </div>
                {/* Contextual menu for side panel options */}
                <Observer sidePanelAnchorElement={this.sidePanelAnchorElement} isPaneOpen={this.isPaneOpen} isExpandedView={this.isExpandedView}>
                    {(props: { sidePanelAnchorElement: HTMLElement | undefined; isPaneOpen: boolean; isExpandedView: boolean }) => {
                        return props.sidePanelAnchorElement ? (
                            <ContextualMenu
                                anchorElement={props.sidePanelAnchorElement}
                                anchorOffset={{ horizontal: 0, vertical: 0 }}
                                anchorOrigin={{ horizontal: Location.start, vertical: Location.end }}
                                onActivate={(menuItem, event) => {
                                    if (event) {
                                        event.preventDefault();
                                        event.stopPropagation();
                                    }
                                    if (menuItem.onActivate) {
                                        const result = menuItem.onActivate(menuItem, event);
                                        return result;
                                    }
                                    return false;
                                }}
                                menuProps={{
                                    id: "side-panel-options",
                                    items: [
                                        {
                                            id: "expand-events-header",
                                            text: "Expand Events",
                                            itemType: MenuItemType.Header,
                                        },
                                        {
                                            id: "expand-events-toggle",
                                            iconProps: {
                                                render: () => (
                                                    <div 
                                                        onMouseDown={(e) => e.stopPropagation()}
                                                        onMouseUp={(e) => e.stopPropagation()}
                                                        onClick={(e) => e.stopPropagation()}
                                                        style={{ display: 'flex', alignItems: 'center' }}
                                                    >
                                                        <Toggle
                                                            checked={props.isExpandedView}
                                                            onText="On"
                                                            offText="Off"
                                                            onChange={(event, checked) => {
                                                                if (event) {
                                                                    event.stopPropagation();
                                                                    event.preventDefault();
                                                                }
                                                                this.isExpandedView.value = checked;
                                                            }}
                                                        />
                                                    </div>
                                                ),
                                            },
                                            onActivate: () => {
                                                this.isExpandedView.value = !this.isExpandedView.value;
                                                return false;
                                            },
                                        },
                                        {
                                            id: "separator",
                                            itemType: MenuItemType.Divider
                                        },
                                        {
                                            id: "side-pane-section",
                                            text: "Side Pane",
                                            itemType: MenuItemType.Header
                                        },
                                        {
                                            id: "events",
                                            iconProps: { 
                                                render: () => (
                                                    <div 
                                                        onMouseDown={(e) => e.stopPropagation()}
                                                        onMouseUp={(e) => e.stopPropagation()}
                                                        onClick={(e) => e.stopPropagation()}
                                                        style={{ display: 'flex', alignItems: 'center', width: '100%' }}
                                                    >
                                                        <Checkbox 
                                                            checked={props.isPaneOpen}
                                                            label="Events"
                                                            onChange={(event, checked) => {
                                                                if (event) {
                                                                    event.stopPropagation();
                                                                    event.preventDefault();
                                                                }
                                                                this.isPaneOpen.value = true;
                                                            }}
                                                        />
                                                    </div>
                                                )
                                            },
                                            onActivate: (menuItem: any, event: any) => {
                                                if (event) {
                                                    event.preventDefault();
                                                    event.stopPropagation();
                                                }
                                                this.isPaneOpen.value = true;
                                                return false;
                                            }
                                        },
                                        {
                                            id: "off",
                                            iconProps: { 
                                                render: () => (
                                                    <div 
                                                        onMouseDown={(e) => e.stopPropagation()}
                                                        onMouseUp={(e) => e.stopPropagation()}
                                                        onClick={(e) => e.stopPropagation()}
                                                        style={{ display: 'flex', alignItems: 'center', width: '100%' }}
                                                    >
                                                        <Checkbox 
                                                            checked={!props.isPaneOpen}
                                                            label="Off"
                                                            onChange={(event, checked) => {
                                                                if (event) {
                                                                    event.stopPropagation();
                                                                    event.preventDefault();
                                                                }
                                                                this.isPaneOpen.value = false;
                                                            }}
                                                        />
                                                    </div>
                                                )
                                            },
                                            onActivate: (menuItem: any, event: any) => {
                                                if (event) {
                                                    event.preventDefault();
                                                    event.stopPropagation();
                                                }
                                                this.isPaneOpen.value = false;
                                                return false;
                                            }
                                        }
                                    ]
                                }}
                                onDismiss={() => {
                                    this.sidePanelAnchorElement.value = undefined;
                                }}
                            />
                        ) : null;
                    }}
                </Observer>
                {/* Content area: Calendar and side panel side by side */}
                <div className="content-row flex-row">
                    <Observer isPaneOpen={this.isPaneOpen} display={this.displayCalendar} isExpandedView={this.isExpandedView}>
                        {(props: { isPaneOpen: boolean; display: boolean; isExpandedView: boolean }) => (
                            <div className={`flex-column scroll-hidden calendar-area ${props.isPaneOpen ? 'pane-open' : 'pane-closed'}`}>
                                {props.display ? (
                                    <div className="calendar-component">
                                        <FullCalendar
                                            defaultView="dayGridMonth"
                                            editable={true}
                                            eventClick={this.onEventClick}
                                            eventDrop={this.onEventDrop}
                                            eventRender={this.eventRender}
                                            eventResize={this.onEventResize}
                                            eventOrder="order,start,-duration,allDay,title"
                                            eventSources={[
                                                { events: this.freeFormEventSource.getEvents },
                                                { events: this.vsoCapacityEventSource.getEvents }
                                            ]}
                                            firstDay={localeData(navigator.language).firstDayOfWeek()}
                                            header={false}
                                            height={this.getCalendarHeight()}
                                            plugins={[dayGridPlugin, timeGridPlugin, interactionPlugin]}
                                            ref={this.calendarComponentRef}
                                            select={this.onSelectCalendarDates}
                                            selectable={true}
                                            eventLimit={!props.isExpandedView}
                                        />
                                    </div>
                                ) : null}
                            </div>
                        )}
                    </Observer>
                    <Observer isPaneOpen={this.isPaneOpen} eventColorMap={this.eventColorMap}>
                        {(props: { isPaneOpen: boolean; eventColorMap: Map<string, string> }) => (
                            props.isPaneOpen ? (
                                <SummaryComponent 
                                    capacityEventSource={this.vsoCapacityEventSource} 
                                    freeFormEventSource={this.freeFormEventSource}
                                    eventColorMap={props.eventColorMap}
                                    onEditDaysOff={this.onEditDaysOff}
                                    onEditEvent={this.onEditFreeFormEvent}
                                    onTogglePane={() => this.isPaneOpen.value = !this.isPaneOpen.value}
                                />
                            ) : null
                        )}
                    </Observer>
                </div>
                <Observer anchorElement={this.anchorElement}>
                    {(props: { anchorElement: HTMLElement | undefined }) => {
                        return props.anchorElement ? (
                            <ContextualMenu
                                anchorElement={props.anchorElement}
                                anchorOffset={{ horizontal: 4, vertical: 4 }}
                                anchorOrigin={{ horizontal: Location.start, vertical: Location.start }}
                                key={this.selectedEndDate!.toString()}
                                menuProps={{
                                    id: "foo",
                                    items: [
                                        { id: "event", text: "Add event", iconProps: { iconName: "Add" }, onActivate: this.onClickAddEvent },
                                        { id: "dayOff", text: "Add days off", iconProps: { iconName: "Clock" }, onActivate: this.onClickAddDaysOff }
                                    ]
                                }}
                                onDismiss={() => {
                                    this.anchorElement.value = undefined;
                                }}
                            />
                        ) : null;
                    }}
                </Observer>
                {/* Color Settings Panel */}
                <Observer isColorPanelOpen={this.isColorPanelOpen} eventColorMap={this.eventColorMap}>
                    {(props: { isColorPanelOpen: boolean; eventColorMap: Map<string, string> }) => {
                        if (!props.isColorPanelOpen) return null;

                        // Get all categories from both event sources
                        const categories: string[] = [];
                        const freeFormCategories = Array.from(this.freeFormEventSource.getCategories());
                        categories.push(...freeFormCategories);
                        
                        // Remove duplicates
                        const uniqueCategories = Array.from(new Set(categories));

                        return (
                            <Panel
                                onDismiss={() => {
                                    this.isColorPanelOpen.value = false;
                                    this.tempColorMap = new Map();
                                }}
                                titleProps={{ text: "Event Colors" }}
                                description="Customize colors for event categories"
                                className="color-settings-panel"
                                footerButtonProps={[
                                    {
                                        text: "Cancel",
                                        onClick: () => {
                                            this.isColorPanelOpen.value = false;
                                            this.tempColorMap = new Map();
                                        }
                                    },
                                    {
                                        text: "Save",
                                        primary: true,
                                        onClick: () => {
                                            // Merge temp colors into existing map instead of replacing
                                            const newColorMap = new Map(this.eventColorMap.value);
                                            this.tempColorMap.forEach((value, key) => {
                                                newColorMap.set(key, value);
                                            });
                                            this.eventColorMap.value = newColorMap;
                                            this.saveColorSettings();
                                            this.isColorPanelOpen.value = false;
                                            this.tempColorMap = new Map();
                                            // Refresh calendar
                                            if (this.calendarComponentRef.current) {
                                                this.getCalendarApi().refetchEvents();
                                            }
                                        }
                                    }
                                ]}
                            >
                                <div className="color-settings-content">
                                    {uniqueCategories.length === 0 ? (
                                        <div style={{ padding: "16px", textAlign: "center", color: "var(--text-secondary-color)" }}>
                                            No event categories yet. Add events to customize their colors.
                                        </div>
                                    ) : (
                                        uniqueCategories.map(category => {
                                            const currentColor = this.tempColorMap.has(category)
                                                ? this.tempColorMap.get(category)
                                                : (props.eventColorMap.get(category) || this.getDefaultColor(category));

                                            return (
                                                <div key={category} style={{ marginBottom: "16px", display: "flex", alignItems: "center", gap: "12px" }}>
                                                    <div style={{ flex: "0 0 120px", fontWeight: 500 }}>
                                                        {category}
                                                    </div>
                                                    <TextField
                                                        value={currentColor}
                                                        onChange={(e, newValue) => {
                                                            this.tempColorMap.set(category, newValue);
                                                            // Force re-render
                                                            this.forceUpdate();
                                                        }}
                                                        placeholder="#3b82f6"
                                                    />
                                                    <div
                                                        className="color-preview-swatch"
                                                        style={{
                                                            backgroundColor: currentColor
                                                        }}
                                                        onClick={() => {
                                                            const colorInput = document.getElementById(`color-picker-${category}`) as HTMLInputElement;
                                                            if (colorInput) {
                                                                colorInput.click();
                                                            }
                                                        }}
                                                        title="Click to open color picker"
                                                    >
                                                        <input
                                                            id={`color-picker-${category}`}
                                                            type="color"
                                                            value={currentColor}
                                                            onChange={(e) => {
                                                                this.tempColorMap.set(category, e.target.value);
                                                                this.forceUpdate();
                                                            }}
                                                            style={{
                                                                opacity: 0,
                                                                position: "absolute",
                                                                width: "100%",
                                                                height: "100%",
                                                                cursor: "pointer"
                                                            }}
                                                        />
                                                    </div>
                                                    <Button
                                                        text="Reset"
                                                        subtle
                                                        onClick={() => {
                                                            this.tempColorMap.set(category, this.getDefaultColor(category));
                                                            this.forceUpdate();
                                                        }}
                                                    />
                                                </div>
                                            );
                                        })
                                    )}
                                </div>
                            </Panel>
                        );
                    }}
                </Observer>
                <Observer dialog={this.openDialog}>
                    {(props: { dialog: Dialogs }) => {
                        return props.dialog === Dialogs.NewDaysOffDialog ? (
                            <AddEditDaysOffDialog
                                calendarApi={this.getCalendarApi()}
                                end={this.selectedEndDate}
                                event={this.eventToEdit}
                                eventSource={this.vsoCapacityEventSource}
                                members={this.members}
                                onDismiss={this.onDialogDismiss}
                                start={this.selectedStartDate}
                            />
                        ) : props.dialog === Dialogs.NewEventDialog ? (
                            <AddEditEventDialog
                                calendarApi={this.getCalendarApi()}
                                end={this.selectedEndDate}
                                eventApi={this.eventApi}
                                eventSource={this.freeFormEventSource}
                                onDismiss={this.onDialogDismiss}
                                start={this.selectedStartDate}
                            />
                        ) : null;
                    }}
                </Observer>
                            </>
                        );
                    }}
                </Observer>
            </Page>
        );
    }

    componentDidMount() {
        SDK.init();
        this.initialize();
        window.addEventListener("resize", this.updateDimensions);
    }

    private calcMonths(current: MonthAndYear, monthDelta: number): MonthAndYear {
        let month = (current.month + monthDelta) % 12;
        let year = current.year + Math.floor((current.month + monthDelta) / 12);
        if (month < 0) {
            month = 12 + month;
        }
        return { month, year };
    }

    /**
     * Edits the rendered event if required
     */
    private eventRender = (arg: { isMirror: boolean; isStart: boolean; isEnd: boolean; event: EventApi; el: HTMLElement; view: View }) => {
        // Apply custom colors based on category
        const category = arg.event.extendedProps?.category;
        if (category && this.eventColorMap.value.has(category)) {
            const customColor = this.eventColorMap.value.get(category);
            arg.el.style.backgroundColor = customColor!;
            arg.el.style.borderColor = customColor!;
        }

        if (arg.event.id.startsWith(DaysOffId) && arg.event.start) {
            // Add days-off class for styling
            arg.el.classList.add('fc-days-off-event');
            
            // get grouped event for that date
            const capacityEvent = this.vsoCapacityEventSource.getGroupedEventForDate(arg.event.start);
            
            if (capacityEvent && capacityEvent.icons && capacityEvent.icons.length > 0) {
                const maxIconsToShow = 4;
                const totalIcons = capacityEvent.icons.length;
                
                capacityEvent.icons.slice(0, maxIconsToShow).forEach(element => {
                    if (element.src) {
                        var img: HTMLImageElement = document.createElement("img");
                        img.src = element.src;
                        img.className = "event-icon";
                        img.title = element.linkedEvent.title;
                        img.onclick = () => {
                            this.eventToEdit = element.linkedEvent;
                            this.openDialog.value = Dialogs.NewDaysOffDialog;
                        };
                        var content = arg.el.querySelector(".fc-content");
                        if (content) {
                            content.appendChild(img);
                        }
                    }
                });
                
                if (totalIcons > maxIconsToShow) {
                    const hiddenPeople = capacityEvent.icons.slice(maxIconsToShow).map(icon => {
                        const member = icon.linkedEvent.member;
                        return member ? member.displayName : icon.linkedEvent.title;
                    });
                    
                    const moreIndicator = document.createElement("span");
                    moreIndicator.className = "event-icon-more";
                    moreIndicator.innerText = `+${totalIcons - maxIconsToShow}`;
                    moreIndicator.title = `${hiddenPeople.join('\n')}`;
                    moreIndicator.style.cursor = 'default';
                    var content = arg.el.querySelector(".fc-content");
                    if (content) {
                        content.appendChild(moreIndicator);
                    }
                }
            }
        } else if (arg.event.id.startsWith(IterationId) && arg.isStart) {
            // iterations are background event, show title for only start
            // Create a span element to hold the text
            const textSpan = document.createElement('span');
            textSpan.innerText = arg.event.title;
            textSpan.classList.add('sprint-label');
            
            // Detect theme using browser's prefers-color-scheme
            const isDarkTheme = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
            
            // Dark mode: white text on dark background
            // Light mode: black text on white background
            const textColor = isDarkTheme ? '#ffffff' : '#000000';
            const bgColor = isDarkTheme ? 'rgba(0, 0, 0, 0.9)' : 'rgba(255, 255, 255, 0.95)';
            
            textSpan.style.cssText = `
                color: ${textColor} !important;
                font-weight: 700 !important;
                font-size: 13px !important;
                padding: 4px 8px !important;
                display: block !important;
                position: absolute !important;
                top: 3px !important;
                left: 6px !important;
                z-index: 10 !important;
                opacity: 1 !important;
                line-height: normal !important;
                pointer-events: none !important;
                white-space: nowrap !important;
                background: ${bgColor} !important;
                border-radius: 3px !important;
            `;
            
            // Make parent td positioned relative
            arg.el.style.position = 'relative';
            
            // Clear any existing content and append the span
            arg.el.innerHTML = '';
            arg.el.appendChild(textSpan);
        }
    };

    private getCalendarApi(): Calendar {
        return this.calendarComponentRef.current!.getApi();
    }

    /**
     * Manually calculates available vertical space for calendar
     */
    private getCalendarHeight(): number {
        var height = document.getElementById("team-calendar");
        if (height) {
            return height.offsetHeight - 150;
        }
        return 200;
    }

    private getDefaultColor(category: string): string {
        // Return default colors for known categories
        const defaultColors: { [key: string]: string } = {
            "Days Off": "#ff6b6b",
            "Iteration": "#4dabf7",
            "Uncategorized": "#868e96"
        };

        if (defaultColors[category]) {
            return defaultColors[category];
        }

        // For custom categories, generate a color based on the category name
        // This ensures consistent colors for the same category
        return generateColor(category);
    }

    private async loadColorSettings() {
        if (!this.dataManager) return;

        try {
            const colorData = await this.dataManager.getValue<{ [key: string]: string }>("eventColors");
            if (colorData) {
                const colorMap = new Map<string, string>();
                for (const key in colorData) {
                    if (colorData.hasOwnProperty(key)) {
                        colorMap.set(key, colorData[key]);
                    }
                }
                this.eventColorMap.value = colorMap;
            }
        } catch (error) {
            // No saved colors yet, use defaults
            console.log("No saved color settings found, using defaults");
        }
    }

    private async saveColorSettings() {
        if (!this.dataManager) return;

        try {
            const colorData: { [key: string]: string } = {};
            this.eventColorMap.value.forEach((value, key) => {
                colorData[key] = value;
            });
            await this.dataManager.setValue("eventColors", colorData);
        } catch (error) {
            console.error("Failed to save color settings:", error);
        }
    }

    private getMonthPickerOptions(): IListBoxItem[] {
        const options: IListBoxItem[] = [];
        const listSize = 3;
        const currentMonth = this.currentMonthAndYear.value;
        for (let i = -listSize; i <= listSize; ++i) {
            const monthAndYear = this.calcMonths(this.currentMonthAndYear.value, i);
            const text = monthAndYearToString(monthAndYear);
            const isCurrent = i === 0;
            options.push({
                data: monthAndYear,
                id: text,
                text: text,
                iconProps: isCurrent ? { iconName: "Calendar" } : undefined
            });
        }
        return options;
    }

    private getTeamPickerOptions(): IListBoxItem[] {
        const options: IListBoxItem[] = [];
        this.teams.value.forEach(function(item) {
            options.push({ data: item, id: item.id, text: item.name });
        });

        return options;
    }

    private async initialize() {
        const dataSvc = await SDK.getService<IExtensionDataService>(CommonServiceIds.ExtensionDataService);
        const projectService = await SDK.getService<IProjectPageService>(CommonServiceIds.ProjectPageService);
        const project = await projectService.getProject();
        const locationService = await SDK.getService<ILocationService>(CommonServiceIds.LocationService);

        this.dataManager = await dataSvc.getExtensionDataManager(SDK.getExtensionContext().id, await SDK.getAccessToken());
        this.navigationService = await SDK.getService<IHostNavigationService>(CommonServiceIds.HostNavigationService);
        this.locationService = locationService;

        const queryParam = await this.navigationService.getQueryParams();
        let selectedTeamId;

        // if URL has team id in it, use that
        if (queryParam && queryParam["team"]) {
            selectedTeamId = queryParam["team"];
        }

        if (project) {
            if (!selectedTeamId) {
                // Nothing in URL - check data service
                selectedTeamId = await this.dataManager.getValue<string>("selected-team-" + project.id, { scopeType: "User" });
            }

            const client = getClient(CoreRestClient);

            const allTeams = [];
            let teams;
            let callCount = 0;
            const fetchCount = 1000;
            do {
                teams = await client.getTeams(project.id, false, fetchCount, callCount * fetchCount);
                allTeams.push(...teams);
                callCount++;
            } while (teams.length === fetchCount);

            this.projectId = project.id;
            this.projectName = project.name;

            allTeams.sort((a, b) => {
                return a.name.toUpperCase().localeCompare(b.name.toUpperCase());
            });

            // if team id wasn't in URL or database use first available team
            if (!selectedTeamId) {
                selectedTeamId = allTeams[0].id;
            }

            if (!queryParam || !queryParam["team"]) {
                // Add team id to URL
                this.navigationService.setQueryParams({ team: selectedTeamId });
            }

            this.hostUrl = await locationService.getServiceLocation();
            try {
                this.selectedTeamName = (await client.getTeam(project.id, selectedTeamId)).name;
            } catch (error) {
                console.error(`Failed to get team with ID ${selectedTeamId}: ${error}`);
              
                
            }
            this.freeFormEventSource.initialize(selectedTeamId, this.dataManager);
            this.vsoCapacityEventSource.initialize(project.id, this.projectName, selectedTeamId, this.selectedTeamName, this.hostUrl, locationService);
            
            const preloadPromises = [
                this.freeFormEventSource.preloadCurrentMonthEvents(),
                this.vsoCapacityEventSource.preloadCurrentIterations(),
                this.loadColorSettings()
            ];
            try {
                await Promise.all(preloadPromises);
            } catch (error) {
                console.error("Error preloading data:", error);
            }

            this.displayCalendar.value = true;
            this.dataManager.setValue<string>("selected-team-" + project.id, selectedTeamId, { scopeType: "User" });
            this.teams.value = allTeams;
            this.members = await client.getTeamMembersWithExtendedProperties(project.id, selectedTeamId);
        }
    }

    private onClickNewItem = () => {
        this.eventApi = undefined;
        const today = new Date();
        this.selectedStartDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
        this.selectedEndDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
        if (this.calendarComponentRef.current) {
            this.openDialog.value = Dialogs.NewEventDialog;
        }
    };

    private onClickAddEvent = () => {
        this.eventApi = undefined;
        this.openDialog.value = Dialogs.NewEventDialog;
    };

    private onClickAddDaysOff = () => {
        this.eventToEdit = undefined;
        this.openDialog.value = Dialogs.NewDaysOffDialog;
    };

    private onDialogDismiss = () => {
        this.openDialog.value = Dialogs.None;
    };

    private onEditDaysOff = (event: ICalendarEvent) => {
        this.eventToEdit = event;
        this.openDialog.value = Dialogs.NewDaysOffDialog;
    };

    private onEditFreeFormEvent = (eventId: string) => {
        if (!eventId) {
            console.error("Event ID is undefined");
            return;
        }
        
        const event = this.freeFormEventSource.eventMap[eventId];
        if (!event) {
            console.error("Event not found in eventMap:", eventId);
            return;
        }
        
        if (!this.calendarComponentRef.current) {
            console.error("Calendar component ref is not available");
            return;
        }
        
        const calendarApi = this.getCalendarApi();
        const calendarEvent = calendarApi.getEventById(FreeFormId + "." + eventId);
        if (calendarEvent) {
            this.eventApi = calendarEvent;
            this.selectedStartDate = calendarEvent.start || new Date();
            this.selectedEndDate = calendarEvent.end ? new Date(calendarEvent.end.getTime() - 86400000) : this.selectedStartDate;
            this.openDialog.value = Dialogs.NewEventDialog;
        } else {
            console.error("Calendar event not found:", FreeFormId + "." + eventId);
        }
    };

    private onEventClick = (arg: { el: HTMLElement; event: EventApi; jsEvent: MouseEvent; view: View }) => {
        if (arg.event.id.startsWith(FreeFormId)) {
            this.eventApi = arg.event;
            this.openDialog.value = Dialogs.NewEventDialog;
        }
    };

    private onEventDrop = (arg: {
        el: HTMLElement;
        event: EventApi;
        oldEvent: EventApi;
        delta: Duration;
        revert: () => void;
        jsEvent: Event;
        view: View;
    }) => {
        if (arg.event.id.startsWith(FreeFormId)) {
            let inclusiveEndDate;
            if (arg.event.end) {
                inclusiveEndDate = new Date(arg.event.end);
                inclusiveEndDate.setDate(arg.event.end.getDate() - 1);
            } else {
                inclusiveEndDate = new Date(arg.event.start!);
            }

            this.freeFormEventSource.updateEvent(
                arg.event.extendedProps.id,
                arg.event.title,
                arg.event.start!,
                inclusiveEndDate,
                arg.event.extendedProps.category,
                arg.event.extendedProps.description
            );
        }
    };

    private onEventResize = (arg: {
        el: HTMLElement;
        startDelta: Duration;
        endDelta: Duration;
        prevEvent: EventApi;
        event: EventApi;
        revert: () => void;
        jsEvent: Event;
        view: View;
    }) => {
        if (arg.event.id.startsWith(FreeFormId)) {
            let inclusiveEndDate;
            if (arg.event.end) {
                inclusiveEndDate = new Date(arg.event.end);
                inclusiveEndDate.setDate(arg.event.end.getDate() - 1);
            } else {
                inclusiveEndDate = new Date(arg.event.start!);
            }

            this.freeFormEventSource.updateEvent(
                arg.event.extendedProps.id,
                arg.event.title,
                arg.event.start!,
                inclusiveEndDate,
                arg.event.extendedProps.category,
                arg.event.extendedProps.description
            );
        }
    };

    private onSelectCalendarDates = (arg: {
        start: Date;
        end: Date;
        startStr: string;
        endStr: string;
        allDay: boolean;
        resource?: any;
        jsEvent: MouseEvent;
        view: View;
    }) => {
        this.selectedEndDate = new Date(arg.end);
        this.selectedEndDate.setDate(arg.end.getDate() - 1);
        this.selectedStartDate = arg.start;
        const dataDate = formatDate(this.selectedEndDate, "YYYY-MM-DD");
        this.anchorElement.value = document.querySelector("td.fc-day-top[data-date='" + dataDate + "']") as HTMLElement;
        
        // Prevent calendar from collapsing when context menu opens
        setTimeout(() => {
            if (this.calendarComponentRef.current) {
                this.calendarComponentRef.current.getApi().updateSize();
            }
        }, 10);
    };

    private onSelectMonthYear = (event: React.SyntheticEvent<HTMLElement, Event>, item: IListBoxItem<{}>) => {
        const date = item.data as MonthAndYear;
        if (this.calendarComponentRef) {
            this.getCalendarApi().gotoDate(new Date(date.year, date.month));
            this.currentMonthAndYear.value = date;
        }
    };

    private onSelectTeam = async (event: React.SyntheticEvent<HTMLElement, Event>, item: IListBoxItem<{}>) => {
        const newTeam = item.data! as WebApiTeam;
        this.selectedTeamName = newTeam.name;
        this.freeFormEventSource.initialize(newTeam.id, this.dataManager!);
        this.vsoCapacityEventSource.initialize(this.projectId, this.projectName, newTeam.id, newTeam.name, this.hostUrl, this.locationService);
        
        const preloadPromises = [
            this.freeFormEventSource.preloadCurrentMonthEvents(),
            this.vsoCapacityEventSource.preloadCurrentIterations()
        ];
        
        await Promise.all(preloadPromises);
        
        this.getCalendarApi().refetchEvents();
        this.dataManager!.setValue<string>("selected-team-" + this.projectId, newTeam.id, { scopeType: "User" });
        this.navigationService!.setQueryParams({ team: newTeam.id });
        this.members = await getClient(CoreRestClient).getTeamMembersWithExtendedProperties(this.projectId, newTeam.id);
    };

    private updateDimensions = () => {
        if (this.calendarComponentRef.current) {
            this.calendarComponentRef.current.getApi().setOption("height", this.getCalendarHeight());
        }
    };
}

function showRootComponent(component: React.ReactElement<any>) {
    ReactDOM.render(component, document.getElementById("team-calendar"));
}

showRootComponent(<ExtensionContent />);
