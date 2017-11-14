import * as React from "react";
import * as ReactDOM from "react-dom";
import * as CalendarView from "./Views";

import "es6-promise/auto";
import * as Calendar_Views from "./Views";

// Webpack-linked stylesheet for this extension
require("./Style/style.scss");

$(() => {
    ReactDOM.render(<CalendarView.CalendarComponent />, document.getElementById("calendarArea"));
    ReactDOM.render(<CalendarView.SummaryComponent />, document.getElementById("summaryArea"));
});
