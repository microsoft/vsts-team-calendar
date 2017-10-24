import "es6-promise/auto";
import * as Calendar_Views from "./Views";

$(() => {
    Calendar_Views.CalendarView.enhance(Calendar_Views.CalendarView, $(".calendar-view"));
    Calendar_Views.SummaryView.enhance(Calendar_Views.SummaryView, $(".summary-view"));
});
