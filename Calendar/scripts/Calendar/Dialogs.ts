/// <reference path='../ref/VSS/VSS.d.ts' />

import Calendar_Contracts = require("Calendar/Contracts");
import Controls = require("VSS/Controls");
import Controls_Common = require("VSS/Controls/Common");
import Controls_Validation = require("VSS/Controls/Validation");
import Utils_Core = require("VSS/Utils/Core");
import Utils_UI = require("VSS/Utils/UI");
import WebApi_Contracts = require("VSS/WebApi/Contracts");

var domElem = Utils_UI.domElem;

/**
 * Base class for a dialog which allows users to add a new event or edit an existing one
*/
export class EditEventDialog extends Controls_Common.ModalDialog {
    private _$container: JQuery;
    protected _$startInput: JQuery;
    protected _$endInput: JQuery;


    private _eventValidationError: Controls_Common.MessageAreaControl;
    protected _calendarEvent: Calendar_Contracts.CalendarEvent;

    public initializeOptions(options?: any) {
        super.initializeOptions($.extend(options, { "height": 250 }));
    }

    public initialize() {
        this._calendarEvent = this._options.event;
        this._createLayout();
        super.initialize();
    }

    /**
     * Returns the title of the dialog
     */
    public getTitle(): string {
        return this._options.title;
    }

    /**
     * Processes the data that the user has entered and either
     * shows an error message, or returns the edited note.
     */
    public onOkClick(): any {
        this._calendarEvent.startDate = this._parseDateValue(this._$startInput.val());
        this._calendarEvent.endDate = this._parseDateValue(this._$endInput.val());

        this._buildCalendarEventFromFields();

        this.processResult(this._calendarEvent);
    }

    protected _buildCalendarEventFromFields() {
    }

    protected _createLayout() {
        this._$container = $(domElem('div')).addClass('edit-event-container').appendTo(this._element);

        this._eventValidationError = <Controls_Common.MessageAreaControl>Controls.BaseControl.createIn(Controls_Common.MessageAreaControl, this._$container, { closeable: false });

        var $editControl = $(domElem('div', 'event-edit-control'));
        var $fieldsContainer = $(domElem('table')).appendTo($editControl);

        this._$startInput = $("<input type='text' id='fieldStartDate' />").val(this._formatDateValue(this._calendarEvent.startDate))
            .on("blur",(e) => {
                this.updateOkButton(this._validate());
        });
        this._$endInput = $("<input type='text' id='fieldEndDate' />")
            .on("blur",(e) => {
                this.updateOkButton(this._validate());
            });

        if (this._calendarEvent.endDate) {
            this._$endInput.val(this._formatDateValue(this._calendarEvent.endDate));
        }
        else {
            this._$endInput.val(this._formatDateValue(this._calendarEvent.startDate));
        }

        // Populate fields container with fields. The form fields array contain pairs of field label and field element itself.
        var fields = this._getFormFields();
        for (var i = 0, l = fields.length; i < l; i += 1) {
            var labelName = fields[i][0];
            var field = fields[i][1];

            var $row = $(domElem("tr"));

            var fieldId = field.attr("id") || $("input", field).attr("id");
            $(domElem("label")).attr("for", fieldId).text(labelName).appendTo($(domElem("td", "label")).appendTo($row));

            field.appendTo($(domElem("td"))
                .appendTo($row));

            $row.appendTo($fieldsContainer);
        }

        this._$container.append($editControl);

        var startCombo = <Controls_Common.Combo>Controls.Enhancement.enhance(Controls_Common.Combo, this._$startInput, {
            type: "date-time"
        });
        var endCombo = <Controls_Common.Combo>Controls.Enhancement.enhance(Controls_Common.Combo, this._$endInput, {
            type: "date-time"
        });

        this._setupValidators(this._$startInput, "Start date must be a valid date");
        this._setupValidators(this._$endInput, "End date must be a valid date", "End date must be equal to or after start date", this._$startInput, DateComparisonOptions.GREATER_OR_EQUAL);
    }

    private _setupValidators($field: JQuery, validDateFormatMessage: string, relativeToErrorMessage?: string, $relativeToField?: JQuery, dateComparisonOptions?: DateComparisonOptions) {
        <Controls_Validation.DateValidator<Controls_Validation.DateValidatorOptions>>Controls.Enhancement.enhance(Controls_Validation.DateValidator, $field, {
            invalidCssClass: "date-invalid",
            group: "default",
            message: validDateFormatMessage
        });

        if (relativeToErrorMessage) {
            <DateRelativeToValidator>Controls.Enhancement.enhance(DateRelativeToValidator, $field, {
                comparison: dateComparisonOptions,
                relativeToField: $relativeToField,
                group: "default",
                message: relativeToErrorMessage
            });
        }
    }

    protected _getFormFields(): any[] {
        var fields = [];

        return fields;
    }

    protected _validate(): boolean {
        return true;
    }

    private _formatDateValue(date: Date): string {
        return date === null ? "" : Utils_Core.DateUtils.format(new Date(date.valueOf()), "d");
    }

    private _parseDateValue(date: string): Date {
        return date === null ? null : Utils_Core.DateUtils.parseDateString(date, "d", true);
    }

    protected _setError(errorMessage: string) {
        this._eventValidationError.setError($("<span />").html(errorMessage));
    }

    protected _clearError() {
        this._eventValidationError.clear();
    }
}

/**
 * A dialog which allows users to add / edit free-form events.
 * In addition to start/end dates, allows user to enter a title and select a category.
*/
export class EditFreeFormEventDialog extends EditEventDialog {
    private _$titleInput: JQuery;
    private _$categoryInput: JQuery;

    public initialize() {
        super.initialize();

        if (this._calendarEvent.title) {
            this._$titleInput.val(this._calendarEvent.title);
            this.updateOkButton(true);
        }

        this._$categoryInput.val(this._calendarEvent.category || "");
    }

    protected _buildCalendarEventFromFields() {
        this._calendarEvent.title = $.trim(this._$titleInput.val());
        this._calendarEvent.category = $.trim(this._$categoryInput.val());
    }

    protected _createLayout() {
        this._$titleInput = $("<input type='text' class='requiredInfoLight' id='fieldTitle'/>")
            .on("input keyup",(e) => {
            if (e.keyCode !== Utils_UI.KeyCode.ENTER) {
                this.updateOkButton(this._validate());
            }
        });

        this._$categoryInput = $("<input type='text' id='fieldCategory' />");

        // Populate categories
        var categories = <IPromise<string[]>>this._options.categories;
        if (categories) {
            categories.then((allCategories: string[]) => {
                Controls.Enhancement.enhance(Controls_Common.Combo, this._$categoryInput, {
                    source: allCategories,
                    dropCount: 3
                });
            });
        }

        super._createLayout();
    }

    protected _getFormFields(): any[] {
        var fields = [];

        fields.push(["Title", this._$titleInput]);
        fields.push(["Start Date", this._$startInput]);
        fields.push(["End Date", this._$endInput]);
        fields.push(["Category", this._$categoryInput]);

        return fields;
    }

    protected _validate(): boolean {
        var title: string = $.trim(this._$titleInput.val());
        if (title.length <= 0) {
            this._clearError();
            return false;
        }
        var validationResult = [];
        var isValid: boolean = Controls_Validation.validateGroup('default', validationResult);
        if (!isValid) {
            this._setError(validationResult[0].getMessage());
            return false;
        }

        this._clearError();
        return true;
    }
}

/**
 * A dialog which allows users to add / edit days off events.
 * In addition to start/end dates, allows user to select a user (or the entire team).
*/
export class EditCapacityEventDialog extends EditEventDialog {
    private _$memberInput: JQuery;
    private _members: WebApi_Contracts.IdentityRef[];
    private static EVERYONE : string = "Everyone";

    public initialize() {
        super.initialize();
        this._$memberInput.val(this._calendarEvent.member.displayName || "");
        if (this._options.isEdit) {
            this._$memberInput.addClass('requiredInfoLight');
            this._$memberInput.prop('disabled', true);
        }
        this.updateOkButton(true);
    }

    protected _buildCalendarEventFromFields() {        
        if (this._members) {
            var displayName = $.trim(this._$memberInput.val());
            var eventMember : any;
            if (displayName === EditCapacityEventDialog.EVERYONE) {
                this._calendarEvent.member.displayName = displayName;
            }
            else {
                this._members.some((member: WebApi_Contracts.IdentityRef, index: number, array: WebApi_Contracts.IdentityRef[]) => {
                    if (member.displayName === displayName) {
                        eventMember = member;
                        return true;
                    }
                    return false;
                });
                this._calendarEvent.member = eventMember;
            }
        }
    }

    protected _createLayout() {
        this._$memberInput = $("<input type='text' id='fieldMember' />");

        if (this._options.membersPromise) {
            this._options.membersPromise.then((members: WebApi_Contracts.IdentityRef[]) => {
                this._members = members;
                var memberNames = [];
                memberNames.push(EditCapacityEventDialog.EVERYONE);
                members.sort((a, b) => { return a.displayName.toLocaleLowerCase().localeCompare(b.displayName.toLocaleLowerCase()); });
                members.forEach((member: WebApi_Contracts.IdentityRef, index: number, array: WebApi_Contracts.IdentityRef[]) => {
                    memberNames.push(member.displayName);
                });

                super._createLayout();

                Controls.Enhancement.enhance(Controls_Common.Combo, this._$memberInput, {
                    source: memberNames,
                    dropCount: 3
                });
            });
        }
        else {
            super._createLayout();
            this._$memberInput.prop('disabled', true);
        }
    }

    protected _getFormFields(): any[] {
        var fields = [];

        fields.push(["Start Date", this._$startInput]);
        fields.push(["End Date", this._$endInput]);
        fields.push(["Team Member", this._$memberInput]);

        return fields;
    }

    protected _validate(): boolean {
        var validationResult = [];
        var isValid: boolean = Controls_Validation.validateGroup('default', validationResult);
        if (!isValid) {
            this._setError(validationResult[0].getMessage());
            return false;
        }

        this._clearError();
        return true;
    }
}

interface DateRelativeToValidatorOptions extends Controls_Validation.BaseValidatorOptions {
    relativeToField: any;
    comparison: DateComparisonOptions;
    message: string;
    group: string;
    parseFormat?: string;
}

const enum DateComparisonOptions {
    GREATER_OR_EQUAL,
    LESS_OR_EQUAL
}

class DateRelativeToValidator extends Controls_Validation.BaseValidator<DateRelativeToValidatorOptions> {

    constructor(options?: DateRelativeToValidatorOptions) {
        super(options);
    }

    public initializeOptions(options?: DateRelativeToValidatorOptions) {
        super.initializeOptions(<DateRelativeToValidatorOptions>$.extend({
            invalidCssClass: "date-relative-to-invalid",
        }, options));
    }

    public isValid(): boolean {
        var fieldText = $.trim(this.getValue()),
            relativeToFieldText = $.trim(this._options.relativeToField.val()),
            fieldDate,
            relativeToFieldDate,
            result = false;

        if (fieldText && relativeToFieldText) {
            fieldDate = Utils_Core.DateUtils.parseDateString(fieldText, this._options.parseFormat, true);
            relativeToFieldDate = Utils_Core.DateUtils.parseDateString(relativeToFieldText, this._options.parseFormat, true);
        }
        else {
            return true;
        }

        if ((fieldDate instanceof Date) && !isNaN(fieldDate) && relativeToFieldDate instanceof Date && !isNaN(relativeToFieldDate)) {
            if (this._options.comparison === DateComparisonOptions.GREATER_OR_EQUAL) {
                result = fieldDate >= relativeToFieldDate;
            }
            else {
                result = fieldDate <= relativeToFieldDate;
            }
        }
        else {
            result = true;
        }

        return result;
    }

    public getMessage() {
        return this._options.message;
    }
}