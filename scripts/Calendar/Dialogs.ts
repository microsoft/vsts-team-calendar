/// <reference path='../../typings/VSS.d.ts' />
/// <reference path='../../typings/moment/moment.d.ts' />

import Calendar_Contracts = require("Calendar/Contracts");
import Calendar_ColorUtils = require("Calendar/Utils/Color");
import Context = require("VSS/Context");
import Controls = require("VSS/Controls");
import Controls_Contributions = require("VSS/Contributions/Controls");
import Controls_Notifications = require("VSS/Controls/Notifications");
import Controls_Combos = require("VSS/Controls/Combos");
import Controls_Dialog = require("VSS/Controls/Dialogs");
import Controls_Menus = require("VSS/Controls/Menus");
import Controls_Popup = require("VSS/Controls/PopupContent");
import Controls_Validation = require("VSS/Controls/Validation");
import Culture = require("VSS/Utils/Culture");
import Utils_Core = require("VSS/Utils/Core");
import Utils_Date = require("VSS/Utils/Date");
import Utils_String = require("VSS/Utils/String");
import Utils_UI = require("VSS/Utils/UI");
import WebApi_Contracts = require("VSS/WebApi/Contracts");
import Work_Contracts = require("TFS/Work/Contracts");
import Q = require("q");

var domElem = Utils_UI.domElem;

export interface IEventControlOptions {    
     calendarEvent: Calendar_Contracts.CalendarEvent;
     title?: string;
     isEdit?: boolean;
     validStateChangedHandler: (valid: boolean) => any;
}

export interface IFreeFormEventControlOptions extends IEventControlOptions { 
    categoriesPromise: () => IPromise<Calendar_Contracts.IEventCategory[]>;
}

export interface ICapacityEventControlOptions extends IEventControlOptions {  
    membersPromise: IPromise<WebApi_Contracts.IdentityRef[]>;
    getIterations: () => IPromise<Work_Contracts.TeamSettingsIteration[]>;
}

export interface IEventDialogOptions extends Controls_Dialog.IModalDialogOptions {
    source: Calendar_Contracts.IEventSource;
    calendarEvent: Calendar_Contracts.CalendarEvent;
    query: Calendar_Contracts.IEventQuery;
    membersPromise?: IPromise<WebApi_Contracts.IdentityRef[]>;
    isEdit?: boolean;
}

export class EditEventDialog extends Controls_Dialog.ModalDialogO<IEventDialogOptions> {
    private _$container: JQuery;
    private _calendarEvent: Calendar_Contracts.CalendarEvent;
    private __etag: number;
    private _source: Calendar_Contracts.IEventSource;
    private _content: Calendar_Contracts.IDialogContent;

    public initializeOptions(options?: any) {
        super.initializeOptions($.extend(options, { "height": 360 }));
    }

    public initialize() {
        super.initialize();
        this._calendarEvent = this._options.calendarEvent;
        this.__etag = this._calendarEvent.__etag;
        this._source = this._options.source;
        this._createLayout();
    }

    /**
     * Processes the data that the user has entered and either
     * shows an error message, or returns the edited note.
     */
    public onOkClick(): any {
        if(this._content){
            this._content.onOkClick().then((event) => {
                this._calendarEvent = event;
                this._calendarEvent.__etag = this.__etag;
                this.processResult(this._calendarEvent);
            });
        }
        else{
            this.onCancelClick();         
        }
    }
        
    private _validate(valid: boolean){
        this.updateOkButton(valid);
    }
    
    private _createDefault(){
        this._content = <EditEventControl<IEventControlOptions>>Controls.Control.createIn(EditEventControl, this._$container, {
            calendarEvent: this._calendarEvent,
            isEdit: this._options.isEdit,
            validStateChangedHandler: this._validate.bind(this),
        })
        this._content.getTitle().then((title: string) => {
            this.setTitle(title);
        });
    }
    
    private _createLayout() {
        this._$container = $(domElem('div')).addClass('edit-event-container').appendTo(this._element);
        var membersPromise = this._options.membersPromise;
        var content;
        if(this._source.getEnhancer) {
            this._source.getEnhancer().then((enhancer) => {
                var options = {
                    calendarEvent: this._calendarEvent,
                    isEdit: this._options.isEdit,
                    categoriesPromise: this._source.getCategories.bind(this, this._options.query),
                    validStateChangedHandler: this._validate.bind(this),
                    membersPromise: this._options.membersPromise,
                    getIterations: !!(<any>this._source).getIterations ? (<any>this._source).getIterations.bind(this._source) : null,
                };
                Controls_Contributions.createContributedControl<Calendar_Contracts.IDialogContent>(
                    this._$container,
                    enhancer.addDialogId,
                    options,
                    Context.getDefaultWebContext()             
                ).then((control: Calendar_Contracts.IDialogContent) => {
                    try {
                        this._content = control;
                        this._content.getTitle().then((title: string) => {
                            this.setTitle(title);
                        });
                    }
                    catch (error) {
                        this._createDefault();
                    }
                }, (error) => {
                    this._createDefault();
                });
            }, (error) => {
                this._createDefault();
            });
        }
        else{
            this._createDefault();
        }
    }  
}

/**
 * Base class for a control to insert in the dialog which allows users to add a new event or edit an existing one
*/
export class EditEventControl<TOptions extends IEventControlOptions> extends Controls.Control<TOptions> implements Calendar_Contracts.IDialogContent {
    private _$container: JQuery;
    protected _$startInput: JQuery;
    protected _$endInput: JQuery;
    protected _$descriptionInput: JQuery;

    private _eventValidationError: Controls_Notifications.MessageAreaControl;
    protected _calendarEvent: Calendar_Contracts.CalendarEvent;
    
    protected _isValid: boolean;
    protected _onValidChange: (valid: boolean) => any;

    public initialize() {
        super.initialize();
        this._calendarEvent = this._options.calendarEvent;
        this._isValid = false;
        this._onValidChange = this._options.validStateChangedHandler;
        this._createLayout();
    }

    /**
     * Returns the title of the control
     */
    public getTitle(): IPromise<string> {
        return Q.resolve(this._options.title || (this._options.isEdit ? "Edit Event" : "Add Event"));
    }

    /**
     * Processes the data that the user has entered and either
     * shows an error message, or returns the edited note.
     */
    public onOkClick(): any {
        this._calendarEvent.startDate = Utils_Date.shiftToLocal(Utils_Date.parseDateString(this._$startInput.val(), Culture.getDateTimeFormat().ShortDatePattern, true)).toISOString();
        this._calendarEvent.endDate = Utils_Date.shiftToLocal(Utils_Date.parseDateString(this._$endInput.val(), Culture.getDateTimeFormat().ShortDatePattern, true)).toISOString();
        this._calendarEvent.description = this._$descriptionInput.val();

        this._buildCalendarEventFromFields();

        return Q.resolve(this._calendarEvent);
    }

    protected _buildCalendarEventFromFields() {
    }

    protected _createLayout() {
        this._$container = $(domElem('div')).addClass('edit-event-container').appendTo(this._element);

        this._eventValidationError = <Controls_Notifications.MessageAreaControl>Controls.BaseControl.createIn(Controls_Notifications.MessageAreaControl, this._$container, { closeable: false });

        var $editControl = $(domElem('div', 'event-edit-control'));
        var $fieldsContainer = $(domElem('table')).appendTo($editControl);

        var startDateString = Utils_Date.localeFormat(Utils_Date.shiftToUTC(new Date(this._calendarEvent.startDate)), Culture.getDateTimeFormat().ShortDatePattern, true);
        var endDateString = startDateString;
        if (this._calendarEvent.endDate) {
            endDateString = Utils_Date.localeFormat(Utils_Date.shiftToUTC(new Date(this._calendarEvent.endDate)), Culture.getDateTimeFormat().ShortDatePattern, true);
        }

        this._$startInput = $("<input type='text' id='fieldStartDate' />").val(startDateString)
            .bind("blur",(e) => {
                this._checkValid();
        });
        this._$endInput = $("<input type='text' id='fieldEndDate' />")
            .bind("blur",(e) => {
                this._checkValid();
            });


        this._$endInput.val(endDateString);
        
        var descriptionString = this._calendarEvent.description || "";
        this._$descriptionInput = $("<textarea rows='3' id='descriptionText' class='event-textarea' />").val(descriptionString);

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

        var startCombo = <Controls_Combos.Combo>Controls.Enhancement.enhance(Controls_Combos.Combo, this._$startInput, {
            type: "date-time"
        });
        
        var endCombo = <Controls_Combos.Combo>Controls.Enhancement.enhance(Controls_Combos.Combo, this._$endInput, {
            type: "date-time"
        });

        this._setupValidators(this._$startInput, "Start date must be a valid date");
        this._setupValidators(this._$endInput, "End date must be a valid date", "End date must be equal to or after start date", this._$startInput, DateComparisonOptions.GREATER_OR_EQUAL);
        this._checkValid();
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
        fields.push(["Start Date", this._$startInput]);
        fields.push(["End Date", this._$endInput]);
        fields.push(["Description", this._$descriptionInput]);

        return fields;
    }

    protected _checkValid() {
        this._isValid = true;
        this._onValidChange(this._isValid);
    }
    
    private _formatDateValue(date: Date): string {
        return date === null ? "" : Utils_Date.format(new Date(date.valueOf()), "d");
    }

    private _parseDateValue(date: string): Date {
        return date === null ? null : Utils_Date.parseDateString(date, "d", true);
    }

    protected _setError(errorMessage: string) {
        this._eventValidationError.setError($("<span />").html(errorMessage));
    }

    protected _clearError() {
        this._eventValidationError.clear();
    }
}

/**
 * A control which allows users to add / edit free-form events.
 * In addition to start/end dates, allows user to enter a title and select a category.
*/
export class EditFreeFormEventControl<TOptions extends IFreeFormEventControlOptions> extends EditEventControl<TOptions> {
    private _$titleInput: JQuery;
    private _$categoryInput: JQuery;
    private _categories: Calendar_Contracts.IEventCategory[];
    
    public initialize() {
        super.initialize();
        if (this._calendarEvent.title) {
            this._$titleInput.val(this._calendarEvent.title);
            this._checkValid();
        }

        this._$categoryInput.val(this._calendarEvent.category ? this._calendarEvent.category.title : "");
    }
    
    protected _buildCalendarEventFromFields() {
        this._calendarEvent.movable = true;
        this._calendarEvent.title = $.trim(this._$titleInput.val());
        if (this._categories) {
            var categoryTitle = $.trim(this._$categoryInput.val())
            var existingCategories = this._categories.filter(cat => cat.title === categoryTitle);
            if(existingCategories.length > 0) {
                this._calendarEvent.category = existingCategories[0];
            }
            else {
                if (!categoryTitle || categoryTitle === "") {
                    categoryTitle = "Uncategorized"
                }
                this._calendarEvent.category = <Calendar_Contracts.IEventCategory> {
                    title: categoryTitle
                }
            }
        }
    }

    protected _createLayout() {
        this._$titleInput = $("<input type='text' class='requiredInfoLight' id='fieldTitle'/>")
            .bind("input keyup",(e) => {
            if (e.keyCode !== Utils_UI.KeyCode.ENTER) {
                this._checkValid();
            }
        });

        this._$categoryInput = $("<input type='text' id='fieldCategory' />").addClass("field-input");
        // Populate categories
        this._options.categoriesPromise().then((categories: Calendar_Contracts.IEventCategory[]) => {
            if (categories) {
                this._categories = categories;
                Controls.Enhancement.enhance(Controls_Combos.Combo, this._$categoryInput, {
                    source: $.map(categories, (category: Calendar_Contracts.IEventCategory, index: number) => { return category.title }),
                    dropCount: 3
                });
            }
        }); 
        super._createLayout();       
    }

    protected _getFormFields(): any[] {
        var fields = [];

        fields.push(["Title", this._$titleInput]);
        fields.push(["Start Date", this._$startInput]);
        fields.push(["End Date", this._$endInput]);
        fields.push(["Category", this._$categoryInput]);
        fields.push(["Description", this._$descriptionInput]);

        return fields;
    }

    protected _checkValid() {
        var title: string = $.trim(this._$titleInput.val());
        if (title.length <= 0) {
            this._clearError();
            if(this._isValid){
                this._isValid = false;
                this._onValidChange(this._isValid);
            }
            return;
        }
        var validationResult = [];
        var isValid: boolean = Controls_Validation.validateGroup('default', validationResult);
        if (!isValid) {
            this._setError(validationResult[0].getMessage());
            if(this._isValid){
                this._isValid = false;
                this._onValidChange(this._isValid);
            }
            return;
        }

        this._clearError();
            if(!this._isValid){
                this._isValid = true;
                this._onValidChange(this._isValid);
            }
            return;
    }
}

/**
 * A control which allows users to add / edit days off events.
 * In addition to start/end dates, allows user to select a user (or the entire team).
*/
export class EditCapacityEventControl<TOptions extends ICapacityEventControlOptions> extends EditEventControl<TOptions> {
    private _$memberInput: JQuery;
    private _members: WebApi_Contracts.IdentityRef[];
    private _$iterationInput: JQuery;
    private _iterations: Work_Contracts.TeamSettingsIteration[];
    private static EVERYONE : string = "Everyone";
    
    public initialize() {
        this._options.getIterations().then((iterations) => {
            this._iterations = iterations;
            super.initialize();
        });
    }
    
    public getTitle(): IPromise<string> {        
        return Q.resolve(this._options.title || (this._options.isEdit ? "Edit Day off" : "Add Day off"));
    }
    
    protected _buildCalendarEventFromFields() {
        if(!this._options.isEdit) {
            if (this._members) {
                var displayName = $.trim(this._$memberInput.val());
                var eventMember : any;
                if (displayName === EditCapacityEventControl.EVERYONE) {
                    eventMember.displayName = displayName;
                }
                else {
                    this._members.some((member: WebApi_Contracts.IdentityRef, index: number, array: WebApi_Contracts.IdentityRef[]) => {
                        if (member.displayName === displayName) {
                            eventMember = member;
                            return true;
                        }
                        return false;
                    });
                }
                this._calendarEvent.member = eventMember;
                var memberId = eventMember.uniqueName || EditCapacityEventControl.EVERYONE;
                this._calendarEvent.id = "daysOff." + memberId + "." + new Date(this._calendarEvent.startDate.valueOf());
            }
        }   
        var iterationVal = this._$iterationInput.val();
        var iteration = this._iterations.filter(i => i.name === iterationVal)[0];
        if(iteration){
            this._calendarEvent.iterationId = iteration.id;
        }
    }

    protected _createLayout() {
        this._$memberInput = $("<input type='text' id='fieldMember' />").addClass("field-input");

        if (this._options.membersPromise) {
            this._options.membersPromise.then((members: WebApi_Contracts.IdentityRef[]) => {
                this._members = members;
                var memberNames = [];
                memberNames.push(EditCapacityEventControl.EVERYONE);
                members.sort((a, b) => { return a.displayName.toLocaleLowerCase().localeCompare(b.displayName.toLocaleLowerCase()); });
                members.forEach((member: WebApi_Contracts.IdentityRef, index: number, array: WebApi_Contracts.IdentityRef[]) => {
                    memberNames.push(member.displayName);
                });
                
                // Populate iterations
                this._$iterationInput = $("<input type='text' id='fieldIteration' />").addClass("field-input")
                    .bind("keydown", (e) => { this._checkValid() });
                Controls.Enhancement.enhance(Controls_Combos.Combo, this._$iterationInput, {
                    source: this._iterations.map((iteration: Work_Contracts.TeamSettingsIteration, index: number) => iteration.name),
                    dropCount: 3
                });
                if(this._iterations && this._iterations.length > 0){
                    this._$iterationInput.val(this._iterations[0].name);
                    var iteration;
                    if(this._calendarEvent.iterationId) {
                         iteration = this._iterations.filter(i => i.id === this._calendarEvent.iterationId)[0];
                    }
                    else{                    
                        iteration = this._getCurrentIteration(new Date(this._calendarEvent.startDate))                
                    }
                    if (iteration) {
                        this._$iterationInput.val(iteration.name);
                    }
                }


                Controls.Enhancement.enhance(Controls_Combos.Combo, this._$memberInput, {
                    source: memberNames,
                    dropCount: 3
                });
                this._$memberInput.val(this._calendarEvent.member.displayName || "");
                if (this._options.isEdit) {
                    this._$memberInput.addClass('requiredInfoLight');
                    this._$memberInput.prop('disabled', true);
                }
                super._createLayout();
            });
        }
        else {
            this._$memberInput.prop('disabled', true);
            super._createLayout();
        }
    }
        
    private _getCurrentIteration(date: Date): Work_Contracts.TeamSettingsIteration {
        return this._iterations.filter((iteration: Work_Contracts.TeamSettingsIteration, index: number, array: Work_Contracts.TeamSettingsIteration[]) => {
            if (iteration.attributes.startDate !== null && iteration.attributes.finishDate !== null && date >= Utils_Date.shiftToUTC(iteration.attributes.startDate) && date <= Utils_Date.shiftToUTC(iteration.attributes.finishDate)) {
                return true;
            }
        })[0];
    }

    protected _getFormFields(): any[] {
        var fields = [];

        fields.push(["Start Date", this._$startInput]);
        fields.push(["End Date", this._$endInput]);
        fields.push(["Team Member", this._$memberInput]);
        fields.push(["Iteration", this._$iterationInput]);
        fields.push(["Description", this._$descriptionInput]);

        return fields;
    }

    protected _checkValid() {
        var validationResult = [];
        
        // Ensure iteration is selected
        var iterationVal = this._$iterationInput.val();
        var validIteration = this._iterations.filter(i => i.name === iterationVal).length > 0;
        if(!validIteration){
            this._clearError();
            if(this._isValid){
                this._isValid = false;
                this._onValidChange(this._isValid);
            }
            return;
        }
        
        var isValid: boolean = Controls_Validation.validateGroup('default', validationResult);
        if (!isValid) {
            this._setError(validationResult[0].getMessage());
            if(this._isValid){
                this._isValid = false;
                this._onValidChange(this._isValid);
            }
            return;
        }
        
        // no error
        this._clearError();
        if(!this._isValid){
            this._isValid = true;
            this._onValidChange(this._isValid);
        }
        return;
    }
}

export interface ColorPickerPopupOptions {
    category: Calendar_Contracts.IEventCategory;
    callback: (color: string) => {};
}

export class ColorPickerPopup <ColorPickerPopupOptions> extends Controls_Popup.RichContentTooltip {
    private _category: Calendar_Contracts.IEventCategory;
    private _allColors: string[];
    private _container: JQuery;
    private _callback: (color:string) => {};
    
    public initializeOptions(options?: any) {
        super.initializeOptions($.extend({
            openCloseOnHover: false
        }, options));
    }
    
    public initialize() {
        super.initialize();
        this._category = this._options.category;
        this._callback = this._options.callback;
        this._allColors = Calendar_ColorUtils.getAllColors();
        this._element.addClass("color-picker-tooltip");
        this._container = this._element.find(".popup-content-container");
        this._drawContent();
    }
    
    private _drawContent() {
        var menuItems: Controls_Menus.IMenuItemSpec[] = [];
        for (var i = 0; i < this._allColors.length; i++){
            var color = Utils_String.format("#{0}", this._allColors[i]);
            var item = <Controls_Menus.IMenuItemSpec>{
                id: i.toString(),
                noIcon: true,
                showText: false,
                title: color,
                html: Utils_String.format("<div class='category-color-choice' style='background-color:{0}'/>", color),
                action: this._options.callback.bind(this, color)
            };
            menuItems.push(item);
        }
        Controls.BaseControl.createIn(Controls_Menus.MenuBar, this._container, <Controls_Menus.MenuOptions> {
            items: menuItems
        });
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
            fieldDate = Utils_Date.parseDateString(fieldText, this._options.parseFormat, true);
            relativeToFieldDate = Utils_Date.parseDateString(relativeToFieldText, this._options.parseFormat, true);
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