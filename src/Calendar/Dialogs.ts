import Calendar_Contracts = require("./Contracts");
import Context = require("VSS/Context");
import Controls = require("VSS/Controls");
import Controls_Contributions = require("VSS/Contributions/Controls");
import Controls_Notifications = require("VSS/Controls/Notifications");
import Controls_Combos = require("VSS/Controls/Combos");
import Controls_Dialog = require("VSS/Controls/Dialogs");
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
    membersPromise: IPromise<WebApi_Contracts.IdentityRef[]>;
    getIterations: () => IPromise<Work_Contracts.TeamSettingsIteration[]>;
    categoriesPromise: () => IPromise<Calendar_Contracts.IEventCategory[]>;
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
    private _source: Calendar_Contracts.IEventSource;
    private _contributedControl: Calendar_Contracts.IDialogContent;
    private _content: Calendar_Contracts.IAddEventContent;
    private _$contributedContent: JQuery;
    
    private _eventValidationError: Controls_Notifications.MessageAreaControl;
    private _contributionsValid: boolean;
        
    private _$titleInput: JQuery;
    private _$startInput: JQuery;
    private _$endInput: JQuery;
    private _$textInputs: JQuery[];
    private _$comboInputs: JQuery[];

    public initializeOptions(options?: any) {
        super.initializeOptions($.extend(options, {
            height: 345,
            coreCssClass: "add-event-dialog"
        }));
    }

    public initialize() {
        super.initialize();
        this._calendarEvent = this._options.calendarEvent;
        this._source = this._options.source;
        this._contributionsValid = true;
        this._$textInputs = [];
        this._$comboInputs = [];
        // default content
        this._content = <Calendar_Contracts.IAddEventContent> {
            title: true,
            start: true,
            end: true
        }
        this._createLayout();
    }

    /**
     * Processes the data that the user has entered and either
     * shows an error message, or returns the edited note.
     */
    public onOkClick(): any {
        this._buildCalendarEventFromFields().then((results) => {
            if(this._contributedControl) {
                this._contributedControl.onOkClick().then((event) => {
                    this._calendarEvent = $.extend(this._calendarEvent, event);
                    this.processResult(this._calendarEvent);
                });
            }
            else {
                this.processResult(this._calendarEvent);         
            }
        });
    }
                
    private _createLayout() {
        this._$container = $(domElem('div')).addClass('edit-event-container').appendTo(this._element);
        this._$contributedContent = $(domElem('div')).addClass('contributed-content-container').appendTo(this._element);
        var membersPromise = this._options.membersPromise;
        var content;
        if(this._source.getEnhancer) {
            this._source.getEnhancer().then((enhancer) => {
                var options = {
                    calendarEvent: this._calendarEvent,
                    isEdit: this._options.isEdit,
                    categoriesPromise: this._source.getCategories.bind(this, this._options.query),
                    validStateChangedHandler: (valid: boolean) => {
                        this._contributionsValid = valid;

                        // Erhm...
                        ((e) => {                        
                            this._validate(true);
                        });
                     },
                    membersPromise: this._options.membersPromise,
                    getIterations: !!(<any>this._source).getIterations ? (<any>this._source).getIterations.bind(this._source) : null,
                };
                Controls_Contributions.createContributedControl<Calendar_Contracts.IDialogContent>(
                    this._$contributedContent,
                    enhancer.addDialogId,
                    $.extend(options, { bowtieVersion: 2 }),
                    Context.getDefaultWebContext()             
                ).then((control: Calendar_Contracts.IDialogContent) => {
                    try {
                        this._contributedControl = control;
                        if(this._contributedControl.getContributedHeight) {
                            this._contributedControl.getContributedHeight().then((height: number) => { this._$contributedContent.find(".external-content-host").css("height", height); })
                        }
                        else {
                            this._$contributedContent.find(".external-content-host").css("height", 0)
                        }
                        this._contributedControl.getTitle().then((title: string) => {
                            this.setTitle(title);
                        });
                        this._contributedControl.getFields().then((fields: Calendar_Contracts.IAddEventContent[]) => {
                            this._content = fields;
                            this._renderContent();
                        });
                    }
                    catch (error) {
                        this._renderContent();
                    }
                }, (error) => {
                    this._renderContent();
                });
            }, (error) => {
                this._renderContent();
            });
        }
        else{
            this._renderContent();
        }
    }
    
    private _renderContent() {
        this._eventValidationError = <Controls_Notifications.MessageAreaControl>Controls.BaseControl.createIn(Controls_Notifications.MessageAreaControl, this._$container, { closeable: false });

        var $editControl = $(domElem('div', 'event-edit-control'));
        var $fieldsContainer = $(domElem('table')).appendTo($editControl);        
        
        // Build title input
        if (this._content.title) {
            this._$titleInput = $("<input type='text' class='requiredInfoLight' id='fieldTitle'/>").val(this._calendarEvent.title)
                .bind("blur", (e) => {                        
                        this._validate(true);
                    });
        }
        
        var startDateString = Utils_Date.localeFormat(Utils_Date.shiftToUTC(new Date(this._calendarEvent.startDate)), Culture.getDateTimeFormat().ShortDatePattern, true);        
        var endDateString = startDateString;
        
        // Build start input
        if (this._content.start) {
            this._$startInput = $("<input type='text' id='fieldStartDate' />").val(startDateString)
                .bind("blur", (e) => {                        
                        this._validate(true);
                    });       
        }
        
        // Build end input
        if (this._content.end) {
            if (this._calendarEvent.endDate) {
                endDateString = Utils_Date.localeFormat(Utils_Date.shiftToUTC(new Date(this._calendarEvent.endDate)), Culture.getDateTimeFormat().ShortDatePattern, true);
            }
            this._$endInput = $("<input type='text' id='fieldEndDate' />")
                .bind("blur", (e) => {                        
                        this._validate(true);
                    });   
            this._$endInput.val(endDateString);
        }
        
        // Build text inputs
        if (this._content.textFields) {
            var textFields = this._content.textFields;
            for (var i = 0; i < textFields.length; i++) {
                var textField = textFields[i]
                var textInput = $(Utils_String.format("<input type='text' class='requiredInfoLight' id='field{0}'/>", textField.label));
                if (textField.checkValid) {
                    textInput.bind("blur", (e) => {                        
                        this._validate(true);
                    });
                }
                if (textField.initialValue) {
                    textInput.val(textField.initialValue);
                }
                if (textField.disabled) {
                    textInput.prop("disabled", true);
                }
                this._$textInputs.push(textInput);           
            }
        }
        
        // Build combo inputs
        if (this._content.comboFields) {
            var comboFields = this._content.comboFields;
            for (var i = 0; i < comboFields.length; i++) {
                var comboField = comboFields[i];
                var comboInput = $(Utils_String.format("<input type='text' class='requiredInfoLight' id='field{0}' />", comboField.label));
                
                if (comboField.initialValue) {
                    comboInput.val(comboField.initialValue);
                }
                if (comboField.checkValid) {                    
                    comboInput.bind("blur",(e) => {
                        this._validate(true);
                    });
                }
                
                this._$comboInputs.push(comboInput);
            }
        }
        
        // Populate fields container with fields. The form fields array contain pairs of field label and field element itself.
        var fields = this._getFormFields();
        for (var i = 0, l = fields.length; i < l; i += 1) {
            var labelName = fields[i][0];
            var field = fields[i][1];
            if(field) {
                if (i === 0) {
                    field.attr("autofocus", true);
                }
                var $row = $(domElem("tr"));

                var fieldId = field.attr("id") || $("input", field).attr("id");
                $(domElem("label")).attr("for", fieldId).text(labelName).appendTo($(domElem("td", "label")).appendTo($row));

                field.appendTo($(domElem("td"))
                    .appendTo($row));

                $row.appendTo($fieldsContainer);
            }
        }
        
        this._$container.append($editControl).bind("keyup", (event) => {
            if (event.keyCode == Utils_UI.KeyCode.ENTER) {
                this.onOkClick();
            }
            else if (event.keyCode == Utils_UI.KeyCode.ESCAPE) {
                this.onClose();
            }
        });
        
        if (this._content.comboFields) {
            this._buildComboControls();
        }
        
        // Add date pickers combos to DOM                 
        <Controls_Combos.Combo>Controls.Enhancement.enhance(Controls_Combos.Combo, this._$startInput, {
            type: "date-time"
        });
                
        <Controls_Combos.Combo>Controls.Enhancement.enhance(Controls_Combos.Combo, this._$endInput, {
            type: "date-time"
        });      
        
        this._setupValidators();
        this._validate();
    }
    
    private _setupValidators() {
        this._setupRequiredValidators(this._$titleInput, "Title cannot be empty");
        this._setupDateValidators(this._$startInput, "Start date must be a valid date");
        this._setupDateValidators(this._$endInput, "End date must be a valid date", "End date must be equal to or after start date", this._$startInput, DateComparisonOptions.GREATER_OR_EQUAL);
        
        // push text input fields
        for (var i = 0; i < this._$textInputs.length; i++) {
            var textField = this._content.textFields[i]
            if (!textField.checkValid && textField.requiredField) {
                this._setupRequiredValidators(this._$textInputs[i], textField.validationErrorMessage)
            }
        }
        // push combo input fields
        for (var i = 0; i < this._$comboInputs.length; i++) {
            var comboField = this._content.comboFields[i]
            if (!comboField.checkValid && comboField.requiredField) {
                this._setupRequiredValidators(this._$comboInputs[i], comboField.validationErrorMessage)
            }
        }
    }
    
    private _getFormFields(): any[]
    {
        var fields = [];
        // push basic fields          
        fields.push(["Title", this._$titleInput]);
        fields.push(["Start Date", this._$startInput]);
        fields.push(["End Date", this._$endInput]);
        
        // push text input fields
        var textFields = this._content.textFields;
        var textInputs = this._$textInputs;
        for (var i = 0; i < textInputs.length; i++) {
            fields.push([textFields[i].label, textInputs[i]]);
        }
        // push combo input fields
        var comboFields = this._content.comboFields;
        var comboInputs = this._$comboInputs;
        for (var i = 0; i < comboInputs.length; i++)         {
            fields.push([comboFields[i].label, comboInputs[i]]);
        }
        return fields;
    }
    
    private _buildComboControls() {
        var comboFields = this._content.comboFields;
        var comboInputs = this._$comboInputs;
        for (var i = 0; i < comboInputs.length; i++) {      
            if (comboFields[i].disabled) {
                comboInputs[i].prop("disabled", true);
            }
            else {
                Controls.Enhancement.enhance(Controls_Combos.Combo, comboInputs[i], {
                    source: comboFields[i].items,
                    dropCount: 3
                });
            }
        }
    }
    
    private _buildCalendarEventFromFields(): IPromise<any> {
        if (this._$startInput) {
            this._calendarEvent.startDate =  Utils_Date.shiftToLocal(Utils_Date.parseDateString(this._$startInput.val(), Culture.getDateTimeFormat().ShortDatePattern, true)).toISOString();
        }
        if (this._$endInput) {            
            this._calendarEvent.endDate = Utils_Date.shiftToLocal(Utils_Date.parseDateString(this._$endInput.val(), Culture.getDateTimeFormat().ShortDatePattern, true)).toISOString();
        }
        if (this._$titleInput) {
            this._calendarEvent.title = $.trim(this._$titleInput.val());
        }
        
        var promises: IPromise<any>[] = [];
        // create event data from text fields
        var textFields = this._content.textFields;
        var textInputs = this._$textInputs;
        for (var i = 0; i < textInputs.length; i++) {
            if (textFields[i].okCallback) {
                promises.push(textFields[i].okCallback(textInputs[i].val()));
            }
            else if (textFields[i].eventProperty) {
                this._calendarEvent[textFields[i].eventProperty] = $.trim(textInputs[i].val());
            }
        }
        // create event data from combo fields
        var comboFields = this._content.comboFields;
        var comboInputs = this._$comboInputs;
        for (var i = 0; i < comboInputs.length; i++) {
            if (comboFields[i].okCallback) {
                promises.push(comboFields[i].okCallback(comboInputs[i].val()));
            }
            else if (comboFields[i].eventProperty) {
                this._calendarEvent[comboFields[i].eventProperty] = $.trim(comboInputs[i].val());
            }
        }
        
        return Q.all(promises);
    }
        
    private _validate(showError?: boolean) {
        if(!this._contributionsValid) {
            this._clearError();
            this.updateOkButton(false);
            return;
        }
        
        var validationResult = [];
        var groupIsValid: boolean = Controls_Validation.validateGroup('default', validationResult);
        if (!groupIsValid) {
            if(showError) {
                this._setError(validationResult[0].getMessage());
            }
            this.updateOkButton(false);
            return;
        }
        
        var errorPromises: IPromise<string>[] = []
        
        // validate text input fields
        var textFields = this._content.textFields;
        var textInputs = this._$textInputs;
        for (var i = 0; i < textInputs.length; i++) {
            var textField = textFields[i]
            if (textField.checkValid) {
                errorPromises.push(textField.checkValid($.trim(textInputs[i].val())).then((isValid: boolean) => {
                    if(!isValid) {
                        return textField.validationErrorMessage;
                    }
                    return "valid"
                }));
            }
        }
        // validate combo input fields
        var comboFields = this._content.comboFields;
        var comboInputs = this._$comboInputs;
        for (var i = 0; i < comboInputs.length; i++) {
            var comboField = comboFields[i]
            if (comboField.checkValid && !comboField.disabled) {
                errorPromises.push(comboField.checkValid($.trim(comboInputs[i].val())).then((isValid: boolean) => {
                    if(!isValid) {
                        return comboField.validationErrorMessage;
                    }
                    return "valid"
                }));
            }
        }
        
        return Q.all(errorPromises).then((results: string[]) => {
            var invalidMessages = results.filter(r => r !== "valid");
            if (invalidMessages && invalidMessages.length > 0) {
                if(showError) {
                    this._setError(invalidMessages[0]);
                }             
                this.updateOkButton(false);
                return;
            }
            this._clearError();
            this.updateOkButton(true);
        })
    }
        
    private _setupDateValidators($field: JQuery, validDateFormatMessage: string, relativeToErrorMessage?: string, $relativeToField?: JQuery, dateComparisonOptions?: DateComparisonOptions) {
        <Controls_Validation.DateValidator<Controls_Validation.DateValidatorOptions>>Controls.Enhancement.enhance(Controls_Validation.DateValidator, $field, {
            invalidCssClass: "date-invalid",
            group: "default",
            message: validDateFormatMessage,
            parseFormat: Culture.getDateTimeFormat().ShortDatePattern
        });
        
        this._setupRequiredValidators($field, validDateFormatMessage);
        
        if (relativeToErrorMessage) {
            <DateRelativeToValidator>Controls.Enhancement.enhance(DateRelativeToValidator, $field, {
                comparison: dateComparisonOptions,
                relativeToField: $relativeToField,
                group: "default",
                message: relativeToErrorMessage,
                parseFormat: Culture.getDateTimeFormat().ShortDatePattern
            });
        }
    }
    
    private _setupRequiredValidators($field: JQuery, requiredFieldMessage: string) {
        <Controls_Validation.RequiredValidator<Controls_Validation.BaseValidatorOptions>>Controls.Enhancement.enhance(Controls_Validation.RequiredValidator, $field, 
        <Controls_Validation.BaseValidatorOptions>{
            invalidCssClass: "field-invalid",
            group: "default",
            message: requiredFieldMessage
        });
    }
    
    private _setupCustomValidators($field: JQuery, validationFunction: (value: string) => boolean, invalidInputMessage: string) {
        <Controls_Validation.CustomValidator<Controls_Validation.CustomValidatorOptions>>Controls.Enhancement.enhance(Controls_Validation.CustomValidator, $field, 
        <Controls_Validation.CustomValidatorOptions>{
            invalidCssClass: "field-invalid",
            group: "default",
            message: invalidInputMessage,
            validate: validationFunction
        });
    }
        
    private _setError(errorMessage: string) {
        if (errorMessage && errorMessage.length !== 0 && errorMessage !== "invalid") {
            this._eventValidationError.setError($("<span />").html(errorMessage));
        }
        else {
            this._clearError();
        }
    }

    private _clearError() {
        this._eventValidationError.clear();
    }
}

/**
 * A control which allows users to add / edit free-form events.
 * In addition to start/end dates, allows user to enter a title and select a category.
*/
export class EditFreeFormEventControl extends Controls.Control<IEventControlOptions> implements Calendar_Contracts.IDialogContent {
    
    private _calendarEvent: Calendar_Contracts.CalendarEvent;
    private _categories: Calendar_Contracts.IEventCategory[];
    private _$descriptionInput: JQuery;
    private static height: number = 50;
        
    public initializeOptions(options?: any) {
        super.initializeOptions($.extend(options, {
            coreCssClass: "edit-freeform-control"
        }));
    }
    
    public initialize() {
        super.initialize();
        this._categories = [];
        this._element.addClass("bowtie-style");
        this._calendarEvent = this._options.calendarEvent;
        this._createLayout();
    }

    public onOkClick(): IPromise<any> {
        return Q.resolve({
            movable: true,
            description: this._$descriptionInput.val(),
            category: this._calendarEvent.category
        });
    }
    
    public getTitle(): IPromise<string> {        
        return Q.resolve(this._options.isEdit ? "Edit Event" : "Add Event");
    }
    
    public getContributedHeight(): IPromise<number> {
        return Q.resolve(EditFreeFormEventControl.height);
    }
    
    public getFields(): IPromise<Calendar_Contracts.IAddEventContent> {
        return this._options.categoriesPromise().then((categories: Calendar_Contracts.IEventCategory[]) => {
            if (categories) {
                this._categories = categories;
                var categoryTitles = categories.map(cat => { return cat.title });
            }
            return (<Calendar_Contracts.IAddEventContent> {
                title: true,
                start: true,
                end: true,
                comboFields: [
                    <Calendar_Contracts.IAddEventComboField> {
                        label: "Category",
                        initialValue: this._calendarEvent.category ? this._calendarEvent.category.title : "",
                        items: categoryTitles || [],
                        okCallback: this._categoryCallback.bind(this)
                    }
                ]
            });
        });
    }
        
    private _createLayout() {
        var $container = $(domElem('table')).appendTo(this._element);
        
        var descriptionString = this._calendarEvent.description || "";
        this._$descriptionInput = $("<textarea rows='3' id='descriptionText' class=''requiredInfoLight' />").val(descriptionString);  
        
        var $row = $(domElem("tr"));

        $(domElem("label")).attr("for", this._$descriptionInput.attr("id")).text("Description").appendTo($(domElem("td", "label")).appendTo($row));

        this._$descriptionInput.appendTo($(domElem("td"))
            .appendTo($row));

        $row.appendTo($container);
    }
    
    private _categoryCallback(value: string): IPromise<any> {
        var title = $.trim(value);
        if(!title || title.length === 0) {
            title = "Uncategorized";
        }
        var category: Calendar_Contracts.IEventCategory = {
            title: title,
            id: Utils_String.format("freeForm.{0}", title)
        }
        var categories = this._categories.filter(cat => cat.title === value);
        if(categories && categories.length > 0) {
            category = categories[0];
        }
        this._calendarEvent.category = category;
        return Q.resolve(null);        
    }
}

/**
 * A control which allows users to add / edit days off events.
 * In addition to start/end dates, allows user to select a user (or the entire team).
*/
export class EditCapacityEventControl extends Controls.Control<IEventControlOptions> implements Calendar_Contracts.IDialogContent{

    private _calendarEvent: Calendar_Contracts.CalendarEvent;
    private _members: WebApi_Contracts.IdentityRef[];
    private _iterations: Work_Contracts.TeamSettingsIteration[];
    private static EVERYONE: string = "Everyone";
    private static height: number = 50;
    
    public initializeOptions(options?: any) {
        super.initializeOptions($.extend(options, {
            coreCssClass: "edit-capacity-control"
        }));
    }
    
    public initialize() {
        super.initialize();
        this._element.addClass("bowtie-style");
        this._calendarEvent = this._options.calendarEvent;
        this._members = [];
        this._iterations = [];
        this._createLayout();
    }

    public onOkClick(): any {
        return Q.resolve({
            iterationId: this._calendarEvent.iterationId,
            member: this._calendarEvent.member
        });
    }
    
    public getTitle(): IPromise<string> {        
        return Q.resolve(this._options.isEdit ? "Edit Day off" : "Add Day off");
    }
    
    public getFields(): IPromise<Calendar_Contracts.IAddEventContent> {
        return this._options.membersPromise.then((members: WebApi_Contracts.IdentityRef[]) => {
            this._members = members;
            var memberNames = [];
            memberNames.push(EditCapacityEventControl.EVERYONE);
            members.sort((a, b) => { return a.displayName.toLocaleLowerCase().localeCompare(b.displayName.toLocaleLowerCase()); });
            members.forEach((member: WebApi_Contracts.IdentityRef, index: number, array: WebApi_Contracts.IdentityRef[]) => {
                memberNames.push(member.displayName);
            });
            
            var initialMemberValue = this._calendarEvent.member.displayName || "";
            var disabled = false;
            if (this._options.isEdit) {
                disabled = true;
            }
            
            return this._options.getIterations().then((iterations) => {
                this._iterations = iterations;
                var initialIterationValue = iterations[0].name;
                var iteration;
                if(this._calendarEvent.iterationId) {
                        iteration = iterations.filter(i => i.id === this._calendarEvent.iterationId)[0];
                }
                else{                    
                    iteration = this._getCurrentIteration(iterations, new Date(this._calendarEvent.startDate))                
                }
                if (iteration) {
                    initialIterationValue = iteration.name;
                }
                return <Calendar_Contracts.IAddEventContent> {
                    start: true,
                    end: true,
                    comboFields: [
                        <Calendar_Contracts.IAddEventComboField> {
                            label: "Team Member", 
                            initialValue: initialMemberValue,
                            items: memberNames,
                            disabled: disabled,
                            checkValid: this._checkMemberValid.bind(this),
                            okCallback: this._memberCallback.bind(this)
                        },
                        <Calendar_Contracts.IAddEventComboField> {
                            label: "Iteration",
                            initialValue: initialIterationValue,
                            items: iterations.map(iteration => { return iteration.name }),
                            checkValid: this._checkIterationValid.bind(this),
                            okCallback: this._iterationCallback.bind(this)
                        }
                    ]
                };
            });
        });
    }
    
    private _createLayout() {
        //no op
    }
        
    private _getCurrentIteration(iterations: Work_Contracts.TeamSettingsIteration[], date: Date): Work_Contracts.TeamSettingsIteration {
        return iterations.filter((iteration: Work_Contracts.TeamSettingsIteration, index: number, array: Work_Contracts.TeamSettingsIteration[]) => {
            if (iteration.attributes.startDate !== null && iteration.attributes.finishDate !== null && date.valueOf() >= iteration.attributes.startDate.valueOf() && date.valueOf() <= iteration.attributes.finishDate.valueOf()) {
                return true;
            }
        })[0];
    }
    
    private _iterationCallback(value: string): IPromise<any> {
        var iteration = this._iterations.filter(i => i.name === $.trim(value))[0];
        if (iteration) {
            this._calendarEvent.iterationId = iteration.id;
        }
        return Q.resolve(null);
    }
    
    private _checkIterationValid(value: string): IPromise<boolean> {
        return Q.resolve(this._iterations.filter(i => i.name === $.trim(value)).length > 0);
    }
    
    private _memberCallback(value: string): IPromise<any>{
        var member = this._members.filter(m => m.displayName === $.trim(value))[0];
        if (member) {
            this._calendarEvent.member = member;
        }
        else {
            this._calendarEvent.member = <WebApi_Contracts.IdentityRef> { displayName: EditCapacityEventControl.EVERYONE };
        }
        return Q.resolve(null);
    }
    
    private _checkMemberValid(value: string): IPromise<boolean> {        
        return Q.resolve(this._members.filter(m => m.displayName === $.trim(value)).length > 0 || value === EditCapacityEventControl.EVERYONE);
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

        if ((fieldDate instanceof Date) && !isNaN(fieldDate.getTime()) && relativeToFieldDate instanceof Date && !isNaN(relativeToFieldDate.getTime())) {
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