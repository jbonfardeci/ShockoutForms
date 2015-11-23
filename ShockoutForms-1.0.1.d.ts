/// <reference path="C:/Users/jbonfardeci/Source/Repos/ShockoutForms/typings/knockout.d.ts" />
/// <reference path="C:/Users/jbonfardeci/Source/Repos/ShockoutForms/typings/jquery.d.ts" />
/// <reference path="C:/Users/jbonfardeci/Source/Repos/ShockoutForms/typings/jquery.ui.datetimepicker.d.ts" />
/// <reference path="C:/Users/jbonfardeci/Source/Repos/ShockoutForms/typings/jqueryui.d.ts" />
declare module Shockout {
    interface IViewModel {
        Id: KnockoutObservable<number>;
        Created: KnockoutObservable<Date>;
        CreatedBy: KnockoutObservable<ISpPerson>;
        Modified: KnockoutObservable<Date>;
        ModifiedBy: KnockoutObservable<ISpPerson>;
        parent: Shockout.SPForm;
        historyItems: KnockoutObservable<Array<IHistoryItem>>;
        attachments: KnockoutObservable<Array<any>>;
        currentUser: KnockoutObservable<any>;
        isValid: KnockoutComputed<boolean>;
        showUserProfiles: KnockoutObservable<boolean>;
        isAuthor(): boolean;
        deleteItem(): void;
        cancel(): void;
        print(): void;
        deleteAttachment(obj: any, event: any): boolean;
        save(model: ViewModel, btn: HTMLElement): void;
        submit(model: ViewModel, btn: HTMLElement): void;
        _allowSave: KnockoutObservable<boolean>;
        _allowPrint: KnockoutObservable<boolean>;
        _allowDelete: KnockoutObservable<boolean>;
        allowSave(): boolean;
        allowPrint(): boolean;
        allowDelete(): boolean;
    }
    class ViewModel implements IViewModel {
        static historyKey: string;
        static historyDescriptionKey: string;
        static historyDateKey: string;
        static isSubmittedKey: string;
        static parent: SPForm;
        Id: KnockoutObservable<number>;
        Created: KnockoutObservable<Date>;
        CreatedBy: KnockoutObservable<ISpPerson>;
        Modified: KnockoutObservable<Date>;
        ModifiedBy: KnockoutObservable<ISpPerson>;
        showUserProfiles: KnockoutObservable<boolean>;
        parent: Shockout.SPForm;
        historyItems: KnockoutObservable<Array<any>>;
        attachments: KnockoutObservableArray<any>;
        currentUser: KnockoutObservable<ICurrentUser>;
        isValid: KnockoutComputed<boolean>;
        deleteAttachment: any;
        _allowSave: KnockoutObservable<boolean>;
        _allowPrint: KnockoutObservable<boolean>;
        _allowDelete: KnockoutObservable<boolean>;
        constructor(instance: Shockout.SPForm);
        isAuthor(): boolean;
        deleteItem(): void;
        cancel(): void;
        print(): void;
        save(model: ViewModel, btn: HTMLElement): void;
        submit(model: ViewModel, btn: HTMLElement): void;
        allowSave(): boolean;
        allowPrint(): boolean;
        allowDelete(): boolean;
    }
}
declare module Shockout {
    class KoHandlers {
        static bindKoHandlers(): void;
    }
}
declare module Shockout {
    class KoComponents {
        static registerKoComponents(): void;
        private static hasErrorCssDiv;
        private static requiredFeedbackSpan;
        static soStaticFieldTemplate: string;
        static soTextFieldTemplate: string;
        static soHtmlFieldTemplate: string;
        static soCheckboxFieldTemplate: string;
        static soSelectFieldTemplate: string;
        static soCheckboxGroupTemplate: string;
        static soRadioGroupTemplate: string;
        static soUsermultiFieldTemplate: string;
        static soCreatedModifiedTemplate: string;
        static soWorkflowHistoryTemplate: string;
    }
}
declare module Shockout {
    class SpApi {
        /**
        * Search the User Information list.
        * @param term: string
        * @param callback: Function
        * @param take?: number = 10
        * @return void
        */
        static peopleSearch(term: string, callback: Function, take?: number): void;
        /**
        * Get a person by their ID from the User Information list.
        * @param id: number
        * @param callback: Function
        * @return void
        */
        static getPersonById(id: number, callback: Function): void;
        static executeRestRequest(url: string, callback: JQueryPromiseCallback<any>, cache?: boolean, type?: string): void;
        /**
        * Get list item via REST services.
        * @param uri: string
        * @param done: JQueryPromiseCallback<any>
        * @param fail?: JQueryPromiseCallback<any> = undefined
        * @param always?: JQueryPromiseCallback<any> = undefined
        * @return void
        */
        static getListItem(listName: string, itemId: number, callback: Function, siteUrl?: string, cache?: boolean, expand?: string): void;
        /**
        * Get list item via REST services.
        * @param uri: string
        * @param done: JQueryPromiseCallback<any>
        * @param fail?: JQueryPromiseCallback<any> = undefined
        * @param always?: JQueryPromiseCallback<any> = undefined
        * @return void
        */
        static getListItems(listName: string, callback: Function, siteUrl?: string, filter?: string, select?: string, orderby?: string, top?: number, cache?: boolean): void;
        static insertListItem(url: string, callback: Function, data?: any): void;
        static updateListItem(item: ISpItem, callback: Function, data?: any): void;
        /**
        * Delete the list item.
        * @param model: IViewModel
        * @param callback?: Function = undefined
        * @return void
        */
        static deleteListItem(item: ISpItem, callback: JQueryPromiseCallback<any>): void;
        /**
        * Delete an attachment.
        */
        static deleteAttachment(att: ISpAttachment, callback: Function): void;
    }
}
declare module Shockout {
    class SpApi15 {
        static getCurrentUser(callback: Function): void;
        static getUsersGroups(userId: number, callback: JQueryPromiseCallback<any>): void;
    }
}
declare module Shockout {
    class SpSoap {
        static getCurrentUser(callback: Function): void;
        static getUsersGroups(loginName: string, callback: Function): void;
        static getListItems(siteUrl: string, listName: string, viewFields: string, query: string, callback: Function, rowLimit?: number): void;
        static getList(siteUrl: string, listName: string, callback: Function): void;
        static checkInFile(pageUrl: string, checkinType: string, comment?: string): void;
        static checkOutFile(pageUrl: string, checkoutToLocal: string, lastmodified: string): void;
        /**
        * Execute SOAP Request
        * @param action: string
        * @param packet: string
        * @param params: Array<any>
        * param self?: SPForm = undefined
        * @param callback?: Function = undefined
        * @param service?: string = 'lists.asmx'
        * @return void
        */
        static executeSoapRequest(action: string, packet: string, params: Array<any>, siteUrl?: string, callback?: Function, service?: string): void;
        /**
        * Update list item via SOAP services.
        * @param listName: string
        * @param fields: Array<Array<any>>
        * @param isNew?: boolean = true
        * param callback?: Function = undefined
        * @param self: SPForm = undefined
        * @return void
        */
        static updateListItem(itemId: number, listName: string, fields: Array<Array<any>>, isNew?: boolean, siteUrl?: string, callback?: Function): void;
        static searchPrincipals(term: string, callback: Function, maxResults?: number, principalType?: string): void;
    }
}
declare module Shockout {
    interface IShockoutObservable<T> extends KnockoutObservable<T> {
        _koName: string;
        _displayName: string;
        _name: string;
        _format: string;
        _required: boolean;
        _readOnly: boolean;
        _description: string;
        _type: string;
        _choices: Array<any>;
        _options: Array<any>;
        _isFillInChoice: boolean;
        _multiChoice: boolean;
    }
    interface IHistoryItem {
        _description: string;
        _dateOccurred: Date;
    }
    class HistoryItem implements IHistoryItem {
        _description: string;
        _dateOccurred: Date;
        constructor(d: string, date: Date);
    }
    interface IFileUploaderSettings {
        element: HTMLElement;
        action: string;
        debug: boolean;
        multiple: boolean;
        maxConnections: number;
        allowedExtensions: Array<string>;
        params: any;
        onSubmit(id: any, fileName: any): any;
        onComplete(id: any, fileName: any, json: any): any;
        template: string;
    }
    interface ISpGroup {
        id: number;
        name: string;
    }
    interface ICurrentUser {
        id: number;
        title: string;
        login: string;
        email: string;
        account: string;
        jobtitle: string;
        department: string;
        isAdmin: boolean;
        groups: Array<ISpGroup>;
    }
    /**
    * JSON "d" wrapper returned from SharePoint /_vti_bin/listdata.svc
    *   Prevents malicious scripts from executing
    */
    interface ISpCollectionWrapper<T> {
        d: ISpCollection<T>;
    }
    interface ISpWrapper<T> {
        d: T;
    }
    interface ISpDeferred {
        uri: string;
    }
    interface ISpDeferred {
        __deferred: ISpDeferred;
    }
    interface ISpMetadata {
        uri: string;
        etag: string;
        type: string;
    }
    interface ISpAttachmentMetadata {
        uri: string;
        type: string;
        edit_media: string;
        media_src: string;
        content_type: string;
        media_etag: string;
    }
    interface ISpCollection<T> {
        results: Array<T>;
    }
    interface ISpPersonSearchResult {
        __metadata: ISpMetadata;
        Id: number;
        Account: string;
        Name: string;
        WorkEMail: string;
    }
    interface ISpPerson {
        __metadata: ISpMetadata;
        ContentTypeID: string;
        Name: string;
        Account: string;
        WorkEMail: string;
        AboutMe: string;
        SIPAddress: string;
        IsSiteAdmin: boolean;
        Deleted: boolean;
        Picture: string;
        Department: string;
        Title: string;
        MobilePhone: string;
        FirstName: string;
        LastName: string;
        WorkPhone: string;
        UserName: string;
        WebSite: string;
        AskMeAbout: string;
        Office: string;
        Id: number;
        ContentType: string;
        Modified: string;
        Created: string;
        CreatedBy: ISpDeferred;
        CreatedById: number;
        ModifiedById: number;
        Owshiddenversion: number;
        Version: string;
        Attachments: ISpDeferred;
        Path: string;
    }
    interface ISpAttachment {
        __metadata: ISpAttachmentMetadata;
        EntitySet: string;
        ItemId: number;
        Name: string;
    }
    class SpAttachment implements ISpAttachment {
        __metadata: ISpAttachmentMetadata;
        EntitySet: string;
        ItemId: number;
        Name: string;
        constructor(rootUrl: string, siteUrl: string, listName: string, itemId: number, fileName: string);
    }
    interface ISpItem {
        __metadata: ISpMetadata;
        Title: string;
        ContentTypeID: string;
        Id: number;
        ContentType: string;
        Modified: any;
        Created: any;
        CreatedBy: ISpPerson;
        CreatedById: number;
        ModifiedBy: ISpPerson;
        ModifiedById: number;
        Owshiddenversion: number;
        Version: string;
        Attachments: ISpDeferred;
        Path: string;
    }
    class SpItem implements ISpItem {
        __metadata: ISpMetadata;
        Title: string;
        ContentTypeID: string;
        Id: number;
        ContentType: string;
        Modified: any;
        Created: any;
        CreatedBy: ISpPerson;
        CreatedById: number;
        ModifiedBy: ISpPerson;
        ModifiedById: number;
        Owshiddenversion: number;
        Version: string;
        Attachments: ISpDeferred;
        Path: string;
        constructor();
    }
    interface ISpMultichoiceValue {
        __metadata: ISpMetadata;
        Value: any;
    }
    interface IPrincipalInfo {
        AccountName: string;
        UserInfoID: number;
        DisplayName: string;
        Email: string;
        Title: string;
        IsResolved: boolean;
        PrincipalType: string;
    }
}
declare module Shockout {
    /**
    * SharePoint 2013 API Data Types
    */
    interface ISpApiMetadata {
        id: string;
        uri: string;
        type: string;
    }
    interface ISpApiDeferred {
        uri: string;
    }
    interface ISpApiUserIdMetadata {
        type: string;
    }
    interface ISpApiUserId {
        __metadata: ISpApiUserIdMetadata;
        NameId: string;
        NameIdIssuer: string;
    }
    interface ISpApiPerson {
        __metadata: ISpApiMetadata;
        Groups: ISpApiDeferred;
        Id: number;
        IsHiddenInUI: boolean;
        LoginName: string;
        Title: string;
        PrincipalType: number;
        Email: string;
        IsSiteAdmin: boolean;
        UserId: ISpApiUserId;
    }
    interface ISpApiUserGroup {
        __metadata: ISpApiMetadata;
        Owner: ISpApiDeferred;
        Users: ISpApiDeferred;
        Id: number;
        IsHiddenInUI: boolean;
        LoginName: string;
        Title: string;
        PrincipalType: number;
        AllowMembersEditMembership: boolean;
        AllowRequestToJoinLeave: boolean;
        AutoAcceptRequestToJoinLeave: boolean;
        Description: string;
        OnlyAllowMembersViewMembership: boolean;
        OwnerTitle: string;
        RequestToJoinLeaveEmailSetting: string;
    }
}
declare module Shockout {
    class Templates {
        static attachmentsTemplate: string;
        static fileuploadTemplate: string;
        static actionTemplate: string;
        static getFileUploadTemplate(): string;
        static getFormAction(): HTMLDivElement;
        static getAttachmentsTemplate(fileuploaderId: string): HTMLElement;
    }
}
declare module Shockout {
    class Utils {
        /**
        * Ensure site url is or ends with '/'
        * @param url: string
        * @return string
        */
        static formatSubsiteUrl(url: any): string;
        /**
        * Convert a name to REST camel case format
        * @param str: string
        * @return string
        */
        static toCamelCase(str: string): string;
        /**
        * Parse a form ID from window.location.hash
        * @return number
        */
        static getIdFromHash(): number;
        /**
        * Set location.hash to form ID `#/id/<ID>`.
        * @return void
        */
        static setIdHash(id: number): void;
        /**
        * Escape column values
        * http://dracoblue.net/dev/encodedecode-special-xml-characters-in-javascript/155/
        */
        static escapeColumnValue(s: any): any;
        static getParent(o: any, num?: number): any;
        static getPrevKOComment(o: any): any;
        static getKoComments(parent: any): any[];
        static getKoContainerlessControls(parent: any): any[];
        static getEditableKoContainerlessControls(parent: any): any[];
        static getEditableKoControlNames(parent: any): any[];
        /**
        * Get the KO names of the edit input controls on a form.
        * @parem parent: HTMLElement
        * @return Array<string>
        */
        static getEditableKoNames(parent: any): any[];
        static getNextSibling(el: any): any;
        /**
        * Extract the Knockout observable name from a field with `data-bind` attribute.
        * If the KO name is `$data`, the method will recursively search for the closest parent element or comment with the `foreach:` binding.
        * @param control: HTMLElement
        * @return string
        */
        static observableNameFromControl(control: any, vm?: IViewModel): string;
        /**
        * Alias for observableNameFromControl()
        */
        static koNameFromControl: typeof Utils.observableNameFromControl;
        static parseJsonDate(d: any): Date;
        static parseIsoDate(d: any): Date;
        static isJsonDateTicks(val: any): boolean;
        static isIsoDateString(val: any): boolean;
        static getQueryParam(p: any): string;
        static clone(objectToBeCloned: any): any;
        static logError(msg: any, errorLogListName: string, siteUrl?: string, debug?: boolean): void;
        static updateKoField(el: HTMLElement, val: any): void;
        static validateSpPerson(person: string): boolean;
        static isTime(val: string): boolean;
        static isDate(val: string): boolean;
        static dateToLocaleString(d: Date): string;
        static toTimeLocaleObject(d: Date): any;
        static toTimeLocaleString(d: any): string;
        static toDateTimeLocaleString(d: any): string;
        /**
        * Parse dates in format: "MM/DD/YYYY", "MM-DD-YYYY", "YYYY-MM-DD", "/Date(1442769001000)/", or YYYY-MM-DDTHH:MM:SSZ
        * @param val: string
        * @return Date
        */
        static parseDate(val: any): Date;
        /**
        * Addapted from accounting.js library. http://josscrowcroft.github.com/accounting.js/
        * Format a number into currency
        *
        * Usage: accounting.formatMoney(number, symbol, precision, thousandsSep, decimalSep, format)
        * defaults: (0, "$", 2, ",", ".", "%s%v")
        *
        * Localise by overriding the symbol, precision, thousand / decimal separators and format
        * Second param can be an object matching `settings.currency` which is the easiest way.
        *
        * To do: tidy up the parameters
        */
        static formatMoney(value: any, symbol?: string, precision?: number): string;
        /**
        * Addapted from accounting.js library. http://josscrowcroft.github.com/accounting.js/
        * Takes a string/array of strings, removes all formatting/cruft and returns the raw float value
        * alias: accounting.`parse(string)`
        *
        * Decimal must be included in the regular expression to match floats (defaults to
        * accounting.settings.number.decimal), so if the number uses a non-standard decimal
        * separator, provide it as the second argument.
        *
        * Also matches bracketed negatives (eg. "$ (1.99)" => -1.99)
        *
        * Doesn't throw any errors (`NaN`s become 0) but this may change in future
        */
        static unformatNumber(value: any): number;
        /**
        * Addapted from accounting.js library. http://josscrowcroft.github.com/accounting.js/
        * Format a number, with comma-separated thousands and custom precision/decimal places
        *
        * Localise by overriding the precision and thousand / decimal separators
        * 2nd parameter `precision` can be an object matching `settings.number`
        */
        static formatNumber(value: any, precision?: number): string;
        /**
         * Tests whether supplied parameter is a string
         * from underscore.js
         */
        static isString(obj: any): boolean;
        /**
        * Addapted from accounting.js library.
        * Implementation of toFixed() that treats floats more like decimals
        *
        * Fixes binary rounding issues (eg. (0.615).toFixed(2) === "0.61") that present
        * problems for accounting- and finance-related software.
        */
        static toFixed(value: any, precision?: number): string;
        /**
        * Addapted from accounting.js library. http://josscrowcroft.github.com/accounting.js/
        * Check and normalise the value of precision (must be positive integer)
        */
        static checkPrecision(val: any): number;
        /**
        * Compares two arrays and returns array of unique matches.
        * @param array1: Array<any>
        * @param array2: Array<any>
        * @return boolean
        */
        static compareArrays(array1: Array<any>, array2: Array<any>): Array<any>;
        static trim(str: string): any;
        static formatPictureUrl(pictureUrl: any): string;
    }
}
declare module Shockout {
    /**
     * http://github.com/valums/file-uploader
     *
     * Multiple file upload component with progress-bar, drag-and-drop.
     * © 2010 Andrew Valums ( andrew(at)valums.com )
     *
     * Licensed under GNU GPL 2 or later, see license.txt.
     */
    var qq: any;
}
/**
* -----------------
* Shockout SP Form
* -----------------
* By John Bonfardeci <john.bonfardeci@gmail.com>
*
* GitHub: https://github.com/jbonfardeci/ShockoutForms
*
* A Replacement for InfoPath and XSLT Forms
* Leverage the power Knockout JS databinding with SharePoint services for modern and dynamic web form development.
*
* Minimum Usage:
* `var spForm = new Shockout.SPForm('My SharePoint List Name', 'my-form-ID', {});`
*
* Dependencies: jQuery 1.72+, jQuery UI<any>, KnockoutJS 3.2+
* Compatible with Bootstrap CSS - http://getbootstrap.com
*
*   Copyright (C) 2015  John T. Bonfardeci
*
*   This program is free software: you can redistribute it and/or modify
*   it under the terms of the GNU Affero General Public License as
*   published by the Free Software Foundation, either version 3 of the
*   License, or (at your option) any later version.
*
*   This program is distributed in the hope that it will be useful,
*   but WITHOUT ANY WARRANTY; without even the implied warranty of
*   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
*   GNU Affero General Public License for more details.
*
*   You should have received a copy of the GNU Affero General Public License
*   along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/
declare module Shockout {
    class SPForm {
        static DEBUG: boolean;
        formId: string;
        listName: string;
        listNameRest: string;
        static errorLogListName: string;
        static errorLogSiteUrl: string;
        static enableErrorLog: boolean;
        $createdInfo: any;
        $dialog: any;
        $form: any;
        $formAction: any;
        $formStatus: any;
        allowDelete: boolean;
        allowPrint: boolean;
        allowSave: boolean;
        allowedExtensions: Array<string>;
        asyncFns: Array<any>;
        attachmentMessage: string;
        confirmationUrl: string;
        debug: boolean;
        dialogOpts: any;
        editableFields: Array<string>;
        enableAttachments: boolean;
        enableErrorLog: boolean;
        errorLogListName: string;
        errorLogSiteUrl: string;
        fieldNames: Array<string>;
        fileHandlerUrl: string;
        fileUploaderSettings: IFileUploaderSettings;
        fileUploader: any;
        includeUserProfiles: boolean;
        includeWorkflowHistory: boolean;
        preRender: Function;
        postRender: Function;
        preSave: Function;
        requireAttachments: boolean;
        siteUrl: string;
        utils: Utils;
        viewModel: IViewModel;
        viewModelIsBound: boolean;
        workflowHistoryListName: string;
        /**
        * Get the current logged in user profile.
        * @return ICurrentUser
        */
        getCurrentUser(): ICurrentUser;
        private currentUser;
        /**
        * Get the default view for the list.
        * @return string
        */
        getDefaultViewUrl(): string;
        private defaultViewUrl;
        /**
        * Get the default mobile view for the list.
        * @return string
        */
        getDefailtMobileViewUrl(): string;
        private defailtMobileViewUrl;
        /**
        * Get a reference to the form element.
        * @return HTMLElement
        */
        getForm(): HTMLElement;
        private form;
        /**
        * Get the SP list item ID number.
        * @return number
        */
        getItemId(): number;
        private itemId;
        /**
        * Get the GUID of the SP list.
        * @return HTMLElement
        */
        getListId(): string;
        private listId;
        /**
        * Get a reference to the original SP list item.
        * @return ISpItem
        */
        getListItem(): ISpItem;
        private listItem;
        /**
        * Requires user to checkout the list item?
        * @return boolean
        */
        private requireCheckout;
        requiresCheckout(): boolean;
        /**
        * Get the SP site root URL
        * @return string
        */
        private rootUrl;
        getRootUrl(): string;
        /**
        * Get the `source` key's value from the querystring.
        * @return string
        */
        private sourceUrl;
        getSourceUrl(): string;
        /**
        * Get a reference to the form's Knockout view model.
        * @return string
        */
        getViewModel(): IViewModel;
        /**
        * Get the version number for this framework.
        * @return string
        */
        getVersion(): string;
        private version;
        queryStringId: string;
        isSp2013: Boolean;
        constructor(listName: string, formId: string, options: Object);
        /**
        * Execute the next asynchronous function from `asyncFns`.
        * @param success?: boolean = undefined
        * @param msg: string = undefined
        * @param args: any = undefined
        * @return void
        */
        nextAsync(success?: boolean, msg?: string, args?: any): void;
        /**
        * Get the current logged in user's profile.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        getCurrentUserAsync(self: SPForm, args?: any): void;
        /**
        * Get metadata about an SP list and the fields to build the Knockout model.
        * Needed to determine the list GUID, if attachments are allowed, and if checkout/in is required.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        getListAsync(self: SPForm, args?: any): void;
        /**
        * Initialize the form.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        initForm(self: SPForm, args?: any): void;
        /**
       * Get the SP list item data and build the Knockout view model.
       * @param self: SPForm
       * @param args?: any = undefined
       * @return void
       */
        getListItemAsync(self: SPForm, args?: any): void;
        /**
        * Get the SP user groups this user is a member of for removing/showing protected form sections.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        getUsersGroupsAsync(self: SPForm, args?: any): void;
        /**
        * Removes form sections the user doesn't have access to from the DOM.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        implementPermissions(self: SPForm, args?: any): void;
        /**
        * Get the workflow history for this form, if any.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        getHistoryAsync(self: SPForm, args?: any): void;
        /**
        * Bind the SP list item values to the view model.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        bindListItemValues(self?: SPForm): void;
        /**
        * Delete the list item.
        * @param model: IViewModel
        * @param callback?: Function = undefined
        * @return void
        */
        deleteListItem(model: IViewModel, callback?: Function, timeout?: number): void;
        /**
        * Save list item via SOAP services.
        * @param vm: IViewModel
        * @param isSubmit?: boolean = false
        * @param refresh?: boolean = false
        * @param customMsg?: string = undefined
        * @return void
        */
        saveListItem(vm: IViewModel, isSubmit?: boolean, refresh?: boolean, customMsg?: string): void;
        /**
        * Add a navigation menu to the form based on parent elements with class `nav-section`
        * @param salef: SPForm
        * @return void
        */
        finalize(self: SPForm): void;
        /**
        * Delete an attachment.
        */
        deleteAttachment(att: ISpAttachment, event: any): void;
        /**
        * Get the form's attachments
        * @param self: SFForm
        * @param callback: Function (optional)
        * @return void
        */
        getAttachments(self?: SPForm, callback?: Function): void;
        /**
        * Log to console in degug mode.
        * @param msg: string
        * @return void
        */
        log(msg: any): void;
        /**
        * Update the form status to display feedback to the user.
        * @param msg: string
        * @param success?: boolean = undefined
        * @return void
        */
        updateStatus(msg: string, success?: boolean): void;
        /**
        * Display a message to the user with jQuery UI Dialog.
        * @param msg: string
        * @param title?: string = undefined
        * @param timeout?: number = undefined
        * @return void
        */
        showDialog(msg: string, title?: string, timeout?: number): void;
        /**
        * Validate the View Model's required fields
        * @param model: IViewModel
        * @param showDialog?: boolean = false
        * @return bool
        */
        formIsValid(model: IViewModel, showDialog?: boolean): boolean;
        /**
        * Get a person by their ID from the User Information list.
        * @param id: number
        * @param callback: Function
        * @return void
        */
        getPersonById(id: number, koField: KnockoutObservable<string>): void;
        /**
        * Keeps track of field names to send back to the server for create and update operations.
        * Skips field names that:
        *   - have already been added to `ediableFields` array;
        *   - begin with an underscore '_' or dollar sign '$';
        *   - that don't exist in `fieldNames` array which includes both writable and read-only SP list field names;
        *
        * @param key: string
        * @return number: length of array or -1 if not added
        */
        pushEditableFieldName(key: string): number;
        /**
        * Log errors to designated SP list.
        * @param msg: string
        * @param self?: SPForm = undefined
        * @return void
        */
        logError(msg: string, e?: any, self?: SPForm): void;
        /**
        * Setup attachments modules.
        * @param self: SPForm = undefined
        * @return number
        */
        setupAttachments(self?: SPForm): number;
        /**
        * Setup Bootstrap validation for required fields.
        * @return number
        */
        setupBootstrapValidation(self?: SPForm): number;
        /**
        * Setup form navigation on sections with class '.nav-section'
        * @return number
        */
        setupNavigation(self?: SPForm): number;
        /**
        * Setup Datepicker fields.
        * @return number
        */
        setupDatePickers(self?: SPForm): number;
        setupHtmlFields(self?: SPForm): number;
        /**
        * Determine if the current user is a member of at least one of list of target SharePoint groups.
        * @param targetGroups: comma delimited string || Array<string>
        * @return boolean
        */
        currentUserIsMemberOfGroups(targetGroups: any): boolean;
    }
}
