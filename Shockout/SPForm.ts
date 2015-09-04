/// <reference path="../typings/knockout.d.ts" />
/// <reference path="../typings/jquery.d.ts" />
/// <reference path="../typings/jquery.ui.datetimepicker.d.ts" />
/// <reference path="../typings/jqueryui.d.ts" />
'use strict';

/**
* -----------------
* Shockout SP Form
* -----------------
* By John Bonfardeci <john.bonfardeci@gmail.com>
*
* A Replacement for InfoPath and XSLT Forms
* Leverage the power Knockout JS databinding with SharePoint services for modern and dynamic web form development. 
*
* Minimum Usage: 
* `var spForm = new Shockout.SPForm('My SharePoint List Name', 'my-form-ID', {});`
* 
* Dependencies: jQuery 1.72+, jQuery UI<any>, KnockoutJS 3.2+
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
module Shockout {

    export class SPForm {

        ///////////////////////////
        // Minimum Required Fields
        ///////////////////////////
        // the ID of the form
        public formId: string;

        // The name of the SP List you're submitting a form to.
        public listName: string;

        /////////////////////
        // Static properties
        /////////////////////
        public static errorLogListName: string;

        //////////////////
        // jQuery objects
        //////////////////
        public $createdInfo;
        public $dialog;
        public $form;
        public $formAction;
        public $formStatus;

        //////////////////
        // Public Options
        //////////////////

        // Allow users to delete a form
        public allowDelete: boolean = false;

        // Allow users to print
        public allowPrint: boolean = true;

        // Enable users to save their form before submitting
        public allowSave: boolean = false;

        // Allowed extensions for file attachments
        public allowedExtensions: Array<string> = ['txt', 'rtf', 'zip', 'pdf', 'doc', 'docx', 'jpg', 'gif', 'png', 'ppt', 'tif', 'pptx', 'csv', 'pub', 'msg'];

        // Message to display if a file attachment is required - good for receipts attached to purchase requisitions and such
        public attachmentMessage: string = 'An attachment is required.';
        
        // Redeirect users after form submission to this page.
        public confirmationUrl: string = '/SitePages/Confirmation.aspx';

        // Run in debug mode with extra logging; disables error logging to SP list.
        public debug: boolean = false;
        
        // Override the SP List fields a user is allowed to submit. 
        public editableFields: Array<string> = [];

        // Enable users to attach files.
        public enableAttachments: boolean = false;

        // Enable error logging to SP List. Good if you want to track and debug errors that users may run into.
        public enableErrorLog: boolean = true;

        // The name of the SP List to log errors to
        public errorLogListName: string = 'Error Log';

        // Override the default fields selected by the SP List Item query.
        public fieldNames: Array<string>;

        // The relative URL of the Handler that attaches fiel uploads to list items.
        public fileHandlerUrl: string = '/_layouts/webster/SPFormFileHandler.ashx';

        // The setting sobject for File Uploader
        public fileUploaderSettings: IFileUploaderSettings;

        // The File Uploader object
        public fileUploader: any;
                
        // Display the user profiles of the users that created and last modified a form. Includes photos. See `Shockout.Templates.getUserProfileTemplate()` in `templates.ts`.
        public includeUserProfiles: boolean = true;

        // Display logs from the workflow history list assigned to form workflows.
        public includeWorkflowHistory: boolean = true;
                
        // Function to execute before rendering templates and before Knockout databinding. Good for inserting your own markup and logic.
        public preRender: Function;

        // Function to execute after a form has rendered all templates and after Knockout binding has taken place.
        public postRender: Function;

        // Function to execute before sacing/submitting a form. Good place to insert extra logic such as extra validation.
        public preSave: Function;

        // Set to true if at least one attachment is required for a form. Good requriring receipts to purchase requisitions and such. 
        public requireAttachments: boolean = false;
        
        // The relative URL of the SP subsite where the target SP list is located.
        public siteUrl: string = '';      

        // Utility methods for internal and external use.
        public utils: Utils = Utils;

        // The form's Knockout view model.
        public viewModel: IViewModel;
           
        // The SP list name of the workflow history list where form workflow entries are stored.
        // Displays workflow history to viewer so they know the status of their form. Depends on writing workflows with good logging.
        // Be careful,. Workflow History lsits can exceed the maximum amount of items regular users are allowed to view. Be sure to implement
        // a good Powershell script to clean up your workflow history lists with Task Scheduler on the server. Good luck doing that with Office 365! 
        public workflowHistoryListName: string = 'Workflow History';
       
        /////////////////////////////////////
        // Private Set Public Get Properties
        /////////////////////////////////////
        /**
        * Get the current logged in user profile.
        * @return ICurrentUser
        */
        public getCurrentUser(): ICurrentUser { return this.currentUser; }
        private currentUser: ICurrentUser;
      
        /**
        * Get the default view for the list.
        * @return string
        */
        public getDefaultViewUrl(): string { return this.defaultViewUrl; }
        private defaultViewUrl: string;
       
        /**
        * Get the default mobile view for the list.
        * @return string
        */
        public getDefailtMobileViewUrl(): string { return this.defailtMobileViewUrl; }
        private defailtMobileViewUrl: string;
        
        /**
        * Get a reference to the form element.
        * @return HTMLElement
        */
        public getForm(): HTMLElement { return this.form; }
        private form: HTMLElement;

        /**
        * Get the SP list item ID number.
        * @return number
        */
        public getItemId(): number { return this.itemId; }
        private itemId: number;

        /**
        * Get the GUID of the SP list.
        * @return HTMLElement
        */
        public getListId(): string { return this.listId; }
        private listId: string;

        /**
        * Get a reference to the original SP list item.
        * @return ISpItem
        */
        public getListItem(): ISpItem { return this.listItem; }
        private listItem: ISpItem;

        /**
        * Requires user to checkout the list item?
        * @return boolean
        */
        private requireCheckout: boolean = false;
        public requiresCheckout(): boolean { return this.requireCheckout; }

        /**
        * Get the SP site root URL
        * @return string
        */
        private rootUrl: string = window.location.protocol + '//' + window.location.hostname + (!!window.location.port ? ':' + window.location.port : '');
        public getRootUrl(): string { return this.rootUrl; }

        /**
        * Get the `source` key's value from the querystring.
        * @return string
        */
        private sourceUrl: string;
        public getSourceUrl(): string { return this.sourceUrl; }

        /**
        * Get a reference to the form's Knockout view model. 
        * @return string
        */
        public getViewModel(): IViewModel { return this.viewModel; }

        /**
        * Get the version number for this framework. 
        * @return string
        */
        public getVersion(): string { return this.version; }
        private version: string = '0.0.1';

        /////////////////////////////
        // Privte GET/SET Properties
        /////////////////////////////
        private asyncFns: Array<any>;
        private viewModelIsBound: boolean = false;
        
        constructor(listName: string, formId: string, options: Object) {
            var self = this;
            var error;

            if (!(this instanceof SPForm)) {
                error = 'You must declare an instance of this class with `new`.';
                alert(error);
                throw error;
                return;
            }

            if (!!!formId || !!!listName) {
                var errors: any = ['Missing required parameters:'];
                if (!!!this.formId) { errors.push(' `formId`') }
                if (!!!this.listName) { errors.push(' `listName`') }
                errors = errors.join('');
                alert(errors);
                throw errors;
                return;
            }

            this.formId = formId;
            this.listName = listName;

            if (!!Utils.getQueryParam("id")) {
                this.itemId = parseInt(Utils.getQueryParam("id"));
            }
            else if (!!Utils.getQueryParam("formid")) {
                this.itemId = parseInt(Utils.getQueryParam("formid"));
            }

            this.sourceUrl = Utils.getQueryParam("source"); //if accessing the form from a SP list, take user back to the list on close

            if (!!this.sourceUrl) {
                this.sourceUrl = decodeURIComponent(this.sourceUrl);
            }

            // override default instance variables with key-value pairs from options
            if (options && options.constructor === Object) {
                for (var p in options) {
                    this[p] = options[p];
                }
            }

            // get the form container element
            this.form = document.getElementById(this.formId);
            this.$form = $(this.form).addClass('sp-form');

            self.$formStatus = $('<div>', { 'class': 'form-status' }).appendTo(self.$form);

            self.$dialog = $('<div>', { 'id': 'formdialog' })
                .appendTo(self.$form)
                .dialog({
                    width: 400,
                    height: 250,
                    autoOpen: false,
                    show: {
                        effect: "blind",
                        duration: 1000
                    },
                    hide: {
                        effect: "explode",
                        duration: 1000
                    }
                });

            SPForm.errorLogListName = this.errorLogListName;

            this.viewModel = new ViewModel(this);

            // Cascading Asynchronous Function Execution (CAFE) Array
            // Don't change the order of these unless you know what you're doing.
            this.asyncFns = [
                function (self: SPForm) {
                    if (self.preRender) {
                        self.preRender(self);
                    }
                    self.nextAsync(true);
                    return;
                }
                , self.initFormAsync
                , self.getListAsync
                , self.getCurrentUserAsync
                , self.getUsersGroupsAsync
                , self.restrictSpGroupElementsAsync
                //, self.initFormAsync
                , self.getListItemAsync
                , self.getAttachmentsAsync
                , self.getHistoryAsync
                , function (self: SPForm) {
                    if (self.postRender) {
                        self.postRender(self);
                    }
                    self.nextAsync(true);
                    return;
                }
                , function (self: SPForm) { self.$formStatus.hide(); }
            ];

            //start CAFE
            this.nextAsync(true, 'Begin initialization...');         
        }

        /**
        * Execute the next asynchronous function from `asyncFns`.
        */
        nextAsync(success: boolean = undefined, msg: string = undefined, args: any = undefined): void {
            var self = this;
            success = success || true;

            if (msg) {
                this.updateStatus(msg, success);
            }

            if (!success) { return; }

            if (this.asyncFns.length == 0) {
                setTimeout(function () {
                    self.$formStatus.slideDown();
                }, 2000);
                return;
            }
            // execute the next function in the array
            this.asyncFns.shift()(self, args);
        }

        /**
        * Initialize the form.
        */
        initFormAsync(self: SPForm, args: any = undefined): void {            
            try {
                self.updateStatus("Initializing dynamic form features...");

                self.$createdInfo = self.$form.find(".created-info");

                // append action buttons
                self.$formAction = $(Templates.getFormAction(self.allowSave, self.allowDelete, self.allowPrint)).appendTo(self.$form);
                
                //append Created/Modified info to predefined section or append to form
                if (!!self.itemId) {
                    self.$createdInfo.html(Templates.getCreatedModifiedHtml());

                    //append Workflow history section
                    if (self.includeWorkflowHistory) {
                        self.$form.append(Templates.getHistoryTemplate());
                    }
                }

                if (self.editableFields.length == 0) {
                    //make array of SP field names and those that are editable from elements w/ data-bind attribute
                    self.$form.find('[data-bind]').each(function (i: number, e: HTMLElement) {
                        var key = Utils.observableNameFromControl(e);

                        //skip observable keys that have already been added or begins with an underscore '_' or dollar sign '$'
                        if (!!!key || self.editableFields.indexOf(key) > -1 || key.match(/^(_|\$)/) != null) { return; }

                        if (e.tagName == 'INPUT' || e.tagName == 'SELECT' || e.tagName == 'TEXTAREA' || $(e).attr('contenteditable') == 'true') {
                            self.editableFields.push(key);
                            self.viewModel[key] = ko.observable(null);
                        }
                    });

                    self.editableFields.sort();
                }

                self.fileUploaderSettings = {
                    element: null,
                    action: self.fileHandlerUrl,
                    debug: self.debug,
                    multiple: false,
                    maxConnections: 3,
                    allowedExtensions: self.allowedExtensions,
                    params: {
                        listId: self.listId,
                        itemId: self.itemId
                    },
                    onSubmit: function(id, fileName){},
                    onComplete: function(id, fileName, json){
                        if (self.itemId == null) {
                            self.viewModel['Id'](json.itemId);
                            self.itemId = json.itemId;
                            self.saveListItem(self.viewModel, false);
                        }
                        if (json.error == null && json.fileName != null) {
                            self.getAttachmentsAsync();
                        }
                    },
                    template: Templates.getFileUploadTemplate()
                }

                //setup attachments module
                self.$form.find(".attachments").each(function (i: number, att: HTMLElement) {
                    var id = 'fileuploader_' + i;
                    $(att).append(Templates.getAttachmentsTemplate(id));
                    self.fileUploaderSettings.element = document.getElementById(id);
                    self.fileUploader = new Shockout.qq.FileUploader(self.fileUploaderSettings);
                });

                // add control validation to Bootstrap form elements
                // http://getbootstrap.com/css/#forms-control-validation 
                self.$form.find('[required], .required').each(function (i: number, el: HTMLElement) {
                    var koName = Utils.observableNameFromControl(el);
                    var $parent = $(el).closest('.form-group')
                        .attr("data-bind", "css: { 'has-error': !!!" + koName + "(), 'has-success has-feedback': !!" + koName + "()}")
                        .append('<span class="glyphicon glyphicon-ok form-control-feedback" aria-hidden="true"></span>');
                });

                self.nextAsync(true, "Form initialized.");
                return;

            }
            catch (e) {
                if (self.debug) {
                    throw e;
                }
                self.logError("initFormAsync: " + e);
                self.nextAsync(false, "Failed to initialize form. " + e);
                return;
            }
        }

        /**
        * Get the current logged in user's profile.
        */
        getCurrentUserAsync(self: SPForm, args: any = undefined): void {
            try {
                var currentUser: ICurrentUser;
                var query = '<Query><Where><Eq><FieldRef Name="ID" /><Value Type="Counter"><UserID /></Value></Eq></Where></Query>';
                var viewFields = '<ViewFields><FieldRef Name="ID" /><FieldRef Name="Name" /><FieldRef Name="EMail" /><FieldRef Name="Department" /><FieldRef Name="JobTitle" /><FieldRef Name="UserName" /><FieldRef Name="Office" /></ViewFields>';

                self.getListItemsSoap('', 'User Information List', viewFields, query, function (xData, Sstatus) {
                    
                    var user: ICurrentUser = {
                        id: null,
                        title: null,
                        login: null,
                        email: null,
                        account: null,
                        jobtitle: null,
                        department: null,
                        groups: []
                    };

                    $(xData.responseXML).find('*').filter(function() {
                        return this.nodeName === 'z:row';
                    }).each(function(i: number, node: any) {
                        user.id = parseInt($(node).attr('ows_ID'));
                        user.title = $(node).attr('ows_Name');
                        user.login = $(node).attr('ows_UserName');
                        user.email = $(node).attr('ows_EMail');
                        user.jobtitle = $(node).attr('ows_JobTitle');
                        user.department = $(node).attr('ows_Department');
                        user.account = user.id + ';#' + user.login;
                    });

                    self.currentUser = user;
                    self.viewModel.currentUser(user);
                    self.nextAsync(true, 'Retrieved your account.');
                    return;
                });
            }
            catch (e) {
                if (self.debug) {
                    throw e;
                }
                self.logError('getCurrentUserAsync():' + e);
                self.nextAsync(false, 'Failed to retrieve your account.');
                return;
            }
        }

        /**
        * Get the SP user groups this user is a member of for removing/showing protected form sections.
        */
        getUsersGroupsAsync(self: SPForm, args: any = undefined): void {
            try {
                var msg = "Retrieved your groups.";

                if (self.$form.find("[user-groups]").length == 0) {
                    self.nextAsync(true, msg);
                    return;
                }

                var packet = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                    '<soap:Body>' +
                    '<GetGroupCollectionFromUser xmlns="http://schemas.microsoft.com/sharepoint/soap/directory/">' +
                    '<userLoginName>' + self.currentUser.login + '</userLoginName>' +
                    '</GetGroupCollectionFromUser>' +
                    '</soap:Body>' +
                    '</soap:Envelope>';

                var $jqXhr: JQueryXHR = $.ajax({
                    url: self.rootUrl + self.siteUrl + '/_vti_bin/usergroup.asmx',
                    type: 'POST',
                    dataType: 'xml',
                    data: packet,
                    contentType: 'text/xml; charset="utf-8"'
                });

                $jqXhr.done(function (doc, statusText, response) {
                    $(response.responseXML).find("Group").each(function (i: number, el: HTMLElement) {
                        self.currentUser.groups.push({
                            id: parseInt($(el).attr("ID")),
                            name: $(el).attr("Name")
                        });
                    });
                    self.nextAsync(true, "Retrieved your groups.");
                    return;
                });

                $jqXhr.fail(function (xData, status) {
                    var msg = "Failed to retrieve your groups: " + status;
                    self.logError(msg);
                    self.nextAsync(false, msg);
                    return;
                });

                self.updateStatus("Retrieving your groups...");
            }
            catch (e) {
                if (self.debug) {
                    throw e;
                }
                self.logError("getUsersGroupsAsync: " + e);
                self.nextAsync(false, "Failed to retrieve your groups.");
                return;
            }
        }

        /**
        * Removes form sections the user doesn't have access to from the DOM.
        */
        restrictSpGroupElementsAsync(self: SPForm, args: any = undefined): void {
            try {
                self.updateStatus("Retrieving your permissions...");

                self.$form.find("[user-groups]").each(function (i, el) {
                    var groups: string = $(el).attr("user-groups");
                    var groupNames: Array<string> = groups.match(/\,/) != null ? groups.split(',') : [groups];
                    var ct = 0;
                    $.each(groupNames, function (i: number, group: string) {
                        group = group.match(/\;#/) != null ? group.split(';')[0] : group; //either id;#groupname or groupname
                        group = $.trim(group);

                        $.each(self.currentUser.groups, function (j: number, g: any) {
                            if (group == g.name || parseInt(group) == g.id) { ct++; }
                        });
                    });

                    if (ct > 0) {
                        $(el).show();
                    }
                    else {
                        $(el).remove();
                    }
                });

                self.nextAsync(true, "Retrieved your permissions.");
                return;
            }
            catch (e) {
                if (self.debug) {
                    throw e;
                }
                self.logError("restrictSpGroupElementsAsync: " + e);
                self.nextAsync(true, "Failed to retrieve your permissions.");
                return;
            }
        }

        /**
        * Get the SP list item data and build the Knockout view model.
        */
        getListItemAsync(self: SPForm, args: any = undefined): void {
            try {
                self.updateStatus("Retrieving form values...");

                if (!!!self.itemId) {

                    // apply Knockout bindings if not already bound.
                    if (!self.viewModelIsBound) {
                        ko.applyBindings(self.viewModel, self.form);
                        self.viewModelIsBound = true;
                    }

                    self.nextAsync(true, "This is a New form.");
                    return;
                }

                var uri = self.rootUrl + self.siteUrl + '/_vti_bin/listdata.svc/' + self.listName.replace(/\s/g, '') + '(' + self.itemId + ')';
                // get the list item data
                self.getListItemsRest(uri, function bind(data: ISpWrapper<ISpItem>, status: string, jqXhr: any): void {
                    self.listItem = Utils.clone(data.d); //store copy of the original SharePoint list item
                    self.bindListItemValues(self);
                    self.nextAsync(true, "Retrieved form data.");
                    return;

                }, function fail(obj: any, status: string, jqXhr: any): void {
                    if (obj.status && obj.status == '404') {
                        var msg = obj.statusText + ". The form may have been deleted by another user."
                    }
                    else {
                        msg = status + ' ' + jqXhr;
                    }
                    self.showDialog(msg);
                    self.nextAsync(false, msg);
                    return;

                });
            }
            catch (e) {
                if (self.debug) {
                    throw e;
                }
                self.nextAsync(false, e);
                return;
            }
        }

        /**
        * Get the workflow history for this form, if any.
        */
        getHistoryAsync(self: SPForm, args: any = undefined): void {
            try {
                if (!!!self.itemId || !self.includeWorkflowHistory) {
                    self.nextAsync(true);
                    return;
                }
                var historyItems: Array<any> = [];
                var uri = self.rootUrl + self.siteUrl + "/_vti_bin/listdata.svc/" + self.workflowHistoryListName.replace(/\s/g, '') +
                    "?$filter=ListID eq '" + self.listId + "' and PrimaryItemID eq " + self.itemId + "&$select=Description,DateOccurred&$orderby=DateOccurred asc";

                self.getListItemsRest(uri, function (data: ISpCollectionWrapper<any>, status: string, jqXhr: any) {
                    $(data.d.results).each(function (i: number, item: any) {
                        historyItems.push(new HistoryItem(item.Description, Utils.parseJsonDate(item.DateOccurred)));
                    });
                    self.viewModel.history(historyItems);
                    self.nextAsync(true, "Retrieved workflow history.");
                    return;
                });
            }
            catch (ex) {
                var wfUrl = self.rootUrl + self.siteUrl + '/Lists/' + self.workflowHistoryListName.replace(/\s/g, '%20');
                self.logError('The Workflow History list may be full at <a href="{url}">{url}</a>. Failed to retrieve workflow history in method, getHistoryAsync(). Error: '
                    .replace(/\{url\}/g, wfUrl) + JSON.stringify(ex));
                self.nextAsync(true, 'Failed to retrieve workflow history.');
                return;
            }
        }

        bindObservable(key: string, val: any, self: SPForm = undefined): void {
            self = self || this;
            if (key in self.viewModel) {
                self.viewModel[key](val);
            }
            else {
                self.viewModel[key] = val != null && /Array/.test(val.constructor)
                    ? ko.observableArray(val)
                    : ko.observable(val);
            }
        }

        /**
        * Bind the SP list item values to the view model.

        - new form - bind input values
        - existing form - get list item values, create observables
        - on saving form - update observables
        */
        bindListItemValues(self: SPForm = undefined): void {
            self = self || this;
            try {
                var item: ISpItem = self.listItem;

                // Exclude these read-only metadata fields from the Knockout view model.
                var rxExclude: RegExp = /^(__metadata|ContentTypeID|ContentType|CreatedBy|ModifiedBy|Owshiddenversion|Version|Attachments|Path)/;
                var isObj: RegExp = /Object/;

                self.viewModel.Id(item.Id);
                self.viewModel.isAuthor(item.CreatedById == self.currentUser.id);

                for (var key in self.viewModel) {

                    console.log('getting: ' + key);

                    if (key in item && !rxExclude.test(key)) {
                        var val = null;

                        if (Utils.isJsonDate(item[key])) {
                            val = Utils.parseJsonDate(item[key]);
                        }
                        // Object types will have a corresponding key name plus the suffix `Value` or `Id` for lookups.
                        // For example: `SupervisorApproval` is an object container for `__deferred` that corresponds to `SupervisorApprovalValue` 
                        // which is an ID or string value.
                        else if (item[key] != null && isObj.test(item[key].constructor + '') && '__deferred' in item[key]) {
                            if (key + 'Value' in item) {
                                val = item[key + 'Value'];
                            }
                            else if (key + 'Id' in item) {
                                val = item[key + 'Id'];
                            }
                        }
                        else {
                            val = item[key];
                        }

                        if ('_choices' in self.viewModel[key]) {
                            self.viewModel[key](val || []);
                        } else {
                            self.viewModel[key](val || null);
                        }

                        if (self.debug) {
                            console.info('assigned value ' + val + ' to ' + key);
                        }
                    }
                }

                var $info = self.$createdInfo.find('.create-mod-info').empty();
                
                // get CreatedBy profile
                self.getListItemsRest(item.CreatedBy.__deferred.uri, function (data: ISpWrapper<ISpPerson>, status: string, jqXhr: any) {
                    var person: ISpPerson = data.d;
                    self.viewModel.CreatedBy(person);
                    self.viewModel.isAuthor(self.currentUser.id == person.Id);
                    self.viewModel.CreatedByName(person.Name);
                    self.viewModel.CreatedByEmail(person.WorkEMail);
                    if (self.includeUserProfiles) {
                        $info.prepend(Templates.getUserProfileTemplate(person, "Created By"));
                    }
                });

                // get ModifiedBy profile
                self.getListItemsRest(item.ModifiedBy.__deferred.uri, function (data: ISpWrapper<ISpPerson>, status: string, jqXhr: any) {
                    var person: ISpPerson = data.d;
                    self.viewModel.ModifiedBy(person);
                    self.viewModel.ModifiedByName(person.Name);
                    self.viewModel.ModifiedByEmail(person.WorkEMail);
                    if (self.includeUserProfiles) {
                        $info.append(Templates.getUserProfileTemplate(person, "Last Modified By"));
                    }
                });
            }
            catch (e) {
                if (self.debug) {
                    throw e;
                }
            }
        }

        /**
        * Delete the list item.
        */
        deleteListItem(model: IViewModel) {
            var self: SPForm = model.parent;
            var item: ISpItem = self.listItem;
            var timeout: number = 3000;

            $.ajax({
                url: item.__metadata.uri,
                type: 'POST',
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'X-Http-Method': 'DELETE',
                    'If-Match': item.__metadata.etag
                },
                success: function (data) {
                    self.showDialog("The form was deleted. You'll be redirected in " + timeout / 1000 + " seconds.");
                    setTimeout(function () {
                        window.location.replace(self.sourceUrl != null ? self.sourceUrl : self.rootUrl);
                    }, timeout);
                },
                error: function (data) {
                    throw data.responseJSON.error;
                }
            });
        }

        /**
        * Save the list item.
        */
        // http://blog.vgrem.com/2014/03/22/list-items-manipulation-via-rest-api-in-sharepoint-2010/
        saveListItem(model: IViewModel, isSubmit: boolean = true, refresh: boolean = true, customMsg: string = undefined): void {

            var self: SPForm = model.parent,
                isNew: boolean = !!!self.itemId,
                timeout: number = 3000,
                saveMsg: string = customMsg || "<p>Your form has been saved.</p>",
                postData = {},
                headers: any = { Accept: 'application/json;odata=verbose' },
                url: string,
                contentType: string = 'application/json';
            
            // run presave action and stop if the presave action returns false
            if (self.preSave) {
                var retVal = self.preSave(self);
                if (typeof (retVal) != 'undefined' && !!!retVal) {
                    return;
                }
            }
            
            // validate the form
            if (isSubmit && !self.formIsValid(model, true)) {
                return;
            }
            
            // prepare data to post
            $.each(self.editableFields, function (i: number, key: string): void {
                postData[key] = model[key]();
            });

            //Only update IsSubmitted if it's != true -- if it was already submitted.
            //Otherwise pressing Save would set it from true back to false - breaking any workflow logic in place!
            if (typeof model[ViewModel.isSubmittedKey] != "undefined" && (model[ViewModel.isSubmittedKey]() == null || model[ViewModel.isSubmittedKey]() == false)) {
                postData[ViewModel.isSubmittedKey] = isSubmit;
            }

            if (isNew) {
                url = self.rootUrl + self.siteUrl + '/_vti_bin/listdata.svc/' + self.listName.replace(/\s/g, '') + '(' + self.itemId + ')';
                //postData = JSON.stringify(postData);
                contentType += ';odata=verbose';
            } else {
                url = self.listItem.__metadata.uri;
                headers['X-HTTP-Method'] = 'MERGE';
                headers['If-Match'] = self.listItem.__metadata.etag;
                // https://msdn.microsoft.com/en-us/library/vstudio/bb383793(v=vs.100).aspx
                
                // JSON.stringify does same thing?
                //postData = window['Sys'].Serialization.JavaScriptSerializer.serialize(postData);
            }

            var $jqXhr: JQueryXHR = $.ajax({
                url: url,
                type: 'POST',
                processData: false,
                contentType: contentType,
                data: JSON.stringify(postData),
                headers: headers
            });

            $jqXhr.done(function (data: ISpWrapper<ISpItem>, status: string, jqXhr: any): void {
                self.listItem = Utils.clone(data.d);
                self.itemId = self.listItem.Id;

                if (isSubmit && !self.debug) {
                    //submitting form
                    self.showDialog("<p>Your form has been submitted. You will be redirected in " + timeout / 1000 + " seconds.</p>", "Form Submission Successful");

                    setTimeout(function () {
                        window.location.href = self.sourceUrl != null ? self.sourceUrl : self.confirmationUrl;
                    }, timeout);
                }
                else {
                    //saving form
                    if (isNew || refresh) {
                        saveMsg += "<p>This page will refresh in " + timeout / 1000 + " seconds.</p>";
                    }

                    self.showDialog(saveMsg, "The form has been saved.", timeout);

                    if (isNew || refresh) {
                        setTimeout(function () {
                            //append list item id to url
                            window.location.search = "?formid=" + self.itemId;
                        }, timeout);
                    }
                    else {
                        // update model values
                        self.bindListItemValues(self);
                        //give WF History list 5 seconds to update
                        setTimeout(function () { self.getHistoryAsync(self); }, 5000);
                    }
                }
            });

            $jqXhr.fail(function(obj: any, status: string, jqXhr: any): void {
                var msg = obj.statusText + '. An error occurred while saving the form.';
                self.showDialog(msg);
                self.logError(msg + ': ' + JSON.stringify(arguments));
            });
        }

        /**
        * Get attachments for this form. 
        */
        getAttachmentsAsync(self: SPForm = undefined, args: any = undefined): void {
            self = self || this;
            
            try {

                if (!!!self.listItem || !self.enableAttachments) {
                    self.nextAsync(true);
                    return;
                }

                var attachments: Array<Attachment> = [];
                self.getListItemsRest(self.listItem.Attachments.__deferred.uri, function (data: ISpCollectionWrapper<ISpAttachment>, status: string, jqXhr: any) {
                    $.each(data.d.results, function (i: number, att: ISpAttachment) {
                        attachments.push(new Attachment(att));
                    });
                    self.viewModel.attachments(attachments);
                    self.nextAsync(true, 'Retrieved attachments.');
                    return;
                });
            }
            catch (e) {
                if (self.debug) {
                    throw e;
                }
            }
        }

        /**
        * Delete an attachment.
        */
        deleteAttachment(att: Attachment): void {
            var self = this
                , model = self.viewModel;
            try {
                var packet = '<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                    '<soap:Body><DeleteAttachment xmlns="http://schemas.microsoft.com/sharepoint/soap/"><listName>' + self.listName + '</listName><listItemID>' + self.itemId + '</listItemID><url>' + att.href + '</url></DeleteAttachment></soap:Body></soap:Envelope>';

                var $jqXhr: JQueryXHR = $.ajax({
                    url: self.rootUrl + self.siteUrl + '/_vti_bin/lists.asmx',
                    type: 'POST',
                    dataType: 'xml',
                    data: packet,
                    contentType: "text/xml; charset='utf-8'",
                    headers: {
                        "SOAPAction": "http://schemas.microsoft.com/sharepoint/soap/DeleteAttachment",
                        "Content-Type": "text/xml; charset=utf-8"
                    }
                });

                $jqXhr.done(function (xData, status) {
                    var attachments: any = model.attachments;
                    attachments.remove(att);
                });

                $jqXhr.fail(function (xData, status) {
                    var msg = "Failed to delete attachment: " + status;
                    self.logError(msg);
                });
            }
            catch (e) {
                self.logError(e);
            }
        }

        /**
        * Get list items via SOAP.
        */
        getListItemsSoap(siteUrl: string, listName: string, viewFields: string, query: string, callback: JQueryPromiseCallback<any>, rowLimit: number = 25, viewName: string = '<ViewName/>', queryOptions: string = '<QueryOptions/>'): void {
            var self = this;
            try {           

                var packet = '<?xml version="1.0" encoding="utf-8"?>' +
                    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                    '<soap:Body>' +
                    '<GetListItems xmlns="http://schemas.microsoft.com/sharepoint/soap/">' +
                    '<listName>' + listName + '</listName>' +
                    //'<viewName>' + viewName + '</viewName>' +
                    '<query>' + query + '</query>' +
                    '<viewFields>' + viewFields + '</viewFields>' +
                    '<rowLimit>' + rowLimit + '</rowLimit>' +
                    //'<queryOptions>' + queryOptions + '</queryOptions>' +
                    '</GetListItems>' +
                    '</soap:Body>' +
                    '</soap:Envelope>';

                var $jqXhr: JQueryXHR = $.ajax({
                    url: siteUrl + '/_vti_bin/lists.asmx',
                    type: 'POST',
                    dataType: 'xml',
                    data: packet,
                    headers: {
                        "SOAPAction": "http://schemas.microsoft.com/sharepoint/soap/GetListItems",
                        "Content-Type": "text/xml; charset=utf-8"
                    }
                });

                $jqXhr.done(callback);

                $jqXhr.fail(function (xData, status) {
                    self.logError('<pre>' + xData + '</pre>');
                });
            }
            catch (e) {
                self.logError(e);
            }
        }

        /**
        * Get metadata about an SP list and the fields to build the Knockout model.
        * Needed to determine the list GUID, if attachments are allowed, and if checkout/in is required.
        */
        getListAsync(self: SPForm, args: any = undefined): void {
            var packet = '<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap:Body><GetList xmlns="http://schemas.microsoft.com/sharepoint/soap/"><listName>{0}</listName></GetList></soap:Body></soap:Envelope>';

            var $jqXhr = $.ajax({
                url: self.rootUrl + self.siteUrl + '/_vti_bin/lists.asmx',
                type: 'POST',
                cache: false,
                dataType: "xml",
                data: packet.replace('{0}', self.listName),
                headers: {
                    "SOAPAction": "http://schemas.microsoft.com/sharepoint/soap/GetList",
                    "Content-Type": "text/xml; charset=utf-8"
                }
            });

            $jqXhr.done(function (xmlDoc, status, jqXhr) {
                var $list = $(xmlDoc).find('List').first();
                self.listId = $list.attr('ID');
                self.requireCheckout = $list.attr('RequireCheckout').toLowerCase() == 'true';
                self.enableAttachments = $list.attr('EnableAttachments').toLowerCase() == 'true';
                self.defaultViewUrl = $list.attr('DefaultViewUrl');
                self.defailtMobileViewUrl = $list.attr('MobileDefaultViewUrl');

                // Find out if the list allows saving before submitting.
                // The field name should be named `IsSubmitted` or anything with the word `submitted` in the name
                var rxAllowsSave = /submitted/i;

                // Determine if the field is a `Choice` or `MultiChoice` field with choices.
                var rxIsChoice = /choice/i;

                // Build the Knockout view model
                $(xmlDoc).find('Field').filter(function (i: number, el: any) {
                    return !!!($(el).attr('Hidden')) && !!($(el).attr('DisplayName')); // exclude hidden fields

                }).each(function (i: number, el: any) {

                    var $el = $(el);
                    var displayName = $el.attr('DisplayName');

                    // convert Display Name to equal format REST returns field names with.
                    // For example, convert 'Computer Name (if applicable)' to 'ComputerNameIfApplicable'.
                    // The we can reference our fields choices with predictable variable names.
                    // So for a field named 'ComputerNameIfApplicable' will have a corresponding observable array names '_options_ComputerNameIfApplicable'.
                    var koName = displayName
                        .replace(/[^A-Za-z0-9\s]/g, '')
                        .replace(/\s[A-Za-z]/g, function (x) {
                            return x[1].toUpperCase();
                        });

                    //if (koName in self.viewModel) { return; }

                    if (rxAllowsSave.test(displayName) && $el.attr('Type') == 'Boolean') {
                        self.allowSave = true;
                        self.$formAction.find('.btn.save').show();
                        ViewModel.isSubmittedKey = koName;
                    }

                    // create the KO object based on the SP type.
                    self.viewModel[koName] = rxIsChoice.test($el.attr('Type')) ? ko.observableArray([]) : ko.observable(null);

                    // add metadata to the KO object
                    self.viewModel[koName]._koName = koName;
                    self.viewModel[koName]._displayName = displayName;
                    self.viewModel[koName]._name = $el.attr('Name');
                    self.viewModel[koName]._format = $el.attr('Format');
                    self.viewModel[koName]._required = $el.attr('Required') == 'True';
                    self.viewModel[koName]._readOnly = !!($el.attr('ReadOnly'));
                    self.viewModel[koName]._description = $el.attr('Description');

                    // Create and attach arrays for the choices in SP field's choice fields.
                    if (rxIsChoice.test($el.attr('Type'))) {
                        self.viewModel[koName]._isFillInChoice = $el.attr('FillInChoice') == 'True'; // allow fill-in choices
                        var choices = [];

                        $el.find('CHOICE').each(function (j: number, choice: any) {
                            choices.push({ 'value': $(choice).text(), 'selected': false });
                        });

                        self.viewModel[koName]._choices = choices;
                        self.viewModel[koName]._multiChoice = $el.attr('Type') == 'MultiChoice';
                    }

                    if (self.debug) {
                        console.info('Created KO object: ' + koName + (!!self.viewModel[koName]._choices ? ', numChoices: ' + self.viewModel[koName]._choices.length : '') );
                    }

                });

                // apply Knockout bindings if not already bound.
                if (!self.viewModelIsBound) {
                    ko.applyBindings(self.viewModel, self.form);
                    self.viewModelIsBound = true;
                }

                self.nextAsync(true);
                return;
            });

            $jqXhr.fail(function () {
                self.nextAsync(false, 'Failed to retrieve list data.');
                return;
            });
        }

        /**
        * Log to console in degug mode.
        */
        log(msg): void {
            if (this.debug) {
                console.log(msg);
            }
        }

        /**
        * Update the form status to display feedback to the user.
        */
        updateStatus(msg: string, success: boolean = undefined): void {
            success = success || true;
            this.$formStatus
                .html(msg)
                .css('color', (success ? "#ff0" : "$f00"))
                .slideUp();
        }

        /**
        * Display a message to the user with jQuery UI Dialog.
        */
        showDialog(msg: string, title: string = undefined, timeout: number = undefined): void {
            var self: SPForm = this;
            title = title || "Form Dialog";
            msg = (msg).toString().match(/<\w>\w*/) == null ? '<p>' + msg + '</p>' : msg; //wrap non-html in <p>
            self.$dialog.html(msg).dialog('open');
            if (timeout) {
                setTimeout(function () { self.$dialog.dialog.close(); }, timeout);
            }
        }

        /**
        * Get list items via REST.
        */
        getListItemsRest(uri: string, done: JQueryPromiseCallback<any>, fail: JQueryPromiseCallback<any> = undefined, always: JQueryPromiseCallback<any> = undefined): void {
            var self = this;

            var $jqXhr: JQueryXHR = $.ajax({
                url: uri,
                type: 'GET',
                cache: false,
                dataType: 'json',
                contentType: 'application/json; charset=utf-8',
                headers: {
                    'Accept': 'application/json'
                }
            });

            $jqXhr.done(done);

            var fail = fail || function (obj: any, status: string, jqXhr: any) {
                if (obj.status && obj.status == '404') {
                    var msg = obj.statusText + ". The data may have been deleted by another user."
                }
                else {
                    msg = status + ' ' + jqXhr;
                }
                self.showDialog(msg);
            }

            $jqXhr.fail(fail);

            if (always) {
                $jqXhr.always(always);
            }
        }

        /**
        * Validate the View Model's required fields
        * @returns: bool
        */
        formIsValid(model: IViewModel, showDialog: boolean = false): boolean {
            var self: SPForm = model.parent,
                labels: Array<string> = [],
                errorCount: number = 0,
                invalidCount: number = 0,
                invalidLabels: Array<string> = []
            ;

            try {

                self.$form.find('.required, [required]').each(function checkRequired(i: number, n: any): void {
                    var p = Utils.observableNameFromControl(n);
                    if (!!p && model[p]) {
                        var val = model[p]();
                        if (val == null || $.trim(val+'').length == 0) {
                            var label = $(n).parent().find('label:first').html();
                            if (!!!label) {
                                $(n).parent().first().html();
                            }
                            if (labels.indexOf(label) < 0) {
                                labels.push(label);
                                errorCount++;
                            }
                        }
                    }
                });

                //check for sp object data errors before saving
                self.$form.find(".invalid").each(function (i: number, el: HTMLElement): void {
                    var $parent = $(el).parent();
                    invalidLabels.push($(parent).first().html());
                    invalidCount++;
                });

                if (invalidCount > 0) {
                    labels.push('<p class="warning">There are validation errors with the following fields. Please correct before saving.</p><p style="color:#f00;">' + invalidLabels.join('<br />') + '</p>');
                }

                //if attachment(s) are required
                if (self.enableAttachments && self.requireAttachments && model.attachments().length == 0) {
                    errorCount++;
                    labels.push(self.attachmentMessage);
                }

                if (errorCount > 0) {
                    if (showDialog) {
                        self.showDialog('<p class="warning">The following are required:</p><p class="error"><strong>' + labels.join('<br/>') + '</strong></p>');
                    }
                    return false;
                }
                return true;
            }
            catch (e) {
                self.logError("Form validation error at formIsValid(): " + JSON.stringify(e));
                return false;
            }
        }

        checkInFile(pageUrl: string, checkinType: string, comment: string = '') {
            var action = 'http://schemas.microsoft.com/sharepoint/soap/CheckInFile';
            var params = [pageUrl, comment, checkinType];
            var packet = '<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap:Body><CheckInFile xmlns="http://schemas.microsoft.com/sharepoint/soap/"><pageUrl>{0}</pageUrl><comment>{1}</comment><CheckinType>{2}</CheckinType></CheckInFile></soap:Body></soap:Envelope>';

            return this.executeRequest(action, packet, params);
        }

        checkOutFile(pageUrl: string, checkoutToLocal: string, lastmodified: string) {
            var action = 'http://schemas.microsoft.com/sharepoint/soap/CheckOutFile';
            var params = [pageUrl, checkoutToLocal, lastmodified];
            var packet = '<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap:Body><CheckOutFile xmlns="http://schemas.microsoft.com/sharepoint/soap/"><pageUrl>{0}</pageUrl><checkoutToLocal>{1}</checkoutToLocal><lastmodified>{2}</lastmodified></CheckOutFile></soap:Body></soap:Envelope>';

            return this.executeRequest(action, packet, params);
        }

        executeRequest = function (action, packet, params, callback: Function = undefined, serviceUrl: string = this.rootUrl + this.siteUrl + '/_vti_bin/lists.asmx'): void {
            try {
                if (params != null) {
                    for (var i = 0; i < params.length; i++) {
                        packet = packet.replace('{' + i.toString() + '}', (params[i] == null ? '' : params[i]));
                    }
                }

                var $jqXhr: JQueryXHR = $.ajax({
                    url: serviceUrl,
                    cache: false,
                    type: 'POST',
                    data: packet,
                    headers: {
                        'Content-Type': 'text/xml; charset=utf-8',
                        'SOAPAction': action
                    }
                });

                if (callback) {
                    $jqXhr.done(<JQueryPromiseCallback<any>>callback);
                }

                $jqXhr.fail(function () { });
            }
            catch (e) {
                if (this.debug) {
                    throw e;
                }
            }
        }

        /**
        * Log errors to designated SP list.
        */
        logError(msg: string, self: SPForm = undefined): void {
            self = self || this;
            self.showDialog('<p>An error has occurred and the web administrator has been notified.</p><p>Error Details: <pre>' + msg + '</pre></p>');
            if (self.enableErrorLog) {
                Utils.logError(msg, self.errorLogListName, self.rootUrl, self.debug);
            }
        }
    }

}