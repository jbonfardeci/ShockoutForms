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
module Shockout {

    export class SPForm {

        ///////////////////////////////////////////
        // Minimum Required Constructor Parameters
        ///////////////////////////////////////////
        // the ID of the form
        public formId: string;

        // The name of the SP List you're submitting a form to.
        public listName: string;

        ////////////////////////////
        // Public Static Properties
        ////////////////////////////
        public static errorLogListName: string;

        /////////////////////////
        // Public jQuery Objects
        /////////////////////////
        public $createdInfo;
        public $dialog;
        public $form;
        public $formAction;
        public $formStatus;

        /////////////////////
        // Public Properties
        /////////////////////

        // Allow users to delete a form
        public allowDelete: boolean = false;

        // Allow users to print
        public allowPrint: boolean = true;

        // Enable users to save their form before submitting
        public allowSave: boolean = false;

        // Allowed extensions for file attachments
        public allowedExtensions: Array<string> = ['txt', 'rtf', 'zip', 'pdf', 'doc', 'docx', 'jpg', 'gif', 'png', 'ppt', 'tif', 'pptx', 'csv', 'pub', 'msg'];

        public asyncFns: Array<any>;

        // Message to display if a file attachment is required - good for receipts attached to purchase requisitions and such
        public attachmentMessage: string = 'An attachment is required.';
        
        // Redeirect users after form submission to this page.
        public confirmationUrl: string = '/SitePages/Confirmation.aspx';

        // Run in debug mode with extra logging; disables error logging to SP list.
        public debug: boolean = false;

        // jQuery UI dialog options
        public dialogOpts: any = {
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
        };
        
        // Override the SP List fields a user is allowed to submit. 
        public editableFields: Array<string> = [];

        // Enable users to attach files.
        public enableAttachments: boolean = false;

        // Enable error logging to SP List. Good if you want to track and debug errors that users may run into.
        public enableErrorLog: boolean = true;

        // The name of the SP List to log errors to
        public errorLogListName: string = 'Error Log';

        public fieldNames: Array<string> = [];

        // The relative URL of the Handler that attaches fiel uploads to list items.
        public fileHandlerUrl: string = '/_layouts/SPFormFileHandler.ashx';

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

        public viewModelIsBound: boolean = false;
           
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
        private currentUser: ICurrentUser = {
            id: null,
            title: null,
            login: null,
            email: null,
            account: null,
            jobtitle: null,
            department: null,
            groups: []
        };
      
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
        private itemId: number = null;

        /**
        * Get the GUID of the SP list.
        * @return HTMLElement
        */
        public getListId(): string { return this.listId; }
        private listId: string = null;

        /**
        * Get a reference to the original SP list item.
        * @return ISpItem
        */
        public getListItem(): ISpItem { return this.listItem; }
        private listItem: ISpItem = null;

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
        private sourceUrl: string = null;
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

        public queryStringId: string = 'formid';

        public sp2013: Boolean = false;

        constructor(listName: string, formId: string, options: Object) {
            var self = this;
            var error;

            // sanity check
            if (!(this instanceof SPForm)) {
                error = 'You must declare an instance of this class with `new`.';
                alert(error);
                throw error;
                return;
            }

            // ensure we have the parameters we require
            if (!!!formId || !!!listName) {
                var errors: any = ['Missing required parameters:'];
                if (!!!this.formId) { errors.push(' `formId`') }
                if (!!!this.listName) { errors.push(' `listName`') }
                errors = errors.join('');
                alert(errors);
                throw errors;
                return;
            }

            // these are the only parameters required
            this.formId = formId; // string ID of the parent form - could be any element you choose.
            this.listName = listName; // the name of the SP List

            // get the form container element
            this.form = <HTMLElement>(typeof formId == 'string' ? document.getElementById(formId) : formId);

            if (!!!this.form) {
                alert('An element with the ID "' + this.formId + '" was not found. Ensure the `formId` parameter in the constructor matches the ID attribute of the form element.');
                return;
            }

            this.$form = $(this.form).addClass('sp-form');

            // Prevent browsers from doing their own validation to allow users to press the `Save` button even when all required fields aren't filled in.
            // We're doing validation ourselves when users presses the `Submit` button.
            $('form').attr({'novalidate': 'novalidate'});

            //if accessing the form from a SP list, take user back to the list on close
            this.sourceUrl = Utils.getQueryParam('source'); 

            if (!!this.sourceUrl) {
                this.sourceUrl = decodeURIComponent(this.sourceUrl);
            }

            // override default instance variables with key-value pairs from options
            if (options && options.constructor === Object) {
                for (var p in options) {
                    this[p] = options[p];
                }
            }

            // try to parse the form ID from the hash or querystring
            this.itemId = Utils.getIdFromHash();
            var idFromQs = Utils.getQueryParam(this.queryStringId);

            if (!!!this.itemId && /\d/.test(idFromQs)) {
                // get the SP list item ID of the form in the querystring
                this.itemId = parseInt(idFromQs);
                Utils.setIdHash(this.itemId);
            }           

            // setup static error log list name
            SPForm.errorLogListName = this.errorLogListName;

            // initialize custom Knockout handlers
            KoHandlers.bindKoHandlers();

            // create instance of the Knockout View Model
            this.viewModel = new ViewModel(this);

            // create element for displaying form load status
            self.$formStatus = $('<div>', { 'class': 'form-status' }).appendTo(self.$form);

            // set the element to display created/modified by info
            self.$createdInfo = self.$form.find(".created-info");

            // create jQuery Dialog for displaying feedback to user
            self.$dialog = $('<div>', { 'id': 'formdialog' })
                .appendTo(self.$form)
                .dialog(self.dialogOpts);

            // Cascading Asynchronous Function Execution (CAFE) Array
            // Don't change the order of these unless you know what you're doing.
            this.asyncFns = [             
                self.getCurrentUserAsync
                , function (self: SPForm) {
                    if (self.preRender) {
                        self.preRender(self);
                    }
                    self.nextAsync(true);
                }  
                , self.getListAsync      
                , self.initFormAsync 
                , function (self: SPForm, args: any = undefined) {
                    // apply Knockout bindings
                    ko.applyBindings(self.viewModel, self.form);
                    self.viewModelIsBound = true;
                    self.nextAsync(true);
                }            
                , self.getListItemAsync
                , self.getUsersGroupsAsync
                , self.implementPermissionsAsync
                , self.getAttachmentsAsync
                , self.getHistoryAsync
                , function (self: SPForm) {
                    if (self.postRender) {
                        self.postRender(self);
                    }
                    self.nextAsync(true);
                }
            ];

            //start CAFE
            this.nextAsync(true, 'Begin initialization...');         
        }

        /**
        * Execute the next asynchronous function from `asyncFns`.
        * @param success?: boolean = undefined
        * @param msg: string = undefined
        * @param args: any = undefined
        * @return void
        */
        nextAsync(success: boolean = true, msg: string = undefined, args: any = undefined): void {
            var self = this;

            if (msg) {
                this.updateStatus(msg, success);
            }

            if (!success) { return; }

            if (this.asyncFns.length == 0) {
                setTimeout(function () {
                    self.$formStatus.hide();
                }, 2000);
                return;
            }
            // execute the next function in the array
            this.asyncFns.shift()(self, args);
        }

        /**
        * Initialize the form.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        initFormAsync(self: SPForm, args: any = undefined): void {            
            try {
                self.updateStatus("Initializing dynamic form features...");

                var vm: IViewModel = self.viewModel;
                var rx: RegExp = /submitted/i;

                // find out of this list allows saving before submitting and triggering workflow approval
                // must have a field with `submitted` in the name and it must be of type `Boolean`
                if (self.fieldNames.indexOf('IsSubmitted') > -1) {
                    self.allowSave = true;
                    ViewModel.isSubmittedKey = 'IsSubmitted';
                }

                // append action buttons
                self.$formAction = $(Templates.getFormAction(self.allowSave, self.allowDelete, self.allowPrint)).appendTo(self.$form);
                if (self.allowSave) {
                    self.$formAction.find('.btn.save').show();
                }

                if (self.enableAttachments) {
                    // set the absolute URI for the file handler 
                    self.fileHandlerUrl = self.rootUrl + self.siteUrl + self.fileHandlerUrl;

                    // file uploader default settings
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
                        onSubmit: function (id, fileName) { },
                        onComplete: function (id, fileName, json) {

                            if (self.debug) {
                                console.warn(json);
                            }

                            if (json.error != null && json.error != "") {
                                self.logError(json.error);
                                if (self.debug) {
                                    console.warn(json.error);
                                }
                                return;
                            }

                            if (self.itemId == null && json.itemId != null) {
                                self.itemId = json.itemId;
                                self.viewModel.Id(json.itemId);
                            }

                            // push a new SP attachment instance to the view model's `attachments` collection
                            self.viewModel.attachments().push(new SpAttachment(self.rootUrl, self.siteUrl, self.listName, self.itemId, fileName));
                            self.viewModel.attachments.valueHasMutated(); // tell KO the array has been updated
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
                }

                // set up HTML editors in the form
                self.$form.find(".rte, [data-bind*='spHtml']").each(function (i: number, el: HTMLElement) {
                    var $el = $(el);
                    var koName = Utils.observableNameFromControl(el);

                    var $rte = $('<div>', {
                        'data-bind': 'spHtmlEditor: ' + koName,
                        'class': 'form-control content-editable',
                        'contenteditable': 'true'
                    });

                    if (!!$el.attr('required') || !!$el.hasClass('required')) {
                        $rte.attr('required', 'required');
                        $rte.addClass('required');
                    }

                    $rte.insertBefore($el);
                    if (!self.debug) {
                        $el.hide();
                    }
                });

                //append Created/Modified info to predefined section or append to form
                if (!!self.itemId) {
                    self.$createdInfo.html(Templates.getCreatedModifiedHtml());

                    //append Workflow history section
                    if (self.includeWorkflowHistory) {
                        self.$form.append(Templates.getHistoryTemplate());
                    }
                }

                // remove elements with attribute `data-edit-only` from the DOM if not editing an existing form - a new form where `itemId == null || undefined`
                if (!!!self.itemId) {
                    self.$form.find('[data-edit-only]').each(function () {
                        $(this).remove();
                    });
                }

                // remove elements with attribute `data-new-only` from the DOM if not a new form - an edit form where `itemId != null`
                if (!!self.itemId) {
                    self.$form.find('[data-new-only]').each(function () {
                        $(this).remove();
                    });
                }
               
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
                    console.warn(e);
                }
                self.logError("initFormAsync: " + e);
                self.nextAsync(false, "Failed to initialize form. " + e);
                return;
            }
        }

        /**
        * Get the current logged in user's profile.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        getCurrentUserAsync(self: SPForm, args: any = undefined): void {
            try {
                var user = self.currentUser;
                var query = '<Query><Where><Eq><FieldRef Name="ID" /><Value Type="Counter"><UserID /></Value></Eq></Where></Query>';
                var viewFields = '<ViewFields><FieldRef Name="ID" /><FieldRef Name="Name" /><FieldRef Name="EMail" /><FieldRef Name="Department" /><FieldRef Name="JobTitle" /><FieldRef Name="UserName" /><FieldRef Name="Office" /></ViewFields>';

                self.getListItemsSoap('', 'User Information List', viewFields, query, function (xmlDoc: any, status: string, jqXhr: any) {
                    
                    /*
                    // Returns
                    <z:row xmlns:z="#RowsetSchema" 
                        ows_ID="1" ows_Name="<DOMAIN\login>" 
                        ows_EMail="<email>" 
                        ows_JobTitle="<job title>" 
                        ows_UserName="<username>" 
                        ows_Office="<office>" 
                        ows__ModerationStatus="0" ows__Level="1" 
                        ows_Title="<Fullname>" 
                        ows_UniqueId="1;#{2AFFA9A1-87D4-44A7-9D4F-618BCBD990D7}" 
                        ows_owshiddenversion="306" ows_FSObjType="1;#0"/>
                    */

                    $(xmlDoc).find('*').filter(function() {
                        return this.nodeName == 'z:row';
                    }).each(function(i: number, node: any) {
                        user.id = parseInt($(node).attr('ows_ID'));
                        user.title = $(node).attr('ows_Name').replace(/\\/, '\\');
                        user.login = $(node).attr('ows_UserName');
                        user.email = $(node).attr('ows_EMail');
                        user.jobtitle = $(node).attr('ows_JobTitle');
                        user.department = $(node).attr('ows_Department');
                        user.account = user.id + ';#' + user.title;
                    });

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
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        getUsersGroupsAsync(self: SPForm, args: any = undefined): void {
            try {
                var msg = "Retrieved your groups.";

                if (self.$form.find("[data-sp-groups], [user-groups]").length == 0) {
                    self.nextAsync(true, msg);
                    return;
                }

                var packet = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                    '<soap:Body>' +
                    '<GetGroupCollectionFromUser xmlns="http://schemas.microsoft.com/sharepoint/soap/directory/">' +
                    '<userLoginName>' + self.currentUser.title + '</userLoginName>' +
                    '</GetGroupCollectionFromUser>' +
                    '</soap:Body>' +
                    '</soap:Envelope>';

                var $jqXhr: JQueryXHR = $.ajax({
                    url: self.rootUrl + '/_vti_bin/usergroup.asmx',
                    type: 'POST',
                    dataType: 'xml',
                    data: packet,
                    contentType: 'text/xml; charset="utf-8"'
                });

                $jqXhr.done(function (xmlDoc, status, jqXhr) {

                    var $errorText = $(xmlDoc).find('errorstring');
                    // catch and handle returned error
                    if (!!$errorText && $errorText.text() != "") {
                        self.logError($errorText.text());
                        return;
                    }

                    $(xmlDoc).find("Group").each(function (i: number, el: HTMLElement) {
                        self.currentUser.groups.push({
                            id: parseInt($(el).attr("ID")),
                            name: $(el).attr("Name")
                        });
                    });
                    self.nextAsync(true, "Retrieved your groups.");
                    return;
                });

                $jqXhr.fail(function (xmlDoc, status) {
                    var msg = "Failed to retrieve your groups: " + status;

                    var $errorText = $(xmlDoc).find('errorstring');
                    // catch and handle returned error
                    if (!!$errorText && $errorText.text() != "") {
                        msg += $errorText.text();
                    }
                   
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
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        implementPermissionsAsync(self: SPForm, args: any = undefined): void {
            try {
                self.updateStatus("Retrieving your permissions...");

                // Remove elements from DOM if current user doesn't belong to any of the SP user groups in an element's attribute `data-sp-groups`.
                self.$form.find("[data-sp-groups], [user-groups]").each(function(i: number, el: HTMLElement): void {

                    // Provide backward compatibility.
                    // Attribute `user-groups` is deprecated and `data-sp-groups` is preferred for HTML5 "correctness."
                    var groups: string = $(el).attr("data-sp-groups");
                    if (!!!groups) {
                        groups = $(el).attr("user-groups");
                    }

                    var groupNames: Array<string> = groups.match(/\,/) != null ? groups.split(',') : [groups];
                    var ct = 0;
                    $.each(groupNames, function (j: number, group: string): void {
                        group = group.match(/\;#/) != null ? group.split(';')[0] : group; //either id;#groupname or groupname
                        group = $.trim(group);

                        $.each(self.currentUser.groups, function (k: number, g: any): void {
                            if (group == g.name || parseInt(group) == g.id) { ct++; }
                        });
                    });

                    if (ct == 0) {
                        $(el).remove();
                    }
                });

                // Remove element if it's restricted to the author only for example, input elements for editing the form. 
                if (!!self.listItem && self.currentUser.id != self.listItem.CreatedById) {
                    self.$form.find('[data-author-only]').each(function (i: number, el: HTMLElement): void {
                        $(this).remove();
                    });
                }
                
                // Remove element if for non-authors only such as read-only elements for viewers of a form. 
                if (!!self.listItem && self.currentUser.id == self.listItem.CreatedById) {
                    self.$form.find('[data-non-authors]').each(function (i: number, el: HTMLElement): void {
                        $(this).remove();
                    });
                }             

                // Build array of SP field names for the input fields remaning on the form.
                // These are the field names to be saved and current user is allows to edit these.
                var rxExcludeInputTypes: RegExp = /(button|submit|cancel|reset)/;
                var rxIncludeInputTags: RegExp = /(input|select|textarea)/i;
                self.$form.find('[data-bind]').each(function (i: number, e: HTMLElement) {                 
                    if (rxIncludeInputTags.test(e.tagName) && !rxExcludeInputTypes.test($(e).attr('type')) || $(e).attr('contenteditable') == 'true') {
                        var key = Utils.observableNameFromControl(e);
                        self.pushEditableFieldName(key);
                    }
                });

                self.editableFields.sort()

                self.nextAsync(true, "Retrieved your permissions.");
            }
            catch (e) {
                if (self.debug) {
                    throw e;
                }
                self.logError("restrictSpGroupElementsAsync: " + e);
                self.nextAsync(true, "Failed to retrieve your permissions.");
            }
        }

        /**
        * Get the SP list item data and build the Knockout view model.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        getListItemAsync(self: SPForm, args: any = undefined): void {
            try {
                self.updateStatus("Retrieving form values...");

                if (!!!self.itemId) {
                    self.nextAsync(true, "This is a New form.");
                    return;
                }

                var uri = self.rootUrl + self.siteUrl + '/_vti_bin/listdata.svc/' + self.listName.replace(/\s/g, '') + '(' + self.itemId + ')';
                // get the list item data
                self.getListItemsRest(uri, function bind(data: ISpWrapper<ISpItem>, status: string, jqXhr: any): void {
                    self.listItem = Utils.clone(data.d); //store copy of the original SharePoint list item
                    self.bindListItemValues(self);
                    self.nextAsync(true, "Retrieved form data.");

                }, function fail(obj: any, status: string, jqXhr: any): void {
                    var msg = null;
                    if (obj.status && obj.status == '404') {
                        msg = obj.statusText + ". A form with ID " + self.itemId + " doesn't exist or it may have been deleted by another user."
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
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        getHistoryAsync(self: SPForm, args: any = undefined): void {

            if (!!!self.itemId || !self.includeWorkflowHistory) {
                self.nextAsync(true);
                return;
            }

            try {
                var historyItems: Array<any> = [];
                var uri = self.rootUrl + self.siteUrl + "/_vti_bin/listdata.svc/" + self.workflowHistoryListName.replace(/\s/g, '') +
                    "?$filter=ListID eq '" + self.listId + "' and PrimaryItemID eq " + self.itemId + "&$select=Description,DateOccurred&$orderby=DateOccurred asc";

                self.getListItemsRest(uri, function (data: ISpCollectionWrapper<any>, status: string, jqXhr: any) {
                    $(data.d.results).each(function (i: number, item: any) {
                        historyItems.push(new HistoryItem(item.Description, Utils.parseJsonDate(item.DateOccurred)));
                    });
                    self.viewModel.history(historyItems);
                    self.nextAsync(true, "Retrieved workflow history.");
                });
            }
            catch (ex) {
                var wfUrl = self.rootUrl + self.siteUrl + '/Lists/' + self.workflowHistoryListName.replace(/\s/g, '%20');
                self.logError('The Workflow History list may be full at <a href="{url}">{url}</a>. Failed to retrieve workflow history in method, getHistoryAsync(). Error: '
                    .replace(/\{url\}/g, wfUrl) + JSON.stringify(ex));
                self.nextAsync(true, 'Failed to retrieve workflow history.');
            }
        }

        /**
        * Bind the SP list item values to the view model.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        bindListItemValues(self: SPForm = undefined): void {
            self = self || this;
            try {
                if (!!!self.itemId) { return; }

                var item: ISpItem = self.listItem;
                var vm: IViewModel = self.viewModel;
                
                // Exclude these read-only metadata fields from the Knockout view model.
                var rxExclude: RegExp = /^(__metadata|ContentTypeID|ContentType|CreatedBy|ModifiedBy|Owshiddenversion|Version|Attachments|Path)/;
                var isObj: RegExp = /Object/;

                self.itemId = item.Id;
                vm.Id(item.Id);
                
                for (var key in self.viewModel) {

                    if (!(key in item) || rxExclude.test(key) || vm[key]._type == 'MultiChoice' || vm[key]._type == 'User' || vm[key]._type == 'Choice')
                    { continue; }

                    if (item[key] != null && Utils.isJsonDate(item[key])) {
                        vm[key](Utils.parseJsonDate(item[key]));
                        continue;
                    }

                    var val = item[key];
                    vm[key](val || null);
                }

                var $info = self.$createdInfo.find('.create-mod-info').empty();
                
                // get CreatedBy profile
                self.getListItemsRest(item.CreatedBy.__deferred.uri, function (data: ISpWrapper<ISpPerson>, status: string, jqXhr: any) {
                    var person: ISpPerson = data.d;
                    vm.CreatedBy(person);
                    vm.CreatedByName(person.Name);
                    vm.CreatedByEmail(person.WorkEMail);
                    if (self.includeUserProfiles) {
                        $info.prepend(Templates.getUserProfileTemplate(person, "Created By"));
                    }
                    vm.isAuthor(self.currentUser.id == person.Id);
                });

                // get ModifiedBy profile
                self.getListItemsRest(item.ModifiedBy.__deferred.uri, function (data: ISpWrapper<ISpPerson>, status: string, jqXhr: any) {
                    var person: ISpPerson = data.d;
                    vm.ModifiedBy(person);
                    vm.ModifiedByName(person.Name);
                    vm.ModifiedByEmail(person.WorkEMail);
                    if (self.includeUserProfiles) {
                        $info.append(Templates.getUserProfileTemplate(person, "Last Modified By"));
                    }
                });

                // Object types `Choice` and `User` will have a corresponding key name plus the suffix `Value` or `Id` for lookups.
                // For example: `SupervisorApproval` is an object container for `__deferred` that corresponds to `SupervisorApprovalValue` which is an ID or string value.

                // query values for the `User` types
                $(self.fieldNames).filter(function (i: number, key: any): boolean {
                    if (!!!self.viewModel[key]) { return false; }
                    return self.viewModel[key]._type == 'User';
                }).each(function (i: number, key: any) {
                    if (!(key+'Id' in item)) { return; }
                    self.getPersonById(parseInt(item[key+'Id']), vm[key]);
                });

                // query values for `Choice` types
                $(self.fieldNames).filter(function (i: number, key: any): boolean {
                    if (!!!self.viewModel[key]) { return false; }
                    return self.viewModel[key]._type == 'Choice';
                }).each(function (i: number, key: any) {
                    if (!(key + 'Value' in item)) { return; }
                    vm[key](item[key+'Value']);
                });

                // query values for `MultiChoice` types
                $(self.fieldNames).filter(function (i: number, key: any): boolean {
                    if (!!!self.viewModel[key]) { return false; }
                    return self.viewModel[key]._type == 'MultiChoice';
                }).each(function (i: number, key: any) {
                    if (!('__deferred' in item[key])) { return; }

                    self.getListItemsRest(item[key].__deferred.uri, function (data: ISpCollectionWrapper<ISpMultichoiceValue>, status: string, jqXhr: any) {
                        var values: Array<any> = [];
                        $.each(data.d.results, function (i: number, choice: ISpMultichoiceValue) {
                            values.push(choice.Value);
                        });
                        vm[key](values);
                    });
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
        * @param model: IViewModel 
        * @param callback?: Function = undefined
        * @return void
        */
        deleteListItem(model: IViewModel, callback: Function = undefined, timeout: number = 3000): void {

            if (!confirm('Are you sure you want to delete this form?')) { return; }

            var self: SPForm = model.parent;
            var item: ISpItem = self.listItem;

            var $jqXhr: JQueryXHR = $.ajax({
                url: item.__metadata.uri,
                type: 'POST',
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'X-Http-Method': 'DELETE',
                    'If-Match': item.__metadata.etag
                }
            });

            $jqXhr.done(function (data: any, status: string, jqXhr: JQueryXHR): void {
                self.showDialog("The form was deleted. You'll be redirected in " + timeout / 1000 + " seconds.");
                if (callback) {
                    callback(data);
                }
                setTimeout(function () {
                    window.location.replace(self.sourceUrl != null ? self.sourceUrl : self.rootUrl);
                }, timeout);
            });

            $jqXhr.fail(function (data: any, status: string, jqXhr: JQueryXHR): void {
                throw data.responseJSON.error;
            });
        }

        /**
        * Save list item via SOAP services.
        * @param vm: IViewModel
        * @param isSubmit?: boolean = false
        * @param refresh?: boolean = false
        * @param customMsg?: string = undefined
        * @return void
        */
        saveListItem(vm: IViewModel, isSubmit: boolean = false, refresh: boolean = false, customMsg: string = undefined): void {
            var self: SPForm = vm.parent;
            var isNew = !!(self.itemId == null)
                , data = []
                , timeout = 3000
                , saveMsg = customMsg || '<p>Your form has been saved.</p>'
                , fields: Array<Array<any>> = []
                ;

            try {
                //override form validation for clicking "Save" as opposed to "Submit" button
                isSubmit = typeof (isSubmit) == "undefined" ? true : isSubmit;

                //run presave action and stop if the presave action returns false
                if (self.preSave) {
                    var retVal = self.preSave(self);
                    if (typeof (retVal) != 'undefined' && !!!retVal) {
                        return;
                    }
                }

                //validate the form
                if (isSubmit && !self.formIsValid(vm)) {
                    return;
                }

                //Only update IsSubmitted if it's != true -- if it was already submitted.
                //Otherwise pressing Save would set it from true back to false - breaking any workflow logic in place!
                var isSubmitted: KnockoutObservable<boolean> = vm[ViewModel.isSubmittedKey];
                if (typeof (isSubmitted) != "undefined" && (isSubmitted() == null || isSubmitted() == false)) {
                    fields.push([ViewModel.isSubmittedKey, (isSubmit ? 1 : 0)]);
                }

                // build the `fields` array 

                $(self.editableFields).each(function(i: number, key: any): void {
                    var val: any = vm[key]();

                    if (typeof (val) == "undefined" || key == ViewModel.isSubmittedKey) { return; }

                    if (val != null && val.constructor === Array) {
                        if (val.length > 0) {
                            val = ';#' + val.join(';#') + ';#';
                        }
                    }
                    else if (val != null && val.constructor == Date) {
                        val = new Date(val).toISOString();
                    }
                    else if (val != null && vm[key]._type == 'Note') {
                        val = '<![CDATA[' + $('<div>').html(val).html() + ']]>';
                    }

                    val = val == null ? '' : val;

                    fields.push([vm[key]._name, val]);
                });

                self.updateListItem(self.listName, fields, isNew, callback);
                 
            }
            catch (e) {
                self.logError(e);
                if (self.debug) { throw e; }                
            }

            function callback(xmlDoc: any, status: string, jqXhr: any): void {

                var itemId: number;

                if (self.debug) {
                    console.log('Callback from saveListItem()...');
                    console.log(status);
                    console.log(xmlDoc);
                }

                /*
                // Error response example
                <?xml version="1.0" encoding="utf-8"?>
                <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
                    <soap:Body>
                        <UpdateListItemsResponse xmlns="http://schemas.microsoft.com/sharepoint/soap/">
                            <UpdateListItemsResult>
                                <Results>
                                    <Result ID="1,New">
                                        <ErrorCode>0x80020005</ErrorCode>
                                        <ErrorText>The operation failed because an unexpected error occurred. (Result Code: 0x80020005)</ErrorText>
                                    </Result>
                                </Results>
                            </UpdateListItemsResult>
                        </UpdateListItemsResponse>
                    </soap:Body>
                </soap:Envelope>
                */

                var $errorText = $(xmlDoc).find('ErrorText');
                // catch and handle returned error
                if (!!$errorText && $errorText.text() != "") {
                    self.logError($errorText.text());
                    return;
                }

                $(xmlDoc).find('*').filter(function(): boolean {
                    return this.nodeName == 'z:row';
                }).each(function (i: number, el: any): void {
                    itemId = parseInt($(el).attr('ows_ID'));

                    if (self.itemId == null) {
                        self.itemId = itemId;
                    }

                    if (self.debug) {
                        console.warn('Item ID returned...');
                        console.warn(itemId);
                    }

                });             

                if (isSubmit && !self.debug) {//submitting form
                    self.showDialog('<p>Your form has been submitted. You will be redirected in ' + timeout / 1000 + ' seconds.</p>', 'Form Submission Successful');
                    setTimeout(function () {
                        window.location.href = self.sourceUrl != null ? self.sourceUrl : self.confirmationUrl;
                    }, timeout);
                }
                else {//saving form
                    self.showDialog(saveMsg, 'The form has been saved.', timeout);

                    // Append list item ID to querystring if this is a new form.
                    if (Utils.getIdFromHash() == null && self.itemId != null) {                     
                        setTimeout(function () {
                            //append list item id to hash
                            Utils.setIdHash(self.itemId);
                        }, 10);
                    }
                    else {
                        // refresh data from the server
                        self.getListItemAsync(self);
                        //give WF History list 5 seconds to update
                        setTimeout(function () { self.getHistoryAsync(self); }, 5000);
                    }
                }
            };     
        }

        /**
        * Save the list item with REST services.
        * UNUSED since saving list items in SP 2010 via REST is ridiculously difficult and chatty for saving multichoice fields.
        * Use saveListItem instead, especiallt for for updating a list item with multichoice fields.
        */
        // http://blog.vgrem.com/2014/03/22/list-items-manipulation-via-rest-api-in-sharepoint-2010/
        saveListItemREST(model: IViewModel, isSubmit: boolean = true, refresh: boolean = true, customMsg: string = undefined): void {

            var self: SPForm = model.parent,
                isNew: boolean = !!!self['itemId'],
                timeout: number = 3000,
                saveMsg: string = customMsg || '<p>Your form has been saved.</p>',
                postData = {},
                headers: any = { 'Accept': 'application/json;odata=verbose' },
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
        getAttachmentsAsync(self: SPForm, args: any = undefined): void {

            if (!!!self.listItem || !self.enableAttachments) {
                self.nextAsync(true);
                return;
            }

            self.getAttachments(self, function () {
                self.nextAsync(true);
                return;
            });
        }

        getAttachments(self: SPForm = undefined, callback: Function = undefined): void {
            self = self || this;

            if (!!!self.listItem || !self.enableAttachments) {
                if (callback) {
                    callback();
                }
                return;
            }

            try {
                var attachments: Array<ISpAttachment> = [];
                self.getListItemsRest(self.listItem.Attachments.__deferred.uri, function (data: ISpCollectionWrapper<ISpAttachment>, status: string, jqXhr: any): void {
                    $.each(data.d.results, function (i: number, att: ISpAttachment) {
                        attachments.push(att);
                    });
                    self.viewModel.attachments(attachments);
                    self.viewModel.attachments.valueHasMutated();
                    if (callback) {
                        callback(attachments);
                    }
                });
            }
            catch (e) {
                self.showDialog("Failed to retrieve attachments: " + JSON.stringify(e));
                if (self.debug) {
                    throw e;
                }
            }
        }

        /**
        * Delete an attachment.
        */
        deleteAttachment(att: ISpAttachment, event: any): void {

            if (!confirm('Are you sure you want to delete ' + att.Name + '? This can\'t be undone.')) { return; }

            try {
                var $jqXhr: JQueryXHR = $.ajax({
                    url: att.__metadata.uri,
                    type: 'POST',
                    dataType: 'json',
                    contentType: "application/json",
                    headers: {
                        'Accept': 'application/json;odata=verbose',
                        'X-HTTP-Method': 'DELETE'
                    }
                });

                $jqXhr.done(function (xData, status) {
                    var attachments: any = ViewModel.parent.viewModel.attachments;
                    attachments.remove(att);
                });

                $jqXhr.fail(function (xData, status) {
                    var msg = "Failed to delete attachment: " + status;
                });
            }
            catch (e) {
                throw e;
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
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        getListAsync(self: SPForm, args: any = undefined): void {
            var packet = '<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap:Body><GetList xmlns="http://schemas.microsoft.com/sharepoint/soap/"><listName>{0}</listName></GetList></soap:Body></soap:Envelope>';

            // Determine if the field is a `Choice` or `MultiChoice` field with choices.
            var rxIsChoice = /choice/i;
            var rxExcludeNames: RegExp = /^(FolderChildCount|ItemChildCount|MetaInfo|ContentType|Edit|Type|LinkTitleNoMenu|LinkTitle|LinkTitle2|Version|Attachments)/;

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

            $jqXhr.done(setupList);

            $jqXhr.fail(function () {
                self.nextAsync(false, 'Failed to retrieve list data.');
                return;
            });

            function setupList(xmlDoc: any, status: string, jqXhr: JQueryXHR): void {
                try {
                    var $list = $(xmlDoc).find('List').first();
                    var listId = $list.attr('ID');
                    self.listId = listId;
                    self.requireCheckout = $list.attr('RequireCheckout') == 'True';
                    self.enableAttachments = $list.attr('EnableAttachments') == 'True';
                    self.defaultViewUrl = $list.attr('DefaultViewUrl');
                    self.defailtMobileViewUrl = $list.attr('MobileDefaultViewUrl');

                    // Build the Knockout view model
                    $(xmlDoc).find('Field').filter(function (i: number, el: any) {
                        return !!!($(el).attr('Hidden')) && !!($(el).attr('DisplayName')) && !rxExcludeNames.test($(el).attr('Name')) && !rxExcludeNames.test($(el).attr('DisplayName'));
                    }).each(setupKoVar);

                    // sort the field names alpha
                    self.fieldNames.sort();

                    self.nextAsync(true);
                    return;
                }
                catch (e) {
                    self.nextAsync(false, 'Failed to initialize list settings.');
                    if (self.debug) {
                        console.warn(e);
                    }
                    return;
                }
            }

            function setupKoVar(i: number, el: any): void {

                if (!!!el) { return; }

                try {
                    var $el = $(el);
                    var displayName: string = $el.attr('DisplayName');
                    var spType: string = $el.attr('Type');
                    var spName: string = $el.attr('Name');
                    var spFormat: string = $el.attr('Format');
                    var spRequired: boolean = $el.attr('Required') == 'True';
                    var spReadOnly: boolean = !!($el.attr('ReadOnly')) && $el.attr('ReadOnly') == 'True';
                    var spDesc: string = $el.attr('Description');
                    var vm: IViewModel = self.viewModel;

                    // Convert the Display Name to equal REST field name conventions.
                    // For example, convert 'Computer Name (if applicable)' to 'ComputerNameIfApplicable'.
                    var koName = Utils.toCamelCase(displayName);

                    // stop and return if it's already a Knockout object
                    if (koName in self.viewModel) { return; }

                    self.fieldNames.push(koName);

                    var defaultValue: any;
                    // find the SP field's default value if exists
                    $el.find('> Default').each(function (j: number, def: any): void {
                        var val: any = $.trim($(def).text());
                        if (val == '[today]' && spType == 'DateTime') {
                            val = new Date();
                        }
                        else if (spType == 'Boolean') {
                            val = val == '0' ? false : true;
                        }
                        else if (spType == 'Number' || spType == 'Currency') {
                            val = val - 0;
                        }
                        defaultValue = val;
                    });

                    var koObj: any = !!spType && spType == 'MultiChoice' ? ko.observableArray([]) : ko.observable(!!defaultValue ? defaultValue : spType == 'Boolean' ? false : null);
                    
                    // add metadata to the KO object
                    koObj._metadata = {
                        koName: koName,
                        displayName: displayName,
                        name: spName,
                        format: spFormat,
                        required: spRequired,
                        readOnly: spReadOnly,
                        description: spDesc,
                        type: spType,
                    };

                    koObj._koName = koName;
                    koObj._displayName = displayName;
                    koObj._name = spName;
                    koObj._format = spFormat;
                    koObj._required = spRequired;
                    koObj._readOnly = spReadOnly;
                    koObj._description = spDesc;
                    koObj._type = spType;

                    if (rxIsChoice.test(spType)) {
                        var isFillIn = $el.attr('FillInChoice');

                        koObj._isFillInChoice = !!isFillIn && isFillIn == 'True'; // allow fill-in choices
                        var choices = [];

                        $el.find('CHOICE').each(function (j: number, choice: any) {
                            choices.push({ 'value': $(choice).text(), 'selected': false });
                        });

                        koObj._choices = choices;
                        koObj._multiChoice = !!spType && spType == 'MultiChoice';

                        koObj._metadata.choices = choices;
                        koObj._metadata.multichoice = koObj._multiChoice;
                    }

                    koObj._metadata.$parent = koObj;
                    vm[koName] = koObj;
                }
                catch (e) {
                    if (self.debug) {
                        console.warn(e);
                    }
                }
            };
        }

        /**
        * Log to console in degug mode.
        * @param msg: string
        * @return void
        */
        log(msg): void {
            if (this.debug) {
                console.log(msg);
            }
        }

        /**
        * Update the form status to display feedback to the user.
        * @param msg: string
        * @param success?: boolean = undefined
        * @return void
        */
        updateStatus(msg: string, success: boolean = true): void {
            var self: SPForm = this;

            this.$formStatus
                .html(msg)
                .css('color', (success ? "#ff0" : "$f00"))
                .show();

            setTimeout(function () { self.$formStatus.hide(); }, 2000);
        }

        /**
        * Display a message to the user with jQuery UI Dialog.
        * @param msg: string
        * @param title?: string = undefined
        * @param timeout?: number = undefined
        * @return void
        */
        showDialog(msg: string, title: string = undefined, timeout: number = undefined): void {
            var self: SPForm = this;
            title = title || "Form Dialog";
            msg = (msg).toString().match(/<\w>\w*/) == null ? '<p>' + msg + '</p>' : msg; //wrap non-html in <p>
            self.$dialog.html(msg).dialog('open');
            if (timeout) {
                setTimeout(function () { self.$dialog.dialog('close'); }, timeout);
            }
        }

        /**
        * Get list items via REST services.
        * @param uri: string
        * @param done: JQueryPromiseCallback<any>
        * @param fail?: JQueryPromiseCallback<any> = undefined
        * @param always?: JQueryPromiseCallback<any> = undefined
        * @return void 
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
        * @param model: IViewModel
        * @param showDialog?: boolean = false
        * @return bool
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

            return this.executeSoapRequest(action, packet, params);
        }

        checkOutFile(pageUrl: string, checkoutToLocal: string, lastmodified: string) {
            var action = 'http://schemas.microsoft.com/sharepoint/soap/CheckOutFile';
            var params = [pageUrl, checkoutToLocal, lastmodified];
            var packet = '<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap:Body><CheckOutFile xmlns="http://schemas.microsoft.com/sharepoint/soap/"><pageUrl>{0}</pageUrl><checkoutToLocal>{1}</checkoutToLocal><lastmodified>{2}</lastmodified></CheckOutFile></soap:Body></soap:Envelope>';

            return this.executeSoapRequest(action, packet, params);
        }

        /**
        * Update list item via SOAP services. 
        * @param listName: string
        * @param fields: Array<Array<any>>
        * @param isNew?: boolean = true
        * param callback?: Function = undefined
        * @param self: SPForm = undefined
        * @return void
        */
        updateListItem = function (listName: string, fields: Array<Array<any>>, isNew: boolean = true, callback: Function = undefined, self: SPForm = undefined): void {
            self = self || this;

            var action = 'http://schemas.microsoft.com/sharepoint/soap/UpdateListItems';
            var packet = '<?xml version="1.0" encoding="utf-8"?>' +
                '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                '<soap:Body>' +
                '<UpdateListItems xmlns="http://schemas.microsoft.com/sharepoint/soap/">' +
                '<listName>{0}</listName>' +
                '<updates>{1}</updates>' +
                '</UpdateListItems>' +
                '</soap:Body>' +
                '</soap:Envelope>';

            var command: string = isNew ? "New" : "Update";
            var params: Array<any> = [listName];
            var soapEnvelope: string = "<Batch OnError='Continue'><Method ID='1' Cmd='" + command + "'>";
            var itemArray: Array<Array<any>> = fields;

            for (var i = 0; i < fields.length; i++) {
                soapEnvelope += "<Field Name='" + fields[i][0] + "'>" + Utils.escapeColumnValue(fields[i][1]) + "</Field>";
            }

            if (command !== "New") {
                soapEnvelope += "<Field Name='ID'>" + self.itemId + "</Field>";
            }
            soapEnvelope += "</Method></Batch>";

            params.push(soapEnvelope);
            
            self.executeSoapRequest(action, packet, params, self, callback);
        }

        /**
        * Execute SOAP Request
        * @param action: string
        * @param packet: string
        * @param params: Array<any>
        * param self?: SPForm = undefined
        * @param callback?: Function = undefined
        * @return void
        */
        executeSoapRequest = function (action: string, packet: string, params: Array<any>, self: SPForm = undefined, callback: Function = undefined): void {
            self = self || this;
            try {
                var serviceUrl: string = self.rootUrl + self.siteUrl + '/_vti_bin/lists.asmx';

                if (params != null) {
                    for (var i = 0; i < params.length; i++) {
                        packet = packet.replace('{' + i + '}', (params[i] == null ? '' : params[i]));
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

                $jqXhr.fail(function (obj: any, status: string, jqXhr: any) {
                    var msg = 'executeSoapRequest() error. Status: ' + obj.statusText + ' ' + status + ' ' + JSON.stringify(jqXhr);
                    Utils.logError(msg, SPForm.errorLogListName);
                    console.warn(msg);
                });
            }
            catch (e) {
                if (self.debug) {
                    console.warn(e);
                }
            }
        }

        /**
        * Get a person by their ID from the User Information list.
        * @param id: number
        * @param callback: Function
        * @return void
        */
        getPersonById(id: number, koField: KnockoutObservable<string>): void {
            var self = this;
            if (!id || id.constructor != Number) {
                return;
            }
            var $jqXhr: JQueryXHR = $.ajax({
                url: this.rootUrl + "/_vti_bin/listdata.svc/UserInformationList(" + id + ")?$select=Id,Account",
                type: 'GET',
                cache: false,
                dataType: 'json',
                contentType: 'application/json; charset=utf-8',
                headers: {
                    'Accept': 'application/json'
                }
            });

            $jqXhr.done(function (data: ISpWrapper<ISpPerson>, status: string, jqXhr: any) {
                var d: ISpPerson = data.d;
                var name: string = d.Id + ';#' + d.Account.replace(/\\/, '\\');
                koField(name);
                if (self.debug) {
                    console.warn('Retrieved person by ID... ' + name);
                }
            });

            $jqXhr.fail(function (obj: any, status: string, jqXhr: any) {
                var msg = 'Get person by ID error. Status: ' + obj.statusText + ' ' + status + ' ' + JSON.stringify(jqXhr);
                Utils.logError(msg, SPForm.errorLogListName);
                if (self.debug) {
                    console.warn(msg);
                }
            });
        }

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
        pushEditableFieldName(key: string): number {
            if (!!!key || this.editableFields.indexOf(key) > -1 || key.match(/^(_|\$)/) != null || this.fieldNames.indexOf(key) < 0 || this.viewModel[key]._readOnly) { return -1; }
            return this.editableFields.push(key);
        }

        /**
        * Log errors to designated SP list.
        * @param msg: string
        * @param self?: SPForm = undefined
        * @return void
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