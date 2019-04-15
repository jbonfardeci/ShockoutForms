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
* Compatible with Bootstrap 3.5.x CSS - http://getbootstrap.com
*    
* The MIT License (MIT) - https://tldrlegal.com/license/mit-license
* Copyright (c) 2015 John T. Bonfardeci
* 
* Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
* 
* The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
* 
* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
* FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
* LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
* DEALINGS IN THE SOFTWARE.
*   
*/

module Shockout {

    export class SPForm {

        public static DEBUG: boolean = false;

        public onCloseAction: Function;

        ///////////////////////////////////////////
        // Minimum Required Constructor Parameters
        ///////////////////////////////////////////
        // the ID of the form
        public formId: string;

        // The name of the SP List you're submitting a form to.
        public listName: string;
        public listNameRest: string;

        ////////////////////////////
        // Public Static Properties
        ////////////////////////////
        public static errorLogListName: string;
        public static errorLogSiteUrl: string;
        public static enableErrorLog: boolean;

        // In people search, select users from People.asmx (true) or from User Information List (false)
        public static searchPrincipals: boolean;

        /////////////////////////
        // Public jQuery Objects
        /////////////////////////
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
        public allowedExtensions: Array<string> = ['txt', 'rtf', 'zip', 'pdf', 'doc', 'docx', 'jpg', 'gif', 'png', 'ppt', 'tif', 'pptx', 'csv', 'pub', 'msg', 'xls', 'xlsx'];

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
        public errorLogSiteUrl: string = '/';

        public fieldNames: Array<string> = [];

        // Display the user profiles of the users that created and last modified a form. Includes photos. See `Shockout.Templates.getUserProfileTemplate()` in `templates.ts`.
        public includeUserProfiles: boolean = true;

        // Display logs from the workflow history list assigned to form workflows.
        public includeWorkflowHistory: boolean = true;

        public includeNavigationMenu: boolean = true;

        // Function to execute before using server results for people picker autocomplete
        public peopleFilter: Function;

        // Function to execute before rendering templates and before Knockout databinding. Good for inserting your own markup and logic.
        public preRender: Function;

        // Function to execute after a form has rendered all templates and after Knockout binding has taken place.
        public postRender: Function;

        // Function to execute before sacing/submitting a form. Good place to insert extra logic such as extra validation.
        public preSave: Function;

        // Set to true if at least one attachment is required for a form. Good requriring receipts to purchase requisitions and such. 
        public requireAttachments: boolean = false;
        
        // In people search, select users from People.asmx (true) or from User Information List (false)
        public searchPrincipals: boolean = true;

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
            isAdmin: false,
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
        public getDefailtMobileViewUrl(): string { return this.defaultMobileViewUrl; }
        private defaultMobileViewUrl: string;
        
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
        public setItemId(id: number): void {
            this.itemId = id;
        }
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
        public listItem: ISpItem = null;

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
        private version: string = '1.0.1';

        public queryStringId: string = 'formid';

        private cafe: ICafe;

        public listItemType: string;

        private listItemMetadata: ISpMetadata;
        public getListItemMetadata(): ISpMetadata{
            return this.listItemMetadata;
        }

        constructor(listName: string, formId: string, options: Object = undefined) {
            var self = this;
            var error;

            // sanity check
            if (!(this instanceof SPForm)) {
                error = 'You must declare an instance of this class with `new`.';
                alert(error);
                throw error;
            }

            // ensure we have the parameters we require
            if (!!!formId || !!!listName) {
                var errors: any = ['Missing required parameters:'];
                if (!!!this.formId) { errors.push(' `formId`') }
                if (!!!this.listName) { errors.push(' `listName`') }
                errors = errors.join('');
                alert(errors);
                throw errors;
            }

            // these are the only parameters required
            this.formId = formId; // string ID of the parent form - could be any element you choose.
            this.listName = listName; // the name of the SP List
            this.listNameRest = Utils.toCamelCase(listName);

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

            SPForm.DEBUG = this.debug;
            SPForm.searchPrincipals = this.searchPrincipals;

            // try to parse the form ID from the hash or querystring
            this.itemId = Utils.getIdFromHash();
            var idFromQs = Utils.getQueryParam(this.queryStringId) || Utils.getQueryParam('ID');

            if (!!!this.itemId && /\d/.test(idFromQs)) {
                // get the SP list item ID of the form in the querystring
                this.itemId = parseInt(idFromQs);
                Utils.setIdHash(this.itemId);
            }           

            // setup static error log list name and site uri
            SPForm.errorLogListName = this.errorLogListName;
            SPForm.errorLogSiteUrl = this.errorLogSiteUrl;
            SPForm.enableErrorLog = this.enableErrorLog;

            // initialize custom Knockout handlers
            KoHandlers.bindKoHandlers();

            // create instance of the Knockout View Model
            this.viewModel = new ViewModel(this);
            this.viewModel.showUserProfiles(this.includeUserProfiles);

            // create element for displaying form load status
            self.$formStatus = $('<div>', { 'class': 'form-status' }).appendTo(self.$form);

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
                        self.preRender(self, self.viewModel);
                    }
                    self.nextAsync(true);
                }  
                , self.getListAsync      
                , self.initForm 
                , function (self: SPForm, args: any = undefined) {
                    // Register Shockout's Knockout Components
                    KoComponents.registerKoComponents();
                    // apply Knockout bindings
                    ko.applyBindings(self.viewModel, self.form);
                    self.viewModelIsBound = true;
                    self.nextAsync(true);
                }            
                , self.getListItemAsync
                , self.getHistoryAsync
                , function (self: SPForm) {
                    if(self.listItem){
                        if(self.listItem.AttachmentFiles){
                            if(self.listItem.AttachmentFiles.results){
                                self.viewModel.attachments(self.listItem.AttachmentFiles.results);
                                self.viewModel.attachments.valueHasMutated();
                            }
                        }
                    }

                    if (self.postRender) {
                        self.postRender(self, self.viewModel);
                    }
                    self.nextAsync(true);
                }
                , self.implementPermissions
                , self.finalize
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
                this.updateStatus(msg, success, self);
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
        * Get the current logged in user's profile.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        getCurrentUserAsync(self: SPForm, args: any = undefined): void {

            self.updateStatus('Retrieving your account...', true, self);
            var success: string = 'Retrieved your account.';

            // If this is SP 2013+, it will return thre current user's account.
            SpApi15.getCurrentUser(/*callback:*/ function (user: ICurrentUser, error: number) {            
                self.currentUser = user;
                self.viewModel.currentUser(user);

                if (self.debug) {
                    console.info('This is the SP 2013 API.');
                    console.info('Current user is...');
                    console.info(self.viewModel.currentUser());
                }

                self.nextAsync(true, success);              
            }, /*expandGroups:*/ true, self.siteUrl);
        }

        /**
        * Get metadata about an SP list and the fields to build the Knockout model.
        * Needed to determine the list GUID, if attachments are allowed, and if checkout/in is required.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        getListAsync(self: SPForm, args: any = undefined): void {
            
            // Determine if the field is a `Choice` or `MultiChoice` field with choices.
            var rxIsChoice = /choice/i;
            var rxExcludeNames: RegExp = /\b(FolderChildCount|ItemChildCount|MetaInfo|ContentType|Edit|Type|LinkTitleNoMenu|LinkTitle|LinkTitle2|Version|Attachments)\b/;

            SpSoap.getList(self.siteUrl, self.listName, function (xmlDoc: XMLDocument, error: string) {
                if (!!error) {
                    var msg = 'Failed to retrieve list data. ' + error;
                    self.nextAsync(false, 'Failed to retrieve list data. ' + error);
                    self.logError(msg);
                    return;
                }

                setupList(xmlDoc);
            });

            function setupList(xmlDoc: XMLDocument): void {
                try {
                    var $list = $(xmlDoc).find('List').first();
                    var listId = $list.attr('ID');
                    self.listId = listId;

                    var requireCheckout = $list.attr('RequireCheckout');
                    self.requireCheckout = !!requireCheckout ? requireCheckout.toLowerCase() == 'true' : false;

                    var enableAttachments = $list.attr('EnableAttachments');
                    self.enableAttachments = !!enableAttachments ? enableAttachments.toLowerCase() == 'true' : false;

                    self.defaultViewUrl = $list.attr('DefaultViewUrl');
                    self.defaultMobileViewUrl = $list.attr('MobileDefaultViewUrl');

                    var rootFolder = $list.attr('RootFolder');
                    self.listItemType = Utils.tail(rootFolder.split('/')).toString();

                    $(xmlDoc).find('Field').filter(function (i: number, el: any) {
                        return !!($(el).attr('StaticName')) && $(el).attr('Hidden') != 'TRUE' && !rxExcludeNames.test($(el).attr('Name'));
                    }).each(setupKoVar);

                    // sort the field names alpha
                    self.fieldNames.sort();

                    if (self.debug) {
                        console.info(self.listName + ' list ID = ' + self.listId);
                        console.info('Field names are...');
                        console.info(self.fieldNames);
                    }

                    self.nextAsync(true, 'Initialized list settings.');
                }
                catch (e) {
                    if (self.debug) { throw e; }
                    var error = 'Failed to initialize list settings.';
                    self.logError(error + ' SPForm.getListAsync.setupList(): ', e);
                    self.nextAsync(false, error);
                }
            }

            function setupKoVar(i: number, el: any): void {

                if (!!!el) { return; }

                try {
                    var $el = $(el);
                    var displayName: string = $el.attr('DisplayName');
                    var spType: string = $el.attr('Type');
                    var spName: string = $el.attr('StaticName');
                    var spFormat: string = $el.attr('Format');
                    var spRequired: boolean = !!($el.attr('Required')) ? $el.attr('Required').toLowerCase() == 'true' : false;
                    var spReadOnly: boolean = !!($el.attr('ReadOnly')) ? $el.attr('ReadOnly').toLowerCase() == 'true' : false;
                    var spDesc: string = $el.attr('Description');
                    var vm: IViewModel = self.viewModel;

                    // Convert the Display Name to equal REST field name conventions.
                    // For example, convert 'Computer Name (if applicable)' to 'ComputerNameIfApplicable'.
                    var koName = Utils.toCamelCase(spName);

                    // stop and return if it's already a Knockout object
                    if (koName in self.viewModel) { return; }

                    self.fieldNames.push(koName);

                    var defaultValue: any;
                    // find the SP field's default value if exists
                    $el.find('> Default').each(function (j: number, def: any): void {
                        var val: any = $.trim($(def).text());

                        if (!!spType) {

                            switch (spType.toLowerCase()) {
                                case 'boolean':
                                    val = val == '0' ? false : true;
                                    break;
                                case 'number':
                                case 'currency':
                                    val = val - 0;
                                    break;
                                case 'datetime':
                                    if (val == '[today]') { val = new Date(); }
                                    break;
                                default:
                                    break;
                            }

                            // TODO: Parse simple SP formulas such as `[me]` etc.

                        }
                        defaultValue = val;
                    });

                    var koObj: any = !!spType && /^multi/i.test(spType)
                        ? ko.observableArray([])
                        : ko.observable(!!defaultValue ? defaultValue : spType == 'Boolean' ? false : null);
                    
                    // Add SP metadata to the KO object.
                    // e.g. <div data-bind="with: koObj._metadata">
                    koObj._metadata = {
                        'koName': koName,
                        'displayName': displayName || null,
                        'name': spName || null,
                        'format': spFormat || null,
                        'required': spRequired || false,
                        'readOnly': spReadOnly || false,
                        'description': spDesc || null,
                        'type': spType
                    };

                    // Also expose these SP metadata properties with an underscore prefix at the first level for convenience in special cases.
                    // e.g. <label data-bind="text: koObj._displayName"></label>
                    for (var p in koObj._metadata) {
                        koObj['_' + p] = koObj._metadata[p];
                    }

                    // Add choices defined in the SP list.
                    if (rxIsChoice.test(spType)) {
                        var isFillIn = $el.attr('FillInChoice');

                        koObj._isFillInChoice = !!isFillIn && isFillIn == 'True'; // allow fill-in choices
                        var choices = [];
                        var options = [];

                        $el.find('CHOICE').each(function (j: number, choice: any) {
                            var txt = $(choice).text();
                            choices.push({ 'value': $(choice).text(), 'selected': false }); // for backward compatibility
                            options.push(txt); // new preferred array to reference in KO foreach binding contexts
                        });

                        koObj._choices = choices;
                        koObj._options = options;
                        koObj._multiChoice = !!spType && spType == 'MultiChoice';

                        koObj._metadata.choices = choices;
                        koObj._metadata.options = options;
                        koObj._metadata.multichoice = koObj._multiChoice;
                    }

                    // Make it convenient to reference the parent KO object from within a KO binding context such as `with` or `foreach`.
                    koObj._metadata.$parent = koObj;
                    vm[koName] = koObj;

                    if (self.debug) {
                        console.info('Created KO object: ' + koName + ', type: ' + spType + ', default val: ' + defaultValue);
                    }
                }
                catch (e) {
                    self.logError('Failed to setup KO object at SPForm.getListAsync.setupKoVar(): ', e);
                    if (self.debug) { throw e; }
                }
            };
        }

        /**
        * Initialize the form.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        initForm(self: SPForm, args: any = undefined): void {            
            try {
                self.updateStatus("Initializing dynamic form features...", true, self);

                var vm: IViewModel = self.viewModel;
                var rx: RegExp = /submitted/i;

                // Register Shockout's Knockout Components
                //KoComponents.registerKoComponents();

                // Find out of this list allows saving before submitting and triggering workflow approval.
                // Must have a field with `submitted` in the name and it must be of type `Boolean`
                if (Utils.indexOf(self.fieldNames, 'IsSubmitted') > -1) {
                    self.allowSave = true;
                    ViewModel.isSubmittedKey = 'IsSubmitted';
                    if (self.debug) {
                        console.info('initFormAsync: IsSubmitted key: ' + ViewModel.isSubmittedKey);
                    }
                }

                // Append action buttons to form.
                self.viewModel.allowSave(self.allowSave);
                self.viewModel.allowPrint(self.allowPrint);
                self.viewModel.allowDelete(self.allowDelete);

                self.$formAction = $(Templates.getFormAction()).appendTo(self.$form);

                // Setup attachments modules.
                if (self.enableAttachments) {
                    self.setupAttachments(self);
                }

                // If error logging is enabled, ensure the list exists and has required columns. Disable if 404.
                if (self.enableErrorLog) {
                    // Send a test query
                    SpApi.getListItems(self.errorLogListName, function (data, error) {
                        if (!!error) {
                            self.enableErrorLog = SPForm.enableErrorLog = false;
                        }
                    }, self.errorLogSiteUrl, null, 'Title,Error', 'Modified', 1, false);
                }

                if (self.includeNavigationMenu) {
                    // add navigation section to top of form
                    var $navMenu: JQuery = self.$form.find(".so-nav-menu, [data-so-nav-menu], [so-nav-menu]");
                    if ($navMenu.length > 0) {
                        $navMenu.replaceWith(Templates.soNavMenuControl);
                    } else {
                        self.$form.prepend(Templates.soNavMenuControl);
                    }
                }

                // set the element to display created/modified by info
                self.$form.find(".created-info, [data-sp-created-info], [data-so-created-info], [sp-created-info], [so-created-info]").replaceWith(Templates.soCreatedModifiedInfoControl);

                // replace/append Workflow history section
                if (self.includeWorkflowHistory) {
                    var $wfControls: JQuery = self.$form.find(".workflow-history, [data-so-workflow-history], [so-workflow-history]");
                    if ($wfControls.length > 0) {
                        $wfControls.replaceWith(Templates.soWorkflowHistoryControl);
                    } else {
                        self.$form.append(Templates.soWorkflowHistoryControl);
                    }
                }
                            
                // Dynamically add/remove elements with attribute `data-new-only` from the DOM if not editing an existing form - a new form where `itemId == null || undefined`.
                self.$form.find('[data-new-only]')
                    .before('<!-- ko ifnot: !!$root.Id() -->')
                    .after('<!-- /ko -->');

                // Dynamically add/remove elements with attribute `data-edit-only` from the DOM if not a new form - an edit form where `itemId != null`.
                self.$form.find('[data-edit-only]')
                    .before('<!-- ko if: !!$root.Id() -->')
                    .after('<!-- /ko -->');

                // Dynamically add/remove elements if it's restricted to the author only for example, input elements for editing the form. 
                self.$form.find('[data-author-only]')
                    .before('<!-- ko if: !!$root.isAuthor() -->')
                    .after('<!-- /ko -->');

                // Dynamically add/remove elements if for non-authors only such as read-only elements for viewers of a form. 
                self.$form.find('[data-non-authors]')
                    .before('<!-- ko ifnot: !!$root.isAuthor() -->')
                    .after('<!-- /ko -->');          

                self.nextAsync(true, "Form initialized.");
                return;
            }
            catch (e) {
                if (self.debug) { throw e; }
                self.logError("Error in SPForm.initFormAsync(): ", e);
                self.nextAsync(false, "Failed to initialize form.");
                return;
            }
        }

        /**
       * Get the SP list item data and build the Knockout view model.
       * @param self: SPForm
       * @param args?: any = undefined
       * @return void
       */
        getListItemAsync(self: SPForm, args: any = undefined): void {
            if (!!!self.itemId) {
                self.nextAsync(true, "This is a New form.");
                return;
            }

            self.updateStatus("Retrieving form values...", true, self);

            SpApi15.GetListItem(self.listName, self.itemId, self.siteUrl, false, 'AttachmentFiles')
            .done((data: ISpItem, error: string = undefined) => {
                if (!!error) {
                    if (/not found/i.test(error + '')) {
                        self.showDialog("The form with ID " + self.itemId + " doesn't exist or it was deleted.");
                    }
                    self.nextAsync(false, error);
                    return;
                }
                self.listItem = data;
                self.listItemMetadata = data.__metadata;
                self.bindListItemValues(self);

            }).then(() => {
                if (self.includeWorkflowHistory) {
                    self.getHistoryAsync(self);
                }
                self.nextAsync(true, "Retrieved form data.");

            }).fail(() => {
                self.nextAsync(false, "Failed to retrieve form data.");
            });
            
        }

        /**
        * Removes form sections the user doesn't have access to from the DOM.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        implementPermissions(self: SPForm, args: any = undefined): void {
            try {
                self.updateStatus("Retrieving your permissions...", true, self);

                // Remove elements from DOM if current user doesn't belong to any of the SP user groups in an element's attribute `data-sp-groups`.
                self.$form.find("[data-sp-groups], [user-groups]").each(function(i: number, el: HTMLElement): void {

                    var groups: string = $(el).attr("data-sp-groups");
                    if (!!!groups) {
                        groups = $(el).attr("user-groups");
                    }

                    $(el).before('<!-- ko if: !!$root.isMember(' + groups + ') -->')
                        .after('<!-- /ko -->');
                });

                self.nextAsync(true, "Retrieved your permissions.");
            }
            catch (e) {
                if (self.debug) { throw e; }
                self.logError("Error in SPForm.implementPermissionsAsync() ", e);
                self.nextAsync(true, "Failed to retrieve your permissions.");
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

            self.updateStatus('Retrieving workflow history...', true, self);

            var filter: string = "ListID eq '" + self.listId + "' and PrimaryItemID eq " + self.itemId;
            var select: string = "Description,DateOccurred";
            var orderby: string = "DateOccurred";
            SpApi.getListItems(self.workflowHistoryListName, callback, self.siteUrl, filter, select, orderby, 25, false);

            function callback(items: Array<any>, error: string) {

                if (self.debug) {
                    console.info('Retrieved workflow history...');
                    console.info(items.length);
                }

                if (!!error || !!!items) {
                    var msg = 'The ' + self.workflowHistoryListName + ' list may be full at <a href="{url}">{url}</a>. Failed to retrieve workflow history in method, getHistoryAsync().'
                        .replace(/\{url\}/g, self.rootUrl + self.siteUrl + '/Lists/' + self.workflowHistoryListName.replace(/\s/g, '%20'));

                    self.logError(msg);
                    self.nextAsync(false, 'Failed to retrieve workflow history. ' + error);
                    return;
                }

                self.viewModel.historyItems([]);
                for (var i = 0; i < items.length; i++) {
                    self.viewModel.historyItems().push(new HistoryItem(items[i].Description, Utils.parseDate(items[i].DateOccurred)));
                }

                self.viewModel.historyItems.valueHasMutated();
                self.nextAsync(true, "Retrieved workflow history.");
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

                if(self.debug){
                    console.info('binding values from list item: ', item);
                }
                
                // Exclude these read-only metadata fields from the Knockout view model.
                var rxExclude: RegExp = /\b(__metadata|ContentTypeID|ContentType|Owshiddenversion|Version|Attachments|AttachmentFiles|Path)\b/;

                self.itemId = item.Id;
                vm.Id(item.Id);
                
                for (var key in self.viewModel) {

                    if (!item[key] || !vm[key]._type || rxExclude.test(key)){ 
                        continue; 
                    }

                    if(self.debug){
                        console.info('binding ko value: ', key, item[key]);
                    }

                    if ((item[key] != null && vm[key]._type == 'DateTime')) {
                        vm[key](Utils.parseDate(item[key]));
                    }
                    else if (vm[key]._type == 'MultiChoice') { //&& 'results' in item[key]                     
                        vm[key](item[key].results);
                    }
                    else {
                        vm[key](item[key] || null);
                    }
                }

                if(!!item.AuthorId || !!item.Author){
                    if(!isNaN(item.Author)){
                        item.AuthorId = item.Author;
                    }
                    Shockout.SpApi15.GetUserById(self.siteUrl, item.AuthorId).done((user: ISpPerson) => {
                        if(!!!user.Picture){
                            user.Picture = Shockout.SpApi15.GetGenericPersonPng();
                        }
                        item.CreatedBy = user;
                        vm.CreatedBy(item.CreatedBy);
                    });
                }

                if(!!item.EditorId || !!item.Editor){
                    if(!isNaN(item.Editor)){
                        item.EditorId = item.Editor;
                    }
                    Shockout.SpApi15.GetUserById(self.siteUrl, item.EditorId).done((user: ISpPerson) => {
                        if(!!!user.Picture){
                            user.Picture = Shockout.SpApi15.GetGenericPersonPng();
                        }
                        item.ModifiedBy = user;
                        vm.ModifiedBy(item.ModifiedBy);
                    });
                }

                vm.Created(Utils.parseDate(item.Created));
                vm.Modified(Utils.parseDate(item.Modified));
            }
            catch (e) {
                if (self.debug) { throw e; }
                self.logError('Failed to bind form values in SPForm.bindListItemValues(): ', e);               
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

            SpApi.deleteListItem(item, function (data: any, error: string) {
                if (!!error) {
                    if (callback) {
                        callback(data, error);
                    }
                    return;
                }

                self.showDialog("The form was deleted. You'll be redirected in " + timeout / 1000 + " seconds.");
                if (callback) {
                    callback(data);
                }
                setTimeout(function () {
                    window.location.replace(self.sourceUrl != null ? self.sourceUrl : self.rootUrl);
                }, timeout);

            });
        }

        /**
        * Save list item via SOAP services.
        * @param vm: IViewModel
        * @param isSubmit?: boolean = false
        * @param customMsg?: string = undefined
        * @return void
        */
        saveListItem(vm: IViewModel, isSubmit: boolean = false, customMsg: string = undefined, showDialog: boolean = true): JQueryPromise<any> {
            let d = $.Deferred();
            let self: SPForm = vm.parent;
            let isNew = !!(self.itemId == null)
                , timeout = 3000
                , saveMsg = customMsg || '<p>Your form has been saved.</p>'
                , fields: Array<Array<any>> = []
                , payload = {};

            try {

                // Build array of SP field names for the input fields remaning on the form.
                // These are the field names to be saved and current user is allowed to edit these.
                var editable = Utils.getEditableKoNames(self.$form);
                $(editable).each(function (i, key: any) {
                    self.pushEditableFieldName(key);
                });

                self.editableFields.sort();

                if (self.debug) {
                    console.info('Editable fields...');
                    console.info(self.editableFields);
                }

                //override form validation for clicking "Save" as opposed to "Submit" button
                isSubmit = typeof (isSubmit) == "undefined" ? true : isSubmit;

                //run presave action and stop if the presave action returns false
                if (self.preSave) {
                    var retVal = self.preSave(self, self.viewModel, isSubmit);
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
                    payload[ViewModel.isSubmittedKey] = isSubmit;
                }

                const $fields = $(self.editableFields);

                $fields.each(function (i: number, key: any): void {
                    if (!('_metadata' in vm[key])) { return; }

                    let val: any = vm[key]();
                    let keyName: string = vm[key]._name;
                    let spType = vm[key]._type || vm[key]._metadata.type;
                    spType = !!spType ? spType.toLowerCase() : null;

                    if(self.debug){
                        console.info('saving: ', keyName, spType);
                    }

                    if (typeof(val) == "undefined" || key == ViewModel.isSubmittedKey) { return; }

                    if (spType == 'datetime') {
                        const d: Date = Utils.parseDate(val);
                        if(!!d){
                            val = d;
                        }
                    }
                    else if (val != null && spType == 'note') {
                        // Clean html/text
                        val = $('<div>').html(val).html();
                    }
                    else if(spType == 'user'){
                        payload[keyName+'Id'] = !!val ? parseInt( val.split(';')[0] ) : null;
                        return;
                    }

                    payload[keyName] = val;
                });

                if(isNew){
                    SpApi15.AddListItem(self.siteUrl, self.listName, self.listItemType, payload).done((data: any) => {
                        saveListItemCallback(vm, self, data.Id).done((listItem) => {
                            d.resolve(listItem);
                        });
                    });
                }
                else{
                    SpApi15.UpdateListItem(self.siteUrl, self.listItem.__metadata, payload).then((data: any) => {
                        saveListItemCallback(vm, self, self.itemId).done((listItem) => {
                            d.resolve(listItem);
                        });
                    });
                }
            }
            catch (e) {
                if (self.debug) { throw e; }  
                self.logError('Error in SpForm.saveListItem(): ', e);
                d.reject(e);                            
            }

            return d.promise();

            function saveListItemCallback(vm: IViewModel, self: SPForm, itemId: number): JQueryPromise<any>{
                let d = $.Deferred();

                if(self.debug){
                    console.info('saveListItemCallback(): ');
                    console.info(self.itemId);
                }

                self.itemId = itemId;
                vm.Id(itemId);

                if (Utils.getIdFromHash() == null && self.itemId != null) {
                    Utils.setIdHash(self.itemId);
                }

                if (isSubmit) {//submitting form
                    if(showDialog){
                        self.showDialog('<p>Your form has been submitted. You will be redirected in ' + timeout / 1000 + ' seconds.</p>', 'Form Submission Successful');
                    }
                    
                    if (self.debug) {
                        console.warn('DEBUG MODE: Would normally redirect user to confirmation page: ' + self.confirmationUrl);
                        self._getListItem(self).done((listItem) => {
                            d.resolve(listItem);
                        });

                    } else {
                        d.resolve(self.listItem);
                        setTimeout(function () {
                            window.location.href = self.sourceUrl != null ? self.sourceUrl : self.confirmationUrl;
                        }, timeout);         
                    }

                    return d.promise();
                }

                //saving form
                if(showDialog){
                    self.showDialog(saveMsg, 'The form has been saved.', timeout);
                }

                self._getListItem(self).done((listItem) => {
                    d.resolve(listItem);
                });

                return d.promise();
            } // function saveListItemCallback
        }

        _getListItem(self): JQueryPromise<any> {
            self = self || this;
            let d = $.Deferred();

            SpApi15.GetListItem(self.listName, self.itemId, self.siteUrl, false, 'AttachmentFiles')
            .done((data: ISpItem, error: string = undefined) => {
                if (!!error) {
                    if (/not found/i.test(error + '')) {
                        self.showDialog("The form with ID " + self.itemId + " doesn't exist or it was deleted.");
                    }
                    self.nextAsync(false, error);
                    return;
                }
                self.listItem = data;
                self.listItemMetadata = data.__metadata;
                self.bindListItemValues(self);
            }).then(() => {
                if (self.includeWorkflowHistory) {
                    self.getHistoryAsync(self);
                }
                d.resolve(self.listItem);
            }).fail(() => {
                d.reject(false);
            });

            return d.promise();
        }

        /**
        * Add a navigation menu to the form based on parent elements with class `nav-section`
        * @param salef: SPForm
        * @return void
        */
        finalize(self: SPForm): void {

            try {
                
                // Setup form navigation on sections with class '.nav-section'
                self.setupNavigation(self);

                // Setup Datepickers.
                self.setupDatePickers(self);

                self.nextAsync(true, 'Finalized form controls.');
            }
            catch (e) {
                if (self.debug) { throw e; }
                self.logError('Error in SpForm.finalize(): ', e);
                self.nextAsync(false, 'Failed to finalize form controls.');
            }       
        }

        /**
        * Delete an attachment.
        */
        deleteAttachment(att: ISpAttachment, event: any): void {
            var self: SPForm = ViewModel.parent;
            var vm: IViewModel = self.viewModel;

            if(self.debug){
                console.info('deleting file: ', att);
            }

            if (!confirm('Are you sure you want to delete ' + att.FileName + '? This can\'t be undone.')) { return; }
            
            SpApi15.DeleteAttachment(self.siteUrl, att)
            .done(function (data: any, status: string, jqXhr: JQueryXHR) {
                if(self.debug){
                    console.info('deleted file: ', att);
                }
                const attachments: any = vm.attachments;
                attachments.remove(att);
                attachments.valueHasMutated();
            })
            .fail(function (jqXhr: JQueryXHR, status: string, error: string) {
                alert("Failed to delete attachment: " + status + ': ' + error);
            });
        }

        /**
        * Get the form's attachments
        * @param self: SFForm
        * @param callback: Function (optional)
        * @return void
        */
        getAttachments(self: SPForm = undefined): JQueryPromise<any> {
            self = self || this;
            let d = $.Deferred();
            let path = `${self.listItem.__metadata.uri}/AttachmentFiles`;

            SpApi15.Get(path, false).done((data: any, status: string, jqXhr: any) => {
                d.resolve(data.results);
            }).fail(() => {
                self.showDialog("Failed to retrieve attachments in SpForm.getAttachments()");
                d.resolve([]);
            });

            return d.promise();
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
        updateStatus(msg: string, success: boolean = true, spForm): void {
            var self: SPForm = spForm;

            self.$formStatus
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

                    var koName = Utils.observableNameFromControl(n, self.viewModel);

                    if (!!koName && model[koName]) {

                        var val = model[koName]();

                        if (val == null || $.trim(val + '').length == 0) {

                            // Try to get the field label text.
                            var labelTxt;
                            var $label = $("label[for='" + koName + "']");

                            if (!!$label) {
                                labelTxt = $label.html();
                            }

                            if (!!!labelTxt) {
                                labelTxt = $(n).closest('.form-group').find('label:first').html();
                            }

                            if (!!!labelTxt) {
                                labelTxt = model[koName]['_displayName'];
                            }

                            if (!!!labelTxt) {
                                $(n).parent().first().html();
                            }

                            if (Utils.indexOf(labels, labelTxt) < 0) {
                                labels.push(labelTxt);
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
                    self.showDialog('<p>The following fields are required or invalid:</p><div class="error">' + labels.join('<br/>') + '</div>');
                    return false;
                }
                return true;
            }
            catch (e) {
                if (self.debug) { throw e; }
                self.logError("Form validation error at SPForm.formIsValid(): ", e);
                return false;
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
            if (!!!id) {
                return;
            }

            SpApi.getPersonById(id, function (person: ISpPerson, error) {
                if (!!error) {
                    var msg = 'Error in SPForm.getPersonById: ' + error;
                    Utils.logError(msg, SPForm.errorLogListName);
                    if (self.debug) {
                        console.warn(msg);
                    }
                    return;
                }

                var name: string = person.Id + ';#' + person.Name.replace(/^i\:0\#\.w\|/, ''); //remove SP 2013 prefix
                koField(name);
                if (self.debug) {
                    console.warn('Retrieved person by ID... ' + name);
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
            if (!!!key || Utils.indexOf(this.editableFields, key) > -1 || key.match(/^(_|\$)/) != null || Utils.indexOf(this.fieldNames, key) < 0 || this.viewModel[key]._readOnly) { return -1; }
            return this.editableFields.push(key);
        }

        /**
        * Log errors to designated SP list.
        * @param msg: string
        * @param self?: SPForm = undefined
        * @return void
        */
        logError(msg: string, e: any = undefined, self: SPForm = undefined): void {
            self = self || this;
            var err: any = [msg];

            if (!!e) {
                err.push(e+'');
            }

            err = err.length > 0 ? err.join('; ') : err.join('');

            if (self.enableErrorLog) {
                Utils.logError(err, self.errorLogListName, self.errorLogSiteUrl, self.debug);
                self.showDialog('<p>An error has occurred and the web administrator has been notified.</p><pre>' + err + '</pre>');
            }
        }

        /** 
        * Setup attachments modules.
        * @param self: SPForm = undefined
        * @return number
        */
        setupAttachments(self: SPForm = undefined): number {
            self = self || this;
            var vm: IViewModel = self.viewModel;
            var count: number = 0;

            if (!self.enableAttachments) { return count; }

            try {
                // update deprecated attachments elements with new so-attachments KO component
                self.$form.find(".attachments, [data-sp-attachments]").each(function (i: number, att: HTMLElement) {
                    $(att).replaceWith('<so-attachments params="val: attachments"></so-attachments>');
                });

                if (self.debug) {
                    console.info('Attachments are enabled.');
                }
            }
            catch (e) {
                if (self.debug) { throw e; }
                self.logError('Error in SPForm.setupAttachments(): ', e);
            }

            return count;
        }

        /**
        * Setup form navigation on sections with class '.nav-section'
        * @return number
        */
        setupNavigation(self: SPForm = undefined): number {
            var self = self || this;
            var count: number = 0;
            if (!self.includeNavigationMenu) { return count; }

            try {
                // Set up a navigation menu at the top of the form if there are elements with the class `nav-section`.
                var $navSections = self.$form.find('.nav-section');

                if ($navSections.length == 0) {
                    return count;
                }

                // add navigation buttons
                self.$form.find(".nav-section:visible").each(function (i, el) {
                    var $el = $(el);
                    var $header = $el.find("> h4");
                    if ($header.length == 0) {
                        return;
                    }
                    var title = $header.html();
                    var anchorName = Utils.toCamelCase(title) + 'Nav';
                    $el.before('<div style="height:1px;" id="' + anchorName + '">&nbsp;</div>');
                    self.viewModel.navMenuItems().push({ 'title': title, 'anchorName': anchorName });
                    count++;
                });

                self.viewModel.navMenuItems.valueHasMutated();

                // add a back-to-top button
                self.$form.append('<a href="#TOP" class="back-to-top"><span class="glyphicon glyphicon-chevron-up"></span></a>');

                // add smooth scrolling to for anchors - animates page navigation
                $('body').delegate('a[href*=#]:not([href=#])', 'click', function () {
                    if (window.location.pathname.replace(/^\//, '') == this.pathname.replace(/^\//, '') && location.hostname == this.hostname) {
                        var target = $(this.hash);
                        target = target.length ? target : $('[name=' + this.hash.slice(1) + ']');
                        if (target.length) {
                            $('html,body').animate({
                                scrollTop: target.offset().top - 50
                            }, 1000);

                            return false;
                        }
                    }
                });              
            }
            catch (e) {
                if (self.debug) { throw e; }
                self.logError('Error in SpForm.setupNavigation(): ', e);
            }

            return count;
        }

        /**
         * Setup Datepicker fields.
         * @param {SPForm = undefined} self
         * @returns
         */
        setupDatePickers(self: SPForm = undefined): number {
            self = self || this;

            // Apply jQueryUI datepickers after all KO bindings have taken place to prevent error: 
            // `Uncaught Missing instance data for this datepicker`
            var $datepickers: Array<JQuery> = self.$form.find('input.datepicker').datepicker();

            if (self.debug) {
                console.info('Bound ' + $datepickers.length + ' jQueryUI datepickers.');
            }

            return $datepickers.length;

        }

        /**
        * Determine if the current user is a member of at least one of list of target SharePoint groups.
        * @param targetGroups: comma delimited string || Array<string>
        * @return boolean
        */
        currentUserIsMemberOfGroups(targetGroups: any): boolean {

            var groupNames: Array<any> = [];

            if (Utils.isString(targetGroups)) {
                groupNames = targetGroups.match(/\,/) != null ? targetGroups.split(',') : [targetGroups];
            }
            else if (targetGroups.constructor === Array) {
                groupNames = targetGroups;
            }
            else {
                return false;
            }

            // return true on first match for efficiency
            for (var i = 0; i < groupNames.length; i++) {
                var group: string = groupNames[i];
                group = group.match(/\;#/) != null ? group.split(';')[0] : group; //either id;#groupname or groupname
                group = Utils.trim(group);

                for (var j = 0; j < this.currentUser.groups.length; j++) {
                    var g: ISpGroup = this.currentUser.groups[j];
                    if (group == g.name || parseInt(group) == g.id) {                       
                        return true;
                    }
                }               
            }

            return false;
        }
    }

}

// Set global alias for Shockout only if it doesn't conflict with another object with the same name.
window['so'] = window['so'] || Shockout;