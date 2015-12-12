/// <reference path="../_references.ts" />
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
var Shockout;
(function (Shockout) {
    var SPForm = (function () {
        function SPForm(listName, formId, options) {
            /////////////////////
            // Public Properties
            /////////////////////
            // Allow users to delete a form
            this.allowDelete = false;
            // Allow users to print
            this.allowPrint = true;
            // Enable users to save their form before submitting
            this.allowSave = false;
            // Allowed extensions for file attachments
            this.allowedExtensions = ['txt', 'rtf', 'zip', 'pdf', 'doc', 'docx', 'jpg', 'gif', 'png', 'ppt', 'tif', 'pptx', 'csv', 'pub', 'msg'];
            // Message to display if a file attachment is required - good for receipts attached to purchase requisitions and such
            this.attachmentMessage = 'An attachment is required.';
            // Redeirect users after form submission to this page.
            this.confirmationUrl = '/SitePages/Confirmation.aspx';
            // Run in debug mode with extra logging; disables error logging to SP list.
            this.debug = false;
            // jQuery UI dialog options
            this.dialogOpts = {
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
            this.editableFields = [];
            // Enable users to attach files.
            this.enableAttachments = false;
            // Enable error logging to SP List. Good if you want to track and debug errors that users may run into.
            this.enableErrorLog = true;
            // The name of the SP List to log errors to
            this.errorLogListName = 'Error Log';
            this.errorLogSiteUrl = '/';
            this.fieldNames = [];
            // The relative URL of the Handler that attaches fiel uploads to list items.
            this.fileHandlerUrl = '/_layouts/SPFormFileHandler.ashx';
            // Display the user profiles of the users that created and last modified a form. Includes photos. See `Shockout.Templates.getUserProfileTemplate()` in `templates.ts`.
            this.includeUserProfiles = true;
            // Display logs from the workflow history list assigned to form workflows.
            this.includeWorkflowHistory = true;
            // Set to true if at least one attachment is required for a form. Good requriring receipts to purchase requisitions and such. 
            this.requireAttachments = false;
            // The relative URL of the SP subsite where the target SP list is located.
            this.siteUrl = '';
            // Utility methods for internal and external use.
            this.utils = Shockout.Utils;
            this.viewModelIsBound = false;
            // The SP list name of the workflow history list where form workflow entries are stored.
            // Displays workflow history to viewer so they know the status of their form. Depends on writing workflows with good logging.
            // Be careful,. Workflow History lsits can exceed the maximum amount of items regular users are allowed to view. Be sure to implement
            // a good Powershell script to clean up your workflow history lists with Task Scheduler on the server. Good luck doing that with Office 365! 
            this.workflowHistoryListName = 'Workflow History';
            this.currentUser = {
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
            this.itemId = null;
            this.listId = null;
            this.listItem = null;
            /**
            * Requires user to checkout the list item?
            * @return boolean
            */
            this.requireCheckout = false;
            /**
            * Get the SP site root URL
            * @return string
            */
            this.rootUrl = window.location.protocol + '//' + window.location.hostname + (!!window.location.port ? ':' + window.location.port : '');
            /**
            * Get the `source` key's value from the querystring.
            * @return string
            */
            this.sourceUrl = null;
            this.version = '1.0.1';
            this.queryStringId = 'formid';
            this.isSp2013 = false;
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
                var errors = ['Missing required parameters:'];
                if (!!!this.formId) {
                    errors.push(' `formId`');
                }
                if (!!!this.listName) {
                    errors.push(' `listName`');
                }
                errors = errors.join('');
                alert(errors);
                throw errors;
                return;
            }
            // these are the only parameters required
            this.formId = formId; // string ID of the parent form - could be any element you choose.
            this.listName = listName; // the name of the SP List
            this.listNameRest = Shockout.Utils.toCamelCase(listName);
            // get the form container element
            this.form = (typeof formId == 'string' ? document.getElementById(formId) : formId);
            if (!!!this.form) {
                alert('An element with the ID "' + this.formId + '" was not found. Ensure the `formId` parameter in the constructor matches the ID attribute of the form element.');
                return;
            }
            this.$form = $(this.form).addClass('sp-form');
            // Prevent browsers from doing their own validation to allow users to press the `Save` button even when all required fields aren't filled in.
            // We're doing validation ourselves when users presses the `Submit` button.
            $('form').attr({ 'novalidate': 'novalidate' });
            //if accessing the form from a SP list, take user back to the list on close
            this.sourceUrl = Shockout.Utils.getQueryParam('source');
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
            // try to parse the form ID from the hash or querystring
            this.itemId = Shockout.Utils.getIdFromHash();
            var idFromQs = Shockout.Utils.getQueryParam(this.queryStringId);
            if (!!!this.itemId && /\d/.test(idFromQs)) {
                // get the SP list item ID of the form in the querystring
                this.itemId = parseInt(idFromQs);
                Shockout.Utils.setIdHash(this.itemId);
            }
            // setup static error log list name and site uri
            SPForm.errorLogListName = this.errorLogListName;
            SPForm.errorLogSiteUrl = this.errorLogSiteUrl;
            SPForm.enableErrorLog = this.enableErrorLog;
            // initialize custom Knockout handlers
            Shockout.KoHandlers.bindKoHandlers();
            // create instance of the Knockout View Model
            this.viewModel = new Shockout.ViewModel(this);
            this.viewModel.showUserProfiles(this.includeUserProfiles);
            // create element for displaying form load status
            self.$formStatus = $('<div>', { 'class': 'form-status' }).appendTo(self.$form);
            // set the element to display created/modified by info
            self.$createdInfo = self.$form.find(".created-info, [data-sp-created-info]");
            // create jQuery Dialog for displaying feedback to user
            self.$dialog = $('<div>', { 'id': 'formdialog' })
                .appendTo(self.$form)
                .dialog(self.dialogOpts);
            // Cascading Asynchronous Function Execution (CAFE) Array
            // Don't change the order of these unless you know what you're doing.
            this.asyncFns = [
                self.getCurrentUserAsync,
                self.getUsersGroupsAsync,
                function (self) {
                    if (self.preRender) {
                        self.preRender(self, self.viewModel);
                    }
                    self.nextAsync(true);
                },
                self.getListAsync,
                self.initForm,
                function (self, args) {
                    if (args === void 0) { args = undefined; }
                    // Register Shockout's Knockout Components
                    Shockout.KoComponents.registerKoComponents();
                    // apply Knockout bindings
                    ko.applyBindings(self.viewModel, self.form);
                    self.viewModelIsBound = true;
                    self.nextAsync(true);
                },
                self.getListItemAsync,
                self.getHistoryAsync,
                function (self) {
                    if (self.postRender) {
                        self.postRender(self, self.viewModel);
                    }
                    self.nextAsync(true);
                },
                self.implementPermissions,
                self.finalize
            ];
            //start CAFE
            this.nextAsync(true, 'Begin initialization...');
        }
        /////////////////////////////////////
        // Private Set Public Get Properties
        /////////////////////////////////////
        /**
        * Get the current logged in user profile.
        * @return ICurrentUser
        */
        SPForm.prototype.getCurrentUser = function () { return this.currentUser; };
        /**
        * Get the default view for the list.
        * @return string
        */
        SPForm.prototype.getDefaultViewUrl = function () { return this.defaultViewUrl; };
        /**
        * Get the default mobile view for the list.
        * @return string
        */
        SPForm.prototype.getDefailtMobileViewUrl = function () { return this.defailtMobileViewUrl; };
        /**
        * Get a reference to the form element.
        * @return HTMLElement
        */
        SPForm.prototype.getForm = function () { return this.form; };
        /**
        * Get the SP list item ID number.
        * @return number
        */
        SPForm.prototype.getItemId = function () { return this.itemId; };
        /**
        * Get the GUID of the SP list.
        * @return HTMLElement
        */
        SPForm.prototype.getListId = function () { return this.listId; };
        /**
        * Get a reference to the original SP list item.
        * @return ISpItem
        */
        SPForm.prototype.getListItem = function () { return this.listItem; };
        SPForm.prototype.requiresCheckout = function () { return this.requireCheckout; };
        SPForm.prototype.getRootUrl = function () { return this.rootUrl; };
        SPForm.prototype.getSourceUrl = function () { return this.sourceUrl; };
        /**
        * Get a reference to the form's Knockout view model.
        * @return string
        */
        SPForm.prototype.getViewModel = function () { return this.viewModel; };
        /**
        * Get the version number for this framework.
        * @return string
        */
        SPForm.prototype.getVersion = function () { return this.version; };
        /**
        * Execute the next asynchronous function from `asyncFns`.
        * @param success?: boolean = undefined
        * @param msg: string = undefined
        * @param args: any = undefined
        * @return void
        */
        SPForm.prototype.nextAsync = function (success, msg, args) {
            if (success === void 0) { success = true; }
            if (msg === void 0) { msg = undefined; }
            if (args === void 0) { args = undefined; }
            var self = this;
            if (msg) {
                this.updateStatus(msg, success);
            }
            if (!success) {
                return;
            }
            if (this.asyncFns.length == 0) {
                setTimeout(function () {
                    self.$formStatus.hide();
                }, 2000);
                return;
            }
            // execute the next function in the array
            this.asyncFns.shift()(self, args);
        };
        /**
        * Get the current logged in user's profile.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        SPForm.prototype.getCurrentUserAsync = function (self, args) {
            if (args === void 0) { args = undefined; }
            self.updateStatus('Retrieving your account...');
            var success = 'Retrieved your account.';
            if (self.debug) {
                console.info('Testing for SP 2013 API...');
            }
            // If this is SP 2013, it will return thre current user's account.
            Shockout.SpApi15.getCurrentUser(/*callback:*/ function (user, error) {
                if (error == 404) {
                    getSp2010User();
                }
                else {
                    self.isSp2013 = true;
                    self.currentUser = user;
                    self.viewModel.currentUser(user);
                    if (self.debug) {
                        console.info('This is the SP 2013 API.');
                        console.info('Current user is...');
                        console.info(self.viewModel.currentUser());
                    }
                    self.nextAsync(true, success);
                }
            }, /*expandGroups:*/ true);
            function getSp2010User() {
                Shockout.SpSoap.getCurrentUser(function (user, error) {
                    if (!!error) {
                        self.nextAsync(false, 'Failed to retrieve your account. ' + error);
                        return;
                    }
                    self.currentUser = user;
                    self.viewModel.currentUser(user);
                    if (self.debug) {
                        console.info('This is SP 2010 REST services.');
                        console.info('Current user is...');
                        console.info(self.viewModel.currentUser());
                    }
                    self.nextAsync(true, success);
                });
            }
        };
        /**
        * Get metadata about an SP list and the fields to build the Knockout model.
        * Needed to determine the list GUID, if attachments are allowed, and if checkout/in is required.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        SPForm.prototype.getListAsync = function (self, args) {
            if (args === void 0) { args = undefined; }
            // Determine if the field is a `Choice` or `MultiChoice` field with choices.
            var rxIsChoice = /choice/i;
            var rxExcludeNames = /\b(FolderChildCount|ItemChildCount|MetaInfo|ContentType|Edit|Type|LinkTitleNoMenu|LinkTitle|LinkTitle2|Version|Attachments)\b/;
            Shockout.SpSoap.getList(self.siteUrl, self.listName, function (xmlDoc, error) {
                if (!!error) {
                    var msg = 'Failed to retrieve list data. ' + error;
                    self.nextAsync(false, 'Failed to retrieve list data. ' + error);
                    self.logError(msg);
                    return;
                }
                setupList(xmlDoc);
            });
            function setupList(xmlDoc) {
                try {
                    var $list = $(xmlDoc).find('List').first();
                    var listId = $list.attr('ID');
                    self.listId = listId;
                    var requireCheckout = $list.attr('RequireCheckout');
                    self.requireCheckout = !!requireCheckout ? requireCheckout.toLowerCase() == 'true' : false;
                    var enableAttachments = $list.attr('EnableAttachments');
                    self.enableAttachments = !!enableAttachments ? enableAttachments.toLowerCase() == 'true' : false;
                    self.defaultViewUrl = $list.attr('DefaultViewUrl');
                    self.defailtMobileViewUrl = $list.attr('MobileDefaultViewUrl');
                    $(xmlDoc).find('Field').filter(function (i, el) {
                        return !!($(el).attr('DisplayName')) && $(el).attr('Hidden') != 'TRUE' && !rxExcludeNames.test($(el).attr('Name'));
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
                    if (self.debug) {
                        throw e;
                    }
                    var error = 'Failed to initialize list settings.';
                    self.logError(error + ' SPForm.getListAsync.setupList(): ', e);
                    self.nextAsync(false, error);
                }
            }
            function setupKoVar(i, el) {
                if (!!!el) {
                    return;
                }
                try {
                    var $el = $(el);
                    var displayName = $el.attr('DisplayName');
                    var spType = $el.attr('Type');
                    var spName = $el.attr('Name');
                    var spFormat = $el.attr('Format');
                    var spRequired = !!($el.attr('Required')) ? $el.attr('Required').toLowerCase() == 'true' : false;
                    var spReadOnly = !!($el.attr('ReadOnly')) ? $el.attr('ReadOnly').toLowerCase() == 'true' : false;
                    var spDesc = $el.attr('Description');
                    var vm = self.viewModel;
                    // Convert the Display Name to equal REST field name conventions.
                    // For example, convert 'Computer Name (if applicable)' to 'ComputerNameIfApplicable'.
                    var koName = Shockout.Utils.toCamelCase(displayName);
                    // stop and return if it's already a Knockout object
                    if (koName in self.viewModel) {
                        return;
                    }
                    self.fieldNames.push(koName);
                    var defaultValue;
                    // find the SP field's default value if exists
                    $el.find('> Default').each(function (j, def) {
                        var val = $.trim($(def).text());
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
                                    if (val == '[today]') {
                                        val = new Date();
                                    }
                                    break;
                                default:
                                    break;
                            }
                        }
                        defaultValue = val;
                    });
                    var koObj = !!spType && /^multi/i.test(spType)
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
                        $el.find('CHOICE').each(function (j, choice) {
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
                    if (self.debug) {
                        throw e;
                    }
                }
            }
            ;
        };
        /**
        * Initialize the form.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        SPForm.prototype.initForm = function (self, args) {
            if (args === void 0) { args = undefined; }
            try {
                self.updateStatus("Initializing dynamic form features...");
                var vm = self.viewModel;
                var rx = /submitted/i;
                // Register Shockout's Knockout Components
                //KoComponents.registerKoComponents();
                // Find out of this list allows saving before submitting and triggering workflow approval.
                // Must have a field with `submitted` in the name and it must be of type `Boolean`
                if (Shockout.Utils.indexOf(self.fieldNames, 'IsSubmitted') > -1) {
                    self.allowSave = true;
                    Shockout.ViewModel.isSubmittedKey = 'IsSubmitted';
                    if (self.debug) {
                        console.info('initFormAsync: IsSubmitted key: ' + Shockout.ViewModel.isSubmittedKey);
                    }
                }
                // Append action buttons to form.
                self.viewModel.allowSave(self.allowSave);
                self.viewModel.allowPrint(self.allowPrint);
                self.viewModel.allowDelete(self.allowDelete);
                self.$formAction = $(Shockout.Templates.getFormAction()).appendTo(self.$form);
                // Setup attachments modules.
                if (self.enableAttachments) {
                    self.setupAttachments(self);
                }
                // If error logging is enabled, ensure the list exists and has required columns. Disable if 404.
                if (self.enableErrorLog) {
                    // Send a test query
                    Shockout.SpApi.getListItems(self.errorLogListName, function (data, error) {
                        if (!!error) {
                            self.enableErrorLog = SPForm.enableErrorLog = false;
                        }
                    }, self.errorLogSiteUrl, null, 'Title,Error', 'Modified', 1, false);
                }
                //append Created/Modified, Workflow History info to predefined section or append to form
                self.$createdInfo.replaceWith(Shockout.KoComponents.soCreatedModifiedTemplate);
                //append Workflow history section
                if (self.includeWorkflowHistory) {
                    self.$form.append(Shockout.KoComponents.soWorkflowHistoryTemplate);
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
                if (self.debug) {
                    throw e;
                }
                self.logError("Error in SPForm.initFormAsync(): ", e);
                self.nextAsync(false, "Failed to initialize form.");
                return;
            }
        };
        /**
       * Get the SP list item data and build the Knockout view model.
       * @param self: SPForm
       * @param args?: any = undefined
       * @return void
       */
        SPForm.prototype.getListItemAsync = function (self, args) {
            if (args === void 0) { args = undefined; }
            if (!!!self.itemId) {
                self.nextAsync(true, "This is a New form.");
                return;
            }
            self.updateStatus("Retrieving form values...");
            var vm = self.viewModel;
            // expand the REST query for MultiChoice types
            // MAXIMUM is 7!!!
            var expand = [];
            //for (var i = 0; i < self.fieldNames.length; i++) {
            //    var key = self.fieldNames[i];
            //    if (!(key in vm) || !('_type' in vm[key])) { continue; }
            //    if (vm[key]._type == 'MultiChoice') {
            //        expand.push(key);
            //    }
            //}
            if (self.enableAttachments) {
                expand.push('Attachments');
            }
            Shockout.SpApi.getListItem(self.listName, self.itemId, callback, self.siteUrl, false, (expand.length > 0 ? expand.join(',') : null));
            function callback(data, error) {
                if (!!error) {
                    if (/not found/i.test(error + '')) {
                        self.showDialog("The form with ID " + self.itemId + " doesn't exist or it was deleted.");
                    }
                    self.nextAsync(false, error);
                    return;
                }
                self.listItem = data;
                self.bindListItemValues(self);
                self.nextAsync(true, "Retrieved form data.");
            }
        };
        /**
        * Get the SP user groups this user is a member of for removing/showing protected form sections.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        SPForm.prototype.getUsersGroupsAsync = function (self, args) {
            if (args === void 0) { args = undefined; }
            if (self.$form.find("[data-sp-groups], [user-groups]").length == 0 || self.isSp2013) {
                self.nextAsync(true);
                return;
            }
            self.updateStatus("Retrieving your groups...");
            if (self.isSp2013) {
                Shockout.SpApi15.getUsersGroups(self.currentUser.id, callback);
            }
            else {
                Shockout.SpSoap.getUsersGroups(self.currentUser.login, callback);
            }
            function callback(groups, error) {
                if (error) {
                    self.nextAsync(false, "Failed to retrieve your groups. " + error);
                    return;
                }
                self.currentUser.groups = groups;
                if (self.debug) {
                    console.info("Retrieved current user's groups...");
                    console.info(self.currentUser.groups);
                }
                self.nextAsync(true, "Retrieved your groups.");
            }
        };
        /**
        * Removes form sections the user doesn't have access to from the DOM.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        SPForm.prototype.implementPermissions = function (self, args) {
            if (args === void 0) { args = undefined; }
            try {
                self.updateStatus("Retrieving your permissions...");
                // Remove elements from DOM if current user doesn't belong to any of the SP user groups in an element's attribute `data-sp-groups`.
                self.$form.find("[data-sp-groups], [user-groups]").each(function (i, el) {
                    // Provide backward compatibility.
                    // Attribute `user-groups` is deprecated and `data-sp-groups` is preferred for HTML5 "correctness."
                    var groups = $(el).attr("data-sp-groups");
                    if (!!!groups) {
                        groups = $(el).attr("user-groups");
                    }
                    var isMember = self.currentUserIsMemberOfGroups(groups);
                    if (self.debug) {
                        console.info('element is restricted to groups...');
                        console.info(groups);
                    }
                    if (!isMember) {
                        $(el).remove();
                    }
                });
                self.nextAsync(true, "Retrieved your permissions.");
            }
            catch (e) {
                if (self.debug) {
                    throw e;
                }
                self.logError("Error in SPForm.implementPermissionsAsync() ", e);
                self.nextAsync(true, "Failed to retrieve your permissions.");
            }
        };
        /**
        * Get the workflow history for this form, if any.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        SPForm.prototype.getHistoryAsync = function (self, args) {
            if (args === void 0) { args = undefined; }
            if (!!!self.itemId || !self.includeWorkflowHistory) {
                self.nextAsync(true);
                return;
            }
            self.updateStatus('Retrieving workflow history...');
            var filter = "ListID eq '" + self.listId + "' and PrimaryItemID eq " + self.itemId;
            var select = "Description,DateOccurred";
            var orderby = "DateOccurred";
            Shockout.SpApi.getListItems(self.workflowHistoryListName, callback, self.siteUrl, filter, select, orderby, 25, false);
            function callback(items, error) {
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
                    self.viewModel.historyItems().push(new Shockout.HistoryItem(items[i].Description, Shockout.Utils.parseDate(items[i].DateOccurred)));
                }
                self.viewModel.historyItems.valueHasMutated();
                self.nextAsync(true, "Retrieved workflow history.");
            }
        };
        /**
        * Bind the SP list item values to the view model.
        * @param self: SPForm
        * @param args?: any = undefined
        * @return void
        */
        SPForm.prototype.bindListItemValues = function (self) {
            if (self === void 0) { self = undefined; }
            self = self || this;
            try {
                if (!!!self.itemId) {
                    return;
                }
                var item = self.listItem;
                var vm = self.viewModel;
                // Exclude these read-only metadata fields from the Knockout view model.
                var rxExclude = /(__metadata|ContentTypeID|ContentType|Owshiddenversion|Version|Attachments|Path)/;
                var rxExcludeTypes = /(MultiChoice|User|Choice)/;
                var isObj = /Object/;
                self.itemId = item.Id;
                vm.Id(item.Id);
                for (var key in self.viewModel) {
                    if (!(key in item) || !('_type' in vm[key]) || rxExclude.test(key) || rxExcludeTypes.test(vm[key]._type)) {
                        continue;
                    }
                    if ((item[key] != null && vm[key]._type == 'DateTime')) {
                        vm[key](Shockout.Utils.parseDate(item[key]));
                    }
                    else if (vm[key]._type == 'MultiChoice' && 'results' in item[key]) {
                        vm[key](item[key].results);
                    }
                    else {
                        vm[key](item[key] || null);
                    }
                }
                if (self.enableAttachments) {
                    self.viewModel.attachments(item.Attachments['results']);
                    self.viewModel.attachments.valueHasMutated();
                }
                // Created/Modified
                var createdBy = item.CreatedBy;
                var modifiedBy = item.ModifiedBy;
                item.CreatedBy.Picture = Shockout.Utils.formatPictureUrl(item.CreatedBy.Picture); //format picture urls
                item.ModifiedBy.Picture = Shockout.Utils.formatPictureUrl(item.ModifiedBy.Picture);
                // Property name shims for variations among SP 2010 & 2013 and User Info List vs. UPS.
                // Email 
                item.CreatedBy.WorkEMail = item.CreatedBy.WorkEMail || item.CreatedBy.EMail || '';
                item.ModifiedBy.WorkEMail = item.ModifiedBy.WorkEMail || item.ModifiedBy.EMail || '';
                // Job Title
                item.CreatedBy.JobTitle = item.CreatedBy.JobTitle || item.CreatedBy.Title || null;
                item.ModifiedBy.JobTitle = item.ModifiedBy.JobTitle || item.ModifiedBy.Title || null;
                // Phone 
                item.CreatedBy.WorkPhone = item.CreatedBy.WorkPhone || createdBy.MobileNumber || null;
                item.ModifiedBy.WorkPhone = item.ModifiedBy.WorkPhone || modifiedBy.MobileNumber || null;
                // Office 
                item.CreatedBy.Office = item.CreatedBy.Office || null;
                item.ModifiedBy.Office = item.ModifiedBy.Office || null;
                vm.CreatedBy(item.CreatedBy);
                vm.ModifiedBy(item.ModifiedBy);
                vm.Created(Shockout.Utils.parseDate(item.Created));
                vm.Modified(Shockout.Utils.parseDate(item.Modified));
                // Object types `Choice` and `User` will have a corresponding key name plus the suffix `Value` or `Id` for lookups.
                // For example: `SupervisorApproval` is an object container for `__deferred` that corresponds to `SupervisorApprovalValue` which is an ID or string value.
                // query values for the `User` types
                $(self.fieldNames).filter(function (i, key) {
                    if (!!!self.viewModel[key]) {
                        return false;
                    }
                    return self.viewModel[key]._type == 'User' && (key + 'Id') in item;
                }).each(function (i, key) {
                    self.getPersonById(parseInt(item[key + 'Id']), vm[key]);
                });
                // query values for `Choice` types
                $(self.fieldNames).filter(function (i, key) {
                    if (!!!self.viewModel[key]) {
                        return false;
                    }
                    return self.viewModel[key]._type == 'Choice' && (key + 'Value' in item);
                }).each(function (i, key) {
                    vm[key](item[key + 'Value']);
                });
                // query values for MultiChoice types
                $(self.fieldNames).filter(function (i, key) {
                    return !!self.viewModel[key] && self.viewModel[key]._type == 'MultiChoice' && '__deferred' in item[key];
                }).each(function (i, key) {
                    Shockout.SpApi.executeRestRequest(item[key].__deferred.uri, function (data, status, jqXhr) {
                        if (self.debug) {
                            console.info('Retrieved MultiChoice data for ' + key + '...');
                            console.info(data);
                        }
                        var values = [];
                        $.each(data.d.results, function (i, choice) {
                            values.push(choice.Value);
                        });
                        vm[key](values);
                    });
                });
                // query values for UserMulti types
                $(self.fieldNames).filter(function (i, key) {
                    return !!self.viewModel[key] && self.viewModel[key]._type == 'UserMulti' && '__deferred' in item[key];
                }).each(function (i, key) {
                    Shockout.SpApi.executeRestRequest(item[key].__deferred.uri, function (data, status, jqXhr) {
                        if (self.debug) {
                            console.info('Retrieved UserMulti data for ' + key + '...');
                            console.info(data);
                        }
                        var values = [];
                        $.each(data.d.results, function (i, p) {
                            values.push(p.Id + ';#' + p.Account);
                        });
                        vm[key](values);
                    });
                });
            }
            catch (e) {
                if (self.debug) {
                    throw e;
                }
                self.logError('Failed to bind form values in SPForm.bindListItemValues(): ', e);
            }
        };
        /**
        * Delete the list item.
        * @param model: IViewModel
        * @param callback?: Function = undefined
        * @return void
        */
        SPForm.prototype.deleteListItem = function (model, callback, timeout) {
            if (callback === void 0) { callback = undefined; }
            if (timeout === void 0) { timeout = 3000; }
            if (!confirm('Are you sure you want to delete this form?')) {
                return;
            }
            var self = model.parent;
            var item = self.listItem;
            Shockout.SpApi.deleteListItem(item, function (data, error) {
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
        };
        /**
        * Save list item via SOAP services.
        * @param vm: IViewModel
        * @param isSubmit?: boolean = false
        * @param refresh?: boolean = false
        * @param customMsg?: string = undefined
        * @return void
        */
        SPForm.prototype.saveListItem = function (vm, isSubmit, refresh, customMsg) {
            if (isSubmit === void 0) { isSubmit = false; }
            if (refresh === void 0) { refresh = false; }
            if (customMsg === void 0) { customMsg = undefined; }
            var self = vm.parent;
            var isNew = !!(self.itemId == null), data = [], timeout = 3000, saveMsg = customMsg || '<p>Your form has been saved.</p>', fields = [];
            try {
                // Build array of SP field names for the input fields remaning on the form.
                // These are the field names to be saved and current user is allowed to edit these.
                var editable = Shockout.Utils.getEditableKoNames(self.$form);
                $(editable).each(function (i, key) {
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
                    var retVal = self.preSave(self, self.viewModel);
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
                var isSubmitted = vm[Shockout.ViewModel.isSubmittedKey];
                if (typeof (isSubmitted) != "undefined" && (isSubmitted() == null || isSubmitted() == false)) {
                    fields.push([Shockout.ViewModel.isSubmittedKey, (isSubmit ? 1 : 0)]);
                }
                $(self.editableFields).each(function (i, key) {
                    if (!('_metadata' in vm[key])) {
                        return;
                    }
                    var val = vm[key]();
                    var spType = vm[key]._type || vm[key]._metadata.type;
                    spType = !!spType ? spType.toLowerCase() : null;
                    if (typeof (val) == "undefined" || key == Shockout.ViewModel.isSubmittedKey) {
                        return;
                    }
                    if (val != null && val.constructor === Array) {
                        if (val.length > 0) {
                            val = val.join(';#') + ';#';
                        }
                    }
                    else if (spType == 'datetime' && Shockout.Utils.parseDate(val) != null) {
                        val = Shockout.Utils.parseDate(val).toISOString();
                    }
                    else if (val != null && spType == 'note') {
                        val = '<![CDATA[' + $('<div>').html(val).html() + ']]>';
                    }
                    val = val == null ? '' : val;
                    fields.push([vm[key]._name, val]);
                });
                Shockout.SpSoap.updateListItem(self.itemId, self.listName, fields, isNew, self.siteUrl, callback);
            }
            catch (e) {
                if (self.debug) {
                    throw e;
                }
                self.logError('Error in SpForm.saveListItem(): ', e);
            }
            function callback(xmlDoc, status, jqXhr) {
                var itemId;
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
                $(xmlDoc).find('*').filter(function () {
                    return this.nodeName == 'z:row';
                }).each(function (i, el) {
                    itemId = parseInt($(el).attr('ows_ID'));
                    if (self.itemId == null) {
                        self.itemId = itemId;
                    }
                    if (self.debug) {
                        console.info('Item ID returned...');
                        console.info(itemId);
                    }
                });
                if (Shockout.Utils.getIdFromHash() == null && self.itemId != null) {
                    Shockout.Utils.setIdHash(self.itemId);
                }
                if (isSubmit) {
                    self.showDialog('<p>Your form has been submitted. You will be redirected in ' + timeout / 1000 + ' seconds.</p>', 'Form Submission Successful');
                    if (self.debug) {
                        console.warn('DEBUG MODE: Would normally redirect user to confirmation page: ' + self.confirmationUrl);
                    }
                    else {
                        setTimeout(function () {
                            window.location.href = self.sourceUrl != null ? self.sourceUrl : self.confirmationUrl;
                        }, timeout);
                    }
                }
                else {
                    self.showDialog(saveMsg, 'The form has been saved.', timeout);
                    // refresh data from the server
                    self.getListItemAsync(self);
                    //give WF History list 5 seconds to update
                    if (self.includeWorkflowHistory) {
                        setTimeout(function () { self.getHistoryAsync(self); }, 5000);
                    }
                }
            }
            ;
        };
        /**
        * Add a navigation menu to the form based on parent elements with class `nav-section`
        * @param salef: SPForm
        * @return void
        */
        SPForm.prototype.finalize = function (self) {
            try {
                // Setup form navigation on sections with class '.nav-section'
                self.setupNavigation(self);
                // Setup Datepickers.
                self.setupDatePickers(self);
                // Setup Bootstrap validation.
                //self.setupBootstrapValidation(self);
                self.nextAsync(true, 'Finalized form controls.');
            }
            catch (e) {
                if (self.debug) {
                    throw e;
                }
                self.logError('Error in SpForm.finalize(): ', e);
                self.nextAsync(false, 'Failed to finalize form controls.');
            }
        };
        /**
        * Delete an attachment.
        */
        SPForm.prototype.deleteAttachment = function (att, event) {
            if (!confirm('Are you sure you want to delete ' + att.Name + '? This can\'t be undone.')) {
                return;
            }
            var self = Shockout.ViewModel.parent;
            var vm = self.viewModel;
            Shockout.SpApi.deleteAttachment(att, function (data, error) {
                if (!!error) {
                    alert("Failed to delete attachment: " + error);
                    return;
                }
                var attachments = vm.attachments;
                attachments.remove(att);
            });
        };
        /**
        * Get the form's attachments
        * @param self: SFForm
        * @param callback: Function (optional)
        * @return void
        */
        SPForm.prototype.getAttachments = function (self, callback) {
            if (self === void 0) { self = undefined; }
            if (callback === void 0) { callback = undefined; }
            self = self || this;
            if (!!!self.listItem || !self.enableAttachments) {
                if (callback) {
                    callback();
                }
                return;
            }
            var attachments = [];
            Shockout.SpApi.executeRestRequest(self.listItem.__metadata.uri + '/Attachments', function (data, status, jqXhr) {
                try {
                    self.viewModel.attachments(data.d.results);
                    self.viewModel.attachments.valueHasMutated();
                    if (callback) {
                        callback(attachments);
                    }
                }
                catch (e) {
                    if (self.debug) {
                        throw e;
                    }
                    self.showDialog("Failed to retrieve attachments in SpForm.getAttachments(): ", e);
                }
            });
        };
        /**
        * Log to console in degug mode.
        * @param msg: string
        * @return void
        */
        SPForm.prototype.log = function (msg) {
            if (this.debug) {
                console.log(msg);
            }
        };
        /**
        * Update the form status to display feedback to the user.
        * @param msg: string
        * @param success?: boolean = undefined
        * @return void
        */
        SPForm.prototype.updateStatus = function (msg, success) {
            if (success === void 0) { success = true; }
            var self = this;
            this.$formStatus
                .html(msg)
                .css('color', (success ? "#ff0" : "$f00"))
                .show();
            setTimeout(function () { self.$formStatus.hide(); }, 2000);
        };
        /**
        * Display a message to the user with jQuery UI Dialog.
        * @param msg: string
        * @param title?: string = undefined
        * @param timeout?: number = undefined
        * @return void
        */
        SPForm.prototype.showDialog = function (msg, title, timeout) {
            if (title === void 0) { title = undefined; }
            if (timeout === void 0) { timeout = undefined; }
            var self = this;
            title = title || "Form Dialog";
            msg = (msg).toString().match(/<\w>\w*/) == null ? '<p>' + msg + '</p>' : msg; //wrap non-html in <p>
            self.$dialog.html(msg).dialog('open');
            if (timeout) {
                setTimeout(function () { self.$dialog.dialog('close'); }, timeout);
            }
        };
        /**
        * Validate the View Model's required fields
        * @param model: IViewModel
        * @param showDialog?: boolean = false
        * @return bool
        */
        SPForm.prototype.formIsValid = function (model, showDialog) {
            if (showDialog === void 0) { showDialog = false; }
            var self = model.parent, labels = [], errorCount = 0, invalidCount = 0, invalidLabels = [];
            try {
                self.$form.find('.required, [required]').each(function checkRequired(i, n) {
                    var koName = Shockout.Utils.observableNameFromControl(n, self.viewModel);
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
                            if (Shockout.Utils.indexOf(labels, labelTxt) < 0) {
                                labels.push(labelTxt);
                                errorCount++;
                            }
                        }
                    }
                });
                //check for sp object data errors before saving
                self.$form.find(".invalid").each(function (i, el) {
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
                if (self.debug) {
                    throw e;
                }
                self.logError("Form validation error at SPForm.formIsValid(): ", e);
                return false;
            }
        };
        /**
        * Get a person by their ID from the User Information list.
        * @param id: number
        * @param callback: Function
        * @return void
        */
        SPForm.prototype.getPersonById = function (id, koField) {
            var self = this;
            if (!!!id) {
                return;
            }
            Shockout.SpApi.getPersonById(id, function (person, error) {
                if (!!error) {
                    var msg = 'Error in SPForm.getPersonById: ' + error;
                    Shockout.Utils.logError(msg, SPForm.errorLogListName);
                    if (self.debug) {
                        console.warn(msg);
                    }
                    return;
                }
                var name = person.Id + ';#' + person.Name;
                koField(name);
                if (self.debug) {
                    console.warn('Retrieved person by ID... ' + name);
                }
            });
        };
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
        SPForm.prototype.pushEditableFieldName = function (key) {
            if (!!!key || Shockout.Utils.indexOf(this.editableFields, key) > -1 || key.match(/^(_|\$)/) != null || Shockout.Utils.indexOf(this.fieldNames, key) < 0 || this.viewModel[key]._readOnly) {
                return -1;
            }
            return this.editableFields.push(key);
        };
        /**
        * Log errors to designated SP list.
        * @param msg: string
        * @param self?: SPForm = undefined
        * @return void
        */
        SPForm.prototype.logError = function (msg, e, self) {
            if (e === void 0) { e = undefined; }
            if (self === void 0) { self = undefined; }
            self = self || this;
            var err = [msg];
            if (!!e) {
                err.push(e + '');
            }
            err = err.length > 0 ? err.join('; ') : err.join('');
            if (self.enableErrorLog) {
                Shockout.Utils.logError(err, self.errorLogListName, self.rootUrl, self.debug);
                self.showDialog('<p>An error has occurred and the web administrator has been notified.</p><pre>' + err + '</pre>');
            }
        };
        /**
        * Setup attachments modules.
        * @param self: SPForm = undefined
        * @return number
        */
        SPForm.prototype.setupAttachments = function (self) {
            if (self === void 0) { self = undefined; }
            self = self || this;
            var vm = self.viewModel;
            var count = 0;
            if (!self.enableAttachments) {
                return count;
            }
            try {
                // set the absolute URI for the file handler 
                var subsiteUrl = Shockout.Utils.formatSubsiteUrl(self.siteUrl); // ensure site url is or ends with '/'
                var fileHandlerUrl = self.fileHandlerUrl.replace(/^\//, '');
                self.fileHandlerUrl = self.rootUrl + subsiteUrl + fileHandlerUrl;
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
                            console.info('Response from SpForm.fileUploaderSettings.onComplete()');
                            console.info(json);
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
                            self.saveListItem(vm, false);
                        }
                        // push a new SP attachment instance to the view model's `attachments` collection
                        self.viewModel.attachments().push(new Shockout.SpAttachment(self.rootUrl, self.siteUrl, self.listName, self.itemId, fileName));
                        self.viewModel.attachments.valueHasMutated(); // tell KO the array has been updated
                    },
                    template: Shockout.Templates.getFileUploadTemplate()
                };
                self.$form.find(".attachments, [data-sp-attachments]").each(function (i, att) {
                    var id = 'so-qq-fileuploader_' + i;
                    $(att).replaceWith(Shockout.Templates.getAttachmentsTemplate(id));
                    self.fileUploaderSettings.element = document.getElementById(id);
                    self.fileUploader = new Shockout.qq.FileUploader(self.fileUploaderSettings);
                    count++;
                });
                if (self.debug) {
                    console.info('Attachments are enabled.');
                }
            }
            catch (e) {
                if (self.debug) {
                    throw e;
                }
                self.logError('Error in SPForm.setupAttachments(): ', e);
            }
            return count;
        };
        /**
        * Setup Bootstrap validation for required fields.
        * @return number
        */
        SPForm.prototype.setupBootstrapValidation = function (self) {
            if (self === void 0) { self = undefined; }
            var count = 0;
            self = self || this;
            try {
                // add control validation to Bootstrap form elements
                // http://getbootstrap.com/css/#forms-control-validation 
                self.$form.find('[required], .required').each(function (i, el) {
                    var $parent = $(el).closest('.form-group');
                    var db = $parent.attr('data-bind');
                    if (/has-error/.test(db)) {
                        return;
                    } // already has the KO bindings
                    var koName = Shockout.Utils.observableNameFromControl(el, self.viewModel);
                    var css = "css:{ 'has-error': !!!" + koName + "(), 'has-success has-feedback': !!" + koName + "()}";
                    // If the parent already has a data-bind attribute, append the css.
                    if (!!db) {
                        var dataBind = $parent.attr("data-bind");
                        $parent.attr("data-bind", dataBind + ", " + css);
                    }
                    else {
                        $parent.attr("data-bind", css);
                    }
                    $parent.append('<span class="glyphicon glyphicon-ok form-control-feedback" aria-hidden="true"></span>');
                    count++;
                });
            }
            catch (e) {
                if (self.debug) {
                    throw e;
                }
                self.logError('Error in SPForm.setupBootstrapValidation(): ', e);
            }
            return count;
        };
        /**
        * Setup form navigation on sections with class '.nav-section'
        * @return number
        */
        SPForm.prototype.setupNavigation = function (self) {
            if (self === void 0) { self = undefined; }
            var self = self || this;
            var count = 0;
            try {
                // Set up a navigation menu at the top of the form if there are elements with the class `nav-section`.
                var $navSections = self.$form.find('.nav-section');
                if ($navSections.length == 0) {
                    return count;
                }
                // add navigation section to top of form
                self.$form.prepend('<section class="no-print" id="TOP">' +
                    '<h4>Navigation</h4>' +
                    '<div class="navigation-buttons"></div>' +
                    '</section>');
                // include the workflow history section
                self.$form.find('#workflowHistory, [data-workflow-history]').addClass('nav-section');
                // add navigation buttons
                self.$form.find(".nav-section:visible").each(function (i, el) {
                    var $el = $(el);
                    var $header = $el.find("> h4");
                    if ($header.length == 0) {
                        return;
                    }
                    var title = $header.text();
                    var anchorName = Shockout.Utils.toCamelCase(title) + 'Nav';
                    $el.before('<div style="height:1px;" id="' + anchorName + '">&nbsp;</div>');
                    self.$form.find(".navigation-buttons").append('<a href="#' + anchorName + '" class="btn btn-sm btn-info">' + title + '</a>');
                    count++;
                });
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
                if (self.debug) {
                    throw e;
                }
                self.logError('Error in SpForm.setupNavigation(): ', e);
            }
            return count;
        };
        /**
        * Setup Datepicker fields.
        * @return number
        */
        SPForm.prototype.setupDatePickers = function (self) {
            if (self === void 0) { self = undefined; }
            self = self || this;
            // Apply jQueryUI datepickers after all KO bindings have taken place to prevent error: 
            // `Uncaught Missing instance data for this datepicker`
            var $datepickers = self.$form.find('input.datepicker').datepicker();
            if (self.debug) {
                console.info('Bound ' + $datepickers.length + ' jQueryUI datepickers.');
            }
            return $datepickers.length;
        };
        // obsolete
        //setupHtmlFields(self: SPForm = undefined): number {
        //    self = self || this;
        //    var count: number = 0;
        //    try {
        //        // set up HTML editors in the form
        //        // This isn't necessary for the Shockout KO Components fields, but included for when a developer creates their own fields.
        //        self.$form.find(".rte, [data-bind*='spHtml'], [data-sp-html]").each(function (i: number, el: HTMLElement) {
        //            var $el = $(el);
        //            var koName = Utils.observableNameFromControl(el, self.viewModel);
        //            var $rte = $('<div>', {
        //                'data-bind': 'spHtmlEditor: ' + koName,
        //                'class': 'form-control content-editable',
        //                'contenteditable': 'true'
        //            });
        //            if (!!$el.attr('required') || !!$el.hasClass('required')) {
        //                $rte.attr('required', 'required');
        //                $rte.addClass('required');
        //            }
        //            $rte.insertBefore($el);
        //            if (!self.debug) {
        //                $el.hide();
        //            }
        //            count++;
        //            if (self.debug) {
        //                console.info('initFormAsync: Created spHtml field: ' + koName);
        //            }
        //        });
        //    }
        //    catch (e) {
        //        if (self.debug) { throw e; }
        //        self.logError('Error in SPForm.setupHtmlFields(): ', e);
        //    }
        //    return count;
        //}
        /**
        * Determine if the current user is a member of at least one of list of target SharePoint groups.
        * @param targetGroups: comma delimited string || Array<string>
        * @return boolean
        */
        SPForm.prototype.currentUserIsMemberOfGroups = function (targetGroups) {
            var groupNames = [];
            if (Shockout.Utils.isString(targetGroups)) {
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
                var group = groupNames[i];
                group = group.match(/\;#/) != null ? group.split(';')[0] : group; //either id;#groupname or groupname
                group = Shockout.Utils.trim(group);
                for (var j = 0; j < this.currentUser.groups.length; j++) {
                    var g = this.currentUser.groups[j];
                    if (group == g.name || parseInt(group) == g.id) {
                        return true;
                    }
                }
            }
            return false;
        };
        SPForm.DEBUG = false;
        return SPForm;
    })();
    Shockout.SPForm = SPForm;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var ViewModel = (function () {
        function ViewModel(instance) {
            // SP List Item Fields
            this.Id = ko.observable(null);
            this.Created = ko.observable(null);
            this.CreatedBy = ko.observable(null);
            this.Modified = ko.observable(null);
            this.ModifiedBy = ko.observable(null);
            this.allowSave = ko.observable(false);
            this.allowPrint = ko.observable(false);
            this.allowDelete = ko.observable(false);
            this.attachments = ko.observableArray();
            this.historyItems = ko.observableArray();
            this.showUserProfiles = ko.observable(false);
            var self = this;
            this.parent = instance;
            ViewModel.parent = instance;
            this.isValid = ko.pureComputed(function () {
                return self.parent.formIsValid(self);
            });
            this.deleteAttachment = instance.deleteAttachment;
            this.currentUser = ko.observable(instance.getCurrentUser());
        }
        ViewModel.prototype.isAuthor = function () {
            if (!!!this.CreatedBy()) {
                return true;
            }
            return this.currentUser().id == this.CreatedBy().Id;
        };
        ViewModel.prototype.deleteItem = function () {
            this.parent.deleteListItem(this);
        };
        ViewModel.prototype.cancel = function () {
            var src = this.parent.getSourceUrl();
            window.location.href = !!src ? src : this.parent.getRootUrl();
        };
        ViewModel.prototype.print = function () {
            window.print();
        };
        ViewModel.prototype.save = function (model, btn) {
            this.parent.saveListItem(model, false);
        };
        ViewModel.prototype.submit = function (model, btn) {
            this.parent.saveListItem(model, true);
        };
        return ViewModel;
    })();
    Shockout.ViewModel = ViewModel;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var KoHandlers = (function () {
        function KoHandlers() {
        }
        KoHandlers.bindKoHandlers = function () {
            bindKoHandlers(ko);
        };
        return KoHandlers;
    })();
    Shockout.KoHandlers = KoHandlers;
    /* Knockout Custom handlers */
    function bindKoHandlers(ko) {
        ko.bindingHandlers['spHtmlEditor'] = {
            init: function (element, valueAccessor, allBindings, vm) {
                var koName = Shockout.Utils.observableNameFromControl(element);
                $(element)
                    .blur(update)
                    .change(update)
                    .keydown(update);
                function update() {
                    vm[koName]($(this).html());
                }
            },
            update: function (element, valueAccessor, allBindings, vm) {
                var value = ko.utils.unwrapObservable(valueAccessor()) || "";
                if (element.innerHTML !== value) {
                    element.innerHTML = value;
                }
            }
        };
        /* SharePoint People Picker */
        ko.bindingHandlers['spPerson'] = {
            init: function (element, valueAccessor, allBindings, bindingContext) {
                try {
                    // stop if not an editable field 
                    if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') {
                        return;
                    }
                    // This will be called when the binding is first applied to an element
                    // Set up any initial state, event handlers, etc. here
                    var viewModel = bindingContext.$data, modelValue = valueAccessor(), person = ko.unwrap(modelValue);
                    var $element = $(element)
                        .addClass('people-picker-control')
                        .attr('placeholder', 'Employee Account Name');
                    //create wrapper for control
                    var $parent = $(element).parent();
                    var $spError = $('<div>', { 'class': 'sp-validation person' });
                    $element.after($spError);
                    var $desc = $('<div>', {
                        'class': 'no-print',
                        'html': '<em>Enter the employee name. The auto-suggest menu will appear below the field. Select the account name.</em>'
                    });
                    $spError.after($desc);
                    //controls
                    var $spValidate = $('<button>', {
                        'html': '<span class="glyphicon glyphicon-user"></span>',
                        'class': 'btn btn-sm btn-default no-print',
                        'title': 'Validate the employee account name.'
                    }).on('click', function () {
                        if ($.trim($element.val()) == '') {
                            $element.removeClass('invalid').removeClass('valid');
                            return false;
                        }
                        if (!Shockout.Utils.validateSpPerson(modelValue())) {
                            $spError.text('Invalid').addClass('error').show();
                            $element.addClass('invalid').removeClass('valid');
                        }
                        else {
                            $spError.text('Valid').removeClass('error');
                            $element.removeClass('invalid').addClass('valid').show();
                        }
                        return false;
                    }).insertAfter($element);
                    var $reset = $('<button>', { 'class': 'btn btn-sm btn-default reset', 'html': 'Reset' })
                        .on('click', function () {
                        modelValue(null);
                        return false;
                    })
                        .insertAfter($spValidate);
                    var autoCompleteOpts = {
                        source: function (request, response) {
                            // Use People.asmx instead of REST services against the User Information List, 
                            // which allows you to search users that haven't logged into SharePoint yet.
                            // Thanks to John Kerski from Definitive Logic for the suggestion.
                            Shockout.SpSoap.searchPrincipals(request.term, function (data) {
                                response($.map(data, function (item) {
                                    return {
                                        label: item.DisplayName + ' (' + item.Email + ')',
                                        value: item.UserInfoID + ';#' + item.AccountName
                                    };
                                }));
                            }, 10, 'User');
                        },
                        minLength: 3,
                        select: function (event, ui) {
                            modelValue(ui.item.value);
                        }
                    };
                    $(element).autocomplete(autoCompleteOpts);
                    $(element).on('focus', function () { $(this).removeClass('valid'); })
                        .on('blur', function () { onChangeSpPersonEvent(this, modelValue); })
                        .on('mouseout', function () { onChangeSpPersonEvent(this, modelValue); });
                }
                catch (e) {
                    if (Shockout.SPForm.DEBUG) {
                        console.info('Error in Knockout handler spPerson init()');
                        console.info(e);
                    }
                }
                function onChangeSpPersonEvent(self, modelValue) {
                    var value = $.trim($(self).val());
                    if (value == '') {
                        modelValue(null);
                        $(self).removeClass('valid').removeClass('invalid');
                        return;
                    }
                    if (Shockout.Utils.validateSpPerson(modelValue())) {
                        $(self).val(modelValue().split('#')[1]);
                        $(self).addClass('valid').removeClass('invalid');
                    }
                    else {
                        $(self).removeClass('valid').addClass('invalid');
                    }
                }
                ;
            },
            update: function (element, valueAccessor, allBindings, bindingContext) {
                // This will be called once when the binding is first applied to an element,
                // and again whenever any observables/computeds that are accessed change
                // Update the DOM element based on the supplied values here.
                try {
                    var viewModel = bindingContext.$data;
                    // First get the latest data that we're bound to
                    var modelValue = valueAccessor();
                    // Next, whether or not the supplied model property is observable, get its current value
                    var person = ko.unwrap(modelValue);
                    // Now manipulate the DOM element
                    var displayName = "";
                    if (Shockout.Utils.validateSpPerson(person)) {
                        displayName = person.split('#')[1];
                        $(element).addClass("valid");
                    }
                    if ('value' in element) {
                        $(element).val(displayName);
                    }
                    else {
                        $(element).text(displayName);
                    }
                }
                catch (e) {
                    if (Shockout.SPForm.DEBUG) {
                        console.info('Error in Knockout handler spPerson update()');
                        console.info(e);
                    }
                }
            }
        };
        ko.bindingHandlers['spMoney'] = {
            'init': function (element, valueAccessor, allBindings, viewModel, bindingContext) {
                /* stop if not an editable field */
                if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') {
                    return;
                }
                viewModel = bindingContext.$data;
                var value = valueAccessor();
                var valueUnwrapped = ko.unwrap(value);
                $(element).on('blur', onChange).on('change', onChange);
                function onChange() {
                    var val = this.value.toString().replace(/[^\d\.\-]/g, '');
                    val = val == '' ? null : (val - 0);
                    value(val);
                }
                ;
            },
            'update': function (element, valueAccessor, allBindings, viewModel, bindingContext) {
                viewModel = bindingContext.$data;
                var value = valueAccessor();
                var valueUnwrapped = ko.unwrap(value);
                if (valueUnwrapped != null) {
                    if (valueUnwrapped < 0) {
                        $(element).addClass('negative');
                    }
                    else {
                        $(element).removeClass('negative');
                    }
                }
                else {
                    valueUnwrapped = 0;
                }
                var formattedValue = Shockout.Utils.formatMoney(valueUnwrapped);
                Shockout.Utils.updateKoField(element, formattedValue);
            }
        };
        ko.bindingHandlers['spDecimal'] = {
            'init': function (element, valueAccessor, allBindings, viewModel, bindingContext) {
                // stop if not an editable field 
                if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') {
                    return;
                }
                viewModel = bindingContext.$data;
                var value = valueAccessor();
                var valueUnwrapped = ko.unwrap(value);
                $(element).on('blur', onChange).on('change', onChange);
                function onChange() {
                    var val = this.value.toString().replace(/[^\d\-\.]/g, '');
                    val = val == '' ? null : (val - 0);
                    value(val);
                }
                ;
            },
            'update': function (element, valueAccessor, allBindings, viewModel, bindingContext) {
                viewModel = bindingContext.$data;
                var value = valueAccessor();
                var valueUnwrapped = ko.unwrap(value);
                var precision = allBindings.get('precision') || 2;
                var formattedValue = Shockout.Utils.toFixed(valueUnwrapped, precision);
                if (valueUnwrapped != null) {
                    if (valueUnwrapped < 0) {
                        $(element).addClass('negative');
                    }
                    else {
                        $(element).removeClass('negative');
                    }
                }
                else {
                    valueUnwrapped = 0;
                }
                Shockout.Utils.updateKoField(element, formattedValue);
            }
        };
        ko.bindingHandlers['spNumber'] = {
            /* executes on load */
            init: function (element, valueAccessor, allBindings, viewModel, bindingContext) {
                /* stop if not an editable field */
                if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') {
                    return;
                }
                viewModel = bindingContext.$data;
                var value = valueAccessor();
                var valueUnwrapped = ko.unwrap(value);
                $(element).on('blur', onChange).on('change', onChange);
                function onChange() {
                    var val = this.value.toString().replace(/[^\d\-]/g, '');
                    val = val == '' ? null : (val - 0);
                    value(val);
                }
                ;
            },
            /* executes on load and on change */
            update: function (element, valueAccessor, allBindings, viewModel, bindingContext) {
                viewModel = bindingContext.$data;
                var value = valueAccessor();
                var valueUnwrapped = ko.unwrap(value);
                valueUnwrapped = valueUnwrapped == null ? 0 : valueUnwrapped;
                valueUnwrapped = valueUnwrapped.constructor == String ? valueUnwrapped = valueUnwrapped.replace(/\D/g) - 0 : valueUnwrapped;
                Shockout.Utils.updateKoField(element, valueUnwrapped);
                if (value.constructor == Function) {
                    value(valueUnwrapped);
                }
            }
        };
        ko.bindingHandlers['spDate'] = {
            after: ['attr'],
            init: function (element, valueAccessor, allBindings, viewModel, bindingContext) {
                var modelValue = valueAccessor();
                if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') {
                    return;
                } // stop if not an editable field
                $(element)
                    .css('display', 'inline-block')
                    .addClass('datepicker med')
                    .attr('placeholder', 'MM/DD/YYYY')
                    .on('blur', onDateChange)
                    .on('change', onDateChange)
                    .after('<span class="glyphicon glyphicon-calendar"></span>');
                $(element).datepicker({
                    changeMonth: true,
                    changeYear: true
                });
                function onDateChange() {
                    modelValue(Shockout.Utils.parseDate(this.value));
                }
                ;
            },
            update: function (element, valueAccessor, allBindings, viewModel, bindingContext) {
                var modelValue = valueAccessor();
                var date = Shockout.Utils.parseDate(ko.unwrap(modelValue));
                var dateStr = '';
                if (!!date && date != null) {
                    dateStr = Shockout.Utils.dateToLocaleString(date);
                }
                if ('value' in element) {
                    $(element).val(dateStr);
                }
                else {
                    $(element).text(dateStr);
                }
            }
        };
        // 1. REST returns UTC
        // 2. getUTCHours converts UTC to Locale
        ko.bindingHandlers['spDateTime'] = {
            after: ['attr'],
            init: function (element, valueAccessor, allBindings, viewModel, bindingContext) {
                if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') {
                    return;
                } // stop if not an editable field
                var modelValue = valueAccessor(), required, $hh, $mm, $tt, $display, $error, $element = $(element), $parent = $element.parent(), $reset;
                try {
                    var currentVal = Shockout.Utils.parseDate(modelValue());
                    modelValue(currentVal); // just in case it was a string date
                    var koName = Shockout.Utils.koNameFromControl(element);
                    $display = $('<span>', { 'class': 'no-print', 'style': 'margin-left:1em;' }).insertAfter($element);
                    $error = $('<span>', { 'class': 'error', 'html': 'Invalid Date-time', 'style': 'display:none;' }).insertAfter($element);
                    element.$display = $display;
                    element.$error = $error;
                    required = $element.hasClass('required') || $element.attr('required') != null;
                    $element.attr({
                        'placeholder': 'MM/DD/YYYY',
                        'maxlength': 10,
                        'class': 'datepicker med form-control'
                    }).css('display', 'inline-block')
                        .on('change', setDateTime)
                        .before('<br />');
                    if (required) {
                        $element.attr('required', 'required');
                    }
                    $element.datepicker({
                        changeMonth: true,
                        changeYear: true
                    });
                    var timeHtml = ['<span class="glyphicon glyphicon-calendar"></span>'];
                    // Hours 
                    timeHtml.push('<select class="form-control select-hours" style="margin-left:1em;width:5em;display:inline-block;">');
                    for (var i = 1; i <= 12; i++) {
                        timeHtml.push('<option value="' + i + '">' + (i < 10 ? '0' + i : i) + '</option>');
                    }
                    timeHtml.push('</select>');
                    timeHtml.push('<span> : </span>');
                    // Minutes     
                    timeHtml.push('<select class="form-control select-minutes" style="width:5em;display:inline-block;">');
                    for (var i = 0; i < 60; i++) {
                        timeHtml.push('<option value="' + i + '">' + (i < 10 ? '0' + i : i) + '</option>');
                    }
                    timeHtml.push('</select>');
                    // TT: AM/PM
                    timeHtml.push('<select class="form-control select-tt" style="margin-left:1em;width:5em;display:inline-block;"><option value="AM">AM</option><option value="PM">PM</option></select>');
                    timeHtml.push('&nbsp;<button class="btn btn-sm btn-default reset">Reset</button>');
                    $element.after(timeHtml.join(''));
                    $hh = $parent.find('.select-hours');
                    $mm = $parent.find('.select-minutes');
                    $tt = $parent.find('.select-tt');
                    $reset = $parent.find('.btn.reset');
                    $hh.on('change', setDateTime); //.on('keydown', onKeyDown);
                    $mm.on('change', setDateTime); //.on('keydown', onKeyDown);
                    $tt.on('change', setDateTime);
                    $reset.on('click', function () {
                        try {
                            modelValue(null);
                            $element.val('');
                            $hh.val('12');
                            $mm.val('0');
                            $tt.val('AM');
                            $display.html('');
                        }
                        catch (e) {
                            console.warn(e);
                        }
                        return false;
                    });
                    element.$hh = $hh;
                    element.$mm = $mm;
                    element.$tt = $tt;
                    // set default time
                    if (!!currentVal) {
                        setDateTime();
                    }
                    else {
                        $element.val('');
                        $hh.val('12');
                        $mm.val('0');
                        $tt.val('AM');
                    }
                }
                catch (e) {
                    if (Shockout.SPForm.DEBUG) {
                        console.warn('Error in Knockout handler spDateTime init()...');
                        console.warn(e);
                    }
                }
                // must convert user's locale date/time to UTC for SP
                function setDateTime() {
                    try {
                        var date = Shockout.Utils.parseDate($element.val());
                        if (!!!date) {
                            date = new Date();
                        }
                        var hrs = parseInt($hh.val());
                        var min = parseInt($mm.val());
                        var tt = $tt.val();
                        if (tt == 'PM' && hrs < 12) {
                            hrs += 12;
                        }
                        else if (tt == 'AM' && hrs > 11) {
                            hrs -= 12;
                        }
                        // SP saves date/time in UTC
                        var curDateTime = new Date();
                        curDateTime.setUTCFullYear(date.getFullYear());
                        curDateTime.setUTCMonth(date.getMonth());
                        curDateTime.setUTCDate(date.getDate());
                        curDateTime.setUTCHours(hrs, min, 0, 0);
                        modelValue(curDateTime);
                    }
                    catch (e) {
                        if (Shockout.SPForm.DEBUG) {
                            console.warn(e);
                        }
                    }
                }
                function onKeyDown() {
                    var val = $(this).val().replace(/\d/g, '');
                    $(this).val(val);
                }
                ;
            },
            update: function (element, valueAccessor, allBindings, viewModel, bindingContext) {
                try {
                    var modelValue = valueAccessor();
                    var date = Shockout.Utils.parseDate(ko.unwrap(modelValue));
                    if (typeof modelValue == 'function') {
                        modelValue(date); // just in case it was a string date 
                    }
                    if (!!date) {
                        var dateTimeStr = Shockout.Utils.toDateTimeLocaleString(date); // convert from UTC to locale
                        // add time zone
                        var timeZone = /\b\s\(\w+\s\w+\s\w+\)/i.exec(date.toString());
                        if (!!timeZone) {
                            // e.g. convert '(Central Daylight Time)' to '(CDT)'
                            dateTimeStr += ' ' + timeZone[0].replace(/\b\w+/g, function (x) {
                                return x[0];
                            }).replace(/\s/g, '');
                        }
                        if (element.tagName.toLowerCase() == 'input') {
                            $(element).val((date.getUTCMonth() + 1) + '/' + date.getUTCDate() + '/' + date.getUTCFullYear());
                            var hrs = date.getUTCHours(); // converts UTC hours to locale hours
                            var min = date.getUTCMinutes();
                            // set TT based on military hours
                            if (hrs > 12) {
                                hrs -= 12;
                                element.$tt.val('PM');
                            }
                            else if (hrs == 0) {
                                hrs = 12;
                                element.$tt.val('AM');
                            }
                            else if (hrs == 12) {
                                element.$tt.val('PM');
                            }
                            else {
                                element.$tt.val('AM');
                            }
                            element.$hh.val(hrs);
                            element.$mm.val(min);
                            element.$display.html(dateTimeStr);
                        }
                        else {
                            $(element).text(dateTimeStr);
                        }
                    }
                }
                catch (e) {
                    if (Shockout.SPForm.DEBUG) {
                        console.warn('Error in Knockout handler spDateTime update()...s');
                        console.warn(e);
                    }
                }
            }
        };
    }
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var KoComponents = (function () {
        function KoComponents() {
        }
        KoComponents.registerKoComponents = function () {
            var uniqueId = (function () {
                var i = 0;
                return function () {
                    return i++;
                };
            })();
            ko.components.register('so-text-field', {
                viewModel: soFieldModel,
                template: KoComponents.soTextFieldTemplate
            });
            ko.components.register('so-html-field', {
                viewModel: soFieldModel,
                template: KoComponents.soHtmlFieldTemplate
            });
            ko.components.register('so-person-field', {
                viewModel: soFieldModel,
                template: KoComponents.soTextFieldTemplate.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spPerson: modelValue')
            });
            ko.components.register('so-date-field', {
                viewModel: soFieldModel,
                template: KoComponents.soTextFieldTemplate.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spDate: modelValue')
            });
            ko.components.register('so-datetime-field', {
                viewModel: soFieldModel,
                template: KoComponents.soTextFieldTemplate.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spDateTime: modelValue')
            });
            ko.components.register('so-money-field', {
                viewModel: soFieldModel,
                template: KoComponents.soTextFieldTemplate.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spMoney: modelValue')
            });
            ko.components.register('so-number-field', {
                viewModel: soFieldModel,
                template: KoComponents.soTextFieldTemplate.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spNumber: modelValue')
            });
            ko.components.register('so-decimal-field', {
                viewModel: soFieldModel,
                template: KoComponents.soTextFieldTemplate.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spDecimal: modelValue')
            });
            ko.components.register('so-checkbox-field', {
                viewModel: soFieldModel,
                template: KoComponents.soCheckboxFieldTemplate
            });
            ko.components.register('so-select-field', {
                viewModel: soFieldModel,
                template: KoComponents.soSelectFieldTemplate
            });
            ko.components.register('so-checkbox-group', {
                viewModel: soFieldModel,
                template: KoComponents.soCheckboxGroupTemplate
            });
            ko.components.register('so-radio-group', {
                viewModel: soFieldModel,
                template: KoComponents.soRadioGroupTemplate
            });
            ko.components.register('so-usermulti-group', {
                viewModel: soUsermultiModel,
                template: KoComponents.soUsermultiFieldTemplate
            });
            ko.components.register('so-static-field', {
                viewModel: soStaticModel,
                template: KoComponents.soStaticFieldTemplate
            });
            ko.components.register('so-static-person', {
                viewModel: soStaticModel,
                template: KoComponents.soStaticFieldTemplate.replace(/data-bind="text: modelValue/g, 'data-bind="spPerson: modelValue')
            });
            ko.components.register('so-static-date', {
                viewModel: soStaticModel,
                template: KoComponents.soStaticFieldTemplate.replace(/data-bind="text: modelValue/g, 'data-bind="spDate: modelValue')
            });
            ko.components.register('so-static-datetime', {
                viewModel: soStaticModel,
                template: KoComponents.soStaticFieldTemplate.replace(/data-bind="text: modelValue/g, 'data-bind="spDateTime: modelValue')
            });
            ko.components.register('so-static-money', {
                viewModel: soStaticModel,
                template: KoComponents.soStaticFieldTemplate.replace(/data-bind="text: modelValue/g, 'data-bind="spMoney: modelValue')
            });
            ko.components.register('so-static-number', {
                viewModel: soStaticModel,
                template: KoComponents.soStaticFieldTemplate.replace(/data-bind="text: modelValue/g, 'data-bind="spNumber: modelValue')
            });
            ko.components.register('so-static-decimal', {
                viewModel: soStaticModel,
                template: KoComponents.soStaticFieldTemplate.replace(/data-bind="text: modelValue/g, 'data-bind="spDecimal: modelValue')
            });
            ko.components.register('so-static-html', {
                viewModel: soStaticModel,
                template: KoComponents.soStaticFieldTemplate.replace(/data-bind="text: modelValue/g, 'data-bind="html: modelValue')
            });
            ko.components.register('so-attachments', {
                viewModel: function (params) {
                    var self = this;
                    if (!params) {
                        throw 'params is undefined in so-attachments';
                        return;
                    }
                    if (!params.val) {
                        throw "Parameter `val` for so-attachments is required!";
                    }
                    this.attachments = params.val;
                    this.title = params.title || 'Attachments';
                    this.id = params.id || 'fileUploader_' + uniqueId();
                    this.description = params.description;
                    // allow for static bool or ko obs
                    this.readOnly = (typeof params.readOnly == 'function') ? params.readOnly : ko.observable(!!params.readOnly || false);
                    this.deleteAttachment = function (att, event) {
                        if (!confirm('Are you sure you want to delete ' + att.Name + '? This can\'t be undone.')) {
                            return;
                        }
                        Shockout.SpApi.deleteAttachment(att, function (data, error) {
                            if (!!error) {
                                alert("Failed to delete attachment: " + error);
                                return;
                            }
                            var attachments = self.attachments;
                            attachments.remove(att);
                        });
                    };
                },
                template: '<section>' +
                    '<h4 data-bind="text: title">Attachments (<span data-bind="text: attachments().length"></span>)</h4>' +
                    '<div data-bind="attr:{id: id}"></div>' +
                    '<div data-bind="foreach: attachments">' +
                    '<div>' +
                    '<a href="" data-bind="attr: {href: __metadata.media_src}"><span class="glyphicon glyphicon-paperclip"></span> <span data-bind="text: Name"></span></a>' +
                    '<!-- ko ifnot: $parent.readOnly() -->' +
                    '<button data-bind="event: {click: $parent.deleteAttachment}" class="btn btn-sm btn-danger" title="Delete Attachment"><span class="glyphicon glyphicon-remove"></span></button>' +
                    '<!-- /ko -->' +
                    '</div>' +
                    '</div>' +
                    '<!-- ko if: description -->' +
                    '<div data-bind="text: description"></div>' +
                    '<!-- /ko -->' +
                    '</section>'
            });
            ko.components.register('so-created-modified-info', {
                viewModel: function (params) {
                    this.CreatedBy = params.createdBy;
                    this.ModifiedBy = params.modifiedBy;
                    this.profiles = ko.observableArray([
                        { header: 'Created By', profile: this.CreatedBy },
                        { header: 'Modified By', profile: this.ModifiedBy }
                    ]);
                    this.Created = params.created;
                    this.Modified = params.modified;
                    this.showUserProfiles = params.showUserProfiles;
                },
                template: '<!-- ko if: showUserProfiles() -->' +
                    '<div class="create-mod-info no-print hidden-xs">' +
                    '<!-- ko foreach: profiles -->' +
                    '<div class="user-profile-card">' +
                    '<h4 data-bind="text: header"></h4>' +
                    '<!-- ko with: profile -->' +
                    '<img data-bind="attr: {src: Picture, alt: Name}" />' +
                    '<ul>' +
                    '<li><label>Name</label><span data-bind="text: Name"></span><li>' +
                    '<li data-bind="visible: !!JobTitle"><label>Job Title</label><span data-bind="text: JobTitle"></span></li>' +
                    '<li data-bind="visible: !!Department"><label>Department</label><span data-bind="text: Department"></span></li>' +
                    '<li><label>Email</label><a data-bind="text: WorkEMail, attr: {href: (\'mailto:\' + WorkEMail)}"></a></li>' +
                    '<li data-bind="visible: !!WorkPhone"><label>Phone</label><span data-bind="text: WorkPhone"></span></li>' +
                    '<li data-bind="visible: !!Office"><label>Office</label><span data-bind="text: Office"></span></li>' +
                    '</ul>' +
                    '<!-- /ko -->' +
                    '</div>' +
                    '<!-- /ko -->' +
                    '</div>' +
                    '<!-- /ko -->' +
                    '<div class="row">' +
                    '<!-- ko with: CreatedBy -->' +
                    '<div class="col-md-3"><label>Created By</label> <a data-bind="text: Name, attr: {href: \'mailto:\' + WorkEMail}" class="email" > </a></div>' +
                    '<!-- /ko -->' +
                    '<div class="col-md-3"><label>Created</label> <span data-bind="spDateTime: Created"></span></div>' +
                    '<!-- ko with: ModifiedBy -->' +
                    '<div class="col-md-3"><label>Modified By</label> <a data-bind="text: Name, attr: {href: \'mailto:\' + WorkEMail}" class="email"></a></div>' +
                    '<!-- /ko -->' +
                    '<div class="col-md-3"><label>Modified</label> <span data-bind="spDateTime: Modified"></span></div>' +
                    '</div>'
            });
            ko.components.register('so-workflow-history', {
                viewModel: function (params) {
                    this.historyItems = (params.val || params.historyItems);
                },
                template: '<div class="row">' +
                    '<div class="col-sm-8"><strong>Description</strong></div>' +
                    '<div class="col-sm-4"><strong>Date</strong></div>' +
                    '</div>' +
                    '<!-- ko foreach: historyItems -->' +
                    '<div class="row">' +
                    '<div class="col-sm-8"><span data-bind="text: _description"></span></div>' +
                    '<div class="col-sm-4"><span data-bind="spDateTime: _dateOccurred"></span></div>' +
                    '</div>' +
                    '<!-- /ko -->'
            });
            function soStaticModel(params) {
                if (!params) {
                    throw 'params is undefined in so-static-field';
                    return;
                }
                var koObj = params.val || params.modelValue;
                if (!koObj) {
                    throw "Parameter `val` or `modelValue` for so-static-field is required!";
                }
                this.modelValue = koObj;
                this.id = params.id || koObj._koName;
                this.label = params.label || koObj._displayName;
                this.description = params.description || koObj._description;
                var labelX = parseInt(params.labelColWidth || 3); // Bootstrap label column width 1-12
                var fieldX = parseInt(params.fieldColWidth || (12 - (labelX - 0))); // Bootstrap field column width 1-12
                this.labelColWidth = 'col-sm-' + labelX;
                this.fieldColWidth = 'col-sm-' + fieldX;
            }
            ;
            function soFieldModel(params) {
                if (!params) {
                    throw 'params is undefined in soFieldModel';
                    return;
                }
                var koObj = params.val || params.modelValue;
                if (!koObj) {
                    throw "Parameter `val` or `modelValue` for so-text-field is required!";
                }
                this.modelValue = koObj;
                this.id = params.id || koObj._koName;
                this.name = params.name || koObj._koName || params.id;
                this.label = params.label = null || params.label == '' ? undefined : params.label || koObj._displayName;
                this.title = params.title;
                this.caption = params.caption;
                this.maxlength = params.maxlength || 255;
                this.placeholder = params.placeholder || params.label || koObj._displayName;
                this.description = (typeof params.description != 'undefined') ? (params.description == null ? undefined : params.description) : koObj._description;
                this.valueUpdate = params.valueUpdate;
                this.editable = !!koObj._koName; // if `_koName` is a prop of our KO var, it's a field we can update in theSharePoint list.
                this.koName = koObj._koName; // include the name of the KO var in case we need to reference it.
                this.options = params.options || koObj._options;
                this.required = (typeof params.required == 'function') ? params.required : ko.observable(!!params.required || false);
                this.inline = params.inline || false;
                var labelX = parseInt(params.labelColWidth || 3); // Bootstrap label column width 1-12
                var fieldX = parseInt(params.fieldColWidth || (12 - (labelX - 0))); // Bootstrap field column width 1-12
                this.labelColWidth = 'col-sm-' + labelX;
                this.fieldColWidth = 'col-sm-' + fieldX;
                // allow for static bool or ko obs
                this.readOnly = (typeof params.readOnly == 'function') ? params.readOnly : ko.observable(!!params.readOnly || false);
            }
            ;
            function soUsermultiModel(params) {
                if (!params) {
                    throw 'params is undefined in soFieldModel';
                    return;
                }
                var self = this;
                var koObj = params.val || params.modelValue;
                if (!koObj) {
                    throw "Parameter `val` or `modelValue` for so-text-field is required!";
                }
                this.modelValue = koObj;
                this.id = params.id || koObj._koName;
                this.name = params.name || koObj._koName || params.id;
                this.label = params.label || koObj._displayName;
                this.title = params.title;
                this.required = params.required;
                this.description = params.description || koObj._description;
                this.editable = !!koObj._koName; // if `_koName` is a prop of our KO var, it's a field we can update in theSharePoint list.
                this.koName = koObj._koName; // include the name of the KO var in case we need to reference it.
                this.person = ko.observable(null);
                // allow for static bool or ko obs
                this.readOnly = (typeof params.readOnly == 'function') ? params.readOnly : ko.observable(!!params.readOnly || false);
                // add a person to KO object People
                this.addPerson = function (model, ctrl) {
                    if (self.modelValue() == null) {
                        self.modelValue([]);
                    }
                    self.modelValue().push(self.person());
                    self.modelValue.valueHasMutated();
                    self.person(null);
                    return false;
                };
                // remove a person from KO object People
                this.removePerson = function (person, event) {
                    self.modelValue.remove(person);
                    return false;
                };
            }
            ;
        };
        ;
        //&& !!required && !readOnly
        KoComponents.hasErrorCssDiv = '<div class="form-group" data-bind="css: {\'has-error\': !!!modelValue() && !!required(), \'has-success has-feedback\': !!modelValue() && !!required()}">';
        KoComponents.requiredFeedbackSpan = '<span class="glyphicon glyphicon-ok form-control-feedback" aria-hidden="true"></span>';
        KoComponents.soStaticFieldTemplate = '<div class="form-group">' +
            '<div class="row">' +
            // field label
            '<!-- ko if: !!label -->' +
            '<div class="col-sm-3" data-bind="attr:{\'class\': labelColWidth}"><label data-bind="html: label"></label></div>' +
            '<!-- /ko -->' +
            // field
            '<div class="col-sm-9" data-bind="text: modelValue, attr:{\'class\': fieldColWidth}"></div>' +
            '</div>' +
            // description
            '<!-- ko if: description -->' +
            '<div class="so-field-description"><p data-bind="html: description"></p></div>' +
            '<!-- /ko -->' +
            '</div>';
        KoComponents.soTextFieldTemplate = KoComponents.hasErrorCssDiv +
            '<div class="row">' +
            // field label
            '<!-- ko if: !!label -->' +
            '<div class="col-sm-3" data-bind="attr:{\'class\': labelColWidth}"><label data-bind="html: label, attr: {for: id}"></label></div>' +
            '<!-- /ko -->' +
            // field
            '<div class="col-sm-9" data-bind="attr:{\'class\': fieldColWidth}">' +
            '<!-- ko if: readOnly() -->' +
            '<div data-bind="text: modelValue"></div>' +
            '<!-- /ko -->' +
            '<!-- ko ifnot: readOnly() -->' +
            '<input type="text" data-bind="value: modelValue, css: {\'so-editable\': editable}, attr: {id: id, placeholder: placeholder, title: title, required: required, maxlength: maxlength, \'ko-name\': koName }" class="form-control" />' +
            '<!-- ko if: !!required() -->' +
            KoComponents.requiredFeedbackSpan +
            '<!-- /ko -->' +
            '<!-- /ko -->' +
            '</div>' +
            '</div>' +
            // description
            '<!-- ko if: description -->' +
            '<div class="so-field-description"><p data-bind="html: description"></p></div>' +
            '<!-- /ko -->' +
            '</div>';
        //'<div data-bind="spHtmlEditor: modelValue" contenteditable="true" class="form-control content-editable"></div>'+
        //'<textarea data-bind="value: modelValue, css: {\'so-editable\': editable}, attr: {id: id, required: required, \'ko-name\': koName }" data-sp-html="" style="display:none;"></textarea>' +
        KoComponents.soHtmlFieldTemplate = KoComponents.hasErrorCssDiv +
            '<div class="row">' +
            // field label
            '<!-- ko if: !!label -->' +
            '<div class="col-sm-3" data-bind="attr:{\'class\': labelColWidth}"><label data-bind="html: label, attr: {for: id}"></label></div>' +
            '<!-- /ko -->' +
            // field
            '<div class="col-sm-9" data-bind="attr:{\'class\': fieldColWidth}">' +
            '<!-- ko if: readOnly() -->' +
            '<div data-bind="html: modelValue"></div>' +
            '<!-- /ko -->' +
            '<!-- ko ifnot: readOnly() -->' +
            '<div data-bind="spHtmlEditor: modelValue" contenteditable="true" class="form-control content-editable"></div>' +
            '<textarea data-bind="value: modelValue, css: {\'so-editable\': editable}, attr: {id: id, required: required, \'ko-name\': koName }" data-sp-html="" style="display:none;"></textarea>' +
            '<!-- ko if: !!required() -->' +
            KoComponents.requiredFeedbackSpan +
            '<!-- /ko -->' +
            '<!-- /ko -->' +
            '</div>' +
            '</div>' +
            // description
            '<!-- ko if: description -->' +
            '<div class="so-field-description"><p data-bind="html: description"></p></div>' +
            '<!-- /ko -->' +
            '</div>';
        KoComponents.soCheckboxFieldTemplate = '<div class="form-group">' +
            '<div class="row">' +
            // field label
            '<!-- ko if: !!label -->' +
            '<div class="col-sm-3" data-bind="attr:{\'class\': labelColWidth}"><label data-bind="html: label"></label></div>' +
            '<!-- /ko -->' +
            // field
            '<div class="col-sm-9" data-bind="attr:{\'class\': fieldColWidth}">' +
            '<!-- ko if: readOnly() -->' +
            '<div data-bind="text: !!modelValue() ? \'Yes\' : \'No\'"></div>' +
            '<!-- /ko -->' +
            '<!-- ko ifnot: readOnly() -->' +
            '<label class="checkbox">' +
            '<input type="checkbox" data-bind="checked: modelValue, css: {\'so-editable\': editable}, attr: {id: id, \'ko-name\': koName}, valueUpdate: valueUpdate" />' +
            '<span data-bind="html: label"></span>' +
            '</label>' +
            '<!-- /ko -->' +
            '</div>' +
            '</div>' +
            // description
            '<!-- ko if: description -->' +
            '<div class="so-field-description"><p data-bind="html: description"></p></div>' +
            '<!-- /ko -->' +
            '</div>';
        KoComponents.soSelectFieldTemplate = KoComponents.hasErrorCssDiv +
            '<div class="row">' +
            // field label
            '<!-- ko if: !!label -->' +
            '<div class="col-sm-3" data-bind="attr:{\'class\': labelColWidth}"><label data-bind="html: label, attr: {for: id}"></label></div>' +
            '<!-- /ko -->' +
            // field
            '<div class="col-sm-9" data-bind="attr:{\'class\': fieldColWidth}">' +
            '<!-- ko if: readOnly() -->' +
            '<div data-bind="text: modelValue"></div>' +
            '<!-- /ko -->' +
            '<!-- ko ifnot: readOnly() -->' +
            '<select data-bind="value: modelValue, options: options, optionsCaption: caption, css: {\'so-editable\': editable}, attr: {id: id, title: title, required: required, \'ko-name\': koName}" class="form-control"></select>' +
            '<!-- ko if: !!required() -->' +
            KoComponents.requiredFeedbackSpan +
            '<!-- /ko -->' +
            '<!-- /ko -->' +
            '</div>' +
            '</div>' +
            // description
            '<!-- ko if: description -->' +
            '<div class="so-field-description"><p data-bind="html: description"></p></div>' +
            '<!-- /ko -->' +
            '</div>';
        KoComponents.soCheckboxGroupTemplate = '<div class="form-group">' +
            // description
            '<!-- ko if: description -->' +
            '<div class="so-field-description"><p data-bind="html: description"></p></div>' +
            '<!-- /ko -->' +
            '<div class="row">' +
            // field label
            '<!-- ko if: !!label -->' +
            '<div><label data-bind="html: label"></label></div>' +
            '<!-- /ko -->' +
            '<div>' +
            // show static elements if inline
            '<!-- ko if: readOnly() -->' +
            // show static unordered list if !inline
            '<!-- ko ifnot: inline -->' +
            '<ul class="list-group">' +
            '<!-- ko foreach: modelValue -->' +
            '<li data-bind="text: $data" class="list-group-item"></li>' +
            '<!-- /ko -->' +
            '<!-- ko if: modelValue().length == 0 -->' +
            '<li class="list-group-item">--None--</li>' +
            '<!-- /ko -->' +
            '</ul>' +
            '<!-- /ko -->' +
            // show static inline elements if inline
            '<!-- ko if: inline -->' +
            '<!-- ko foreach: modelValue -->' +
            '<span data-bind="text: $data"></span>' +
            '<!-- ko if: $index() < $parent.modelValue().length-1 -->,&nbsp;<!-- /ko -->' +
            '<!-- /ko -->' +
            '<!-- ko if: modelValue().length == 0 -->' +
            '<span>--None--</span>' +
            '<!-- /ko -->' +
            '<!-- /ko -->' +
            '<!-- /ko -->' +
            // show input field if not readOnly
            '<!-- ko ifnot: readOnly() -->' +
            '<input type="hidden" data-bind="value: modelValue, attr:{required: !!required}" /><p data-bind="visible: !!required" class="req">(Required)</p>' +
            '<!-- ko foreach: options -->' +
            '<label data-bind="css:{\'checkbox\': !$parent.inline, \'checkbox-inline\': $parent.inline}">' +
            '<input type="checkbox" data-bind="checked: $parent.modelValue, css: {\'so-editable\': $parent.editable}, attr: {\'ko-name\': $parent.koName, \'value\': $data}" />' +
            '<span data-bind="text: $data"></span>' +
            '</label>' +
            '<!-- /ko -->' +
            '<!-- /ko -->' +
            '</div>' +
            '</div>';
        KoComponents.soRadioGroupTemplate = '<div class="form-group">' +
            // description
            '<!-- ko if: description -->' +
            '<div class="so-field-description"><p data-bind="html: description"></p></div>' +
            '<!-- /ko -->' +
            '<div class="row">' +
            // field label
            '<!-- ko if: !!label -->' +
            '<div class="col-sm-3" data-bind="attr:{\'class\': labelColWidth}"><label data-bind="html: label"></label></div>' +
            '<!-- /ko -->' +
            '<div class="col-sm-9" data-bind="attr:{\'class\': fieldColWidth}">' +
            // show static field if readOnly
            '<!-- ko if: readOnly() -->' +
            '<div data-bind="text: modelValue"></div>' +
            '<!-- /ko -->' +
            // show input field if not readOnly
            '<!-- ko ifnot: readOnly() -->' +
            '<!-- ko foreach: options -->' +
            '<label data-bind="css:{\'radio\': !$parent.inline, \'radio-inline\': $parent.inline}">' +
            '<input type="radio" data-bind="checked: $parent.modelValue, attr:{value: $data, name: $parent.name, \'ko-name\': $parent.koName}, css:{\'so-editable\': $parent.editable}" />' +
            '<span data-bind="text: $data"></span>' +
            '</label>' +
            '<!-- /ko -->' +
            '<!-- /ko -->' +
            '</div>' +
            '</div>' +
            '</div>';
        KoComponents.soUsermultiFieldTemplate = '<div class="form-group">' +
            // show input field if not readOnly
            '<!-- ko ifnot: readOnly() -->' +
            '<input type="hidden" data-bind="value: modelValue, css: {\'so-editable\': editable}, attr: {id: id, \'ko-name\': koName, required: required}" />' +
            '<div class="row">' +
            '<div class="col-md-6 col-xs-6">' +
            '<input type="text" data-bind="spPerson: person, attr: {placeholder: placeholder}" />' +
            '<button class="btn btn-success" data-bind="click: addPerson, attr: {\'disabled\': person() == null}"><span>Add</span></button>' +
            '</div>' +
            '<!-- ko if: required && modelValue() == null && !readOnly -->' +
            '<div class="col-md-6 col-xs-6">' +
            '<p class="error">This field is required.</p>' +
            '</div>' +
            '<!-- /ko -->' +
            '</div>' +
            '<!-- /ko -->' +
            '<!-- ko foreach: modelValue -->' +
            '<div class="row">' +
            '<div class="col-md-10 col-xs-10" data-bind="spPerson: $data"></div>' +
            '<!-- ko ifnot: readOnly() -->' +
            '<div class="col-md-2 col-xs-2">' +
            '<button class="btn btn-xs btn-danger" data-bind="click: $parent.removePerson"><span class="glyphicon glyphicon-trash"></span></button>' +
            '</div>' +
            '<!-- /ko -->' +
            '</div>' +
            '<!-- /ko -->' +
            '<!-- ko if: description -->' +
            '<div class="so-field-description"><p data-bind="html: description"></p></div>' +
            '<!-- /ko -->' +
            '</div>';
        KoComponents.soCreatedModifiedTemplate = '<!-- ko if: !!CreatedBy && CreatedBy() != null --><section><so-created-modified-info params="created: Created, createdBy: CreatedBy, modified: Modified, modifiedBy: ModifiedBy, showUserProfiles: showUserProfiles"></so-created-modified-info></section><!-- /ko -->';
        KoComponents.soWorkflowHistoryTemplate = '<!-- ko if: !!Id() --><section id="workflowHistory" class="nav-section"><h4>Workflow History</h4><so-workflow-history params="val: historyItems"></so-workflow-history></section><!-- /ko -->';
        return KoComponents;
    })();
    Shockout.KoComponents = KoComponents;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var SpApi = (function () {
        function SpApi() {
        }
        /**
        * Search the User Information list.
        * @param term: string
        * @param callback: Function
        * @param take?: number = 10
        * @return void
        */
        SpApi.peopleSearch = function (term, callback, take) {
            if (take === void 0) { take = 10; }
            var filter = "startswith(Name,'{0}')".replace(/\{0\}/g, term);
            var select = null;
            var orderby = "Name";
            var top = 10;
            var cache = true;
            SpApi.getListItems('UserInformationList', fn, '/', filter, select, orderby, top, cache);
            function fn(data, error) {
                if (!!error) {
                    callback(null, error);
                    return;
                }
                callback(data, error);
            }
            ;
        };
        /**
        * Get a person by their ID from the User Information list.
        * @param id: number
        * @param callback: Function
        * @return void
        */
        SpApi.getPersonById = function (id, callback) {
            SpApi.getListItem('UserInformationList', id, function (data, error) {
                if (!!error) {
                    callback(null, error);
                }
                callback(data);
            }, '/', true);
        };
        /**
        * General REST request method.
        * @param url: string
        * @param callback: Function
        * @param cache?: boolean = false
        * @param type?: string = 'GET'
        * @return void
        */
        SpApi.executeRestRequest = function (url, callback, cache, type) {
            if (cache === void 0) { cache = false; }
            if (type === void 0) { type = 'GET'; }
            var $jqXhr = $.ajax({
                url: url,
                type: type,
                cache: cache,
                dataType: 'json',
                contentType: 'application/json; charset=utf-8',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                }
            });
            $jqXhr.done(function (data, status, jqXhr) {
                callback(data);
            });
            $jqXhr.fail(function (jqXhr, status, error) {
                if (!!status && status == '404') {
                    var msg = status + ". The data may have been deleted by another user.";
                }
                else {
                    msg = status + ' ' + error;
                }
                callback(null, msg);
            });
        };
        /**
        * Get list item via REST services.
        * @param uri: string
        * @param done: JQueryPromiseCallback<any>
        * @param fail?: JQueryPromiseCallback<any> = undefined
        * @param always?: JQueryPromiseCallback<any> = undefined
        * @return void
        */
        SpApi.getListItem = function (listName, itemId, callback, siteUrl, cache, expand) {
            if (siteUrl === void 0) { siteUrl = '/'; }
            if (cache === void 0) { cache = false; }
            if (expand === void 0) { expand = null; }
            siteUrl = Shockout.Utils.formatSubsiteUrl(siteUrl);
            var url = siteUrl + '_vti_bin/listdata.svc/' + Shockout.Utils.toCamelCase(listName) + '(' + itemId + ')?$expand=CreatedBy,ModifiedBy' + (!!expand ? ',' + expand : '');
            SpApi.executeRestRequest(url, fn, cache, 'GET');
            function fn(data, error) {
                if (!!error) {
                    callback(data, error);
                    return;
                }
                if (!!data) {
                    if (data.d) {
                        callback(data.d);
                    }
                    else {
                        callback(data);
                    }
                }
            }
            ;
        };
        /**
        * Get list item via REST services.
        * @param uri: string
        * @param done: JQueryPromiseCallback<any>
        * @param fail?: JQueryPromiseCallback<any> = undefined
        * @param always?: JQueryPromiseCallback<any> = undefined
        * @return void
        */
        SpApi.getListItems = function (listName, callback, siteUrl, filter, select, orderby, top, cache) {
            if (siteUrl === void 0) { siteUrl = '/'; }
            if (filter === void 0) { filter = null; }
            if (select === void 0) { select = null; }
            if (orderby === void 0) { orderby = null; }
            if (top === void 0) { top = 10; }
            if (cache === void 0) { cache = false; }
            siteUrl = Shockout.Utils.formatSubsiteUrl(siteUrl);
            var url = [siteUrl + '_vti_bin/listdata.svc/' + Shockout.Utils.toCamelCase(listName)];
            if (!!filter) {
                url.push('$filter=' + filter);
            }
            if (!!select) {
                url.push('$select=' + select);
            }
            if (!!orderby) {
                url.push('$orderby=' + orderby);
            }
            url.push('$top=' + top);
            SpApi.executeRestRequest(url.join('&').replace(/\&/, '\?'), fn, cache, 'GET');
            function fn(data, error) {
                var data = !!data && 'd' in data ? data.d : data;
                var results = null;
                if (!!data) {
                    results = 'results' in data ? data.results : data;
                }
                callback(results, error);
            }
            ;
        };
        /**
        * Insert a list item with REST service.
        * @param item: ISpItem
        * @param callback: Function
        * @param data?: Object<any> = undefined
        * @return void
        */
        SpApi.insertListItem = function (url, callback, data) {
            if (data === void 0) { data = undefined; }
            var $jqXhr = $.ajax({
                url: url,
                type: 'POST',
                processData: false,
                contentType: 'application/json',
                data: !!data ? JSON.stringify(data) : null,
                headers: {
                    'Accept': 'application/json;odata=verbose'
                }
            });
            $jqXhr.done(function (data, status, jqXhr) {
                callback(data);
            });
            $jqXhr.fail(function (jqXhr, status, error) {
                callback(null, status + ': ' + error);
            });
        };
        /**
        * Update a list item with REST service.
        * @param item: ISpItem
        * @param callback: Function
        * @param data?: Object<any> = undefined
        * @return void
        */
        SpApi.updateListItem = function (item, callback, data) {
            if (data === void 0) { data = undefined; }
            var $jqXhr = $.ajax({
                url: item.__metadata.uri,
                type: 'POST',
                processData: false,
                contentType: 'application/json',
                data: data ? JSON.stringify(data) : null,
                beforeSend: function (xhr) {
                    xhr.setRequestHeader('X-HTTP-Method', 'MERGE');
                    xhr.setRequestHeader('If-Match', item.__metadata.etag);
                }
            });
            $jqXhr.done(function (data, status, jqXhr) {
                callback(data);
            });
            $jqXhr.fail(function (jqXhr, status, error) {
                callback(null, status + ': ' + error);
            });
        };
        /**
        * Delete the list item with REST service.
        * @param model: IViewModel
        * @param callback?: Function = undefined
        * @return void
        */
        SpApi.deleteListItem = function (item, callback) {
            var $jqXhr = $.ajax({
                url: item.__metadata.uri,
                type: 'POST',
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'X-Http-Method': 'DELETE',
                    'If-Match': item.__metadata.etag
                }
            });
            $jqXhr.done(function (data, status, jqXhr) {
                callback(data);
            });
            $jqXhr.fail(function (jqXhr, status, error) {
                callback(null, error);
            });
        };
        /**
        * Delete an attachment with REST service.
        * @param att: ISpAttachment
        * @param callback: Function
        * @return void
        */
        SpApi.deleteAttachment = function (att, callback) {
            var $jqXhr = $.ajax({
                url: att.__metadata.uri,
                type: 'POST',
                dataType: 'json',
                contentType: "application/json",
                headers: {
                    'Accept': 'application/json;odata=verbose',
                    'X-HTTP-Method': 'DELETE'
                }
            });
            $jqXhr.done(function (data, status, jqXhr) {
                callback(data);
            });
            $jqXhr.fail(function (jqXhr, status, error) {
                callback(null, status + ': ' + error);
            });
        };
        return SpApi;
    })();
    Shockout.SpApi = SpApi;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var SpApi15 = (function () {
        function SpApi15() {
        }
        /**
        * Get the current user.
        * @param callback: Function
        * @param expandGroups?: boolean = false
        * @return void
        */
        SpApi15.getCurrentUser = function (callback, expandGroups) {
            if (expandGroups === void 0) { expandGroups = false; }
            var $jqXhr = $.ajax({
                url: '/_api/Web/CurrentUser' + (expandGroups ? '?$expand=Groups' : ''),
                type: 'GET',
                cache: true,
                dataType: 'json',
                contentType: 'application/json; charset=utf-8',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                }
            });
            $jqXhr.done(function (data, status, jqXhr) {
                var user = data.d;
                var currentUser = {
                    account: user.LoginName,
                    department: null,
                    email: user.Email,
                    groups: [],
                    id: user.Id,
                    jobtitle: null,
                    login: user.LoginName,
                    title: user.Title
                };
                if (expandGroups) {
                    var groups = data.d.Groups;
                    $(groups.results).each(function (i, group) {
                        currentUser.groups.push({ id: group.Id, name: group.Title });
                    });
                }
                callback(currentUser);
            });
            $jqXhr.fail(function (jqXhr, status, error) {
                callback(null, jqXhr.status); // '404'
            });
        };
        /**
        * Get user's groups.
        * @param iserId: number
        * @param callback: Function
        * @return void
        */
        SpApi15.getUsersGroups = function (userId, callback) {
            var $jqXhr = $.ajax({
                url: '/_api/Web/GetUserById(' + userId + ')/Groups',
                type: 'GET',
                cache: true,
                dataType: 'json',
                contentType: 'application/json; charset=utf-8',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                }
            });
            $jqXhr.done(function (data, status, jqXhr) {
                var groups = [];
                for (var i = 0; i < data.d.results.length; i++) {
                    var group = data.d.results[i];
                    groups.push({ id: group.Id, name: group.Title });
                }
                callback(groups);
            });
            $jqXhr.fail(function (jqXhr, status, error) {
                callback(null, error);
            });
        };
        return SpApi15;
    })();
    Shockout.SpApi15 = SpApi15;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var SpSoap = (function () {
        function SpSoap() {
        }
        /**
        * Get the current user via SOAP.
        * @param callback: Function
        * @return void
        */
        SpSoap.getCurrentUser = function (callback) {
            var user = {};
            var query = '<Query><Where><Eq><FieldRef Name="ID" /><Value Type="Counter"><UserID /></Value></Eq></Where></Query>';
            var viewFields = '<ViewFields><FieldRef Name="ID" /><FieldRef Name="Name" /><FieldRef Name="EMail" /><FieldRef Name="Department" /><FieldRef Name="JobTitle" /><FieldRef Name="UserName" /><FieldRef Name="Office" /></ViewFields>';
            SpSoap.getListItems('', 'User Information List', viewFields, query, function (xmlDoc, status, jqXhr) {
                $(xmlDoc).find('*').filter(function () {
                    return this.nodeName == 'z:row';
                }).each(function (i, node) {
                    user.id = parseInt($(node).attr('ows_ID'));
                    user.title = $(node).attr('ows_Title');
                    user.login = $(node).attr('ows_Name');
                    user.email = $(node).attr('ows_EMail');
                    user.jobtitle = $(node).attr('ows_JobTitle');
                    user.department = $(node).attr('ows_Department');
                    user.account = user.id + ';#' + user.title;
                    user.groups = [];
                });
                callback(user);
            });
            /*
            // Returns
            <z:row xmlns:z="#RowsetSchema"
                ows_ID="1"
                ows_Name="<DOMAIN\login>"
                ows_EMail="<email>"
                ows_JobTitle="<job title>"
                ows_UserName="<username>"
                ows_Office="<office>"
                ows__ModerationStatus="0"
                ows__Level="1"
                ows_Title="<Fullname>"
                ows_Dapartment="<Department>"
                ows_UniqueId="1;#{2AFFA9A1-87D4-44A7-9D4F-618BCBD990D7}"
                ows_owshiddenversion="306"
                ows_FSObjType="1;#0"/>
            */
        };
        /**
        * Get the a user's groups via SOAP.
        * @param loginName: string (DOMAIN\loginName)
        * @param callback: Function
        * @return void
        */
        SpSoap.getUsersGroups = function (loginName, callback) {
            var packet = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                '<soap:Body>' +
                '<GetGroupCollectionFromUser xmlns="http://schemas.microsoft.com/sharepoint/soap/directory/">' +
                '<userLoginName>' + loginName + '</userLoginName>' +
                '</GetGroupCollectionFromUser>' +
                '</soap:Body>' +
                '</soap:Envelope>';
            var $jqXhr = $.ajax({
                url: '/_vti_bin/usergroup.asmx',
                type: 'POST',
                dataType: 'xml',
                data: packet,
                contentType: 'text/xml; charset="utf-8"'
            });
            $jqXhr.done(cb);
            $jqXhr.fail(cb);
            function cb(xmlDoc, status, jqXhr) {
                var $errorText = $(xmlDoc).find('errorstring');
                // catch and handle returned error
                if (!!$errorText && $errorText.text() != "") {
                    callback(null, $errorText.text());
                    return;
                }
                var groups = [];
                $(xmlDoc).find("Group").each(function (i, el) {
                    groups.push({
                        id: parseInt($(el).attr("ID")),
                        name: $(el).attr("Name")
                    });
                });
                callback(groups);
            }
        };
        /**
        * Get list items via SOAP.
        * @param siteUrl: string
        * @param listName: string
        * @param viewFields: string (XML)
        * @param query?: string (XML)
        * @param callback?: Function
        * @param rowLimit?: number = 25
        * @return void
        */
        SpSoap.getListItems = function (siteUrl, listName, viewFields, query, callback, rowLimit) {
            if (rowLimit === void 0) { rowLimit = 25; }
            siteUrl = Shockout.Utils.formatSubsiteUrl(siteUrl);
            if (!!!listName) {
                Shockout.Utils.logError("Parameter `listName` is null or undefined in method SpSoap.getListItems()", Shockout.SPForm.errorLogListName, Shockout.SPForm.errorLogSiteUrl);
            }
            var packet = '<?xml version="1.0" encoding="utf-8"?>' +
                '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                '<soap:Body>' +
                '<GetListItems xmlns="http://schemas.microsoft.com/sharepoint/soap/">' +
                '<listName>' + listName + '</listName>' +
                '<query>' + query + '</query>' +
                '<viewFields>' + viewFields + '</viewFields>' +
                '<rowLimit>' + rowLimit + '</rowLimit>' +
                '</GetListItems>' +
                '</soap:Body>' +
                '</soap:Envelope>';
            var $jqXhr = $.ajax({
                url: siteUrl + '_vti_bin/lists.asmx',
                type: 'POST',
                dataType: 'xml',
                data: packet,
                headers: {
                    "SOAPAction": "http://schemas.microsoft.com/sharepoint/soap/GetListItems",
                    "Content-Type": "text/xml; charset=utf-8"
                }
            });
            $jqXhr.done(function (xmlDoc, status, error) {
                callback(xmlDoc);
            });
            $jqXhr.fail(function (jqXhr, status, error) {
                callback(null, status + ': ' + error);
            });
        };
        /**
        * Get list definition
        * @param siteUrl: string
        * @param listName: string
        * @param callback: Function
        * @return void
        */
        SpSoap.getList = function (siteUrl, listName, callback) {
            siteUrl = Shockout.Utils.formatSubsiteUrl(siteUrl);
            var packet = '<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap:Body><GetList xmlns="http://schemas.microsoft.com/sharepoint/soap/"><listName>{0}</listName></GetList></soap:Body></soap:Envelope>'
                .replace('{0}', listName);
            var $jqXhr = $.ajax({
                url: siteUrl + '_vti_bin/lists.asmx',
                type: 'POST',
                cache: false,
                dataType: "xml",
                data: packet,
                headers: {
                    "SOAPAction": "http://schemas.microsoft.com/sharepoint/soap/GetList",
                    "Content-Type": "text/xml; charset=utf-8"
                }
            });
            $jqXhr.done(function (xmlDoc, status, jqXhr) {
                callback(xmlDoc);
            });
            $jqXhr.fail(function (jqXhr, status, error) {
                callback(null, status + ': ' + error);
            });
        };
        /**
        * Check in file.
        * @param pageUrl: string
        * @param checkinType: string
        * @param callback: Function
        * @param comment?: string = ''
        * @return void
        */
        SpSoap.checkInFile = function (pageUrl, checkinType, callback, comment) {
            if (comment === void 0) { comment = ''; }
            var action = 'http://schemas.microsoft.com/sharepoint/soap/CheckInFile';
            var params = [pageUrl, comment, checkinType];
            var packet = '<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap:Body><CheckInFile xmlns="http://schemas.microsoft.com/sharepoint/soap/"><pageUrl>{0}</pageUrl><comment>{1}</comment><CheckinType>{2}</CheckinType></CheckInFile></soap:Body></soap:Envelope>';
            return this.executeSoapRequest(action, packet, params, null, callback);
        };
        /**
        * Check out file.
        * @param pageUrl: string
        * @param checkoutToLocal: string
        * @param lastmodified: string
        * @param callback: Function
        * @return void
        */
        SpSoap.checkOutFile = function (pageUrl, checkoutToLocal, lastmodified, callback) {
            var action = 'http://schemas.microsoft.com/sharepoint/soap/CheckOutFile';
            var params = [pageUrl, checkoutToLocal, lastmodified];
            var packet = '<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap:Body><CheckOutFile xmlns="http://schemas.microsoft.com/sharepoint/soap/"><pageUrl>{0}</pageUrl><checkoutToLocal>{1}</checkoutToLocal><lastmodified>{2}</lastmodified></CheckOutFile></soap:Body></soap:Envelope>';
            return this.executeSoapRequest(action, packet, params, null, callback);
        };
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
        SpSoap.executeSoapRequest = function (action, packet, params, siteUrl, callback, service) {
            if (siteUrl === void 0) { siteUrl = '/'; }
            if (callback === void 0) { callback = undefined; }
            if (service === void 0) { service = 'lists.asmx'; }
            siteUrl = Shockout.Utils.formatSubsiteUrl(siteUrl);
            try {
                var serviceUrl = siteUrl + '_vti_bin/' + service;
                if (params != null) {
                    for (var i = 0; i < params.length; i++) {
                        packet = packet.replace('{' + i + '}', (params[i] == null ? '' : params[i]));
                    }
                }
                var $jqXhr = $.ajax({
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
                    $jqXhr.done(callback);
                }
                $jqXhr.fail(function (jqXhr, status, error) {
                    var msg = 'Error in SpSoap.executeSoapRequest. ' + status + ': ' + error + ' ';
                    Shockout.Utils.logError(msg, Shockout.SPForm.errorLogListName);
                    console.warn(msg);
                });
            }
            catch (e) {
                Shockout.Utils.logError('Error in SpSoap.executeSoapRequest.', JSON.stringify(e), Shockout.SPForm.errorLogListName);
                console.warn(e);
            }
        };
        /**
        * Update list item via SOAP services.
        * @param listName: string
        * @param fields: Array<Array<any>>
        * @param isNew?: boolean = true
        * param callback?: Function = undefined
        * @param self: SPForm = undefined
        * @return void
        */
        SpSoap.updateListItem = function (itemId, listName, fields, isNew, siteUrl, callback) {
            if (isNew === void 0) { isNew = true; }
            if (siteUrl === void 0) { siteUrl = '/'; }
            if (callback === void 0) { callback = undefined; }
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
            var command = isNew ? "New" : "Update";
            var params = [listName];
            var soapEnvelope = "<Batch OnError='Continue'><Method ID='1' Cmd='" + command + "'>";
            var itemArray = fields;
            for (var i = 0; i < fields.length; i++) {
                soapEnvelope += "<Field Name='" + fields[i][0] + "'>" + Shockout.Utils.escapeColumnValue(fields[i][1]) + "</Field>";
            }
            if (command !== "New") {
                soapEnvelope += "<Field Name='ID'>" + itemId + "</Field>";
            }
            soapEnvelope += "</Method></Batch>";
            params.push(soapEnvelope);
            SpSoap.executeSoapRequest(action, packet, params, siteUrl, callback);
        };
        /**
        * People search.
        * @param term: string
        * @param callback: Function
        * @param maxResults?: number = 10
        * @param principalType?: string = 'User'
        */
        SpSoap.searchPrincipals = function (term, callback, maxResults, principalType) {
            if (maxResults === void 0) { maxResults = 10; }
            if (principalType === void 0) { principalType = 'User'; }
            var action = 'http://schemas.microsoft.com/sharepoint/soap/SearchPrincipals';
            var params = [term, maxResults, principalType];
            var packet = '<?xml version="1.0" encoding="utf-8"?>' +
                '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                '<soap:Body>' +
                '<SearchPrincipals xmlns="http://schemas.microsoft.com/sharepoint/soap/">' +
                '<searchText>{0}</searchText>' +
                '<maxResults>{1}</maxResults>' +
                '<principalType>{2}</principalType>' +
                '</SearchPrincipals>' +
                '</soap:Body>' +
                '</soap:Envelope>';
            SpSoap.executeSoapRequest(action, packet, params, '/', cb, 'People.asmx');
            function cb(xmlDoc, status, jqXhr) {
                var results = [];
                $(xmlDoc).find('PrincipalInfo').each(function (i, n) {
                    results.push({
                        AccountName: $('AccountName', n).text(),
                        UserInfoID: parseInt($('UserInfoID', n).text()),
                        DisplayName: $('DisplayName', n).text(),
                        Email: $('Email', n).text(),
                        Title: $('Title', n).text(),
                        IsResolved: $('IsResolved', n).text() == 'true' ? !0 : !1,
                        PrincipalType: $('PrincipalType', n).text()
                    });
                });
                callback(results);
            }
        };
        return SpSoap;
    })();
    Shockout.SpSoap = SpSoap;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var HistoryItem = (function () {
        function HistoryItem(d, date) {
            this._description = d || null;
            this._dateOccurred = date || null;
        }
        return HistoryItem;
    })();
    Shockout.HistoryItem = HistoryItem;
    // recreate the SP REST object for an attachment
    var SpAttachment = (function () {
        function SpAttachment(rootUrl, siteUrl, listName, itemId, fileName) {
            var entitySet = listName.replace(/\s/g, '');
            siteUrl = Shockout.Utils.formatSubsiteUrl(siteUrl);
            var uri = rootUrl + siteUrl + "_vti_bin/listdata.svc/Attachments(EntitySet='{0}',ItemId={1},Name='{2}')";
            uri = uri.replace(/\{0\}/, entitySet).replace(/\{1\}/, itemId + '').replace(/\{2\}/, fileName);
            this.__metadata = {
                uri: uri,
                content_type: "application/octetstream",
                edit_media: uri + "/$value",
                media_etag: null,
                media_src: rootUrl + siteUrl + "/Lists/" + listName + "/Attachments/" + itemId + "/" + fileName,
                type: "Microsoft.SharePoint.DataService.AttachmentsItem"
            };
            this.EntitySet = entitySet;
            this.ItemId = itemId;
            this.Name = fileName;
        }
        return SpAttachment;
    })();
    Shockout.SpAttachment = SpAttachment;
    var SpItem = (function () {
        function SpItem() {
        }
        return SpItem;
    })();
    Shockout.SpItem = SpItem;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var Templates = (function () {
        function Templates() {
        }
        Templates.getFileUploadTemplate = function () {
            return Templates.fileuploadTemplate;
        };
        Templates.getFormAction = function () {
            var div = document.createElement('div');
            div.className = 'form-action no-print';
            div.innerHTML = Templates.actionTemplate;
            return div;
        };
        Templates.getAttachmentsTemplate = function (fileuploaderId) {
            var section = document.createElement('section');
            section.innerHTML = Templates.attachmentsTemplate.replace(/\{0\}/, fileuploaderId);
            return section;
        };
        Templates.attachmentsTemplate = '<h4>Attachments <span data-bind="text: attachments().length" class="badge"></span></h4>' +
            '<div id="{0}"></div>' +
            '<!-- ko foreach: attachments -->' +
            '<div class="so-attachment">' +
            '<a href="" data-bind="attr: {href: __metadata.media_src}"><span class="glyphicon glyphicon-paperclip"></span> <span data-bind="text: Name"></span></a>' +
            '&nbsp;&nbsp;<button data-bind="event: {click: $root.deleteAttachment}" class="btn btn-sm btn-danger" title="Delete Attachment"><span class="glyphicon glyphicon-trash"></span></button>' +
            '</div>' +
            '<!-- /ko -->';
        Templates.fileuploadTemplate = '<div class="qq-uploader" data-author-only>' +
            '<div class="qq-upload-drop-area"><span>Drop files here to upload</span></div>' +
            '<div class="btn btn-primary qq-upload-button"><span class="glyphicon glyphicon-paperclip"></span> Attach File</div>' +
            '<ul class="qq-upload-list"></ul>' +
            '</div>';
        Templates.actionTemplate = '<label>Logged in as:</label><span data-bind="text: currentUser().title" class="current-user"></span>' +
            '<button class="btn btn-default cancel" data-bind="event: { click: cancel }" title="Close"><span class="glyphicon glyphicon-remove"></span><span class="hidden-xs">Close</span></button>' +
            '<!-- ko if: allowPrint() -->' +
            '<button class="btn btn-primary print" data-bind="visible: Id() != null, event: {click: print}" title="Print"><span class="glyphicon glyphicon-print"></span><span class="hidden-xs">Print</span></button>' +
            '<!-- /ko -->' +
            '<!-- ko if: allowDelete() -->' +
            '<button class="btn btn-warning delete" data-bind="visible: Id() != null, event: {click: deleteItem}" title="Delete"><span class="glyphicon glyphicon-remove"></span><span class="hidden-xs">Delete</span></button>' +
            '<!-- /ko -->' +
            '<!-- ko if: allowSave() -->' +
            '<button class="btn btn-success save" data-bind="event: { click: save }" title="Save your work."><span class="glyphicon glyphicon-floppy-disk"></span><span class="hidden-xs">Save</span></button>' +
            '<!-- /ko -->' +
            '<button class="btn btn-danger submit" data-bind="event: { click: submit }" title="Submit for routing."><span class="glyphicon glyphicon-floppy-open"></span><span class="hidden-xs">Submit</span></button>';
        return Templates;
    })();
    Shockout.Templates = Templates;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var Utils = (function () {
        function Utils() {
        }
        /**
        * Returns the index of a value in an array. Returns -1 if not found. Use for IE8 browser compatibility.
        * @param a: Array<any>
        * @param value: any
        * @return number
        */
        Utils.indexOf = function (a, value) {
            // use the native Array.indexOf method if exists
            if (!!Array.prototype.indexOf) {
                return Array.prototype.indexOf.apply(a, [value]);
            }
            for (var i = 0; i < a.length; i++) {
                if (a[i] === value) {
                    return i;
                }
            }
            return -1;
        };
        /**
        * Ensure site url is or ends with '/'
        * @param url: string
        * @return string
        */
        Utils.formatSubsiteUrl = function (url) {
            return !!!url ? '/' : !/\/$/.test(url) ? url + '/' : url;
        };
        /**
        * Convert a name to REST camel case format
        * @param str: string
        * @return string
        */
        Utils.toCamelCase = function (str) {
            return str.toString()
                .replace(/\s*\b\w/g, function (x) {
                return (x[1] || x[0]).toUpperCase();
            }).replace(/\s/g, '')
                .replace(/\'s/, 'S')
                .replace(/[^A-Za-z0-9\s]/g, '');
        };
        /**
        * Parse a form ID from window.location.hash
        * @return number
        */
        Utils.getIdFromHash = function () {
            // example: parse ID from a URI `http://<mysite>/Forms/form.aspx/#/id/1`
            var rxHash = /\/id\/\d+/i;
            var exec = rxHash.exec(window.location.hash);
            var id = !!exec ? exec[0].replace(/\D/g, '') : null;
            return /\d/.test(id) ? parseInt(id) : null;
        };
        /**
        * Set location.hash to form ID `#/id/<ID>`.
        * @return void
        */
        Utils.setIdHash = function (id) {
            window.location.hash = '#/id/' + id;
        };
        /**
        * Escape column values
        * http://dracoblue.net/dev/encodedecode-special-xml-characters-in-javascript/155/
        */
        Utils.escapeColumnValue = function (s) {
            if (typeof s === "string") {
                return s.replace(/&(?![a-zA-Z]{1,8};)/g, "&amp;");
            }
            else {
                return s;
            }
        };
        Utils.getParent = function (o, num) {
            if (num === void 0) { num = 1; }
            for (var i = 0; i < num; i++) {
                if (!!!o) {
                    continue;
                }
                o = o.parentNode;
            }
            return o;
        };
        Utils.getPrevKOComment = function (o) {
            do {
                o = o.previousSibling;
            } while (o && o.nodeType != 8 && !/^\s*ko/.test(o.textContent)); // a KO comment is node type 8 and starts with 'ko'
            return o;
        };
        Utils.getKoComments = function (parent) {
            var koNames = [];
            parent = parent || $('body');
            $(parent).contents().filter(function (i, e) {
                return e.nodeType == 8 && /^\s*ko/.test(e.nodeValue);
            }).each(function (i, e) {
                koNames.push(e.nodeValue.replace(/\s*ko\s*foreach\s*:\s*(\$root\.|)/, '').replace(/\s/g, ''));
            });
            return koNames;
        };
        Utils.getKoContainerlessControls = function (parent) {
            var a = [];
            parent = parent || document.body;
            // need jQuery as it does a great job at selecting comment elements
            $(parent).contents().filter(function (i, e) {
                return e.nodeType == 8 && /^\s*ko\s*foreach:/.test(e.nodeValue);
            }).each(function (i, e) {
                a.push(e);
            });
            return a;
        };
        Utils.getEditableKoContainerlessControls = function (parent) {
            parent = parent || document.body;
            var comments = Utils.getKoContainerlessControls(parent);
            var a = [];
            var rxNotTypes = /(^button|submit|cancel|reset)/i;
            var rxTagNames = /(input|textarea)/i;
            var rxIsContext = /\$data/;
            for (var i = 0; i < comments.length; i++) {
                var next = Utils.getNextSibling(comments[i]);
                // when next sibling is the input
                var db = next.getAttribute('data-bind');
                if (!!db && rxTagNames.test(next.tagName) && rxIsContext.test(db) && rxNotTypes.test(next.getAttribute('type'))) {
                    a.push(comments[i]);
                    continue;
                }
                //otherwise the input control is a child of the next sibling
                var bindings = next.querySelectorAll("input[data-bind*='$data']:enabled, textarea[data-bind*='$data']:enabled");
                if (bindings.length > 0) {
                    a.push(comments[i]);
                }
            }
            return a;
        };
        Utils.getEditableKoControlNames = function (parent) {
            var a = [];
            var rxNotTypes = /(button|submit|cancel|reset)/;
            var rx = /\s*:\s*(\$root.|)\w*\b/;
            var replace = $(parent).find('[data-bind]').filter(':input').filter(function (i, e) {
                return !rxNotTypes.test($(e).attr('type'));
            }).each(clean);
            $(parent).find('[data-bind][contenteditable="true"]').each(clean);
            function clean(i, e) {
                var exec = rx.exec($(e).attr('data-bind'));
                var koName = !!exec ? exec[0]
                    .replace(/:(\s+|)/g, '')
                    .replace(/\$root\./, '')
                    .replace(/\s/g, '') : null;
                if (koName != null) {
                    a.push(koName);
                }
            }
            return a;
        };
        /**
        * Get the KO names of the edit input controls on a form.
        * @parem parent: HTMLElement
        * @return Array<string>
        */
        Utils.getEditableKoNames = function (parent) {
            parent = parent || document.body;
            var a = [];
            var rxExcludeInputTypes = /(button|submit|cancel|reset)/;
            $(parent).find('.so-editable[ko-name]').each(function (i, el) {
                var n = $(el).attr('ko-name');
                if (Utils.indexOf(a, n) < 0) {
                    a.push(n);
                }
            });
            // get KO containerless control names
            var comments = Utils.getEditableKoContainerlessControls(parent);
            for (var i = 0; i < comments.length; i++) {
                var n = comments[i].nodeValue
                    .replace(/\s*ko\s*foreach\s*:\s*(\$root\.|)/, '')
                    .replace(/\s/g, '');
                if (Utils.indexOf(a, n) < 0) {
                    a.push(n);
                }
            }
            // get KO input controls
            var koNames = Utils.getEditableKoControlNames(parent);
            for (var i = 0; i < koNames.length; i++) {
                var n = koNames[i];
                if (Utils.indexOf(a, n) < 0) {
                    a.push(n);
                }
            }
            return a;
        };
        Utils.getNextSibling = function (el) {
            do {
                el = el.nextSibling;
            } while (el.nodeType != 1);
            return el;
        };
        /**
        * Extract the Knockout observable name from a field with `data-bind` attribute.
        * If the KO name is `$data`, the method will recursively search for the closest parent element or comment with the `foreach:` binding.
        * @param control: HTMLElement
        * @return string
        */
        Utils.observableNameFromControl = function (control, vm) {
            if (vm === void 0) { vm = undefined; }
            var db = control.getAttribute('data-bind');
            if (!!!db) {
                return null;
            }
            var koName = $(control).attr('ko-name');
            if (!!koName) {
                return koName;
            }
            var rx = /(\b:\s*|\$root\.)\w*\b/;
            var exec = rx.exec(db);
            koName = !!exec ? exec[0]
                .replace(/:(\s+|)/g, '')
                .replace(/\$root\./, '')
                .replace(/\s/g, '') : null;
            return koName;
        };
        Utils.parseJsonDate = function (d) {
            if (!Utils.isJsonDateTicks(d)) {
                return null;
            }
            return new Date(parseInt(d.replace(/\D/g, '')));
        };
        Utils.parseIsoDate = function (d) {
            if (!Utils.isIsoDateString(d)) {
                return null;
            }
            return new Date(d);
        };
        Utils.isJsonDateTicks = function (val) {
            // `/Date(1442769001000)/`
            if (!!!val) {
                return false;
            }
            return /\/Date\(\d+\)\//.test(val + '');
        };
        Utils.isIsoDateString = function (val) {
            // `2015-09-23T16:21:24Z`
            if (!!!val) {
                return false;
            }
            return /\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z/.test(val + '');
        };
        Utils.getQueryParam = function (p) {
            var escape = window["escape"], unescape = window["unescape"];
            p = escape(unescape(p));
            var regex = new RegExp("[?&]" + p + "(?:=([^&]*))?", "i");
            var match = regex.exec(window.location.search);
            return match != null ? match[1] : null;
        };
        // https://developer.mozilla.org/en-US/docs/Web/Guide/API/DOM/The_structured_clone_algorithm
        Utils.clone = function (objectToBeCloned) {
            // Basis.
            if (!(objectToBeCloned instanceof Object)) {
                return objectToBeCloned;
            }
            var objectClone;
            // Filter out special objects.
            var Constructor = objectToBeCloned.constructor;
            switch (Constructor) {
                // Implement other special objects here.
                case RegExp:
                    objectClone = new Constructor(objectToBeCloned);
                    break;
                case Date:
                    objectClone = new Constructor(objectToBeCloned.getTime());
                    break;
                default:
                    objectClone = new Constructor();
            }
            // Clone each property.
            for (var prop in objectToBeCloned) {
                objectClone[prop] = Utils.clone(objectToBeCloned[prop]);
            }
            return objectClone;
        };
        Utils.logError = function (msg, errorLogListName, siteUrl, debug) {
            if (siteUrl === void 0) { siteUrl = '/'; }
            if (debug === void 0) { debug = false; }
            if (debug || !Shockout.SPForm.enableErrorLog) {
                throw msg;
                return;
            }
            siteUrl = Utils.formatSubsiteUrl(siteUrl);
            var loc = window.location.href;
            var errorMsg = '<p>An error occurred at <a href="' + loc + '" target="_blank">' + loc + '</a></p><p>Message: ' + msg + '</p>';
            var $jqXhr = $.ajax({
                url: siteUrl + "_vti_bin/listdata.svc/" + errorLogListName.replace(/\s/g, ''),
                type: "POST",
                processData: false,
                contentType: "application/json;odata=verbose",
                data: JSON.stringify({ "Title": "Web Form Error", "Error": errorMsg }),
                headers: {
                    "Accept": "application/json;odata=verbose"
                }
            });
            $jqXhr.fail(function (data) {
                throw data.responseJSON.error;
            });
        };
        /* update a KO observable whether it's an update or text field */
        Utils.updateKoField = function (el, val) {
            if (el.tagName.toLowerCase() == "input") {
                $(el).val(val);
            }
            else {
                $(el).html(val);
            }
        };
        //validate format ID;#UserName
        Utils.validateSpPerson = function (person) {
            return person != null && person.toString().match(/^\d*;#/) != null;
        };
        Utils.isTime = function (val) {
            if (!!!val) {
                return false;
            }
            var rx = /\d{1,2}:\d{2}(:\d{2}|)\s{0,1}(AM|PM)/;
            return rx.test(val);
        };
        Utils.isDate = function (val) {
            if (!!!val) {
                return false;
            }
            var rx = /\d{1,2}\/\d{1,2}\/\d{4}/;
            return rx.test(val.toString());
        };
        Utils.dateToLocaleString = function (d) {
            try {
                var dd = d.getUTCDate();
                dd = dd < 10 ? "0" + dd : dd;
                var mo = d.getUTCMonth() + 1;
                mo = mo < 10 ? "0" + mo : mo;
                return mo + '/' + dd + '/' + d.getUTCFullYear();
            }
            catch (e) {
                return 'Invalid Date';
            }
        };
        Utils.toTimeLocaleObject = function (d) {
            var hours = 0;
            var minutes;
            var tt;
            hours = d.getUTCHours();
            minutes = d.getUTCMinutes();
            tt = hours > 11 ? 'PM' : 'AM';
            if (minutes < 10) {
                minutes = '0' + minutes;
            }
            if (hours > 12) {
                hours -= 12;
            }
            return {
                hours: hours,
                minutes: minutes,
                tt: tt
            };
        };
        Utils.toTimeLocaleString = function (d) {
            var str = '12:00 AM';
            var hours = d.getUTCHours();
            var minutes = d.getUTCMinutes();
            var tt = hours > 11 ? 'PM' : 'AM';
            if (minutes < 10) {
                minutes = '0' + minutes;
            }
            if (hours > 12) {
                hours -= 12;
            }
            else if (hours == 0) {
                hours = 12;
            }
            return hours + ':' + minutes + ' ' + tt;
        };
        Utils.toDateTimeLocaleString = function (d) {
            var time = Utils.toTimeLocaleString(d);
            return Utils.dateToLocaleString(d) + ' ' + time;
        };
        /**
        * Parse dates in format: "MM/DD/YYYY", "MM-DD-YYYY", "YYYY-MM-DD", "/Date(1442769001000)/", or YYYY-MM-DDTHH:MM:SSZ
        * @param val: string
        * @return Date
        */
        Utils.parseDate = function (val) {
            if (!!!val) {
                return null;
            }
            if (typeof val == 'object' && val.constructor == Date) {
                return val;
            }
            var rxSlash = /\d{1,2}\/\d{1,2}\/\d{2,4}/, // "09/29/2015" 
            rxHyphen = /\d{1,2}-\d{1,2}-\d{2,4}/, // "09-29-2015"
            rxIsoDate = /\d{4}-\d{1,2}-\d{1,2}/, // "2015-09-29"
            rxTicks = /(\/|)\d{13}(\/|)/, // "/1442769001000/"
            rxIsoDateTime = /\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z/, tmp, m, d, y, date = null;
            val = rxIsoDate.test(val) ? val : (val + '').replace(/[^0-9\/\-]/g, '');
            if (val == '') {
                return null;
            }
            if (rxSlash.test(val) || rxHyphen.test(val)) {
                tmp = rxSlash.test(val) ? val.split('/') : val.split('-');
                m = parseInt(tmp[0]) - 1;
                d = parseInt(tmp[1]);
                y = parseInt(tmp[2]);
                y = y < 100 ? 2000 + y : y;
                date = new Date(y, m, d, 0, 0, 0, 0);
            }
            else if (rxIsoDate.test(val)) {
                tmp = val.split('-');
                y = parseInt(tmp[0]);
                m = parseInt(tmp[1]) - 1;
                d = parseInt(tmp[2]);
                y = y < 100 ? 2000 + y : y;
                date = new Date(y, m, d, 0, 0, 0, 0);
            }
            else if (rxIsoDateTime.test(val)) {
                date = new Date(val);
            }
            else if (rxTicks.test(val)) {
                date = new Date(parseInt(val.replace(/\D/g, '')));
            }
            return date;
        };
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
        Utils.formatMoney = function (value, symbol, precision) {
            if (symbol === void 0) { symbol = '$'; }
            if (precision === void 0) { precision = 2; }
            // Clean up number:
            var num = Utils.unformatNumber(value), format = '%s%v', neg = format.replace('%v', '-%v'), useFormat = num > 0 ? format : num < 0 ? neg : format, // Choose which format to use for this value:
            numFormat = Utils.formatNumber(Math.abs(num), Utils.checkPrecision(precision));
            // Return with currency symbol added:
            return useFormat
                .replace('%s', symbol)
                .replace('%v', numFormat);
        };
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
        Utils.unformatNumber = function (value) {
            // Return the value as-is if it's already a number:
            if (typeof value === "number")
                return value;
            // Build regex to strip out everything except digits, decimal point and minus sign:
            var unformatted = parseFloat((value + '')
                .replace(/\((.*)\)/, '-$1') // replace parenthesis for negative numbers
                .replace(/[^0-9-.]/g, ''));
            return !isNaN(unformatted) ? unformatted : 0;
        };
        /**
        * Addapted from accounting.js library. http://josscrowcroft.github.com/accounting.js/
        * Format a number, with comma-separated thousands and custom precision/decimal places
        *
        * Localise by overriding the precision and thousand / decimal separators
        * 2nd parameter `precision` can be an object matching `settings.number`
        */
        Utils.formatNumber = function (value, precision) {
            if (precision === void 0) { precision = 0; }
            // Build options object from second param (if object) or all params, extending defaults:
            var num = Utils.unformatNumber(value), usePrecision = Utils.checkPrecision(precision), negative = num < 0 ? "-" : "", base = parseInt(Utils.toFixed(Math.abs(num || 0), usePrecision), 10) + "", mod = base.length > 3 ? base.length % 3 : 0;
            // Format the number:
            return negative + (mod ? base.substr(0, mod) + ',' : '') + base.substr(mod).replace(/(\d{3})(?=\d)/g, '$1,') + (usePrecision ? '.' + Utils.toFixed(Math.abs(num), usePrecision).split('.')[1] : "");
        };
        /**
         * Tests whether supplied parameter is a string
         * from underscore.js
         */
        Utils.isString = function (obj) {
            return !!(obj === '' || (obj && obj.charCodeAt && obj.substr));
        };
        /**
        * Addapted from accounting.js library.
        * Implementation of toFixed() that treats floats more like decimals
        *
        * Fixes binary rounding issues (eg. (0.615).toFixed(2) === "0.61") that present
        * problems for accounting- and finance-related software.
        */
        Utils.toFixed = function (value, precision) {
            if (precision === void 0) { precision = 0; }
            precision = Utils.checkPrecision(precision);
            var power = Math.pow(10, precision);
            // Multiply up by precision, round accurately, then divide and use native toFixed():
            return (Math.round(Utils.unformatNumber(value) * power) / power).toFixed(precision);
        };
        /**
        * Addapted from accounting.js library. http://josscrowcroft.github.com/accounting.js/
        * Check and normalise the value of precision (must be positive integer)
        */
        Utils.checkPrecision = function (val) {
            val = Math.round(Math.abs(val));
            return isNaN(val) ? 0 : val;
        };
        /**
        * Compares two arrays and returns array of unique matches.
        * @param array1: Array<any>
        * @param array2: Array<any>
        * @return boolean
        */
        Utils.compareArrays = function (array1, array2) {
            var results = [];
            for (var i = 0; i < array1.length; i++) {
                for (var j = 0; j < array2.length; j++) {
                    if (array1[i] == array2[j] && Utils.indexOf(results, array2[j]) < 0) {
                        results.push(array2[j]);
                    }
                }
            }
            return results;
        };
        Utils.trim = function (str) {
            if (!Utils.isString(str)) {
                return str;
            }
            return str.replace(/(^\s+|\s+$)/g, '');
        };
        Utils.formatPictureUrl = function (pictureUrl) {
            return pictureUrl == null
                ? '/_layouts/images/person.gif'
                : pictureUrl.indexOf(',') > -1
                    ? pictureUrl.split(',')[0]
                    : pictureUrl;
        };
        /**
        * Alias for observableNameFromControl()
        */
        Utils.koNameFromControl = Utils.observableNameFromControl;
        return Utils;
    })();
    Shockout.Utils = Utils;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    /**
     * http://github.com/valums/file-uploader
     *
     * Multiple file upload component with progress-bar, drag-and-drop.
     * © 2010 Andrew Valums ( andrew(at)valums.com )
     *
     * Licensed under GNU GPL 2 or later, see license.txt.
     */
    //
    // Helper functions
    //
    Shockout.qq = Shockout.qq || {};
    /**
     * Adds all missing properties from second obj to first obj
     */
    Shockout.qq.extend = function (first, second) {
        for (var prop in second) {
            first[prop] = second[prop];
        }
    };
    /**
     * Searches for a given element in the array, returns -1 if it is not present.
     * @param {Number} [from] The index at which to begin the search
     */
    Shockout.qq.indexOf = function (arr, elt, from) {
        if (arr.indexOf)
            return arr.indexOf(elt, from);
        from = from || 0;
        var len = arr.length;
        if (from < 0)
            from += len;
        for (; from < len; from++) {
            if (from in arr && arr[from] === elt) {
                return from;
            }
        }
        return -1;
    };
    Shockout.qq.getUniqueId = (function () {
        var id = 0;
        return function () { return id++; };
    })();
    //
    // Events
    Shockout.qq.attach = function (element, type, fn) {
        if (element.addEventListener) {
            element.addEventListener(type, fn, false);
        }
        else if (element.attachEvent) {
            element.attachEvent('on' + type, fn);
        }
    };
    Shockout.qq.detach = function (element, type, fn) {
        if (element.removeEventListener) {
            element.removeEventListener(type, fn, false);
        }
        else if (element.attachEvent) {
            element.detachEvent('on' + type, fn);
        }
    };
    Shockout.qq.preventDefault = function (e) {
        if (e.preventDefault) {
            e.preventDefault();
        }
        else {
            e.returnValue = false;
        }
    };
    //
    // Node manipulations
    /**
     * Insert node a before node b.
     */
    Shockout.qq.insertBefore = function (a, b) {
        b.parentNode.insertBefore(a, b);
    };
    Shockout.qq.remove = function (element) {
        element.parentNode.removeChild(element);
    };
    Shockout.qq.contains = function (parent, descendant) {
        // compareposition returns false in this case
        if (parent == descendant)
            return true;
        if (parent.contains) {
            return parent.contains(descendant);
        }
        else {
            return !!(descendant.compareDocumentPosition(parent) & 8);
        }
    };
    /**
     * Creates and returns element from html string
     * Uses innerHTML to create an element
     */
    Shockout.qq.toElement = (function () {
        var div = document.createElement('div');
        return function (html) {
            div.innerHTML = html;
            var element = div.firstChild;
            div.removeChild(element);
            return element;
        };
    })();
    //
    // Node properties and attributes
    /**
     * Sets styles for an element.
     * Fixes opacity in IE6-8.
     */
    Shockout.qq.css = function (element, styles) {
        if (styles.opacity != null) {
            if (typeof element.style.opacity != 'string' && typeof (element.filters) != 'undefined') {
                styles.filter = 'alpha(opacity=' + Math.round(100 * styles.opacity) + ')';
            }
        }
        Shockout.qq.extend(element.style, styles);
    };
    Shockout.qq.hasClass = function (element, name) {
        var re = new RegExp('(^| )' + name + '( |$)');
        return re.test(element.className);
    };
    Shockout.qq.addClass = function (element, name) {
        if (!Shockout.qq.hasClass(element, name)) {
            element.className += ' ' + name;
        }
    };
    Shockout.qq.removeClass = function (element, name) {
        var re = new RegExp('(^| )' + name + '( |$)');
        element.className = element.className.replace(re, ' ').replace(/^\s+|\s+$/g, "");
    };
    Shockout.qq.setText = function (element, text) {
        element.innerText = text;
        element.textContent = text;
    };
    //
    // Selecting elements
    Shockout.qq.children = function (element) {
        var children = [], child = element.firstChild;
        while (child) {
            if (child.nodeType == 1) {
                children.push(child);
            }
            child = child.nextSibling;
        }
        return children;
    };
    Shockout.qq.getByClass = function (element, className) {
        if (element.querySelectorAll) {
            return element.querySelectorAll('.' + className);
        }
        var result = [];
        var candidates = element.getElementsByTagName("*");
        var len = candidates.length;
        for (var i = 0; i < len; i++) {
            if (Shockout.qq.hasClass(candidates[i], className)) {
                result.push(candidates[i]);
            }
        }
        return result;
    };
    /**
     * obj2url() takes a json-object as argument and generates
     * a querystring. pretty much like jQuery.param()
     *
     * how to use:
     *
     *    `qq.obj2url({a:'b',c:'d'},'http://any.url/upload?otherParam=value');`
     *
     * will result in:
     *
     *    `http://any.url/upload?otherParam=value&a=b&c=d`
     *
     * @param  Object JSON-Object
     * @param  String current querystring-part
     * @return String encoded querystring
     */
    Shockout.qq.obj2url = function (obj, temp, prefixDone) {
        var uristrings = [], prefix = '&', add = function (nextObj, i) {
            var nextTemp = temp
                ? (/\[\]$/.test(temp)) // prevent double-encoding
                    ? temp
                    : temp + '[' + i + ']'
                : i;
            if ((nextTemp != 'undefined') && (i != 'undefined')) {
                uristrings.push((typeof nextObj === 'object')
                    ? Shockout.qq.obj2url(nextObj, nextTemp, true)
                    : (Object.prototype.toString.call(nextObj) === '[object Function]')
                        ? encodeURIComponent(nextTemp) + '=' + encodeURIComponent(nextObj())
                        : encodeURIComponent(nextTemp) + '=' + encodeURIComponent(nextObj));
            }
        };
        if (!prefixDone && temp) {
            prefix = (/\?/.test(temp)) ? (/\?$/.test(temp)) ? '' : '&' : '?';
            uristrings.push(temp);
            uristrings.push(Shockout.qq.obj2url(obj));
        }
        else if ((Object.prototype.toString.call(obj) === '[object Array]') && (typeof obj != 'undefined')) {
            // we wont use a for-in-loop on an array (performance)
            for (var i = 0, len = obj.length; i < len; ++i) {
                add(obj[i], i);
            }
        }
        else if ((typeof obj != 'undefined') && (obj !== null) && (typeof obj === "object")) {
            // for anything else but a scalar, we will use for-in-loop
            for (var p in obj) {
                add(obj[p], p);
            }
        }
        else {
            uristrings.push(encodeURIComponent(temp) + '=' + encodeURIComponent(obj));
        }
        return uristrings.join(prefix)
            .replace(/^&/, '')
            .replace(/%20/g, '+');
    };
    //
    //
    // Uploader Classes
    //
    //
    /**
     * Creates upload button, validates upload, but doesn't create file list or dd.
     */
    Shockout.qq.FileUploaderBasic = function (o) {
        this._options = {
            // set to true to see the server response
            debug: false,
            action: '/server/upload',
            params: {},
            button: null,
            multiple: true,
            maxConnections: 3,
            // validation        
            allowedExtensions: [],
            sizeLimit: 0,
            minSizeLimit: 0,
            // events
            // return false to cancel submit
            onSubmit: function (id, fileName) { },
            onProgress: function (id, fileName, loaded, total) { },
            onComplete: function (id, fileName, responseJSON) { },
            onCancel: function (id, fileName) { },
            // messages                
            messages: {
                typeError: "{file} has invalid extension. Only {extensions} are allowed.",
                sizeError: "{file} is too large, maximum file size is {sizeLimit}.",
                minSizeError: "{file} is too small, minimum file size is {minSizeLimit}.",
                emptyError: "{file} is empty, please select files again without it.",
                onLeave: "The files are being uploaded, if you leave now the upload will be cancelled."
            },
            showMessage: function (message) {
                alert(message);
            }
        };
        Shockout.qq.extend(this._options, o);
        // number of files being uploaded
        this._filesInProgress = 0;
        this._handler = this._createUploadHandler();
        if (this._options.button) {
            this._button = this._createUploadButton(this._options.button);
        }
        this._preventLeaveInProgress();
    };
    Shockout.qq.FileUploaderBasic.prototype = {
        setParams: function (params) {
            this._options.params = params;
        },
        getInProgress: function () {
            return this._filesInProgress;
        },
        _createUploadButton: function (element) {
            var self = this;
            return new Shockout.qq.UploadButton({
                element: element,
                multiple: this._options.multiple && Shockout.qq.UploadHandlerXhr.isSupported(),
                onChange: function (input) {
                    self._onInputChange(input);
                }
            });
        },
        _createUploadHandler: function () {
            var self = this, handlerClass;
            if (Shockout.qq.UploadHandlerXhr.isSupported()) {
                handlerClass = 'UploadHandlerXhr';
            }
            else {
                handlerClass = 'UploadHandlerForm';
            }
            var handler = new Shockout.qq[handlerClass]({
                debug: this._options.debug,
                action: this._options.action,
                maxConnections: this._options.maxConnections,
                onProgress: function (id, fileName, loaded, total) {
                    self._onProgress(id, fileName, loaded, total);
                    self._options.onProgress(id, fileName, loaded, total);
                },
                onComplete: function (id, fileName, result) {
                    self._onComplete(id, fileName, result);
                    self._options.onComplete(id, fileName, result);
                },
                onCancel: function (id, fileName) {
                    self._onCancel(id, fileName);
                    self._options.onCancel(id, fileName);
                }
            });
            return handler;
        },
        _preventLeaveInProgress: function () {
            var self = this;
            Shockout.qq.attach(window, 'beforeunload', function (e) {
                if (!self._filesInProgress) {
                    return;
                }
                var e = e || window.event;
                // for ie, ff
                e.returnValue = self._options.messages.onLeave;
                // for webkit
                return self._options.messages.onLeave;
            });
        },
        _onSubmit: function (id, fileName) {
            this._filesInProgress++;
        },
        _onProgress: function (id, fileName, loaded, total) {
        },
        _onComplete: function (id, fileName, result) {
            this._filesInProgress--;
            if (result.error) {
                this._options.showMessage(result.error);
            }
        },
        _onCancel: function (id, fileName) {
            this._filesInProgress--;
        },
        _onInputChange: function (input) {
            if (this._handler instanceof Shockout.qq.UploadHandlerXhr) {
                this._uploadFileList(input.files);
            }
            else {
                if (this._validateFile(input)) {
                    this._uploadFile(input);
                }
            }
            this._button.reset();
        },
        _uploadFileList: function (files) {
            for (var i = 0; i < files.length; i++) {
                if (!this._validateFile(files[i])) {
                    return;
                }
            }
            for (var i = 0; i < files.length; i++) {
                this._uploadFile(files[i]);
            }
        },
        _uploadFile: function (fileContainer) {
            var id = this._handler.add(fileContainer);
            var fileName = this._handler.getName(id);
            if (this._options.onSubmit(id, fileName) !== false) {
                this._onSubmit(id, fileName);
                this._handler.upload(id, this._options.params);
            }
        },
        _validateFile: function (file) {
            var name, size;
            if (file.value) {
                // it is a file input            
                // get input value and remove path to normalize
                name = file.value.replace(/.*(\/|\\)/, "");
            }
            else {
                // fix missing properties in Safari
                name = file.fileName != null ? file.fileName : file.name;
                size = file.fileSize != null ? file.fileSize : file.size;
            }
            if (!this._isAllowedExtension(name)) {
                this._error('typeError', name);
                return false;
            }
            else if (size === 0) {
                this._error('emptyError', name);
                return false;
            }
            else if (size && this._options.sizeLimit && size > this._options.sizeLimit) {
                this._error('sizeError', name);
                return false;
            }
            else if (size && size < this._options.minSizeLimit) {
                this._error('minSizeError', name);
                return false;
            }
            return true;
        },
        _error: function (code, fileName) {
            var message = this._options.messages[code];
            function r(name, replacement) { message = message.replace(name, replacement); }
            r('{file}', this._formatFileName(fileName));
            r('{extensions}', this._options.allowedExtensions.join(', '));
            r('{sizeLimit}', this._formatSize(this._options.sizeLimit));
            r('{minSizeLimit}', this._formatSize(this._options.minSizeLimit));
            this._options.showMessage(message);
        },
        _formatFileName: function (name) {
            if (name.length > 33) {
                name = name.slice(0, 19) + '...' + name.slice(-13);
            }
            return name;
        },
        _isAllowedExtension: function (fileName) {
            var ext = (-1 !== fileName.indexOf('.')) ? fileName.replace(/.*[.]/, '').toLowerCase() : '';
            var allowed = this._options.allowedExtensions;
            if (!allowed.length) {
                return true;
            }
            for (var i = 0; i < allowed.length; i++) {
                if (allowed[i].toLowerCase() == ext) {
                    return true;
                }
            }
            return false;
        },
        _formatSize: function (bytes) {
            var i = -1;
            do {
                bytes = bytes / 1024;
                i++;
            } while (bytes > 99);
            return Math.max(bytes, 0.1).toFixed(1) + ['kB', 'MB', 'GB', 'TB', 'PB', 'EB'][i];
        }
    };
    /**
     * Class that creates upload widget with drag-and-drop and file list
     * @inherits qq.FileUploaderBasic
     */
    Shockout.qq.FileUploader = function (o) {
        // call parent constructor
        Shockout.qq.FileUploaderBasic.apply(this, arguments);
        // additional options    
        Shockout.qq.extend(this._options, {
            element: null,
            // if set, will be used instead of qq-upload-list in template
            listElement: null,
            template: '<div class="qq-uploader">' +
                '<div class="qq-upload-drop-area"><span>Drop files here to upload</span></div>' +
                '<div class="qq-upload-button">Attach File</div>' +
                '<ul class="qq-upload-list"></ul>' +
                '</div>',
            // template for one item in file list
            fileTemplate: '<li>' +
                '<span class="qq-upload-file"></span>' +
                '<span class="qq-upload-spinner"></span>' +
                '<span class="qq-upload-size"></span>' +
                '<a class="qq-upload-cancel" href="#">Cancel</a>' +
                '<span class="qq-upload-failed-text">Failed</span>' +
                '</li>',
            classes: {
                // used to get elements from templates
                button: 'qq-upload-button',
                drop: 'qq-upload-drop-area',
                dropActive: 'qq-upload-drop-area-active',
                list: 'qq-upload-list',
                file: 'qq-upload-file',
                spinner: 'qq-upload-spinner',
                size: 'qq-upload-size',
                cancel: 'qq-upload-cancel',
                // added to list item when upload completes
                // used in css to hide progress spinner
                success: 'qq-upload-success',
                fail: 'qq-upload-fail'
            }
        });
        // overwrite options with user supplied    
        Shockout.qq.extend(this._options, o);
        this._element = this._options.element;
        this._element.innerHTML = this._options.template;
        this._listElement = this._options.listElement || this._find(this._element, 'list');
        this._classes = this._options.classes;
        this._button = this._createUploadButton(this._find(this._element, 'button'));
        this._bindCancelEvent();
        this._setupDragDrop();
    };
    // inherit from Basic Uploader
    Shockout.qq.extend(Shockout.qq.FileUploader.prototype, Shockout.qq.FileUploaderBasic.prototype);
    Shockout.qq.extend(Shockout.qq.FileUploader.prototype, {
        /**
         * Gets one of the elements listed in this._options.classes
         **/
        _find: function (parent, type) {
            var element = Shockout.qq.getByClass(parent, this._options.classes[type])[0];
            if (!element) {
                throw new Error('element not found ' + type);
            }
            return element;
        },
        _setupDragDrop: function () {
            var self = this, dropArea = this._find(this._element, 'drop');
            var dz = new Shockout.qq.UploadDropZone({
                element: dropArea,
                onEnter: function (e) {
                    Shockout.qq.addClass(dropArea, self._classes.dropActive);
                    e.stopPropagation();
                },
                onLeave: function (e) {
                    e.stopPropagation();
                },
                onLeaveNotDescendants: function (e) {
                    Shockout.qq.removeClass(dropArea, self._classes.dropActive);
                },
                onDrop: function (e) {
                    dropArea.style.display = 'none';
                    Shockout.qq.removeClass(dropArea, self._classes.dropActive);
                    self._uploadFileList(e.dataTransfer.files);
                }
            });
            dropArea.style.display = 'none';
            Shockout.qq.attach(document, 'dragenter', function (e) {
                if (!dz._isValidFileDrag(e))
                    return;
                dropArea.style.display = 'block';
            });
            Shockout.qq.attach(document, 'dragleave', function (e) {
                if (!dz._isValidFileDrag(e))
                    return;
                var relatedTarget = document.elementFromPoint(e.clientX, e.clientY);
                // only fire when leaving document out
                if (!relatedTarget || relatedTarget.nodeName == "HTML") {
                    dropArea.style.display = 'none';
                }
            });
        },
        _onSubmit: function (id, fileName) {
            Shockout.qq.FileUploaderBasic.prototype._onSubmit.apply(this, arguments);
            this._addToList(id, fileName);
        },
        _onProgress: function (id, fileName, loaded, total) {
            Shockout.qq.FileUploaderBasic.prototype._onProgress.apply(this, arguments);
            var item = this._getItemByFileId(id);
            var size = this._find(item, 'size');
            size.style.display = 'inline';
            var text;
            if (loaded != total) {
                text = Math.round(loaded / total * 100) + '% from ' + this._formatSize(total);
            }
            else {
                text = this._formatSize(total);
            }
            Shockout.qq.setText(size, text);
        },
        _onComplete: function (id, fileName, result) {
            Shockout.qq.FileUploaderBasic.prototype._onComplete.apply(this, arguments);
            // mark completed
            var item = this._getItemByFileId(id);
            Shockout.qq.remove(this._find(item, 'cancel'));
            Shockout.qq.remove(this._find(item, 'spinner'));
            if (result.success) {
                Shockout.qq.addClass(item, this._classes.success);
            }
            else {
                Shockout.qq.addClass(item, this._classes.fail);
            }
        },
        _addToList: function (id, fileName) {
            var item = Shockout.qq.toElement(this._options.fileTemplate);
            item.qqFileId = id;
            var fileElement = this._find(item, 'file');
            Shockout.qq.setText(fileElement, this._formatFileName(fileName));
            this._find(item, 'size').style.display = 'none';
            this._listElement.appendChild(item);
        },
        _getItemByFileId: function (id) {
            var item = this._listElement.firstChild;
            // there can't be txt nodes in dynamically created list
            // and we can  use nextSibling
            while (item) {
                if (item.qqFileId == id)
                    return item;
                item = item.nextSibling;
            }
        },
        /**
         * delegate click event for cancel link
         **/
        _bindCancelEvent: function () {
            var self = this, list = this._listElement;
            Shockout.qq.attach(list, 'click', function (e) {
                e = e || window.event;
                var target = e.target || e.srcElement;
                if (Shockout.qq.hasClass(target, self._classes.cancel)) {
                    Shockout.qq.preventDefault(e);
                    var item = target.parentNode;
                    self._handler.cancel(item.qqFileId);
                    Shockout.qq.remove(item);
                }
            });
        }
    });
    Shockout.qq.UploadDropZone = function (o) {
        this._options = {
            element: null,
            onEnter: function (e) { },
            onLeave: function (e) { },
            // is not fired when leaving element by hovering descendants   
            onLeaveNotDescendants: function (e) { },
            onDrop: function (e) { }
        };
        Shockout.qq.extend(this._options, o);
        this._element = this._options.element;
        this._disableDropOutside();
        this._attachEvents();
    };
    Shockout.qq.UploadDropZone.prototype = {
        _disableDropOutside: function (e) {
            // run only once for all instances
            if (!Shockout.qq.UploadDropZone.dropOutsideDisabled) {
                Shockout.qq.attach(document, 'dragover', function (e) {
                    if (e.dataTransfer) {
                        e.dataTransfer.dropEffect = 'none';
                        e.preventDefault();
                    }
                });
                Shockout.qq.UploadDropZone.dropOutsideDisabled = true;
            }
        },
        _attachEvents: function () {
            var self = this;
            Shockout.qq.attach(self._element, 'dragover', function (e) {
                if (!self._isValidFileDrag(e))
                    return;
                var effect = e.dataTransfer.effectAllowed;
                if (effect == 'move' || effect == 'linkMove') {
                    e.dataTransfer.dropEffect = 'move'; // for FF (only move allowed)    
                }
                else {
                    e.dataTransfer.dropEffect = 'copy'; // for Chrome
                }
                e.stopPropagation();
                e.preventDefault();
            });
            Shockout.qq.attach(self._element, 'dragenter', function (e) {
                if (!self._isValidFileDrag(e))
                    return;
                self._options.onEnter(e);
            });
            Shockout.qq.attach(self._element, 'dragleave', function (e) {
                if (!self._isValidFileDrag(e))
                    return;
                self._options.onLeave(e);
                var relatedTarget = document.elementFromPoint(e.clientX, e.clientY);
                // do not fire when moving a mouse over a descendant
                if (Shockout.qq.contains(this, relatedTarget))
                    return;
                self._options.onLeaveNotDescendants(e);
            });
            Shockout.qq.attach(self._element, 'drop', function (e) {
                if (!self._isValidFileDrag(e))
                    return;
                e.preventDefault();
                self._options.onDrop(e);
            });
        },
        _isValidFileDrag: function (e) {
            var dt = e.dataTransfer, 
            // do not check dt.types.contains in webkit, because it crashes safari 4            
            isWebkit = navigator.userAgent.indexOf("AppleWebKit") > -1;
            // dt.effectAllowed is none in Safari 5
            // dt.types.contains check is for firefox            
            return dt && dt.effectAllowed != 'none' &&
                (dt.files || (!isWebkit && dt.types.contains && dt.types.contains('Files')));
        }
    };
    Shockout.qq.UploadButton = function (o) {
        this._options = {
            element: null,
            // if set to true adds multiple attribute to file input      
            multiple: false,
            // name attribute of file input
            name: 'file',
            onChange: function (input) { },
            hoverClass: 'qq-upload-button-hover',
            focusClass: 'qq-upload-button-focus'
        };
        Shockout.qq.extend(this._options, o);
        this._element = this._options.element;
        // make button suitable container for input
        Shockout.qq.css(this._element, {
            position: 'relative',
            overflow: 'hidden',
            // Make sure browse button is in the right side
            // in Internet Explorer
            direction: 'ltr'
        });
        this._input = this._createInput();
    };
    Shockout.qq.UploadButton.prototype = {
        /* returns file input element */
        getInput: function () {
            return this._input;
        },
        /* cleans/recreates the file input */
        reset: function () {
            if (this._input.parentNode) {
                Shockout.qq.remove(this._input);
            }
            Shockout.qq.removeClass(this._element, this._options.focusClass);
            this._input = this._createInput();
        },
        _createInput: function () {
            var input = document.createElement("input");
            if (this._options.multiple) {
                input.setAttribute("multiple", "multiple");
            }
            input.setAttribute("type", "file");
            input.setAttribute("name", this._options.name);
            Shockout.qq.css(input, {
                position: 'absolute',
                // in Opera only 'browse' button
                // is clickable and it is located at
                // the right side of the input
                right: 0,
                top: 0,
                fontFamily: 'Arial',
                // 4 persons reported this, the max values that worked for them were 243, 236, 236, 118
                fontSize: '118px',
                margin: 0,
                padding: 0,
                cursor: 'pointer',
                opacity: 0
            });
            this._element.appendChild(input);
            var self = this;
            Shockout.qq.attach(input, 'change', function () {
                self._options.onChange(input);
            });
            Shockout.qq.attach(input, 'mouseover', function () {
                Shockout.qq.addClass(self._element, self._options.hoverClass);
            });
            Shockout.qq.attach(input, 'mouseout', function () {
                Shockout.qq.removeClass(self._element, self._options.hoverClass);
            });
            Shockout.qq.attach(input, 'focus', function () {
                Shockout.qq.addClass(self._element, self._options.focusClass);
            });
            Shockout.qq.attach(input, 'blur', function () {
                Shockout.qq.removeClass(self._element, self._options.focusClass);
            });
            // IE and Opera, unfortunately have 2 tab stops on file input
            // which is unacceptable in our case, disable keyboard access
            if (window["attachEvent"]) {
                // it is IE or Opera
                input.setAttribute('tabIndex', "-1");
            }
            return input;
        }
    };
    /**
     * Class for uploading files, uploading itself is handled by child classes
     */
    Shockout.qq.UploadHandlerAbstract = function (o) {
        this._options = {
            debug: false,
            action: '/upload.php',
            // maximum number of concurrent uploads        
            maxConnections: 999,
            onProgress: function (id, fileName, loaded, total) { },
            onComplete: function (id, fileName, response) { },
            onCancel: function (id, fileName) { }
        };
        Shockout.qq.extend(this._options, o);
        this._queue = [];
        // params for files in queue
        this._params = [];
    };
    Shockout.qq.UploadHandlerAbstract.prototype = {
        log: function (str) {
            if (this._options.debug && window.console)
                console.log('[uploader] ' + str);
        },
        /**
         * Adds file or file input to the queue
         * @returns id
         **/
        add: function (file) { },
        /**
         * Sends the file identified by id and additional query params to the server
         */
        upload: function (id, params) {
            var len = this._queue.push(id);
            var copy = {};
            Shockout.qq.extend(copy, params);
            this._params[id] = copy;
            // if too many active uploads, wait...
            if (len <= this._options.maxConnections) {
                this._upload(id, this._params[id]);
            }
        },
        /**
         * Cancels file upload by id
         */
        cancel: function (id) {
            this._cancel(id);
            this._dequeue(id);
        },
        /**
         * Cancells all uploads
         */
        cancelAll: function () {
            for (var i = 0; i < this._queue.length; i++) {
                this._cancel(this._queue[i]);
            }
            this._queue = [];
        },
        /**
         * Returns name of the file identified by id
         */
        getName: function (id) { },
        /**
         * Returns size of the file identified by id
         */
        getSize: function (id) { },
        /**
         * Returns id of files being uploaded or
         * waiting for their turn
         */
        getQueue: function () {
            return this._queue;
        },
        /**
         * Actual upload method
         */
        _upload: function (id) { },
        /**
         * Actual cancel method
         */
        _cancel: function (id) { },
        /**
         * Removes element from queue, starts upload of next
         */
        _dequeue: function (id) {
            var i = Shockout.qq.indexOf(this._queue, id);
            this._queue.splice(i, 1);
            var max = this._options.maxConnections;
            if (this._queue.length >= max) {
                var nextId = this._queue[max - 1];
                this._upload(nextId, this._params[nextId]);
            }
        }
    };
    /**
     * Class for uploading files using form and iframe
     * @inherits qq.UploadHandlerAbstract
     */
    Shockout.qq.UploadHandlerForm = function (o) {
        Shockout.qq.UploadHandlerAbstract.apply(this, arguments);
        this._inputs = {};
    };
    // @inherits qq.UploadHandlerAbstract
    Shockout.qq.extend(Shockout.qq.UploadHandlerForm.prototype, Shockout.qq.UploadHandlerAbstract.prototype);
    Shockout.qq.extend(Shockout.qq.UploadHandlerForm.prototype, {
        add: function (fileInput) {
            fileInput.setAttribute('name', 'qqfile');
            var id = 'qq-upload-handler-iframe' + Shockout.qq.getUniqueId();
            this._inputs[id] = fileInput;
            // remove file input from DOM
            if (fileInput.parentNode) {
                Shockout.qq.remove(fileInput);
            }
            return id;
        },
        getName: function (id) {
            // get input value and remove path to normalize
            return this._inputs[id].value.replace(/.*(\/|\\)/, "");
        },
        _cancel: function (id) {
            this._options.onCancel(id, this.getName(id));
            delete this._inputs[id];
            var iframe = document.getElementById(id);
            if (iframe) {
                // to cancel request set src to something else
                // we use src="javascript:false;" because it doesn't
                // trigger ie6 prompt on https
                iframe.setAttribute('src', 'javascript:false;');
                Shockout.qq.remove(iframe);
            }
        },
        _upload: function (id, params) {
            var input = this._inputs[id];
            if (!input) {
                throw new Error('file with passed id was not added, or already uploaded or cancelled');
            }
            var fileName = this.getName(id);
            var iframe = this._createIframe(id);
            var form = this._createForm(iframe, params);
            form.appendChild(input);
            var self = this;
            this._attachLoadEvent(iframe, function () {
                self.log('iframe loaded');
                var response = self._getIframeContentJSON(iframe);
                self._options.onComplete(id, fileName, response);
                self._dequeue(id);
                delete self._inputs[id];
                // timeout added to fix busy state in FF3.6
                setTimeout(function () {
                    Shockout.qq.remove(iframe);
                }, 1);
            });
            form.submit();
            Shockout.qq.remove(form);
            return id;
        },
        _attachLoadEvent: function (iframe, callback) {
            Shockout.qq.attach(iframe, 'load', function () {
                // when we remove iframe from dom
                // the request stops, but in IE load
                // event fires
                if (!iframe.parentNode) {
                    return;
                }
                // fixing Opera 10.53
                if (iframe.contentDocument &&
                    iframe.contentDocument.body &&
                    iframe.contentDocument.body.innerHTML == "false") {
                    // In Opera event is fired second time
                    // when body.innerHTML changed from false
                    // to server response approx. after 1 sec
                    // when we upload file with iframe
                    return;
                }
                callback();
            });
        },
        /**
         * Returns json object received by iframe from server.
         */
        _getIframeContentJSON: function (iframe) {
            // iframe.contentWindow.document - for IE<7
            var doc = iframe.contentDocument ? iframe.contentDocument : iframe.contentWindow.document, response;
            this.log("converting iframe's innerHTML to JSON");
            this.log("innerHTML = " + doc.body.innerHTML);
            try {
                response = eval("(" + doc.body.innerHTML + ")");
            }
            catch (err) {
                response = {};
            }
            return response;
        },
        /**
         * Creates iframe with unique name
         */
        _createIframe: function (id) {
            // We can't use following code as the name attribute
            // won't be properly registered in IE6, and new window
            // on form submit will open
            // var iframe = document.createElement('iframe');
            // iframe.setAttribute('name', id);
            var iframe = Shockout.qq.toElement('<iframe src="javascript:false;" name="' + id + '" />');
            // src="javascript:false;" removes ie6 prompt on https
            iframe.setAttribute('id', id);
            iframe.style.display = 'none';
            document.body.appendChild(iframe);
            return iframe;
        },
        /**
         * Creates form, that will be submitted to iframe
         */
        _createForm: function (iframe, params) {
            // We can't use the following code in IE6
            // var form = document.createElement('form');
            // form.setAttribute('method', 'post');
            // form.setAttribute('enctype', 'multipart/form-data');
            // Because in this case file won't be attached to request
            var form = Shockout.qq.toElement('<form method="post" enctype="multipart/form-data"></form>');
            var queryString = Shockout.qq.obj2url(params, this._options.action);
            form.setAttribute('action', queryString);
            form.setAttribute('target', iframe.name);
            form.style.display = 'none';
            document.body.appendChild(form);
            return form;
        }
    });
    /**
     * Class for uploading files using xhr
     * @inherits qq.UploadHandlerAbstract
     */
    Shockout.qq.UploadHandlerXhr = function (o) {
        Shockout.qq.UploadHandlerAbstract.apply(this, arguments);
        this._files = [];
        this._xhrs = [];
        // current loaded size in bytes for each file 
        this._loaded = [];
    };
    // static method
    Shockout.qq.UploadHandlerXhr.isSupported = function () {
        var input = document.createElement('input');
        input.type = 'file';
        return ('multiple' in input &&
            typeof File != "undefined" &&
            typeof (new XMLHttpRequest()).upload != "undefined");
    };
    // @inherits qq.UploadHandlerAbstract
    Shockout.qq.extend(Shockout.qq.UploadHandlerXhr.prototype, Shockout.qq.UploadHandlerAbstract.prototype);
    Shockout.qq.extend(Shockout.qq.UploadHandlerXhr.prototype, {
        /**
         * Adds file to the queue
         * Returns id to use with upload, cancel
         **/
        add: function (file) {
            if (!(file instanceof File)) {
                throw new Error('Passed obj in not a File (in qq.UploadHandlerXhr)');
            }
            return this._files.push(file) - 1;
        },
        getName: function (id) {
            var file = this._files[id];
            // fix missing name in Safari 4
            return file.fileName != null ? file.fileName : file.name;
        },
        getSize: function (id) {
            var file = this._files[id];
            return file.fileSize != null ? file.fileSize : file.size;
        },
        /**
         * Returns uploaded bytes for file identified by id
         */
        getLoaded: function (id) {
            return this._loaded[id] || 0;
        },
        /**
         * Sends the file identified by id and additional query params to the server
         * @param {Object} params name-value string pairs
         */
        _upload: function (id, params) {
            var file = this._files[id], name = this.getName(id), size = this.getSize(id);
            this._loaded[id] = 0;
            var xhr = this._xhrs[id] = new XMLHttpRequest();
            var self = this;
            xhr.upload.onprogress = function (e) {
                if (e.lengthComputable) {
                    self._loaded[id] = e.loaded;
                    self._options.onProgress(id, name, e.loaded, e.total);
                }
            };
            xhr.onreadystatechange = function () {
                if (xhr.readyState == 4) {
                    self._onComplete(id, xhr);
                }
            };
            // build query string
            params = params || {};
            params['qqfile'] = name;
            var queryString = Shockout.qq.obj2url(params, this._options.action);
            xhr.open("POST", queryString, true);
            xhr.setRequestHeader("X-Requested-With", "XMLHttpRequest");
            xhr.setRequestHeader("X-File-Name", encodeURIComponent(name));
            xhr.setRequestHeader("Content-Type", "application/octet-stream");
            xhr.send(file);
        },
        _onComplete: function (id, xhr) {
            // the request was aborted/cancelled
            if (!this._files[id])
                return;
            var name = this.getName(id);
            var size = this.getSize(id);
            this._options.onProgress(id, name, size, size);
            if (xhr.status == 200) {
                this.log("xhr - server response received");
                this.log("responseText = " + xhr.responseText);
                var response;
                try {
                    response = eval("(" + xhr.responseText + ")");
                }
                catch (err) {
                    response = {};
                }
                this._options.onComplete(id, name, response);
            }
            else {
                this._options.onComplete(id, name, {});
            }
            this._files[id] = null;
            this._xhrs[id] = null;
            this._dequeue(id);
        },
        _cancel: function (id) {
            this._options.onCancel(id, this.getName(id));
            this._files[id] = null;
            if (this._xhrs[id]) {
                this._xhrs[id].abort();
                this._xhrs[id] = null;
            }
        }
    });
})(Shockout || (Shockout = {}));
/// <reference path="Shockout/a_spform.ts" />
/// <reference path="Shockout/b_viewmodel.ts" />
/// <reference path="Shockout/c_kohandlers.ts" />
/// <reference path="Shockout/d_kocomponents.ts" />
/// <reference path="Shockout/e_spapi.ts" />
/// <reference path="Shockout/f_spapi15.ts" />
/// <reference path="Shockout/g_spsoap.ts" />
/// <reference path="Shockout/h_spdatatypes.ts" />
/// <reference path="Shockout/i_spdatatypes15.ts" />
/// <reference path="Shockout/j_templates.ts" />
/// <reference path="Shockout/k_utils.ts" />
/// <reference path="Shockout/z_qqfileuploader.ts" />
//# sourceMappingURL=ShockoutForms-1.0.1.js.map