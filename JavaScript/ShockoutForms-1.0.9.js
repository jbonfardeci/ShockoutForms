'use strict';
var Shockout;
(function (Shockout) {
    var SPForm = (function () {
        function SPForm(listName, formId, options) {
            if (options === void 0) { options = undefined; }
            this.allowDelete = false;
            this.allowPrint = true;
            this.allowSave = false;
            this.allowedExtensions = ['txt', 'rtf', 'zip', 'pdf', 'doc', 'docx', 'jpg', 'gif', 'png', 'ppt', 'tif', 'pptx', 'csv', 'pub', 'msg'];
            this.attachmentMessage = 'An attachment is required.';
            this.confirmationUrl = '/SitePages/Confirmation.aspx';
            this.debug = false;
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
            this.editableFields = [];
            this.enableAttachments = false;
            this.enableErrorLog = true;
            this.errorLogListName = 'Error Log';
            this.errorLogSiteUrl = '/';
            this.fieldNames = [];
            this.includeUserProfiles = true;
            this.includeWorkflowHistory = true;
            this.includeNavigationMenu = true;
            this.requireAttachments = false;
            this.searchPrincipals = true;
            this.siteUrl = '';
            this.utils = Shockout.Utils;
            this.viewModelIsBound = false;
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
            this.requireCheckout = false;
            this.rootUrl = window.location.protocol + '//' + window.location.hostname + (!!window.location.port ? ':' + window.location.port : '');
            this.sourceUrl = null;
            this.version = '1.0.9';
            this.queryStringId = 'formid';
            this.isSp2013 = false;
            var self = this;
            var error;
            if (!(this instanceof SPForm)) {
                error = 'You must declare an instance of this class with `new`.';
                alert(error);
                throw error;
            }
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
            }
            this.formId = formId;
            this.listName = listName;
            this.listNameRest = Shockout.Utils.toCamelCase(listName);
            this.form = (typeof formId == 'string' ? document.getElementById(formId) : formId);
            if (!!!this.form) {
                alert('An element with the ID "' + this.formId + '" was not found. Ensure the `formId` parameter in the constructor matches the ID attribute of the form element.');
                return;
            }
            this.$form = $(this.form).addClass('sp-form');
            $('form').attr({ 'novalidate': 'novalidate' });
            this.sourceUrl = Shockout.Utils.getQueryParam('source');
            if (!!this.sourceUrl) {
                this.sourceUrl = decodeURIComponent(this.sourceUrl);
            }
            if (options && options.constructor === Object) {
                for (var p in options) {
                    this[p] = options[p];
                }
            }
            SPForm.DEBUG = this.debug;
            SPForm.searchPrincipals = this.searchPrincipals;
            this.itemId = Shockout.Utils.getIdFromHash();
            var idFromQs = Shockout.Utils.getQueryParam(this.queryStringId);
            if (!!!this.itemId && /\d/.test(idFromQs)) {
                this.itemId = parseInt(idFromQs);
                Shockout.Utils.setIdHash(this.itemId);
            }
            SPForm.errorLogListName = this.errorLogListName;
            SPForm.errorLogSiteUrl = this.errorLogSiteUrl;
            SPForm.enableErrorLog = this.enableErrorLog;
            Shockout.KoHandlers.bindKoHandlers();
            this.viewModel = new Shockout.ViewModel(this);
            this.viewModel.showUserProfiles(this.includeUserProfiles);
            self.$formStatus = $('<div>', { 'class': 'form-status' }).appendTo(self.$form);
            self.$dialog = $('<div>', { 'id': 'formdialog' })
                .appendTo(self.$form)
                .dialog(self.dialogOpts);
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
                    Shockout.KoComponents.registerKoComponents();
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
            this.nextAsync(true, 'Begin initialization...');
        }
        SPForm.prototype.getCurrentUser = function () { return this.currentUser; };
        SPForm.prototype.getDefaultViewUrl = function () { return this.defaultViewUrl; };
        SPForm.prototype.getDefailtMobileViewUrl = function () { return this.defaultMobileViewUrl; };
        SPForm.prototype.getForm = function () { return this.form; };
        SPForm.prototype.getItemId = function () { return this.itemId; };
        SPForm.prototype.setItemId = function (id) {
            this.itemId = id;
        };
        SPForm.prototype.getListId = function () { return this.listId; };
        SPForm.prototype.getListItem = function () { return this.listItem; };
        SPForm.prototype.requiresCheckout = function () { return this.requireCheckout; };
        SPForm.prototype.getRootUrl = function () { return this.rootUrl; };
        SPForm.prototype.getSourceUrl = function () { return this.sourceUrl; };
        SPForm.prototype.getViewModel = function () { return this.viewModel; };
        SPForm.prototype.getVersion = function () { return this.version; };
        SPForm.prototype.nextAsync = function (success, msg, args) {
            if (success === void 0) { success = true; }
            if (msg === void 0) { msg = undefined; }
            if (args === void 0) { args = undefined; }
            var self = this;
            if (msg) {
                this.updateStatus(msg, success, self);
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
            this.asyncFns.shift()(self, args);
        };
        SPForm.prototype.getCurrentUserAsync = function (self, args) {
            if (args === void 0) { args = undefined; }
            self.updateStatus('Retrieving your account...', true, self);
            var success = 'Retrieved your account.';
            if (self.debug) {
                console.info('Testing for SP 2013 API...');
            }
            Shockout.SpApi15.getCurrentUser(function (user, error) {
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
            }, true);
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
        SPForm.prototype.getListAsync = function (self, args) {
            if (args === void 0) { args = undefined; }
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
                    self.defaultMobileViewUrl = $list.attr('MobileDefaultViewUrl');
                    $(xmlDoc).find('Field').filter(function (i, el) {
                        return !!($(el).attr('DisplayName')) && $(el).attr('Hidden') != 'TRUE' && !rxExcludeNames.test($(el).attr('Name'));
                    }).each(setupKoVar);
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
                    var koName = Shockout.Utils.toCamelCase(displayName);
                    if (koName in self.viewModel) {
                        return;
                    }
                    self.fieldNames.push(koName);
                    var defaultValue;
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
                    for (var p in koObj._metadata) {
                        koObj['_' + p] = koObj._metadata[p];
                    }
                    if (rxIsChoice.test(spType)) {
                        var isFillIn = $el.attr('FillInChoice');
                        koObj._isFillInChoice = !!isFillIn && isFillIn == 'True';
                        var choices = [];
                        var options = [];
                        $el.find('CHOICE').each(function (j, choice) {
                            var txt = $(choice).text();
                            choices.push({ 'value': $(choice).text(), 'selected': false });
                            options.push(txt);
                        });
                        koObj._choices = choices;
                        koObj._options = options;
                        koObj._multiChoice = !!spType && spType == 'MultiChoice';
                        koObj._metadata.choices = choices;
                        koObj._metadata.options = options;
                        koObj._metadata.multichoice = koObj._multiChoice;
                    }
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
        SPForm.prototype.initForm = function (self, args) {
            if (args === void 0) { args = undefined; }
            try {
                self.updateStatus("Initializing dynamic form features...", true, self);
                var vm = self.viewModel;
                var rx = /submitted/i;
                if (Shockout.Utils.indexOf(self.fieldNames, 'IsSubmitted') > -1) {
                    self.allowSave = true;
                    Shockout.ViewModel.isSubmittedKey = 'IsSubmitted';
                    if (self.debug) {
                        console.info('initFormAsync: IsSubmitted key: ' + Shockout.ViewModel.isSubmittedKey);
                    }
                }
                self.viewModel.allowSave(self.allowSave);
                self.viewModel.allowPrint(self.allowPrint);
                self.viewModel.allowDelete(self.allowDelete);
                self.$formAction = $(Shockout.Templates.getFormAction()).appendTo(self.$form);
                if (self.enableAttachments) {
                    self.setupAttachments(self);
                }
                if (self.enableErrorLog) {
                    Shockout.SpApi.getListItems(self.errorLogListName, function (data, error) {
                        if (!!error) {
                            self.enableErrorLog = SPForm.enableErrorLog = false;
                        }
                    }, self.errorLogSiteUrl, null, 'Title,Error', 'Modified', 1, false);
                }
                if (self.includeNavigationMenu) {
                    var $navMenu = self.$form.find(".so-nav-menu, [data-so-nav-menu], [so-nav-menu]");
                    if ($navMenu.length > 0) {
                        $navMenu.replaceWith(Shockout.Templates.soNavMenuControl);
                    }
                    else {
                        self.$form.prepend(Shockout.Templates.soNavMenuControl);
                    }
                }
                self.$form.find(".created-info, [data-sp-created-info], [data-so-created-info], [sp-created-info], [so-created-info]").replaceWith(Shockout.Templates.soCreatedModifiedInfoControl);
                if (self.includeWorkflowHistory) {
                    var $wfControls = self.$form.find(".workflow-history, [data-so-workflow-history], [so-workflow-history]");
                    if ($wfControls.length > 0) {
                        $wfControls.replaceWith(Shockout.Templates.soWorkflowHistoryControl);
                    }
                    else {
                        self.$form.append(Shockout.Templates.soWorkflowHistoryControl);
                    }
                }
                self.$form.find('[data-new-only]')
                    .before('<!-- ko ifnot: !!$root.Id() -->')
                    .after('<!-- /ko -->');
                self.$form.find('[data-edit-only]')
                    .before('<!-- ko if: !!$root.Id() -->')
                    .after('<!-- /ko -->');
                self.$form.find('[data-author-only]')
                    .before('<!-- ko if: !!$root.isAuthor() -->')
                    .after('<!-- /ko -->');
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
        SPForm.prototype.getListItemAsync = function (self, args) {
            if (args === void 0) { args = undefined; }
            if (!!!self.itemId) {
                self.nextAsync(true, "This is a New form.");
                return;
            }
            self.updateStatus("Retrieving form values...", true, self);
            var vm = self.viewModel;
            var expand = [];
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
        SPForm.prototype.getUsersGroupsAsync = function (self, args) {
            if (args === void 0) { args = undefined; }
            if (self.isSp2013) {
                self.nextAsync(true);
                return;
            }
            self.updateStatus("Retrieving your groups...", true, self);
            Shockout.SpSoap.getUsersGroups(self.currentUser.login, function callback(groups, error) {
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
            });
        };
        SPForm.prototype.implementPermissions = function (self, args) {
            if (args === void 0) { args = undefined; }
            try {
                self.updateStatus("Retrieving your permissions...", true, self);
                self.$form.find("[data-sp-groups], [user-groups]").each(function (i, el) {
                    var groups = $(el).attr("data-sp-groups");
                    if (!!!groups) {
                        groups = $(el).attr("user-groups");
                    }
                    $(el).before('<!-- ko if: !!$root.isMember(' + groups + ') -->')
                        .after('<!-- /ko -->');
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
        SPForm.prototype.getHistoryAsync = function (self, args) {
            if (args === void 0) { args = undefined; }
            if (!!!self.itemId || !self.includeWorkflowHistory) {
                self.nextAsync(true);
                return;
            }
            self.updateStatus('Retrieving workflow history...', true, self);
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
        SPForm.prototype.bindListItemValues = function (self) {
            if (self === void 0) { self = undefined; }
            self = self || this;
            try {
                if (!!!self.itemId) {
                    return;
                }
                var item = self.listItem;
                var vm = self.viewModel;
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
                var createdBy = item.CreatedBy;
                var modifiedBy = item.ModifiedBy;
                if (!!createdBy && !!modifiedBy) {
                    item.CreatedBy.Picture = Shockout.Utils.formatPictureUrl(item.CreatedBy.Picture);
                    item.ModifiedBy.Picture = Shockout.Utils.formatPictureUrl(item.ModifiedBy.Picture);
                    item.CreatedBy.WorkEMail = item.CreatedBy.WorkEMail || item.CreatedBy.EMail || '';
                    item.ModifiedBy.WorkEMail = item.ModifiedBy.WorkEMail || item.ModifiedBy.EMail || '';
                    item.CreatedBy.JobTitle = item.CreatedBy.JobTitle || item.CreatedBy.Title || null;
                    item.ModifiedBy.JobTitle = item.ModifiedBy.JobTitle || item.ModifiedBy.Title || null;
                    item.CreatedBy.WorkPhone = item.CreatedBy.WorkPhone || createdBy.MobileNumber || null;
                    item.ModifiedBy.WorkPhone = item.ModifiedBy.WorkPhone || modifiedBy.MobileNumber || null;
                    item.CreatedBy.Office = item.CreatedBy.Office || null;
                    item.ModifiedBy.Office = item.ModifiedBy.Office || null;
                    vm.CreatedBy(item.CreatedBy);
                    vm.ModifiedBy(item.ModifiedBy);
                }
                vm.Created(Shockout.Utils.parseDate(item.Created));
                vm.Modified(Shockout.Utils.parseDate(item.Modified));
                $(self.fieldNames).filter(function (i, key) {
                    if (!!!self.viewModel[key]) {
                        return false;
                    }
                    return self.viewModel[key]._type == 'User' && (key + 'Id') in item;
                }).each(function (i, key) {
                    self.getPersonById(parseInt(item[key + 'Id']), vm[key]);
                });
                $(self.fieldNames).filter(function (i, key) {
                    if (!!!self.viewModel[key]) {
                        return false;
                    }
                    return self.viewModel[key]._type == 'Choice' && (key + 'Value' in item);
                }).each(function (i, key) {
                    vm[key](item[key + 'Value']);
                });
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
                $(self.fieldNames).filter(function (i, key) {
                    return !!self.viewModel[key] && self.viewModel[key]._type == 'UserMulti' && '__deferred' in item[key];
                }).each(function (i, key) {
                    Shockout.SpApi.executeRestRequest(item[key].__deferred.uri, function (data, status, jqXhr) {
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
        SPForm.prototype.saveListItem = function (vm, isSubmit, customMsg, callback) {
            if (isSubmit === void 0) { isSubmit = false; }
            if (customMsg === void 0) { customMsg = undefined; }
            if (callback === void 0) { callback = undefined; }
            var self = vm.parent;
            var isNew = !!(self.itemId == null), data = [], timeout = 3000, saveMsg = customMsg || '<p>Your form has been saved.</p>', fields = [];
            try {
                var editable = Shockout.Utils.getEditableKoNames(self.$form);
                $(editable).each(function (i, key) {
                    self.pushEditableFieldName(key);
                });
                self.editableFields.sort();
                if (self.debug) {
                    console.info('Editable fields...');
                    console.info(self.editableFields);
                }
                isSubmit = typeof (isSubmit) == "undefined" ? true : isSubmit;
                if (self.preSave) {
                    var retVal = self.preSave(self, self.viewModel, isSubmit);
                    if (typeof (retVal) != 'undefined' && !!!retVal) {
                        return;
                    }
                }
                if (isSubmit && !self.formIsValid(vm)) {
                    return;
                }
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
                Shockout.SpSoap.updateListItem(self.itemId, self.listName, fields, isNew, self.siteUrl, cb);
            }
            catch (e) {
                if (self.debug) {
                    throw e;
                }
                self.logError('Error in SpForm.saveListItem(): ', e);
            }
            function cb(xmlDoc, status, jqXhr) {
                var itemId;
                if (self.debug) {
                    console.log('Callback from saveListItem()...');
                    console.log(status);
                    console.log(xmlDoc);
                }
                var $errorText = $(xmlDoc).find('ErrorText');
                if (!!$errorText && $errorText.text() != "") {
                    self.logError($errorText.text());
                    return;
                }
                $(xmlDoc).find('*').filter(function () {
                    return Shockout.Utils.isZrow(this);
                }).each(function (i, el) {
                    itemId = parseInt($(el).attr('ows_ID'));
                    if (self.itemId == null) {
                        self.itemId = itemId;
                        vm.Id(itemId);
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
                    if (callback) {
                        callback(self.itemId);
                    }
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
                    self.getListItemAsync(self);
                    if (callback) {
                        callback(self.itemId);
                    }
                    if (self.includeWorkflowHistory) {
                        setTimeout(function () { self.getHistoryAsync(self); }, 5000);
                    }
                }
            }
            ;
        };
        SPForm.prototype.finalize = function (self) {
            try {
                self.setupNavigation(self);
                self.setupDatePickers(self);
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
        SPForm.prototype.log = function (msg) {
            if (this.debug) {
                console.log(msg);
            }
        };
        SPForm.prototype.updateStatus = function (msg, success, spForm) {
            if (success === void 0) { success = true; }
            var self = spForm;
            self.$formStatus
                .html(msg)
                .css('color', (success ? "#ff0" : "$f00"))
                .show();
            setTimeout(function () { self.$formStatus.hide(); }, 2000);
        };
        SPForm.prototype.showDialog = function (msg, title, timeout) {
            if (title === void 0) { title = undefined; }
            if (timeout === void 0) { timeout = undefined; }
            var self = this;
            title = title || "Form Dialog";
            msg = (msg).toString().match(/<\w>\w*/) == null ? '<p>' + msg + '</p>' : msg;
            self.$dialog.html(msg).dialog('open');
            if (timeout) {
                setTimeout(function () { self.$dialog.dialog('close'); }, timeout);
            }
        };
        SPForm.prototype.formIsValid = function (model, showDialog) {
            if (showDialog === void 0) { showDialog = false; }
            var self = model.parent, labels = [], errorCount = 0, invalidCount = 0, invalidLabels = [];
            try {
                self.$form.find('.required, [required]').each(function checkRequired(i, n) {
                    var koName = Shockout.Utils.observableNameFromControl(n, self.viewModel);
                    if (!!koName && model[koName]) {
                        var val = model[koName]();
                        if (val == null || $.trim(val + '').length == 0) {
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
                self.$form.find(".invalid").each(function (i, el) {
                    var $parent = $(el).parent();
                    invalidLabels.push($(parent).first().html());
                    invalidCount++;
                });
                if (invalidCount > 0) {
                    labels.push('<p class="warning">There are validation errors with the following fields. Please correct before saving.</p><p style="color:#f00;">' + invalidLabels.join('<br />') + '</p>');
                }
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
                var name = person.Id + ';#' + person.Name.replace(/^i\:0\#\.w\|/, '');
                koField(name);
                if (self.debug) {
                    console.warn('Retrieved person by ID... ' + name);
                }
            });
        };
        SPForm.prototype.pushEditableFieldName = function (key) {
            if (!!!key || Shockout.Utils.indexOf(this.editableFields, key) > -1 || key.match(/^(_|\$)/) != null || Shockout.Utils.indexOf(this.fieldNames, key) < 0 || this.viewModel[key]._readOnly) {
                return -1;
            }
            return this.editableFields.push(key);
        };
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
                Shockout.Utils.logError(err, self.errorLogListName, self.errorLogSiteUrl, self.debug);
                self.showDialog('<p>An error has occurred and the web administrator has been notified.</p><pre>' + err + '</pre>');
            }
        };
        SPForm.prototype.setupAttachments = function (self) {
            if (self === void 0) { self = undefined; }
            self = self || this;
            var vm = self.viewModel;
            var count = 0;
            if (!self.enableAttachments) {
                return count;
            }
            try {
                self.$form.find(".attachments, [data-sp-attachments]").each(function (i, att) {
                    $(att).replaceWith('<so-attachments params="val: attachments"></so-attachments>');
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
        SPForm.prototype.setupNavigation = function (self) {
            if (self === void 0) { self = undefined; }
            var self = self || this;
            var count = 0;
            if (!self.includeNavigationMenu) {
                return count;
            }
            try {
                var $navSections = self.$form.find('.nav-section');
                if ($navSections.length == 0) {
                    return count;
                }
                self.$form.find(".nav-section:visible").each(function (i, el) {
                    var $el = $(el);
                    var $header = $el.find("> h4");
                    if ($header.length == 0) {
                        return;
                    }
                    var title = $header.html();
                    var anchorName = Shockout.Utils.toCamelCase(title) + 'Nav';
                    $el.before('<div style="height:1px;" id="' + anchorName + '">&nbsp;</div>');
                    self.viewModel.navMenuItems().push({ 'title': title, 'anchorName': anchorName });
                    count++;
                });
                self.viewModel.navMenuItems.valueHasMutated();
                self.$form.append('<a href="#TOP" class="back-to-top"><span class="glyphicon glyphicon-chevron-up"></span></a>');
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
        SPForm.prototype.setupDatePickers = function (self) {
            if (self === void 0) { self = undefined; }
            self = self || this;
            var $datepickers = self.$form.find('input.datepicker').datepicker();
            if (self.debug) {
                console.info('Bound ' + $datepickers.length + ' jQueryUI datepickers.');
            }
            return $datepickers.length;
        };
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
            for (var i = 0; i < groupNames.length; i++) {
                var group = groupNames[i];
                group = group.match(/\;#/) != null ? group.split(';')[0] : group;
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
    }());
    Shockout.SPForm = SPForm;
})(Shockout || (Shockout = {}));
window['so'] = window['so'] || Shockout;
var Shockout;
(function (Shockout) {
    var ViewModel = (function () {
        function ViewModel(instance) {
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
            this.navMenuItems = ko.observableArray();
            var self = this;
            this.parent = instance;
            ViewModel.parent = instance;
            this.isValid = ko.pureComputed(function () {
                return self.parent.formIsValid(self);
            });
            this.deleteAttachment = instance.deleteAttachment;
            this.currentUser = ko.observable(instance.getCurrentUser());
            this.attachments.getViewModel = function () {
                return self;
            };
            this.attachments.getSpForm = function () {
                return self.parent;
            };
            this.isMember = ko.pureComputed({
                read: function () {
                    return false;
                },
                write: function (groups) {
                    return self.parent.currentUserIsMemberOfGroups(groups);
                },
                owner: this
            });
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
    }());
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
    }());
    Shockout.KoHandlers = KoHandlers;
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
        ko.bindingHandlers['spPerson'] = {
            init: function (element, valueAccessor, allBindings, bindingContext) {
                try {
                    if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') {
                        return;
                    }
                    var viewModel = bindingContext.$data, modelValue = valueAccessor(), person = ko.unwrap(modelValue);
                    var $element = $(element)
                        .addClass('people-picker-control')
                        .attr('placeholder', 'Employee Account Name');
                    var $parent = $(element).parent();
                    var $spError = $('<div>', { 'class': 'sp-validation person' });
                    $element.after($spError);
                    var $desc = $('<div>', {
                        'class': 'no-print',
                        'html': '<em>Enter the employee name. The auto-suggest menu will appear below the field. Select the account name.</em>'
                    });
                    $spError.after($desc);
                    var $spValidate = $('<button>', {
                        'html': Shockout.Templates.personIcon,
                        'class': Shockout.Templates.buttonDefault,
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
                    var $reset = $('<button>', { 'class': Shockout.Templates.resetButton, 'html': 'Reset' })
                        .on('click', function () {
                        modelValue(null);
                        return false;
                    })
                        .insertAfter($spValidate);
                    var autoCompleteOpts = {
                        source: function () {
                            var src = Shockout.SPForm.searchPrincipals ? searchPrincipals : searchUserInformationList;
                            return $.isFunction(Shockout.SPForm['peopleFilter']) ? Shockout.SPForm['peopleFilter'](src) : src;
                        },
                        minLength: 3,
                        select: function (event, ui) {
                            modelValue(ui.item.value);
                        }
                    };
                    $(element).autocomplete(autoCompleteOpts);
                    $(element)
                        .on('focus', function () { $(this).removeClass('valid'); })
                        .on('blur', function () { onChangeSpPersonEvent(this, modelValue); })
                        .on('mouseout', function () { onChangeSpPersonEvent(this, modelValue); });
                }
                catch (e) {
                    if (Shockout.SPForm.DEBUG) {
                        console.info('Error in Knockout handler spPerson init()');
                        console.info(e);
                    }
                }
                function searchPrincipals(request, response) {
                    Shockout.SpSoap.searchPrincipals(request.term, function (data) {
                        var mapped = $.map(data, function (user) {
                            var account = user.AccountName;
                            if (account.indexOf('|') > -1) {
                                account = account.split('|')[1];
                            }
                            return {
                                label: user.DisplayName + " (" + user.Email + ")",
                                value: user.UserInfoID + ";#" + account
                            };
                        });
                        response(mapped);
                    }, 10, 'User');
                }
                ;
                function searchUserInformationList(request, response) {
                    Shockout.SpApi.searchUsers(request.term, function (data) {
                        var users = [];
                        for (var i = 0; i < data.length; i++) {
                            var user = data[i];
                            var account = user.Account;
                            if (account.indexOf('|') > -1) {
                                account = account.split('|')[1];
                            }
                            users.push({
                                label: user.Name + " (" + user.WorkEmail + ")",
                                value: user.Id + ";#" + account
                            });
                        }
                        response(users);
                    });
                }
                ;
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
                try {
                    var viewModel = bindingContext.$data;
                    var modelValue = valueAccessor();
                    var person = ko.unwrap(modelValue);
                    var displayName = "";
                    if (Shockout.Utils.validateSpPerson(person)) {
                        displayName = person.indexOf('|') > -1 ? person.split('|')[1] : person.split('#')[1];
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
            init: function (element, valueAccessor, allBindings, viewModel, bindingContext) {
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
                }
                $(element)
                    .css('display', 'inline-block')
                    .addClass('datepicker med')
                    .attr('placeholder', 'MM/DD/YYYY')
                    .on('blur', onDateChange)
                    .on('change', onDateChange)
                    .after(Shockout.Templates.calendarIcon);
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
        ko.bindingHandlers['spDateTime'] = {
            after: ['attr'],
            init: function (element, valueAccessor, allBindings, viewModel, bindingContext) {
                if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') {
                    return;
                }
                try {
                    var modelValue = valueAccessor();
                    var model = new Shockout.DateTimeModel(element, modelValue);
                    element['$$model'] = model;
                }
                catch (e) {
                    if (Shockout.SPForm.DEBUG) {
                        console.warn('Error in Knockout handler spDateTime init()...');
                        throw e;
                    }
                }
            },
            update: function (element, valueAccessor, allBindings, viewModel, bindingContext) {
                try {
                    var modelValue = valueAccessor();
                    var date = Shockout.Utils.parseDate(ko.unwrap(modelValue));
                    if (element.tagName.toLowerCase() == 'input') {
                        var model = element['$$model'];
                        if (!!model && model.constructor == Shockout.DateTimeModel) {
                            model.setDisplayValue(modelValue);
                        }
                    }
                    else {
                        $(element).text(Shockout.DateTimeModel.toString(modelValue));
                    }
                }
                catch (e) {
                    if (Shockout.SPForm.DEBUG) {
                        console.warn('Error in Knockout handler spDateTime update()...');
                        throw e;
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
                template: Shockout.Templates.soTextField
            });
            ko.components.register('so-html-field', {
                viewModel: soFieldModel,
                template: Shockout.Templates.soHtmlFieldTemplate
            });
            ko.components.register('so-person-field', {
                viewModel: soFieldModel,
                template: Shockout.Templates.soTextField.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spPerson: modelValue')
            });
            ko.components.register('so-date-field', {
                viewModel: soFieldModel,
                template: Shockout.Templates.soTextField.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spDate: modelValue')
            });
            ko.components.register('so-datetime-field', {
                viewModel: soFieldModel,
                template: Shockout.Templates.soTextField.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spDateTime: modelValue')
            });
            ko.components.register('so-money-field', {
                viewModel: soFieldModel,
                template: Shockout.Templates.soTextField.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spMoney: modelValue')
            });
            ko.components.register('so-number-field', {
                viewModel: soFieldModel,
                template: Shockout.Templates.soTextField.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spNumber: modelValue')
            });
            ko.components.register('so-decimal-field', {
                viewModel: soFieldModel,
                template: Shockout.Templates.soTextField.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spDecimal: modelValue')
            });
            ko.components.register('so-checkbox-field', {
                viewModel: soFieldModel,
                template: Shockout.Templates.soCheckboxField
            });
            ko.components.register('so-select-field', {
                viewModel: soFieldModel,
                template: Shockout.Templates.soSelectField
            });
            ko.components.register('so-checkbox-group', {
                viewModel: soFieldModel,
                template: Shockout.Templates.soCheckboxGroup
            });
            ko.components.register('so-radio-group', {
                viewModel: soFieldModel,
                template: Shockout.Templates.soRadioGroup
            });
            ko.components.register('so-usermulti-group', {
                viewModel: soUsermultiModel,
                template: Shockout.Templates.soUsermultiField
            });
            ko.components.register('so-static-field', {
                viewModel: soStaticModel,
                template: Shockout.Templates.soStaticField
            });
            ko.components.register('so-static-person', {
                viewModel: soStaticModel,
                template: Shockout.Templates.soStaticField.replace(/data-bind="text: modelValue/g, 'data-bind="spPerson: modelValue')
            });
            ko.components.register('so-static-date', {
                viewModel: soStaticModel,
                template: Shockout.Templates.soStaticField.replace(/data-bind="text: modelValue/g, 'data-bind="spDate: modelValue')
            });
            ko.components.register('so-static-datetime', {
                viewModel: soStaticModel,
                template: Shockout.Templates.soStaticField.replace(/data-bind="text: modelValue/g, 'data-bind="spDateTime: modelValue')
            });
            ko.components.register('so-static-money', {
                viewModel: soStaticModel,
                template: Shockout.Templates.soStaticField.replace(/data-bind="text: modelValue/g, 'data-bind="spMoney: modelValue')
            });
            ko.components.register('so-static-number', {
                viewModel: soStaticModel,
                template: Shockout.Templates.soStaticField.replace(/data-bind="text: modelValue/g, 'data-bind="spNumber: modelValue')
            });
            ko.components.register('so-static-decimal', {
                viewModel: soStaticModel,
                template: Shockout.Templates.soStaticField.replace(/data-bind="text: modelValue/g, 'data-bind="spDecimal: modelValue')
            });
            ko.components.register('so-static-html', {
                viewModel: soStaticModel,
                template: Shockout.Templates.soStaticField.replace(/data-bind="text: modelValue/g, 'data-bind="html: modelValue')
            });
            ko.components.register('so-attachments', {
                viewModel: soAttachmentsModel,
                template: Shockout.Templates.soAttachments
            });
            ko.components.register('so-created-modified-info', {
                viewModel: soCreatedModifiedInfoModel,
                template: Shockout.Templates.soCreatedModifiedInfo
            });
            ko.components.register('so-nav-menu', {
                viewModel: soNavMenuModel,
                template: Shockout.Templates.soNavMenu
            });
            ko.components.register('so-workflow-history', {
                viewModel: function (params) {
                    this.historyItems = (params.val || params.historyItems);
                },
                template: Shockout.Templates.soWorkflowHistory
            });
            function soCreatedModifiedInfoModel(params) {
                this.CreatedBy = params.createdBy;
                this.ModifiedBy = params.modifiedBy;
                this.profiles = ko.observableArray([
                    { header: 'Created By', profile: this.CreatedBy },
                    { header: 'Modified By', profile: this.ModifiedBy }
                ]);
                this.Created = params.created;
                this.Modified = params.modified;
                this.showUserProfiles = params.showUserProfiles;
            }
            ;
            function soStaticModel(params) {
                if (!params) {
                    throw 'params is undefined in so-static-field';
                }
                var koObj = params.val || params.modelValue;
                if (!koObj) {
                    throw "Parameter `val` or `modelValue` for so-static-field is required!";
                }
                this.modelValue = koObj;
                this.id = params.id || koObj._koName;
                this.label = params.label || koObj._displayName;
                this.description = params.description || koObj._description;
                var labelX = parseInt(params.labelColWidth || 3);
                var fieldX = parseInt(params.fieldColWidth || (12 - (labelX - 0)));
                this.labelColWidth = 'col-sm-' + labelX;
                this.fieldColWidth = 'col-sm-' + fieldX;
            }
            ;
            function soFieldModel(params) {
                if (!params) {
                    throw 'params is undefined in soFieldModel';
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
                this.editable = !!koObj._koName;
                this.koName = koObj._koName;
                this.options = params.options || koObj._options;
                this.required = (typeof params.required == 'function') ? params.required : ko.observable(!!params.required || false);
                this.inline = params.inline || false;
                this.multiline = params.multiline || false;
                var labelX = parseInt(params.labelColWidth || 3);
                var fieldX = parseInt(params.fieldColWidth || (12 - (labelX - 0)));
                this.labelColWidth = 'col-sm-' + labelX;
                this.fieldColWidth = 'col-sm-' + fieldX;
                this.readOnly = (typeof params.readOnly == 'function') ? params.readOnly : ko.observable(!!params.readOnly || false);
            }
            ;
            function soAttachmentsModel(params) {
                var self = this;
                var w = window;
                this.errorMsg = ko.observable(null);
                if (!!!params) {
                    this.errorMsg('`params` is undefined in component so-attachments');
                    throw this.errorMsg();
                }
                if (!!!params.val) {
                    this.errorMsg('Parameter `val` for component so-attachments is required!');
                    throw this.errorMsg();
                }
                var spForm = params.val.getSpForm();
                var vm = spForm.getViewModel();
                var allowedExtensions = params.allowedExtensions || spForm.allowedExtensions;
                var reader;
                var cafe;
                var asyncFns;
                this.attachments = params.val;
                this.label = params.label || 'Attach Files';
                this.drop = params.drop || true;
                this.dropLabel = params.dropLabel || '...or Drag and Drop Files Here';
                this.className = params.className || 'btn btn-primary';
                this.title = params.title || 'Attachments';
                this.description = params.description;
                this.readOnly = (typeof params.readOnly == 'function') ? params.readOnly : ko.observable(params.readOnly || false);
                this.length = ko.pureComputed(function () { return self.attachments().length; });
                this.fileUploads = ko.observableArray();
                this.hasFileReader = ko.observable(w.File && w.FileReader && w.FileList && w.Blob);
                if (!this.hasFileReader()) {
                    this.errorMsg('This browser does not support the FileReader class required for uplaoding files. You may be using IE 9 or another unsupported browser.');
                }
                this.id = params.id || 'so_fileUploader_' + uniqueId();
                this.deleteAttachment = function (att, event) {
                    if (!confirm('Are you sure you want to delete ' + att.Name + '? This can\'t be undone.')) {
                        return;
                    }
                    Shockout.SpApi.deleteAttachment(att, function (data, error) {
                        if (!!error) {
                            alert("Failed to delete attachment: " + error);
                            return;
                        }
                        self.attachments.remove(att);
                    });
                };
                this.fileHandler = function (e) {
                    var files = document.getElementById(self.id)['files'];
                    readFiles(files);
                };
                this.onSelect = function (e) {
                    cancel(e);
                    document.getElementById(self.id).click();
                };
                this.onDrop = function (localViewModel, e) {
                    cancel(e);
                    if (spForm.debug) {
                        console.info('dropped files over dropzone, arguments are...');
                        console.info(arguments);
                    }
                    var dt = (e.originalEvent || e).dataTransfer;
                    var files = dt.files;
                    if (!!!files) {
                        console.warn('Error in so-attachments - event.dataTransfer.files is ' + typeof files);
                        return false;
                    }
                    else {
                        readFiles(files);
                    }
                };
                function readFiles(files) {
                    asyncFns = [];
                    var fileArray = Array.prototype.slice.call(files, 0);
                    fileArray.map(function (file, i) {
                        asyncFns.push(function () {
                            readFile(file);
                        });
                    });
                    cafe = new Shockout.Cafe(asyncFns);
                    if (vm.Id() == null) {
                        spForm.saveListItem(vm, false, undefined, function (itemId) {
                            if (vm.Id() == null && !!itemId && itemId.toFixed) {
                                vm.Id(itemId);
                            }
                            setTimeout(function () {
                                cafe.next(true);
                            }, 1000);
                        });
                    }
                    else {
                        cafe.next(true);
                    }
                }
                ;
                function readFile(file) {
                    if (spForm.debug) {
                        console.info('uploading file...');
                        console.info(file);
                    }
                    var fileName = file.name.replace(/[^a-zA-Z0-9_\-\.]/g, '');
                    var ext = /\.\w{2,4}$/.exec(fileName)[0];
                    var rootName = fileName.replace(new RegExp(ext + '$'), '');
                    var allowedExtension = new RegExp("^(\\.|)(" + allowedExtensions.join('|') + ")$", "i").test(ext);
                    if (!allowedExtension) {
                        self.errorMsg('Only files with the extensions: ' + allowedExtensions.join(', ') + ' are allowed.');
                        return;
                    }
                    for (var i = 0; i < self.attachments().length; i++) {
                        if (new RegExp(fileName, 'i').test(self.attachments()[i].Name)) {
                            fileName = rootName + '-1' + ext;
                            break;
                        }
                    }
                    var fileUpload = new Shockout.FileUpload(fileName, file.size);
                    self.fileUploads().push(fileUpload);
                    self.fileUploads.valueHasMutated();
                    reader = new FileReader();
                    reader.onerror = function errorHandler(e) {
                        var evt = e;
                        var className = fileUpload.className();
                        fileUpload.className(className.replace('-success', '-danger'));
                        switch (evt.target.error.code) {
                            case evt.target.error.NOT_FOUND_ERR:
                                self.errorMsg = 'File Not Found!';
                                break;
                            case evt.target.error.NOT_READABLE_ERR:
                                self.errorMsg = 'File is not readable.';
                                break;
                            case evt.target.error.ABORT_ERR:
                                break;
                            default:
                                self.errorMsg = 'An error occurred reading this file.';
                        }
                        ;
                    };
                    reader.onprogress = function (e) {
                        updateProgress(e, fileUpload);
                    };
                    reader.onabort = function (e) {
                        self.errorMsg('File read cancelled');
                    };
                    reader.onloadstart = function (e) {
                        fileUpload.progress(0);
                    };
                    reader.onload = function (e) {
                        var event = e;
                        fileUpload.progress(100);
                        Shockout.SpSoap.addAttachment(event.target.result, fileName, spForm.listName, spForm.viewModel.Id(), spForm.siteUrl, callback);
                    };
                    reader.onloadend = function (loadend) {
                    };
                    reader.readAsDataURL(file);
                    function callback() {
                        var status = arguments[1];
                        if (spForm.debug) {
                            console.info('so-html5-attachments.onFileUploadComplete()...');
                            console.info(arguments);
                        }
                        if (!!!status && status == 'error') {
                            var jqXhr = arguments[0];
                            var responseXml = jqXhr.responseXML;
                            var errorString = $(jqXhr.responseXML).find('errorstring').text();
                            if (!!errorString) {
                                spForm.$dialog.html('Error on file upload. Message from server: ' + (errorString || jqXhr.statusText)).dialog('open');
                            }
                            fileUpload.className(fileUpload.className().replace('-success', '-danger'));
                            cafe.next(false);
                        }
                        else if (status == 'success') {
                            var att = new Shockout.SpAttachment(spForm.getRootUrl(), spForm.siteUrl, spForm.listName, spForm.getItemId(), fileName);
                            self.attachments().push(att);
                            self.attachments.valueHasMutated();
                            cafe.next(true);
                        }
                        setTimeout(function () {
                            self.fileUploads.remove(fileUpload);
                        }, 1000);
                    }
                }
                ;
                this.onDragenter = cancel;
                this.onDragover = cancel;
                function updateProgress(e, fileUpload) {
                    if (e.lengthComputable) {
                        var percentLoaded = Math.round((e.loaded / e.total) * 100);
                        if (percentLoaded < 100) {
                            fileUpload.progress(percentLoaded);
                        }
                    }
                }
                ;
                function cancel(e) {
                    if (e.preventDefault) {
                        e.preventDefault();
                    }
                    if (e.stopPropagation) {
                        e.stopPropagation();
                    }
                }
                ;
                if (!spForm.enableAttachments) {
                    this.errorMsg('Attachments are disabled for this form or SharePoint list.');
                    this.readOnly(true);
                }
            }
            function soUsermultiModel(params) {
                if (!params) {
                    throw 'params is undefined in soFieldModel';
                }
                var self = this;
                var koObj = params.val || params.modelValue;
                if (!koObj) {
                    throw "Parameter `val` or `modelValue` for so-text-field is required!";
                }
                this.modelValue = koObj;
                this.placeholder = "";
                this.id = params.id || koObj._koName;
                this.name = params.name || koObj._koName || params.id;
                this.label = params.label || koObj._displayName;
                this.title = params.title;
                this.required = params.required;
                this.description = params.description || koObj._description;
                this.editable = !!koObj._koName;
                this.koName = koObj._koName;
                this.person = ko.observable(null);
                this.readOnly = (typeof params.readOnly == 'function') ? params.readOnly : ko.observable(!!params.readOnly || false);
                this.addPerson = function (model, ctrl) {
                    if (self.modelValue() == null) {
                        self.modelValue([]);
                    }
                    var isAlreadyInArray = false;
                    ko.utils.arrayForEach(self.modelValue(), function (item) {
                        if (item == self.person()) {
                            isAlreadyInArray = true;
                        }
                        return;
                    });
                    if (!isAlreadyInArray) {
                        self.modelValue().push(self.person());
                        self.modelValue.valueHasMutated();
                        self.person(null);
                    }
                    else {
                        this.shake(ctrl);
                    }
                    return false;
                };
                this.removePerson = function (person, event) {
                    try {
                        self.modelValue.remove(person);
                    }
                    catch (err) {
                        var index = self.modelValue().indexOf(person);
                        if (index > -1) {
                            self.modelValue().splice(self.modelValue().indexOf(person), 1);
                            self.modelValue.valueHasMutated();
                        }
                    }
                    return false;
                };
                this.showRequiredText = ko.pureComputed(function () {
                    if (self.required) {
                        if (!!self.modelValue()) {
                            return self.modelValue().length < 1;
                        }
                        return true;
                    }
                    return false;
                });
                this.shake = function (element) {
                    var $el = $('button[id=' + element.currentTarget.id + ']');
                    var shakes = 3;
                    var distance = 5;
                    var duration = 200;
                    $el.css("position", "relative");
                    for (var x = 1; x <= shakes; x++) {
                        $el.removeClass("btn-success")
                            .addClass("btn-danger")
                            .animate({ left: (distance * -1) }, (((duration / shakes) / 4)))
                            .animate({ left: distance }, ((duration / shakes) / 2))
                            .animate({ left: 0 }, (((duration / shakes) / 4)));
                        setTimeout(function () {
                            $el.removeClass("btn-danger btn-warning").addClass("btn-success");
                        }, 1000);
                    }
                };
            }
            ;
            function soNavMenuModel(params) {
                this.navMenuItems = params.val;
                this.title = params.title || 'Navigation';
            }
            ;
        };
        ;
        return KoComponents;
    }());
    Shockout.KoComponents = KoComponents;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var SpApi = (function () {
        function SpApi() {
        }
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
        SpApi.searchUsers = function (term, callback, select, maxResults) {
            if (select === void 0) { select = 'Id,Account,Name,WorkEmail'; }
            if (maxResults === void 0) { maxResults = 10; }
            var url = "/_vti_bin/listdata.svc/UserInformationList?\n            $filter=startswith(WorkEmail, '" + term + "') or startswith(Name, '" + term + "') or startswith(LastName, '" + term + "') and ContentType eq 'Person'\n            &$orderby=Name\n            &$top=" + maxResults;
            SpApi.executeRestRequest(url, fn, true, 'GET');
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
        SpApi.getPersonById = function (id, callback) {
            if (isNaN(parseInt(id + ''))) {
                return;
            }
            var url = '/_vti_bin/listdata.svc/UserInformationList(' + id + ')';
            SpApi.executeRestRequest(url, fn, true, 'GET');
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
                if (!!status && parseInt(status) == 404) {
                    var msg = status + ". The data may have been deleted by another user.";
                }
                else {
                    msg = status + ' ' + error;
                }
                callback(null, msg);
            });
        };
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
    }());
    Shockout.SpApi = SpApi;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var SpApi15 = (function () {
        function SpApi15() {
        }
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
                var account = user.LoginName.replace(/^i\:0\#\.w\|/, '');
                var currentUser = {
                    account: user.Id + ';#' + user.Title,
                    department: null,
                    email: user.Email,
                    groups: [],
                    id: user.Id,
                    jobtitle: null,
                    login: account,
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
                callback(null, jqXhr.status);
            });
        };
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
    }());
    Shockout.SpApi15 = SpApi15;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var SpSoap = (function () {
        function SpSoap() {
        }
        SpSoap.getCurrentUser = function (callback) {
            var user = {};
            var query = '<Query><Where><Eq><FieldRef Name="ID" /><Value Type="Counter"><UserID /></Value></Eq></Where></Query>';
            var viewFields = '<ViewFields><FieldRef Name="ID" /><FieldRef Name="Name" /><FieldRef Name="EMail" /><FieldRef Name="Department" /><FieldRef Name="JobTitle" /><FieldRef Name="UserName" /><FieldRef Name="Office" /></ViewFields>';
            SpSoap.getListItems('', 'User Information List', viewFields, query, function (xmlDoc, status, jqXhr) {
                $(xmlDoc).find('*').filter(function () {
                    return Shockout.Utils.isZrow(this);
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
        };
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
        SpSoap.checkInFile = function (pageUrl, checkinType, callback, comment) {
            if (comment === void 0) { comment = ''; }
            var action = 'http://schemas.microsoft.com/sharepoint/soap/CheckInFile';
            var params = [pageUrl, comment, checkinType];
            var packet = '<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap:Body><CheckInFile xmlns="http://schemas.microsoft.com/sharepoint/soap/"><pageUrl>{0}</pageUrl><comment>{1}</comment><CheckinType>{2}</CheckinType></CheckInFile></soap:Body></soap:Envelope>';
            return this.executeSoapRequest(action, packet, params, null, callback);
        };
        SpSoap.checkOutFile = function (pageUrl, checkoutToLocal, lastmodified, callback) {
            var action = 'http://schemas.microsoft.com/sharepoint/soap/CheckOutFile';
            var params = [pageUrl, checkoutToLocal, lastmodified];
            var packet = '<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap:Body><CheckOutFile xmlns="http://schemas.microsoft.com/sharepoint/soap/"><pageUrl>{0}</pageUrl><checkoutToLocal>{1}</checkoutToLocal><lastmodified>{2}</lastmodified></CheckOutFile></soap:Body></soap:Envelope>';
            return this.executeSoapRequest(action, packet, params, null, callback);
        };
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
                    if (Shockout.SPForm.enableErrorLog && !Shockout.SPForm.DEBUG) {
                        Shockout.Utils.logError(msg, Shockout.SPForm.errorLogListName);
                    }
                    if (callback) {
                        callback(jqXhr, status, error);
                    }
                });
            }
            catch (e) {
                Shockout.Utils.logError('Error in SpSoap.executeSoapRequest.', JSON.stringify(e), Shockout.SPForm.errorLogListName);
                console.warn(e);
            }
        };
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
                    var id = parseInt($('UserInfoID', n).text());
                    if (id > 0) {
                        var account = $('AccountName', n).text().replace(/^i\:0\#\.w\|/, '');
                        results.push({
                            AccountName: account,
                            UserInfoID: id,
                            DisplayName: $('DisplayName', n).text(),
                            Email: $('Email', n).text(),
                            Title: $('Title', n).text(),
                            IsResolved: $('IsResolved', n).text() == 'true' ? !0 : !1,
                            PrincipalType: $('PrincipalType', n).text()
                        });
                    }
                });
                callback(results);
            }
        };
        SpSoap.addAttachment = function (base64Data, fileName, listName, listItemId, siteUrl, callback) {
            var strData = base64Data.indexOf(',') > -1 ? base64Data.split(',')[1] : base64Data;
            var action = 'http://schemas.microsoft.com/sharepoint/soap/AddAttachment';
            var packet = '<?xml version="1.0" encoding="utf-8"?>' +
                '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                '<soap:Body>' +
                '<AddAttachment xmlns="http://schemas.microsoft.com/sharepoint/soap/">' +
                '<listName>{0}</listName>' +
                '<listItemID>{1}</listItemID>' +
                '<fileName>{2}</fileName>' +
                '<attachment>{3}</attachment>' +
                '</AddAttachment>' +
                '</soap:Body>' +
                '</soap:Envelope>';
            this.executeSoapRequest(action, packet, [listName, listItemId, fileName, strData], siteUrl, callback, 'lists.asmx');
        };
        return SpSoap;
    }());
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
    }());
    Shockout.HistoryItem = HistoryItem;
    var SpItem = (function () {
        function SpItem() {
        }
        return SpItem;
    }());
    Shockout.SpItem = SpItem;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var Templates = (function () {
        function Templates() {
        }
        Templates.getFormAction = function () {
            var div = document.createElement('div');
            div.className = 'form-action no-print';
            div.innerHTML = Templates.soFormAction;
            return div;
        };
        Templates.buttonDefault = 'btn btn-sm btn-default no-print';
        Templates.calendarIcon = '<span class="glyphicon glyphicon-calendar"></span>';
        Templates.personIcon = '<span class="glyphicon glyphicon-user"></span>';
        Templates.resetButton = 'btn btn-sm btn-default no-print reset';
        Templates.timeControlsHtml = "<span class=\"glyphicon glyphicon-calendar\"></span>\n            <select class=\"form-control so-select-hours\" style=\"margin-left:1em; max-width:5em; display:inline-block;\">{0}</select><span> : </span>\n            <select class=\"form-control so-select-minutes\" style=\"width:5em; display:inline-block;\">{1}</select>\n            <select class=\"form-control so-select-tt\" style=\"margin-left:1em; max-width:5em; display:inline-block;\"><option value=\"AM\">AM</option><option value=\"PM\">PM</option></select>\n            <button class=\"btn btn-sm btn-default reset\" style=\"margin-left:1em;\">Reset</button>\n            <span class=\"error no-print\" style=\"display:none;\">Invalid Date-time</span>\n            <span class=\"so-datetime-display no-print\" style=\"margin-left:1em;\"></span>";
        Templates.getTimeControlsHtml = function () {
            var hrsOpts = [];
            for (var i = 1; i <= 12; i++) {
                hrsOpts.push('<option value="' + i + '">' + (i < 10 ? '0' + i : i) + '</option>');
            }
            var mmOpts = [];
            for (var i = 0; i < 60; i++) {
                mmOpts.push('<option value="' + i + '">' + (i < 10 ? '0' + i : i) + '</option>');
            }
            return Templates.timeControlsHtml.replace('{0}', hrsOpts.join('')).replace('{1}', mmOpts.join(''));
        };
        Templates.soFormAction = "<div class=\"row\">\n            <div class=\"col-sm-8 col-sm-offset-4 text-right\">\n                <label>Logged in as:</label><span data-bind=\"text: currentUser().title\" class=\"current-user\"></span>\n                <button class=\"btn btn-default cancel\" data-bind=\"event: { click: cancel }\" title=\"Close\"><span class=\"glyphicon glyphicon-remove\"></span><span class=\"hidden-xs\">Close</span></button>\n                <!-- ko if: allowPrint() -->\n                    <button class=\"btn btn-primary print\" data-bind=\"event: {click: print}\" title=\"Print\"><span class=\"glyphicon glyphicon-print\"></span><span class=\"hidden-xs\">Print</span></button>\n                <!-- /ko -->\n                <!-- ko if: allowDelete() && Id() != null -->\n                    <button class=\"btn btn-warning delete\" data-bind=\"event: {click: deleteItem}\" title=\"Delete\"><span class=\"glyphicon glyphicon-remove\"></span><span class=\"hidden-xs\">Delete</span></button>\n                <!-- /ko -->\n                <!-- ko if: allowSave() -->\n                    <button class=\"btn btn-success save\" data-bind=\"event: { click: save }\" title=\"Save your work.\"><span class=\"glyphicon glyphicon-floppy-disk\"></span><span class=\"hidden-xs\">Save</span></button>\n                <!-- /ko -->\n                <button class=\"btn btn-danger submit\" data-bind=\"event: { click: submit }\" title=\"Submit for routing.\"><span class=\"glyphicon glyphicon-floppy-open\"></span><span class=\"hidden-xs\">Submit</span></button>\n            </div>\n        </div>";
        Templates.hasErrorCssDiv = '<div class="form-group" data-bind="css: {\'has-error\': !!!modelValue() && !!required(), \'has-success has-feedback\': !!modelValue() && !!required()}">';
        Templates.requiredFeedbackSpan = '<span class="glyphicon glyphicon-ok form-control-feedback" aria-hidden="true"></span>';
        Templates.soNavMenuControl = '<so-nav-menu params="val: navMenuItems"></so-nav-menu>';
        Templates.soNavMenu = "<nav class=\"navbar navbar-default no-print\" id=\"TOP\">\n            <div class=\"navbar-header\">\n                <button type=\"button\" class=\"navbar-toggle collapsed\" data-toggle=\"collapse\" data-target=\"#bs-example-navbar-collapse-1\" aria-expanded=\"false\">\n                <span class=\"sr-only\">Toggle navigation</span>\n                <span class=\"icon-bar\"></span>\n                <span class=\"icon-bar\"></span>\n                <span class=\"icon-bar\"></span>\n                </button>\n            </div>\n            <div class=\"collapse navbar-collapse\" id=\"bs-example-navbar-collapse-1\">\n                <ul class=\"nav navbar-nav\" data-bind=\"foreach: navMenuItems\">\n                    <li><a href=\"#\" data-bind=\"html: title, attr: {'href': '#' + anchorName }\"></a></li>\n                </ul>\n            </div>\n        </nav>";
        Templates.soStaticField = "<div class=\"form-group\">\n            <div class=\"row\">\n                <!-- ko if: !!label -->\n                    <div class=\"col-sm-3\" data-bind=\"attr:{'class': labelColWidth}\"><label data-bind=\"html: label\"></label></div>\n                <!-- /ko -->\n                <div class=\"col-sm-9\" data-bind=\"text: modelValue, attr:{'class': fieldColWidth}\"></div>\n            </div>\n            <!-- ko if: description -->\n            <div class=\"so-field-description\"><p data-bind=\"html: description\"></p></div>\n            <!-- /ko -->\n        </div>";
        Templates.soTextField = Templates.hasErrorCssDiv + "\n        <div class=\"row\">\n\t        <!-- ko if: !!label -->\n\t\t        <div class=\"col-sm-3\" data-bind=\"attr:{'class': labelColWidth}\"><label data-bind=\"html: label, attr: {for: id}\"></label></div>\n\t        <!-- /ko -->\n\t        <div class=\"col-sm-9\" data-bind=\"attr:{'class': fieldColWidth}\">\n\t\t        <!-- ko if: readOnly() -->\n\t\t\t        <div data-bind=\"text: modelValue\"></div>\n\t\t        <!-- /ko -->\n\t\t        <!-- ko ifnot: readOnly() -->\n\t\t\t        <!-- ko if: multiline -->\n\t\t\t\t        <textarea data-bind=\"value: modelValue, css: {'so-editable': editable}, attr: {id: id, placeholder: placeholder, title: title, required: required, 'ko-name': koName }\" class=\"form-control\"></textarea>\n\t\t\t        <!-- /ko -->\n\t\t\t        <!-- ko ifnot: multiline -->\n\t\t\t\t        <input type=\"text\" data-bind=\"value: modelValue, css: {'so-editable': editable}, attr: {id: id, placeholder: placeholder, title: title, required: required, maxlength: maxlength, 'ko-name': koName }\" class=\"form-control\" />\n\t\t\t        <!-- /ko -->\n\t\t\t        <!-- ko if: !!required() -->\n\t\t\t\t        " + Templates.requiredFeedbackSpan + "\n\t\t\t        <!-- /ko -->\n\t\t        <!-- /ko -->\n\t        </div>\n\t        <!-- ko if: description -->\n\t\t    <div class=\"so-field-description\"><p data-bind=\"html: description\"></p></div>\n\t        <!-- /ko -->\n        </div>";
        Templates.soAttachments = "<section class=\"nav-section\">\n            <h4><span data-bind=\"text: title\"></span> &ndash; <span data-bind=\"text: length\" class=\"badge\"></span></h4>\n            <div data-bind=\"visible: !!errorMsg()\" class=\"alert alert-danger alert-dismissable\">\n                <button type=\"button\" class=\"close\" data-dismiss=\"alert\" aria-label=\"Close\"><span aria-hidden=\"true\">&times;</span></button>\n                <span class=\"glyphicon glyphicon-exclamation-sign\"></span>&nbsp;<span data-bind=\"text: errorMsg\"></span>\n            </div>\n            <!-- ko ifnot: hasFileReader() -->\n            <div data-bind=\"visible: !!!readOnly(), attr: {id: this.qqFileUploaderId}\"></div>\n            <!-- /ko -->\n            <!-- ko if: !readOnly() && hasFileReader() -->\n                <div class=\"row\">\n                    <div class=\"col-md-2 so-attach-files-btn\">\n                        <input type=\"file\" data-bind=\"attr: {'id': id}, event: {'change': fileHandler}\" multiple class=\"form-control\" style=\"display:none;\" />\n                        <div data-bind=\"attr:{'class': className}, event: {'click': onSelect}\"><span class=\"glyphicon glyphicon-paperclip\"></span>&nbsp;<span data-bind=\"text: label\"></span></div>\n                    </div>\n                    <div class=\"col-md-10\">\n                        <!-- ko if: drop -->\n                            <div class=\"so-file-dropzone\" data-bind=\"event: {'dragenter': onDragenter, 'dragover': onDragover, 'drop': onDrop}\">\n                                <div><span data-bind=\"text: dropLabel\"></span> <span class=\"glyphicon glyphicon-upload\"></span></div>\n                            </div>\n                        <!-- /ko -->\n                    </div>\n                </div>\n                <!-- ko foreach: fileUploads -->\n                    <div class=\"progress\">\n                        <div data-bind=\"attr: {'aria-valuenow': progress(), 'style': 'width:' + progress() + '%;', 'class': className() }\" role=\"progressbar\" aria-valuemin=\"0\" aria-valuemax=\"100\">\n                            <span data-bind=\"text: fileName\"></span>\n                        </div>\n                    </div>\n                <!-- /ko -->\n            <!-- /ko -->\n            <div data-bind=\"foreach: attachments\" style=\"margin:1em auto;\">\n                <div class=\"so-attachment\">\n                    <a href=\"\" data-bind=\"attr: {href: __metadata.media_src}\"><span class=\"glyphicon glyphicon-paperclip\"></span>&nbsp;<span data-bind=\"text: Name\"></span></a>\n                    <!-- ko ifnot: $parent.readOnly() -->\n                    <button data-bind=\"event: {click: $parent.deleteAttachment}\" class=\"btn btn-sm btn-danger delete\" title=\"Delete Attachment\"><span class=\"glyphicon glyphicon-remove\"></span></button>\n                    <!-- /ko -->\n                </div>\n            </div>\n            <!-- ko if: length() == 0 && readOnly() -->\n                <p>No attachments have been included.</p>\n            <!-- /ko -->\n            <!-- ko if: description -->\n                <div data-bind=\"text: description\"></div>\n            <!-- /ko -->\n        </section>";
        Templates.soHtmlFieldTemplate = Templates.hasErrorCssDiv + "\n        <div class=\"row\">\n            <!-- ko if: !!label -->\n                <div class=\"col-sm-3\" data-bind=\"attr:{'class': labelColWidth}\"><label data-bind=\"html: label, attr: {for: id}\"></label></div>\n            <!-- /ko -->\n            <div class=\"col-sm-9\" data-bind=\"attr:{'class': fieldColWidth}\">\n                <!-- ko if: readOnly() -->\n                    <div data-bind=\"html: modelValue\"></div>\n                <!-- /ko -->\n                <!-- ko ifnot: readOnly() -->\n                    <div data-bind=\"spHtmlEditor: modelValue\" contenteditable=\"true\" class=\"form-control content-editable\"></div>\n                    <textarea data-bind=\"value: modelValue, css: {'so-editable': editable}, attr: {id: id, required: required, 'ko-name': koName }\" data-sp-html=\"\" style=\"display:none;\"></textarea>\n                    <!-- ko if: !!required() -->\n                        " + Templates.requiredFeedbackSpan + "\n                    <!-- /ko -->\n                <!-- /ko -->\n                </div>\n            </div>\n            <!-- ko if: description -->\n                <div class=\"so-field-description\"><p data-bind=\"html: description\"></p></div>\n            <!-- /ko -->\n        </div>";
        Templates.soCheckboxField = "<div class=\"form-group\">\n            <div class=\"row\">\n                <div class=\"col-sm-3\" data-bind=\"attr:{'class': labelColWidth}\">\n                    <!-- ko if: readOnly() -->\n                    <label data-bind=\"html: label, attr: {for: id}\"></label>\n                    <!-- /ko -->\n                </div>\n                <div class=\"col-sm-9\" data-bind=\"attr:{'class': fieldColWidth}\">\n                    <!-- ko if: readOnly() -->\n                        <div data-bind=\"text: !!modelValue() ? 'Yes' : 'No'\"></div>\n                    <!-- /ko -->\n                    <!-- ko ifnot: readOnly() -->\n                        <label class=\"checkbox\">\n                            <input type=\"checkbox\" data-bind=\"checked: modelValue, css: {'so-editable': editable}, attr: {id: id, 'ko-name': koName}, valueUpdate: valueUpdate\" />\n                            <span data-bind=\"html: label\" style=\"margin-left:1em;\"></span>\n                        </label>\n                    <!-- /ko -->\n                </div>\n            </div>\n            <!-- ko if: description -->\n                <div class=\"so-field-description\"><p data-bind=\"html: description\"></p></div>\n            <!-- /ko -->\n        </div>";
        Templates.soSelectField = Templates.hasErrorCssDiv + "<div class=\"row\">\n                <!-- ko if: !!label -->\n                    <div class=\"col-sm-3\" data-bind=\"attr:{'class': labelColWidth}\"><label data-bind=\"html: label, attr: {for: id}\"></label></div>\n                <!-- /ko -->\n                <div class=\"col-sm-9\" data-bind=\"attr:{'class': fieldColWidth}\">\n                    <!-- ko if: readOnly() -->\n                        <div data-bind=\"text: modelValue\"></div>\n                    <!-- /ko -->\n                    <!-- ko ifnot: readOnly() -->\n                        <select data-bind=\"value: modelValue, options: options, optionsCaption: caption, css: {'so-editable': editable}, attr: {id: id, title: title, required: required, 'ko-name': koName}\" class=\"form-control\"></select>\n                        <!-- ko if: !!required() -->\n                            " + Templates.requiredFeedbackSpan + "\n                        <!-- /ko -->\n                    <!-- /ko -->\n                </div>\n            </div>\n            <!-- ko if: description -->\n                <div class=\"so-field-description\"><p data-bind=\"html: description\"></p></div>\n            <!-- /ko -->\n        </div>";
        Templates.soCheckboxGroup = "<div class=\"form-group\">\n            <!-- ko if: description -->\n\t            <div class=\"so-field-description\"><p data-bind=\"html: description\"></p></div>\n            <!-- /ko -->\n            <div class=\"row\">\n\t            <!-- ko if: !!label -->\n\t\t            <div><label data-bind=\"html: label\"></label></div>\n\t            <!-- /ko -->\n\t            <div>\n\t\t            <!-- ko if: readOnly() -->\n\t\t\t            <!-- ko ifnot: inline -->\n\t\t\t\t            <ul class=\"list-group\">\n\t\t\t\t\t            <!-- ko foreach: modelValue -->\n\t\t\t\t\t\t            <li data-bind=\"text: $data\" class=\"list-group-item\"></li>\n\t\t\t\t\t            <!-- /ko -->\n\t\t\t\t\t            <!-- ko if: modelValue().length == 0 -->\n\t\t\t\t\t\t            <li class=\"list-group-item\">--None--</li>\n\t\t\t\t\t            <!-- /ko -->\n\t\t\t\t            </ul>\n\t\t\t            <!-- /ko -->\n\t\t\t            <!-- ko if: inline -->\n\t\t\t\t            <!-- ko foreach: modelValue -->\n\t\t\t\t\t            <span data-bind=\"text: $data\"></span>\n\t\t\t\t\t            <!-- ko if: $index() < $parent.modelValue().length-1 -->,&nbsp;<!-- /ko -->\n\t\t\t\t            <!-- /ko -->\n\t\t\t\t            <!-- ko if: modelValue().length == 0 -->\n\t\t\t\t\t            <span>--None--</span>\n\t\t\t\t            <!-- /ko -->\n\t\t\t            <!-- /ko -->\n\t\t            <!-- /ko -->\n\t\t            <!-- ko ifnot: readOnly() -->\n\t\t\t            <input type=\"hidden\" data-bind=\"value: modelValue, attr:{required: !!required}\" /><p data-bind=\"visible: !!required\" class=\"req\">(Required)</p>\n\t\t\t            <!-- ko foreach: options -->\n\t\t\t\t            <label data-bind=\"css:{'checkbox': !$parent.inline, 'checkbox-inline': $parent.inline}\">\n\t\t\t\t\t            <input type=\"checkbox\" data-bind=\"checked: $parent.modelValue, css: {'so-editable': $parent.editable}, attr: {'ko-name': $parent.koName, 'value': $data}\" />\n\t\t\t\t\t            <span data-bind=\"text: $data\"></span>\n\t\t\t\t            </label>\n\t\t\t            <!-- /ko -->\n\t\t            <!-- /ko -->\n\t            </div>\n            </div>";
        Templates.soRadioGroup = "<div class=\"form-group\">\n\t        <!-- ko if: description -->\n\t\t        <div class=\"so-field-description\"><p data-bind=\"html: description\"></p></div>\n\t        <!-- /ko -->\n\t        <div class=\"row\">\n\t\t        <!-- ko if: !!label -->\n\t\t\t        <div class=\"col-sm-3\" data-bind=\"attr:{'class': labelColWidth}\"><label data-bind=\"html: label\"></label></div>\n\t\t        <!-- /ko -->\n\t\t        <div class=\"col-sm-9\" data-bind=\"attr:{'class': fieldColWidth}\">\n\t\t\t        <!-- ko if: readOnly() -->\n\t\t\t\t        <div data-bind=\"text: modelValue\"></div>\n\t\t\t        <!-- /ko -->\n\t\t\t        <!-- ko ifnot: readOnly() -->\n\t\t\t\t        <!-- ko foreach: options -->\n\t\t\t\t\t        <label data-bind=\"css:{'radio': !$parent.inline, 'radio-inline': $parent.inline}\">\n\t\t\t\t\t\t        <input type=\"radio\" data-bind=\"checked: $parent.modelValue, attr:{value: $data, name: $parent.name, 'ko-name': $parent.koName}, css:{'so-editable': $parent.editable}\" />\n\t\t\t\t\t\t        <span data-bind=\"text: $data\"></span>\n\t\t\t\t\t        </label>\n\t\t\t\t        <!-- /ko -->\n\t\t\t        <!-- /ko -->\n\t\t        </div>\n\t        </div>\n        </div>";
        Templates.soUsermultiField = "<div class=\"form-group\">\n\t        <input type=\"hidden\" data-bind=\"value: modelValue, css: {'so-editable': editable}, attr: {id: id, 'ko-name': koName, required: required}\" />\n\t        <div class=\"row\">\n\t\t        <div class=\"col-md-3 col-xs-3\">\n\t\t\t        <label data-bind=\"html: label\"></label>\n\t\t        </div>\n\t\t        <div class=\"col-md-9 col-xs-9\">\n\t\t\t        <!-- ko ifnot: readOnly -->\n\t\t\t\t        <input type=\"text\" data-bind=\"spPerson: person, attr: {placeholder: placeholder}\" />\n\t\t\t\t        <button class=\"btn btn-success\" data-bind=\"click: addPerson, attr: {'disabled': person() == null, id: koName + '_AddButton' }\"><span>Add</span></button>\n\t\t\t\t        <!-- ko if: showRequiredText -->\n\t\t\t\t\t        <div class=\"col-md-6 col-xs-6\">\n\t\t\t\t\t\t        <p class=\"text-danger\">At least one person must be added.</p>\n\t\t\t\t\t        </div>\n\t\t\t\t        <!-- /ko -->\n\t\t\t        <!-- /ko -->\n\t\t\t        <!-- ko foreach: modelValue -->\n\t\t\t\t        <div class=\"row\">\n\t\t\t\t\t        <div class=\"col-md-10 col-xs-10\" data-bind=\"spPerson: $data\"></div>\n\t\t\t\t\t        <!-- ko ifnot: $parent.readOnly() -->\n\t\t\t\t\t\t        <div class=\"col-md-2 col-xs-2\">\n\t\t\t\t\t\t\t        <button class=\"btn btn-xs btn-danger\" data-bind=\"click: $parent.removePerson\"><span class=\"glyphicon glyphicon-trash\"></span></button>\n\t\t\t\t\t\t        </div>\n\t\t\t\t\t        <!-- /ko -->\n\t\t\t\t        </div>\n\t\t\t        <!-- /ko -->\n\t\t\t        <!-- ko if: description -->\n\t\t\t\t        <div class=\"so-field-description\"><p data-bind=\"html: description\"></p></div>\n\t\t\t        <!-- /ko -->\n\t\t        </div>\n\t        </div>\n        </div>";
        Templates.soWorkflowHistoryControl = "<!-- ko if: !!Id() && historyItems().length > 0 -->\n            <section id=\"workflowHistory\" class=\"nav-section\">\n                <h4>Workflow History</h4>\n                <so-workflow-history params=\"val: historyItems\"></so-workflow-history>\n            </section>\n        <!-- /ko -->";
        Templates.soWorkflowHistory = "<div class=\"row\">\n            <div class=\"col-sm-8\"><strong>Description</strong></div>\n            <div class=\"col-sm-4\"><strong>Date</strong></div>\n        </div>\n        <!-- ko foreach: historyItems -->\n            <div class=\"row\">\n                <div class=\"col-sm-8\"><span data-bind=\"text: _description\"></span></div>\n                <div class=\"col-sm-4\"><span data-bind=\"spDateTime: _dateOccurred\"></span></div>\n            </div>\n        <!-- /ko -->";
        Templates.soCreatedModifiedInfoControl = "<!-- ko if: !!CreatedBy && CreatedBy() != null -->\n            <section class=\"nav-section\">\n                <h4>Created/Modified By</h4>\n                <so-created-modified-info params=\"created: Created, createdBy: CreatedBy, modified: Modified, modifiedBy: ModifiedBy, showUserProfiles: showUserProfiles\"></so-created-modified-info>\n            </section>\n        <!-- /ko -->";
        Templates.soCreatedModifiedInfo = "<!-- ko if: showUserProfiles() -->\n            <div class=\"create-mod-info no-print hidden-xs\">\n                <!-- ko foreach: profiles -->\n                    <div class=\"user-profile-card\">\n                        <h4 data-bind=\"text: header\"></h4>\n                        <!-- ko with: profile -->\n                            <img data-bind=\"attr: {src: Picture, alt: Name}\" />\n                            <ul>\n                                <li><label>Name</label><span data-bind=\"text: Name\"></span><li>\n                                <li><label>Job Title</label><span data-bind=\"text: JobTitle\"></span></li>\n                                <li><label>Department</label><span data-bind=\"text: Department\"></span></li>\n                                <li><label>Email</label><a data-bind=\"text: WorkEMail, attr: {href: ('mailto:' + WorkEMail)}\"></a></li>\n                                <li><label>Phone</label><span data-bind=\"text: WorkPhone\"></span></li>\n                                <li><label>Office</label><span data-bind=\"text: Office\"></span></li>\n                            </ul>\n                        <!-- /ko -->\n                    </div>\n                <!-- /ko -->\n            </div>\n        <!-- /ko -->\n        <div class=\"row\">\n            <!-- ko with: CreatedBy -->\n                <div class=\"col-md-3\"><label>Created By</label> <a data-bind=\"text: Name, attr: {href: 'mailto:' + WorkEMail}\" class=\"email\"> </a></div>\n            <!-- /ko -->\n            <div class=\"col-md-3\"><label>Created</label> <span data-bind=\"spDateTime: Created\"></span></div>\n            <!-- ko with: ModifiedBy -->\n                <div class=\"col-md-3\"><label>Modified By</label> <a data-bind=\"text: Name, attr: {href: 'mailto:' + WorkEMail}\" class=\"email\"></a></div>\n            <!-- /ko -->\n            <div class=\"col-md-3\"><label>Modified</label> <span data-bind=\"spDateTime: Modified\"></span></div>\n        </div>";
        return Templates;
    }());
    Shockout.Templates = Templates;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var Utils = (function () {
        function Utils() {
        }
        Utils.indexOf = function (a, value) {
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
        Utils.formatSubsiteUrl = function (url) {
            return !!!url ? '/' : !/\/$/.test(url) ? url + '/' : url;
        };
        Utils.toCamelCase = function (str) {
            return str.toString()
                .replace(/\s*\b\w/g, function (x) {
                return (x[1] || x[0]).toUpperCase();
            }).replace(/\s/g, '')
                .replace(/\'s/, 'S')
                .replace(/[^A-Za-z0-9\s_]/g, '');
        };
        Utils.getIdFromHash = function () {
            var rxHash = /\/id\/\d+/i;
            var exec = rxHash.exec(window.location.hash);
            var id = !!exec ? exec[0].replace(/\D/g, '') : null;
            return /\d/.test(id) ? parseInt(id) : null;
        };
        Utils.setIdHash = function (id) {
            window.location.hash = '#/id/' + id;
        };
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
            } while (o && o.nodeType != 8 && !/^\s*ko/.test(o.textContent));
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
                var db = next.getAttribute('data-bind');
                if (!!db && rxTagNames.test(next.tagName) && rxIsContext.test(db) && rxNotTypes.test(next.getAttribute('type'))) {
                    a.push(comments[i]);
                    continue;
                }
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
            var comments = Utils.getEditableKoContainerlessControls(parent);
            for (var i = 0; i < comments.length; i++) {
                var n = comments[i].nodeValue
                    .replace(/\s*ko\s*foreach\s*:\s*(\$root\.|)/, '')
                    .replace(/\s/g, '');
                if (Utils.indexOf(a, n) < 0) {
                    a.push(n);
                }
            }
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
            if (!!!val) {
                return false;
            }
            return /\/Date\(\d+\)\//.test(val + '');
        };
        Utils.isIsoDateString = function (val) {
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
        Utils.clone = function (objectToBeCloned) {
            if (!(objectToBeCloned instanceof Object)) {
                return objectToBeCloned;
            }
            var objectClone;
            var Constructor = objectToBeCloned.constructor;
            switch (Constructor) {
                case RegExp:
                    objectClone = new Constructor(objectToBeCloned);
                    break;
                case Date:
                    objectClone = new Constructor(objectToBeCloned.getTime());
                    break;
                default:
                    objectClone = new Constructor();
            }
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
        Utils.updateKoField = function (el, val) {
            if (el.tagName.toLowerCase() == "input") {
                $(el).val(val);
            }
            else {
                $(el).html(val);
            }
        };
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
        Utils.parseDate = function (val) {
            if (!!!val) {
                return null;
            }
            if (typeof val == 'object' && val.constructor == Date) {
                return val;
            }
            var rxSlash = /\d{1,2}\/\d{1,2}\/\d{2,4}/, rxHyphen = /\d{1,2}-\d{1,2}-\d{2,4}/, rxIsoDate = /\d{4}-\d{1,2}-\d{1,2}/, rxTicks = /(\/|)\d{13}(\/|)/, rxIsoDateTime = /\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z/, tmp, m, d, y, date = null;
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
        Utils.formatMoney = function (value, symbol, precision) {
            if (symbol === void 0) { symbol = '$'; }
            if (precision === void 0) { precision = 2; }
            var num = Utils.unformatNumber(value), format = '%s%v', neg = format.replace('%v', '-%v'), useFormat = num > 0 ? format : num < 0 ? neg : format, numFormat = Utils.formatNumber(Math.abs(num), Utils.checkPrecision(precision));
            return useFormat
                .replace('%s', symbol)
                .replace('%v', numFormat);
        };
        Utils.unformatNumber = function (value) {
            if (typeof value === "number")
                return value;
            var unformatted = parseFloat((value + '')
                .replace(/\((.*)\)/, '-$1')
                .replace(/[^0-9-.]/g, ''));
            return !isNaN(unformatted) ? unformatted : 0;
        };
        Utils.formatNumber = function (value, precision) {
            if (precision === void 0) { precision = 0; }
            var num = Utils.unformatNumber(value), usePrecision = Utils.checkPrecision(precision), negative = num < 0 ? "-" : "", base = parseInt(Utils.toFixed(Math.abs(num || 0), usePrecision), 10) + "", mod = base.length > 3 ? base.length % 3 : 0;
            return negative + (mod ? base.substr(0, mod) + ',' : '') + base.substr(mod).replace(/(\d{3})(?=\d)/g, '$1,') + (usePrecision ? '.' + Utils.toFixed(Math.abs(num), usePrecision).split('.')[1] : "");
        };
        Utils.isString = function (obj) {
            return !!(obj === '' || (obj && obj.charCodeAt && obj.substr));
        };
        Utils.toFixed = function (value, precision) {
            if (precision === void 0) { precision = 0; }
            precision = Utils.checkPrecision(precision);
            var power = Math.pow(10, precision);
            return (Math.round(Utils.unformatNumber(value) * power) / power).toFixed(precision);
        };
        Utils.checkPrecision = function (val) {
            val = Math.round(Math.abs(val));
            return isNaN(val) ? 0 : val;
        };
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
        Utils.isZrow = function (node) {
            return node.nodeName.toLowerCase() == 'z:row';
        };
        Utils.koNameFromControl = Utils.observableNameFromControl;
        return Utils;
    }());
    Shockout.Utils = Utils;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
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
    }());
    Shockout.SpAttachment = SpAttachment;
    var Cafe = (function () {
        function Cafe(asyncFns) {
            if (asyncFns === void 0) { asyncFns = undefined; }
            if (asyncFns) {
                this.asyncFns = asyncFns;
            }
            return this;
        }
        Cafe.prototype.start = function (msg) {
            if (msg === void 0) { msg = undefined; }
            this.next(true, msg);
        };
        Cafe.prototype.complete = function (fn) {
            this._complete = fn;
            return this;
        };
        ;
        Cafe.prototype.fail = function (fn) {
            this._fail = fn;
            return this;
        };
        Cafe.prototype.finally = function (fn) {
            this._finally = fn;
            return this;
        };
        Cafe.prototype.next = function (success, msg, args) {
            if (success === void 0) { success = true; }
            if (msg === void 0) { msg = undefined; }
            if (args === void 0) { args = undefined; }
            if (!this.asyncFns) {
                throw "Error in Cafe: The required parameter `asyncFns` of type (Array<Function>) is undefined. Don't forget to instantiate Cafe with this parameter or set its value after instantiation.";
            }
            if (this._complete) {
                this._complete(msg, success, args);
            }
            if (!success) {
                if (this._fail) {
                    this._fail(msg, success, args);
                }
                return;
            }
            if (this.asyncFns.length == 0) {
                if (this._finally) {
                    this._finally(msg, success, args);
                }
                return;
            }
            this.asyncFns.shift()(this, args);
        };
        return Cafe;
    }());
    Shockout.Cafe = Cafe;
    var FileUpload = (function () {
        function FileUpload(fileName, bytes) {
            var self = this;
            this.label = ko.observable(null);
            this.progress = ko.observable(0);
            this.fileName = ko.observable(fileName);
            this.kb = ko.observable((bytes / 1024));
            this.className = ko.observable('progress-bar progress-bar-info progress-bar-striped active');
            this.getProgress = ko.pureComputed(function () {
                return self.fileName() + ' ' + self.progress() + '%';
            }, this);
        }
        return FileUpload;
    }());
    Shockout.FileUpload = FileUpload;
    var DateTimeModel = (function () {
        function DateTimeModel(element, modelValue) {
            var self = this;
            var date = ko.unwrap(modelValue);
            this.$element = $(element);
            this.$parent = this.$element.parent();
            if (Shockout.Utils.isJsonDateTicks(date)) {
                modelValue(Shockout.Utils.parseJsonDate(date));
            }
            this.required = this.$element.hasClass('required') || this.$element.attr('required') != null;
            this.$element.attr({
                'placeholder': 'MM/DD/YYYY',
                'maxlength': 10,
                'class': 'datepicker med form-control'
            }).css('display', 'inline-block')
                .on('change', onChange)
                .datepicker({
                changeMonth: true,
                changeYear: true
            });
            if (this.required) {
                this.$element.attr('required', '');
            }
            this.$element.after(Shockout.Templates.getTimeControlsHtml());
            this.$display = this.$parent.find('.so-datetime-display');
            this.$error = this.$parent.find('.error');
            this.$hh = this.$parent.find('.so-select-hours').val('12').on('change', onChange);
            this.$mm = this.$parent.find('.so-select-minutes').val('0').on('change', onChange);
            this.$tt = this.$parent.find('.so-select-tt').val('AM').on('change', onChange);
            this.$reset = this.$parent.find('.btn.reset')
                .on('click', function () {
                try {
                    self.$element.val('');
                    self.$hh.val('12');
                    self.$mm.val('0');
                    self.$tt.val('AM');
                    self.$display.html('');
                    modelValue(null);
                }
                catch (e) {
                    console.warn(e);
                }
                return false;
            });
            function onChange() {
                self.setModelValue(modelValue);
            }
        }
        DateTimeModel.prototype.setModelValue = function (modelValue) {
            try {
                var val = this.$element.val();
                if (!!!$.trim(this.$element.val())) {
                    return;
                }
                var hrs = parseInt(this.$hh.val());
                var min = parseInt(this.$mm.val());
                var tt = this.$tt.val();
                if (tt == 'PM' && hrs < 12) {
                    hrs += 12;
                }
                else if (tt == 'AM' && hrs > 11) {
                    hrs -= 12;
                }
                var curDateTime = new Date(val);
                curDateTime.setUTCHours(hrs, min, 0, 0);
                modelValue(curDateTime);
            }
            catch (e) {
                if (Shockout.SPForm.DEBUG) {
                    throw e;
                }
            }
        };
        DateTimeModel.prototype.setDisplayValue = function (modelValue) {
            try {
                var date = ko.unwrap(modelValue);
                if (!!!date || date.constructor != Date) {
                    return;
                }
                var date = ko.unwrap(modelValue);
                var hrs = date.getUTCHours();
                var min = date.getUTCMinutes();
                this.$element.val((date.getUTCMonth() + 1) + '/' + date.getUTCDate() + '/' + date.getUTCFullYear());
                if (hrs > 12) {
                    hrs -= 12;
                    this.$tt.val('PM');
                }
                else if (hrs == 0) {
                    hrs = 12;
                    this.$tt.val('AM');
                }
                else if (hrs == 12) {
                    this.$tt.val('PM');
                }
                else {
                    this.$tt.val('AM');
                }
                this.$hh.val(hrs.toString());
                this.$mm.val(min.toString());
                this.$display.html(this.toString(modelValue));
            }
            catch (e) {
                if (Shockout.SPForm.DEBUG) {
                    throw e;
                }
            }
        };
        DateTimeModel.prototype.toString = function (modelValue) {
            return DateTimeModel.toString(modelValue);
        };
        DateTimeModel.toString = function (modelValue) {
            try {
                var date = ko.unwrap(modelValue);
                if (date == null || date.constructor != Date) {
                    return;
                }
                var dateTimeStr = Shockout.Utils.toDateTimeLocaleString(date);
                var timeZone = /\b\s\(\w+\s\w+\s\w+\)/i.exec(date.toString());
                if (!!timeZone) {
                    dateTimeStr += ' ' + timeZone[0].replace(/\b\w+/g, function (x) {
                        return x[0];
                    }).replace(/\s/g, '');
                }
                return dateTimeStr;
            }
            catch (e) {
                if (Shockout.SPForm.DEBUG) {
                    throw e;
                }
            }
            return null;
        };
        return DateTimeModel;
    }());
    Shockout.DateTimeModel = DateTimeModel;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var ClassNameDict = (function () {
        function ClassNameDict() {
        }
        return ClassNameDict;
    }());
    Shockout.ClassNameDict = ClassNameDict;
})(Shockout || (Shockout = {}));
