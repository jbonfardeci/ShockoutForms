/// <reference path="../typings/knockout.d.ts" />
/// <reference path="../typings/jquery.d.ts" />
/// <reference path="../typings/jquery.ui.datetimepicker.d.ts" />
/// <reference path="../typings/jqueryui.d.ts" />
'use strict';
var Shockout;
(function (Shockout) {
    var SPForm = (function () {
        function SPForm(listName, formId, options) {
            // public options
            this.allowDelete = false;
            this.allowPrint = true;
            this.allowSave = false;
            this.allowedExtensions = ['txt', 'rtf', 'zip', 'pdf', 'doc', 'docx', 'jpg', 'gif', 'png', 'ppt', 'tif', 'pptx', 'csv', 'pub', 'msg'];
            this.attachmentMessage = 'An attachment is required.';
            this.confirmationUrl = '/SitePages/Confirmation.aspx';
            this.debug = false;
            this.editableFields = [];
            this.enableErrorLog = true;
            this.errorLogListName = 'Error Log';
            this.fileHandlerUrl = '/_layouts/webster/SPFormFileHandler.ashx';
            this.hasAttachments = true;
            this.includeUserProfiles = true;
            this.includeWorkflowHistory = true;
            this.requireAttachments = false;
            this.rootUrl = window.location.protocol + '//' + window.location.hostname + (!!window.location.port ? ':' + window.location.port : '');
            this.siteUrl = '';
            this.viewModelIsBound = false;
            this.utils = Shockout.Utils;
            this.workflowHistoryListName = 'Workflow History';
            this.enableAttachments = false;
            this.requireCheckout = false;
            this.version = '0.0.1';
            var self = this;
            var error;
            if (!(this instanceof SPForm)) {
                error = 'You must declare an instance of this class with `new`.';
                alert(error);
                throw error;
                return;
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
                return;
            }
            this.formId = formId;
            this.listName = listName;
            if (!!Shockout.Utils.getQueryParam("id")) {
                this.itemId = parseInt(Shockout.Utils.getQueryParam("id"));
            }
            else if (!!Shockout.Utils.getQueryParam("formid")) {
                this.itemId = parseInt(Shockout.Utils.getQueryParam("formid"));
            }
            this.sourceUrl = Shockout.Utils.getQueryParam("source"); //if accessing the form from a SP list, take user back to the list on close
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
            this.$form = $(this.form);
            self.$formStatus = $('<div>', { 'class': 'form-status' }).appendTo(self.$form);
            self.$dialog = $('<div>', { 'id': 'formdialog' })
                .appendTo(self.$form)
                .dialog({
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
            this.viewModel = new Shockout.ViewModel(this);
            //Cascading Asynchronous Function Execution (CAFE) Array
            this.asyncFns = [
                self.getListAsync,
                function () {
                    if (self.preRender) {
                        self.preRender(self);
                    }
                    self.nextAsync(true);
                },
                self.getCurrentUserAsync,
                self.getUsersGroupsAsync,
                self.restrictSpGroupElementsAsync,
                self.initFormAsync,
                self.getListItemAsync,
                self.getAttachmentsAsync,
                self.getHistoryAsync,
                function () {
                    if (self.postRender) {
                        self.postRender(self);
                    }
                    self.nextAsync(true);
                }
            ];
            //start CAFE
            this.nextAsync(true, 'Begin initialization...');
        }
        SPForm.prototype.nextAsync = function (success, msg, args) {
            if (success === void 0) { success = undefined; }
            if (msg === void 0) { msg = undefined; }
            if (args === void 0) { args = undefined; }
            var self = this;
            success = success || true;
            if (msg) {
                this.updateStatus(msg, success);
            }
            if (!success) {
                return;
            }
            if (this.asyncFns.length == 0) {
                setTimeout(function () {
                    self.$formStatus.slideDown();
                }, 2000);
                return;
            }
            // execute the next function in the array
            this.asyncFns.shift()(self, args);
        };
        SPForm.prototype.initFormAsync = function (self, args) {
            if (args === void 0) { args = undefined; }
            try {
                self.updateStatus("Initializing dynamic form features...");
                self.$createdInfo = self.$form.find(".created-info");
                // append action buttons
                self.$formAction = $(Shockout.Templates.getFormAction(self.allowSave, self.allowDelete, self.allowPrint)).appendTo(self.$form);
                //append Created/Modified info to predefined section or append to form
                if (!!self.itemId) {
                    self.$createdInfo.html(Shockout.Templates.getCreatedModifiedInfo().innerHTML);
                    //append Workflow history section
                    if (self.includeWorkflowHistory) {
                        self.$form.append(Shockout.Templates.getHistoryTemplate());
                    }
                }
                if (self.editableFields.length == 0) {
                    //make array of SP field names and those that are editable from elements w/ data-bind attribute
                    self.$form.find('[data-bind]').each(function (i, e) {
                        var key = Shockout.Utils.observableNameFromControl(e);
                        //skip observable keys that have already been added or begins with an underscore '_' or dollar sign '$'
                        if (!!!key || self.editableFields.indexOf(key) > -1 || key.match(/^(_|\$)/) != null) {
                            return;
                        }
                        if (e.tagName == 'INPUT' || e.tagName == 'SELECT' || e.tagName == 'TEXTAREA' || $(e).attr('contenteditable') == 'true') {
                            self.editableFields.push(key);
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
                    onSubmit: function (id, fileName) { },
                    onComplete: function (id, fileName, json) {
                        if (self.itemId == null) {
                            self.viewModel['Id'](json.itemId);
                            self.itemId = json.itemId;
                            self.saveListItem(self.viewModel, false);
                        }
                        if (json.error == null && json.fileName != null) {
                            self.getAttachmentsAsync();
                        }
                    },
                    template: Shockout.Templates.getFileUploadTemplate()
                };
                //setup attachments module
                self.$form.find(".attachments").each(function (i, att) {
                    var id = 'fileuploader_' + i;
                    $(att).append(Shockout.Templates.getAttachmentsTemplate(id));
                    self.fileUploaderSettings.element = document.getElementById(id);
                    self.fileUploader = new Shockout.qq.FileUploader(self.fileUploaderSettings);
                });
                self.$form.find('[required]').addClass('required');
                self.nextAsync(true, "Form initialized.");
            }
            catch (e) {
                if (self.debug) {
                    console.warn(e);
                }
                self.logError("initFormAsync: " + e);
                self.nextAsync(false, "Failed to initialize form. " + e);
            }
        };
        SPForm.prototype.getCurrentUserAsync = function (self, args) {
            if (args === void 0) { args = undefined; }
            try {
                var currentUser;
                var query = '<Query><Where><Eq><FieldRef Name="ID" /><Value Type="Counter"><UserID /></Value></Eq></Where></Query>';
                var viewFields = '<ViewFields><FieldRef Name="ID" /><FieldRef Name="Name" /><FieldRef Name="EMail" /><FieldRef Name="Department" /><FieldRef Name="JobTitle" /><FieldRef Name="UserName" /><FieldRef Name="Office" /></ViewFields>';
                self.getListItemsSoap('', 'User Information List', viewFields, query, function (xData, Sstatus) {
                    var user = {
                        id: null,
                        title: null,
                        login: null,
                        email: null,
                        account: null,
                        jobtitle: null,
                        department: null,
                        groups: []
                    };
                    $(xData.responseXML).find('*').filter(function () {
                        return this.nodeName === 'z:row';
                    }).each(function (i, node) {
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
                });
            }
            catch (e) {
                if (self.debug) {
                    console.warn(e);
                }
                self.logError('getCurrentUserAsync():' + e);
                self.nextAsync(false, 'Failed to retrieve your account.');
            }
        };
        SPForm.prototype.getUsersGroupsAsync = function (self, args) {
            if (args === void 0) { args = undefined; }
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
                var $jqXhr = $.ajax({
                    url: self.rootUrl + self.siteUrl + '/_vti_bin/usergroup.asmx',
                    type: 'POST',
                    dataType: 'xml',
                    data: packet,
                    contentType: 'text/xml; charset="utf-8"'
                });
                $jqXhr.done(function (doc, statusText, response) {
                    $(response.responseXML).find("Group").each(function (i, el) {
                        self.currentUser.groups.push({
                            id: parseInt($(el).attr("ID")),
                            name: $(el).attr("Name")
                        });
                    });
                    self.nextAsync(true, "Retrieved your groups.");
                });
                $jqXhr.fail(function (xData, status) {
                    var msg = "Failed to retrieve your groups: " + status;
                    self.logError(msg);
                    self.nextAsync(false, msg);
                });
                self.updateStatus("Retrieving your groups...");
            }
            catch (e) {
                if (self.debug) {
                    console.warn(e);
                }
                self.logError("getUsersGroupsAsync: " + e);
                self.nextAsync(false, "Failed to retrieve your groups.");
            }
        };
        SPForm.prototype.restrictSpGroupElementsAsync = function (self, args) {
            if (args === void 0) { args = undefined; }
            try {
                self.updateStatus("Retrieving your permissions...");
                self.$form.find("[user-groups]").each(function (i, el) {
                    var groups = $(el).attr("user-groups");
                    var groupNames = groups.match(/\,/) != null ? groups.split(',') : [groups];
                    var ct = 0;
                    $.each(groupNames, function (i, group) {
                        group = group.match(/\;#/) != null ? group.split(';')[0] : group; //either id;#groupname or groupname
                        group = $.trim(group);
                        $.each(self.currentUser.groups, function (j, g) {
                            if (group == g.name || parseInt(group) == g.id) {
                                ct++;
                            }
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
            }
            catch (e) {
                if (self.debug) {
                    console.warn(e);
                }
                self.logError("restrictSpGroupElementsAsync: " + e);
                self.nextAsync(true, "Failed to retrieve your permissions.");
            }
        };
        SPForm.prototype.getListItemAsync = function (self, args) {
            if (args === void 0) { args = undefined; }
            try {
                self.updateStatus("Retrieving form values...");
                if (!!!self.itemId) {
                    self.nextAsync(true, "This is a New form.");
                    return;
                }
                var uri = self.rootUrl + self.siteUrl + '/_vti_bin/listdata.svc/' + self.listName.replace(/\s/g, '') + '(' + self.itemId + ')';
                // get the list item data
                self.getListItemsRest(uri, bind, fail);
            }
            catch (e) {
                if (self.debug) {
                    console.warn(e);
                }
            }
            function bind(data, status, jqXhr) {
                self.listItem = Shockout.Utils.clone(data.d); //store copy of the original SharePoint list item
                self.bindListItemValues(self);
                self.nextAsync(true, "Retrieved form data.");
            }
            ;
            function fail(obj, status, jqXhr) {
                if (obj.status && obj.status == '404') {
                    var msg = obj.statusText + ". The form may have been deleted by another user.";
                }
                else {
                    msg = status + ' ' + jqXhr;
                }
                self.showDialog(msg);
                self.nextAsync(false, msg);
            }
            ;
        };
        SPForm.prototype.getHistoryAsync = function (self, args) {
            if (args === void 0) { args = undefined; }
            try {
                if (!!!self.itemId || !self.includeWorkflowHistory) {
                    self.nextAsync(true);
                    return;
                }
                var historyItems = [];
                var uri = self.rootUrl + self.siteUrl + "/_vti_bin/listdata.svc/" + self.workflowHistoryListName.replace(/\s/g, '') +
                    "?$filter=ListID eq '" + self.listId + "' and PrimaryItemID eq " + self.itemId + "&$select=Description,DateOccurred&$orderby=DateOccurred asc";
                self.getListItemsRest(uri, function (data, status, jqXhr) {
                    $(data.d.results).each(function (i, item) {
                        historyItems.push(new Shockout.HistoryItem(item.Description, Shockout.Utils.parseJsonDate(item.DateOccurred)));
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
        };
        SPForm.prototype.bindListItemValues = function (self) {
            if (self === void 0) { self = undefined; }
            self = self || this;
            // Exclude these read-only metadata fields from the Knockout view model.
            var rxExclude = new RegExp("^(__metadata|ContentTypeID|ContentType|CreatedBy|ModifiedBy|Owshiddenversion|Version|Attachments|Path)");
            var item = self.listItem;
            for (var key in item) {
                if (rxExclude.test(key) || !!self.viewModel[key]) {
                    continue;
                }
                // Object types will have a corresponding key name plus the suffix `Value` or `Id` for lookups.
                // For example: `SupervisorApproval` is an object container for `__deferred` that corresponds to `SupervisorApprovalValue` 
                // which is an ID or string value.
                if (item[key] != null && item[key].constructor === Object && item[key]['__deferred']) {
                    if (item[key + 'Value']) {
                        self.viewModel[key] = ko.observable(item[key + 'Value']);
                    }
                    else if (item[key + 'Id']) {
                        self.viewModel[key] = ko.observable(item[key + 'Id']);
                    }
                }
                else if (item[key] != null && Shockout.Utils.isJsonDate(item[key])) {
                    // parse JSON dates
                    self.viewModel[key] = ko.observable(Shockout.Utils.parseJsonDate(item[key]));
                }
                else {
                    // if there is a boolean field for storing the state of a form's submission status 
                    if (/submitted/i.test(key)) {
                        self.allowSave = true;
                        self.$formAction.find('.btn.save').show();
                        self.isSubmittedKey = key;
                    }
                    self.viewModel[key] = ko.observable(item[key]);
                }
            }
            // apply Knockout bindings
            ko.applyBindings(self.viewModel, self.form);
            self.viewModelIsBound = true;
            // get CreatedBy profile
            self.getListItemsRest(item.CreatedBy.__deferred.uri, function (data, status, jqXhr) {
                var person = data.d;
                self.viewModel.CreatedBy(person);
                self.viewModel.isAuthor(self.currentUser.id == person.Id);
                self.viewModel.CreatedByName(person.Name);
                self.viewModel.CreatedByEmail(person.WorkEMail);
                if (self.includeUserProfiles) {
                    self.$createdInfo.find('.create-mod-info').prepend(Shockout.Templates.getUserProfileTemplate(person, "Created By"));
                }
            });
            // get ModifiedBy profile
            self.getListItemsRest(item.ModifiedBy.__deferred.uri, function (data, status, jqXhr) {
                var person = data.d;
                self.viewModel.ModifiedBy(person);
                self.viewModel.ModifiedByName(person.Name);
                self.viewModel.ModifiedByEmail(person.WorkEMail);
                if (self.includeUserProfiles) {
                    self.$createdInfo.find('.create-mod-info').append(Shockout.Templates.getUserProfileTemplate(person, "Last Modified By"));
                }
            });
        };
        // http://blog.vgrem.com/2014/03/22/list-items-manipulation-via-rest-api-in-sharepoint-2010/
        SPForm.prototype.deleteListItem = function (model) {
            var self = model.parent;
            var item = self.listItem;
            var timeout = 3000;
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
        };
        // http://blog.vgrem.com/2014/03/22/list-items-manipulation-via-rest-api-in-sharepoint-2010/
        SPForm.prototype.saveListItem = function (model, isSubmit, refresh, customMsg) {
            if (isSubmit === void 0) { isSubmit = true; }
            if (refresh === void 0) { refresh = true; }
            if (customMsg === void 0) { customMsg = undefined; }
            var self = model.parent, isNew = !!!self.itemId, timeout = 3000, saveMsg = customMsg || "<p>Your form has been saved.</p>", postData = {}, headers = { Accept: 'application/json;odata=verbose' }, url, contentType = 'application/json';
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
            $.each(self.editableFields, function (i, key) {
                postData[key] = model[key]();
            });
            //Only update IsSubmitted if it's != true -- if it was already submitted.
            //Otherwise pressing Save would set it from true back to false - breaking any workflow logic in place!
            if (typeof model[self.isSubmittedKey] != "undefined" && (model[self.isSubmittedKey]() == null || model[self.isSubmittedKey]() == false)) {
                postData[self.isSubmittedKey] = isSubmit;
            }
            if (isNew) {
                url = self.rootUrl + self.siteUrl + '/_vti_bin/listdata.svc/' + self.listName.replace(/\s/g, '') + '(' + self.itemId + ')';
                //postData = JSON.stringify(postData);
                contentType += ';odata=verbose';
            }
            else {
                url = self.listItem.__metadata.uri;
                headers['X-HTTP-Method'] = 'MERGE';
                headers['If-Match'] = self.listItem.__metadata.etag;
            }
            var $jqXhr = $.ajax({
                url: url,
                type: 'POST',
                processData: false,
                contentType: contentType,
                data: JSON.stringify(postData),
                headers: headers
            });
            $jqXhr.done(function (data, status, jqXhr) {
                self.listItem = Shockout.Utils.clone(data.d);
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
            $jqXhr.fail(function (obj, status, jqXhr) {
                var msg = obj.statusText + '. An error occurred while saving the form.';
                self.showDialog(msg);
                self.logError(msg + ': ' + JSON.stringify(arguments));
            });
        };
        SPForm.prototype.getAttachmentsAsync = function (self, args) {
            if (self === void 0) { self = undefined; }
            if (args === void 0) { args = undefined; }
            self = self || this;
            if (!!!self.listItem || !self.enableAttachments) {
                self.nextAsync(true);
                return;
            }
            try {
                self.getListItemsRest(self.listItem.Attachments.__deferred.uri, function (data, status, jqXhr) {
                    $.each(data.d.results, function (i, att) {
                        self.viewModel.attachments().push(new Shockout.Attachment(att));
                    });
                    self.nextAsync(true, 'Retrieved attachments.');
                });
            }
            catch (e) {
                if (self.debug) {
                    console.warn(e);
                }
            }
        };
        SPForm.prototype.deleteAttachment = function (att) {
            var self = this, model = self.viewModel;
            try {
                var packet = '<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                    '<soap:Body><DeleteAttachment xmlns="http://schemas.microsoft.com/sharepoint/soap/"><listName>' + self.listName + '</listName><listItemID>' + self.itemId + '</listItemID><url>' + att.href + '</url></DeleteAttachment></soap:Body></soap:Envelope>';
                var $jqXhr = $.ajax({
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
                    var attachments = model.attachments;
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
        };
        SPForm.prototype.getListItemsSoap = function (siteUrl, listName, viewFields, query, callback, rowLimit, viewName, queryOptions) {
            if (rowLimit === void 0) { rowLimit = 25; }
            if (viewName === void 0) { viewName = '<ViewName/>'; }
            if (queryOptions === void 0) { queryOptions = '<QueryOptions/>'; }
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
                var $jqXhr = $.ajax({
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
        };
        SPForm.prototype.getListAsync = function (self, args) {
            if (args === void 0) { args = undefined; }
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
                self.nextAsync(true);
            });
            $jqXhr.fail(function () {
                self.nextAsync(false, 'Failed to retrieve list data.');
            });
        };
        SPForm.prototype.log = function (msg) {
            if (this.debug) {
                console.log(msg);
            }
        };
        SPForm.prototype.updateStatus = function (msg, success) {
            if (success === void 0) { success = undefined; }
            success = success || true;
            this.$formStatus
                .html(msg)
                .css('color', (success ? "#ff0" : "$f00"))
                .slideUp();
        };
        SPForm.prototype.showDialog = function (msg, title, timeout) {
            if (title === void 0) { title = undefined; }
            if (timeout === void 0) { timeout = undefined; }
            var self = this;
            title = title || "Form Dialog";
            msg = (msg).toString().match(/<\w>\w*/) == null ? '<p>' + msg + '</p>' : msg; //wrap non-html in <p>
            self.$dialog.html(msg).dialog('open');
            if (timeout) {
                setTimeout(function () { self.$dialog.dialog.close(); }, timeout);
            }
        };
        SPForm.prototype.getListItemsRest = function (uri, done, fail, always) {
            if (fail === void 0) { fail = undefined; }
            if (always === void 0) { always = undefined; }
            var self = this;
            var $jqXhr = $.ajax({
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
            var fail = fail || function (obj, status, jqXhr) {
                if (obj.status && obj.status == '404') {
                    var msg = obj.statusText + ". The data may have been deleted by another user.";
                }
                else {
                    msg = status + ' ' + jqXhr;
                }
                self.showDialog(msg);
            };
            $jqXhr.fail(fail);
            if (always) {
                $jqXhr.always(always);
            }
        };
        /**
        * Validate the View Model's required fields
        * @returns: bool
        */
        SPForm.prototype.formIsValid = function (model, showDialog) {
            if (showDialog === void 0) { showDialog = false; }
            var self = model.parent, labels = [], errorCount = 0, invalidCount = 0, invalidLabels = [];
            try {
                self.$form.find('.required, [required]').each(function checkRequired(i, n) {
                    var p = Shockout.Utils.observableNameFromControl(n);
                    if (!!p && model[p]) {
                        var val = model[p]();
                        if (val == null) {
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
                self.$form.find(".invalid").each(function (i, el) {
                    var $parent = $(el).parent();
                    invalidLabels.push($(parent).first().html());
                    invalidCount++;
                });
                if (invalidCount > 0) {
                    labels.push('<p class="warning">There are validation errors with the following fields. Please correct before saving.</p><p style="color:#f00;">' + invalidLabels.join('<br />') + '</p>');
                }
                //if attachment(s) are required
                if (self.hasAttachments && self.requireAttachments && model.attachments().length == 0) {
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
        };
        SPForm.prototype.getVersion = function () {
            return this.version;
        };
        SPForm.prototype.checkoutRequired = function () {
            return this.requireCheckout;
        };
        SPForm.prototype.attachmentsEnabled = function () {
            return this.enableAttachments;
        };
        SPForm.prototype.logError = function (msg, self) {
            if (self === void 0) { self = undefined; }
            self = self || this;
            self.showDialog('<p>An error has occurred and the web administrator has been notified.</p><p>Error Details: <pre>' + msg + '</pre></p>');
            if (self.enableErrorLog) {
                Shockout.Utils.logError(msg, self.errorLogListName, self.rootUrl, self.debug);
            }
        };
        return SPForm;
    })();
    Shockout.SPForm = SPForm;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var HistoryItem = (function () {
        function HistoryItem(description, date) {
            this.description = description;
            this.date = date;
        }
        return HistoryItem;
    })();
    Shockout.HistoryItem = HistoryItem;
    var Attachment = (function () {
        function Attachment(att) {
            this.title = att.name;
            this.href = att.__metadata.media_src;
            this.ext = att.name.match(/./) != null ? att.name.substring(att.name.lastIndexOf('.') + 1, att.name.length) : '';
        }
        return Attachment;
    })();
    Shockout.Attachment = Attachment;
    var SpItem = (function () {
        function SpItem() {
        }
        return SpItem;
    })();
    Shockout.SpItem = SpItem;
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
var Shockout;
(function (Shockout) {
    /* Knockout Custom handlers */
    (function bindKoHandlers(ko) {
        //http://stackoverflow.com/questions/7904522/knockout-content-editable-custom-binding?lq=1
        ko.bindingHandlers['spContentHtml'] = {
            init: function (element, valueAccessor, allBindingsAccessor) {
                //ko.utils.registerEventHandler(element, "blur", update);
                //ko.utils.registerEventHandler(element, "keydown", update);
                //ko.utils.registerEventHandler(element, "change", update);
                //ko.utils.registerEventHandler(element, "mousedown", update);
                $(element).on('blur', update)
                    .on('keydown', update)
                    .on('change', update)
                    .on('mousedown', update);
                function update() {
                    var modelValue = valueAccessor();
                    var elementValue = element.innerHTML;
                    if (ko.isWriteableObservable(modelValue)) {
                        modelValue(elementValue);
                    }
                    else {
                        var allBindings = allBindingsAccessor();
                        if (allBindings['_ko_property_writers'] && allBindings['_ko_property_writers']['spContentHtml']) {
                            allBindings['_ko_property_writers']['spContentHtml'](elementValue);
                        }
                    }
                }
            },
            update: function (element, valueAccessor) {
                var value = ko.utils.unwrapObservable(valueAccessor()) || "";
                if (element.innerHTML !== value) {
                    element.innerHTML = value;
                }
            }
        };
        ko.bindingHandlers['spContentEditor'] = {
            init: function (element, valueAccessor, allBindings, bindingContext) {
                // This will be called when the binding is first applied to an element
                // Set up any initial state, event handlers, etc. here
                var viewModel = bindingContext.$data, modelValue = valueAccessor(), person = ko.unwrap(modelValue), $element = $(element);
                var key = Shockout.Utils.observableNameFromControl(element);
                if (!!!key) {
                    return;
                }
                var $rte = $('<div>', {
                    'data-bind': 'spContentHtml: ' + key,
                    'class': 'content-editable',
                    'contenteditable': true
                });
                if (!!$element.attr('required') && !!!$element.hasClass('required')) {
                    $rte.attr('required', '');
                    $rte.addClass('required');
                }
                $rte.insertBefore($element);
                $element.hide();
            }
        };
        /* SharePoint People Picker */
        ko.bindingHandlers['spPerson'] = {
            init: function (element, valueAccessor, allBindings, bindingContext) {
                try {
                    if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') {
                        return;
                    } /*stop if not an editable field */
                    // This will be called when the binding is first applied to an element
                    // Set up any initial state, event handlers, etc. here
                    var viewModel = bindingContext.$data, modelValue = valueAccessor(), person = ko.unwrap(modelValue);
                    $(element).attr('placeholder', 'Employee Account Name').addClass('people-picker-control');
                    //create wrapper for control
                    var $parent = $(element).parent();
                    //controls
                    var $spValidate = $('<button>', { 'html': '<span>Validate</span>', 'class': 'sp-validate-person', 'title': 'Validate the employee account name.' }).on('click', function () {
                        if ($.trim($(element).val()) == '') {
                            $(element).removeClass('invalid').removeClass('valid');
                            return false;
                        }
                        if (!Shockout.Utils.validateSpPerson(modelValue())) {
                            $spError.text('Invalid').addClass('error');
                            $(element).addClass('invalid').removeClass('valid');
                        }
                        else {
                            $spError.text('Valid').removeClass('error');
                            $(element).removeClass('invalid').addClass('valid');
                        }
                        return false;
                    });
                    $parent.append($spValidate);
                    var $spError = $('<span>', { 'class': 'sp-validation person' });
                    $parent.append($spError);
                    var $desc = $('<div>', { 'class': 'no-print', 'html': '<em>Enter the employee name. The auto-suggest menu will appear below the field. Select the account name.</em>' });
                    $parent.append($desc);
                    $(element).autocomplete({
                        source: function (request, response) {
                            Shockout.Utils.peopleSearch(request.term, function (data) {
                                response($.map(data, function (item) {
                                    return {
                                        label: item.Name + ' (' + item.WorkEMail + ')',
                                        value: item.Id + ';#' + item.Account
                                    };
                                }));
                            });
                        },
                        minLength: 3,
                        select: function (event, ui) {
                            modelValue(ui.item.value);
                        }
                    }).on('focus', function () { $(this).removeClass('valid'); })
                        .on('blur', function () { onChangeSpPersonEvent(this, modelValue); })
                        .on('mouseout', function () { onChangeSpPersonEvent(this, modelValue); });
                }
                catch (e) {
                    var msg = 'Error in Knockout handler spPerson init(): ' + JSON.stringify(e);
                    Shockout.Utils.logError(msg, Shockout.SPForm.errorLogListName);
                    throw msg;
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
                    var msg = 'Error in Knockout handler spPerson update(): ' + JSON.stringify(e);
                    Shockout.Utils.logError(msg, Shockout.SPForm.errorLogListName);
                    throw msg;
                }
            }
        };
        ko.bindingHandlers['spDate'] = {
            init: function (element, valueAccessor, allBindings, bindingContext) {
                var modelValue = valueAccessor();
                if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') {
                    return;
                } /*stop if not an editable field */
                $(element).datepicker().addClass('datepicker med').on('blur', onDateChange).on('change', onDateChange);
                $(element).attr('placeholder', 'MM/DD/YYYY');
                function onDateChange() {
                    try {
                        if ($.trim(this.value) == '') {
                            modelValue(null);
                            return;
                        }
                        modelValue(new Date(this.value));
                    }
                    catch (e) {
                        modelValue(null);
                        this.value = '';
                    }
                }
                ;
            },
            update: function (element, valueAccessor, allBindings, bindingContext) {
                var viewModel = bindingContext.$data;
                var modelValue = valueAccessor();
                var value = ko.unwrap(modelValue);
                var dateStr = '';
                if (value && value != null) {
                    var d = new Date(value);
                    dateStr = Shockout.Utils.dateToLocaleString(d);
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
            init: function (element, valueAccessor, allBindings, bindingContext) {
                if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') {
                    return;
                } /*stop if not an editable field */
                var viewModel = bindingContext.$data, modelValue = valueAccessor(), value = ko.unwrap(modelValue), required, $time, $display, $error, $element = $(element);
                try {
                    $display = $('<span>', { 'class': 'no-print' }).insertAfter($element);
                    $error = $('<span>', { 'class': 'error', 'html': 'Invalid Date-time', 'style': 'display:none;' }).insertAfter($element);
                    element.$display = $display;
                    element.$error = $error;
                    required = $element.hasClass('required') || $element.attr('required') != null;
                    $element.attr({
                        'placeholder': 'MM/DD/YYYY',
                        'maxlength': 10,
                        'class': 'datepicker med'
                    }).datepicker().on('change', function () {
                        try {
                            $error.hide();
                            if (!Shockout.Utils.isDate(this.value)) {
                                $error.show();
                                return;
                            }
                            var val = this.value;
                            val = val.split('/');
                            var m = val[0] - 1;
                            var d = val[1] - 0;
                            var y = val[2] - 0;
                            var date = modelValue() == null ? new Date(y, m, d) : modelValue();
                            date.setMonth(m);
                            date.setDate(d);
                            date.setYear(y);
                            modelValue(date);
                            $display.html(Shockout.Utils.toDateTimeLocaleString(date));
                        }
                        catch (e) {
                            $error.show();
                        }
                    });
                    $time = $("<input>", {
                        'type': 'text',
                        'maxlength': 8,
                        'style': 'width:6em;',
                        'class': (required ? 'required' : ''),
                        'placeholder': 'HH:MM PM'
                    }).insertAfter($element)
                        .on('change', function () {
                        try {
                            $error.hide();
                            var time = this.value.toString().toUpperCase().replace(/[^\d\:AMP\s]/g, '');
                            this.value = time;
                            if (modelValue() == null) {
                                return;
                            }
                            if (!Shockout.Utils.isTime(time)) {
                                $error.show();
                                return;
                            }
                            var d = modelValue();
                            var tt = time.replace(/[^AMP]/g, ''); // AM/PM
                            var t = time.replace(/[^\d\:]/g, '').split(':');
                            var h = t[0] - 0; //hours
                            var m = t[1] - 0; //minutes
                            if (tt == 'PM' && h < 12) {
                                h += 12; //convert to military time
                            }
                            else if (tt == 'AM' && h == 12) {
                                h = 0; //convert to military midnight
                            }
                            d.setHours(h);
                            d.setMinutes(m);
                            modelValue(d);
                            $display.html(Shockout.Utils.toDateTimeLocaleString(d));
                            $error.hide();
                        }
                        catch (e) {
                            $display.html(e);
                            $error.show();
                        }
                    });
                    $time.before('<span> Time: </span>').after('<span class="no-print"> (HH:MM PM) </span>');
                    element.$time = $time;
                    if (modelValue() == null) {
                        $element.val('');
                        $time.val('');
                    }
                }
                catch (e) {
                    var msg = 'Error in Knockout handler spDateTime init(): ' + JSON.stringify(e);
                    Shockout.Utils.logError(msg, Shockout.SPForm.errorLogListName);
                }
            },
            update: function (element, valueAccessor, allBindings, bindingContext) {
                var viewModel = bindingContext.$data, modelValue = valueAccessor(), value = ko.unwrap(modelValue);
                try {
                    if (value && value != null) {
                        var d = new Date(value);
                        var dateStr = Shockout.Utils.dateToLocaleString(d);
                        var timeStr = Shockout.Utils.toTimeLocaleString(d);
                        if (element.tagName.toLowerCase() == 'input') {
                            element.value = dateStr;
                            element.$time.val(timeStr);
                            element.$display.html(dateStr + ' ' + timeStr);
                        }
                        else {
                            element.innerHTML = dateStr + ' ' + timeStr;
                        }
                    }
                }
                catch (e) {
                    var msg = 'Error in Knockout handler spDateTime update(): ' + JSON.stringify(e);
                    Shockout.Utils.logError(msg, Shockout.SPForm.errorLogListName);
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
    })(ko);
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var Templates = (function () {
        function Templates() {
        }
        Templates.getFileUploadTemplate = function () {
            return '<div class="qq-uploader">' +
                '<div class="qq-upload-drop-area"><span>Drop files here to upload</span></div>' +
                '<div class="btn qq-upload-button">Attach Files</div>' +
                '<ul class="qq-upload-list"></ul>' +
                '</div>';
        };
        Templates.getCreatedModifiedInfo = function () {
            var template = '<h4>Created/Modified Information</h4>' +
                '<ul>' +
                '<li class="create-mod-info no-print"></li>' +
                '<li><label>Created By</label><a data-bind="text: {0}, attr:{href: \'mailto:\'+{1}()}" class="email"></a></li>' +
                '<li><label>Created</label><span data-bind="spDateTime: {2}"></span></li>' +
                '<li><label>Modified By</label><a data-bind="text: {3}, attr:{href: \'mailto:\'+{4}()}" class="email"></a></li>' +
                '<li><label>Modified</label><span data-bind="spDateTime: {5}"></span></li>' +
                '</ul>';
            var section = document.createElement('section');
            section.className = 'created-mod-info';
            section.innerHTML = template
                .replace(/\{0\}/g, Shockout.ViewModel.createdByKey)
                .replace(/\{1\}/g, Shockout.ViewModel.createdByEmailKey)
                .replace(/\{2\}/g, Shockout.ViewModel.createdKey)
                .replace(/\{3\}/g, Shockout.ViewModel.modifiedByKey)
                .replace(/\{4\}/g, Shockout.ViewModel.modifiedByEmailKey)
                .replace(/\{5\}/g, Shockout.ViewModel.modifiedKey);
            return section;
        };
        Templates.getHistoryTemplate = function () {
            var template = '<h4>Workflow History</h4>' +
                '<table border="1" cellpadding="5" cellspacing="0" class="data-table" style="width:100%;border-collapse:collapse;">' +
                '<thead>' +
                '<tr><th>Description</th><th>Date</th></tr>' +
                '</thead>' +
                '<tbody data-bind="foreach: {0}">' +
                '<tr><td data-bind="text: {1}"></td><td data-bind="text: {2}"></td></tr>' +
                '</tbody>' +
                '</table>';
            var section = document.createElement('section');
            section.setAttribute('data-bind', 'visible: {0}.length > 0'.replace(/\{0\}/i, Shockout.ViewModel.historyKey));
            section.innerHTML = template
                .replace(/\{0\}/g, Shockout.ViewModel.historyKey)
                .replace(/\{1\}/g, Shockout.ViewModel.historyDescriptionKey)
                .replace(/\{2\}/g, Shockout.ViewModel.historyDateKey);
            return section;
        };
        Templates.getFormAction = function (allowSave, allowDelete, allowPrint) {
            if (allowSave === void 0) { allowSave = true; }
            if (allowDelete === void 0) { allowDelete = true; }
            if (allowPrint === void 0) { allowPrint = true; }
            var template = [];
            template.push('<div class="form-breadcrumbs"><a href="/">Home</a> &gt; eForms</div>');
            template.push('<button class="btn cancel" data-bind="event: { click: cancel }"><span>Close</span></button>');
            if (allowPrint) {
                template.push('<button class="btn print" data-bind="visible: Id() != null, event: {click: print}"><span>Print</span></button>');
            }
            if (allowDelete) {
                template.push('<button class="btn delete" data-bind="visible: Id() != null, event: {click: deleteItem}"><span>Delete</span></button>');
            }
            if (allowSave) {
                template.push('<button class="btn save" data-bind="event: { click: save }"><span>Save</span></button>');
            }
            template.push('<button class="btn submit" data-bind="event: { click: submit }, visible: isValid"><span>Submit</span></button>');
            var div = document.createElement('div');
            div.className = 'form-action';
            div.innerHTML = template.join('');
            return div;
        };
        Templates.getAttachmentsTemplate = function (fileuploaderId) {
            var template = '<h4>Attachments</h4>' +
                '<div id="{0}"></div>' +
                '<table class="attachments-table">' +
                '<tbody data-bind="foreach: attachments">' +
                '<tr>' +
                '<td><a href="" data-bind="text: title, attr: {href: href, \'class\': ext}"></a></td>' +
                '<td><button data-bind="event: {click: $root.deleteAttachment}" class="btn del" title="Delete"><span>Delete</span></button></td>' +
                '</tr>' +
                '</tbody>' +
                '</table>';
            var div = document.createElement('div');
            div.innerHTML = template.replace(/\{0\}/, fileuploaderId);
            return div;
        };
        Templates.getUserProfileTemplate = function (profile, headerTxt) {
            var template = '<h4>{header}</h4>' +
                '<img src="{pictureurl}" alt="{name}" />' +
                '<ul>' +
                '<li><label>Name</label>{name}</li>' +
                '<li><label>Title</label>{jobtitle}</li>' +
                '<li><label>Department</label>{department}</li>' +
                '<li><label>Email</label><a href="mailto:{workemail}">{workemail}</a></li>' +
                '<li><label>Phone</label>{workphone}</li>' +
                '<li><label>Office</label>{office}</li>' +
                '</ul>';
            var div = document.createElement("div");
            div.className = "user-profile-card";
            div.innerHTML = template
                .replace(/\{header\}/g, headerTxt)
                .replace(/\{pictureurl\}/g, (profile.Picture.indexOf(',') > 0 ? profile.Picture.split(',')[0] : profile.Picture))
                .replace(/\{name\}/g, (profile.Name || ''))
                .replace(/\{jobtitle\}/g, profile.Title || '')
                .replace(/\{department\}/g, profile.Department || '')
                .replace(/\{workemail\}/g, profile.WorkEMail || '')
                .replace(/\{workphone\}/g, profile.WorkPhone || '')
                .replace(/\{office\}/g, profile.Office || '');
            return div;
        };
        return Templates;
    })();
    Shockout.Templates = Templates;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var Utils = (function () {
        function Utils() {
        }
        Utils.peopleSearch = function (term, callback, take, siteUrl) {
            if (take === void 0) { take = 10; }
            if (siteUrl === void 0) { siteUrl = ''; }
            var page = !!take ? take.toString() : '10';
            // Allowed system query options are $filter, $select, $orderby, $skip, $top, $count, $search, $expand, and $levels.
            var uri = siteUrl + "/_vti_bin/listdata.svc/UserInformationList?$filter=startswith(Name,'{0}')&$select=Id,Account,Name,WorkEMail&$orderby=Name&$top={1}"
                .replace(/\{0\}/, term).replace(/\{1\}/, page);
            var $jqXhr = $.ajax({
                url: uri,
                type: 'GET',
                cache: false,
                dataType: 'json',
                contentType: 'application/json; charset=utf-8',
                headers: {
                    'Accept': 'application/json'
                }
            });
            $jqXhr.done(function (data, status, jqXhr) {
                callback(data.d.results);
            });
            $jqXhr.fail(function (obj, status, jqXhr) {
                var msg = 'People Search error. Status: ' + obj.statusText + ' ' + status + ' ' + JSON.stringify(jqXhr);
                Utils.logError(msg, Shockout.SPForm.errorLogListName);
                throw msg;
            });
        };
        /**
        * Extract the Knockout observable name from a field with `data-bind` attribute
        * @param control: HTMLElement
        * @return string
        */
        Utils.observableNameFromControl = function (control) {
            var attr = $(control).attr("data-bind");
            if (!!!attr) {
                return null;
            }
            var rx = new RegExp("\\b:(\\s+|)\\w*\\b");
            var exec = rx.exec(attr);
            var result = !!exec ? exec[0].replace(/:(\s+|)/gi, "") : null;
            return result;
        };
        Utils.parseJsonDate = function (d) {
            if (!Utils.isJsonDate(d)) {
                return null;
            }
            return new Date(parseInt(d.replace(/\D/g, '')));
        };
        Utils.isJsonDate = function (val) {
            if (!!!val) {
                return false;
            }
            return /\/Date\(\d+\)\//.test(val + '');
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
                objectClone[prop] = this.clone(objectToBeCloned[prop]);
            }
            return objectClone;
        };
        Utils.logError = function (msg, errorLogListName, siteUrl, debug) {
            if (siteUrl === void 0) { siteUrl = ''; }
            if (debug === void 0) { debug = false; }
            if (debug) {
                console.warn(msg);
                return;
            }
            var loc = window.location.href;
            var errorMsg = '<p>An error occurred at <a href="' + loc + '" target="_blank">' + loc + '</a></p><p>Message: ' + msg + '</p>';
            $.ajax({
                url: siteUrl + "/_vti_bin/listdata.svc/" + errorLogListName.replace(/\s/g, ''),
                type: "POST",
                processData: false,
                contentType: "application/json;odata=verbose",
                data: JSON.stringify({ "Title": "Web Form Error", "Error": errorMsg }),
                headers: {
                    "Accept": "application/json;odata=verbose"
                },
                success: function () {
                },
                error: function (data) {
                    throw data.responseJSON.error;
                }
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
            var rx = new RegExp("\\d{1,2}:\\d{2}\\s{0,1}(AM|PM)");
            return rx.test(val);
        };
        Utils.isDate = function (val) {
            if (!!!val) {
                return false;
            }
            var rx = new RegExp("\\d{1,2}\/\\d{1,2}\/\\d{4}");
            return rx.test(val.toString());
        };
        Utils.dateToLocaleString = function (d) {
            try {
                var dd = d.getDate();
                dd = dd < 10 ? "0" + dd : dd;
                var mo = d.getMonth() + 1;
                mo = mo < 10 ? "0" + mo : mo;
                return mo + '/' + dd + '/' + d.getFullYear();
            }
            catch (e) {
                return 'Invalid Date';
            }
        };
        Utils.toTimeLocaleObject = function (d) {
            var hours = 0;
            var minutes;
            var tt;
            hours = d.getHours();
            minutes = d.getMinutes();
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
            var hours = d.getHours();
            var minutes = d.getMinutes();
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
            var num = Utils.unformatNumber(value), format = '%s%v', neg = format.replace('%v', '-%v'), useFormat = num > 0 ? format : num < 0 ? neg : format, numFormat = Utils.formatNumber(Math.abs(num), Utils.checkPrecision(precision));
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
        return Utils;
    })();
    Shockout.Utils = Utils;
})(Shockout || (Shockout = {}));
var Shockout;
(function (Shockout) {
    var ViewModel = (function () {
        function ViewModel(instance) {
            //public Title: KnockoutObservable<string> = ko.observable(null);
            this.CreatedBy = ko.observable(null);
            this.CreatedByName = ko.observable(null);
            this.CreatedByEmail = ko.observable(null);
            this.ModifiedBy = ko.observable(null);
            this.ModifiedByName = ko.observable(null);
            this.ModifiedByEmail = ko.observable(null);
            this.history = ko.observableArray([]);
            this.attachments = ko.observableArray([]);
            this.isAuthor = ko.observable(false);
            this.currentUser = ko.observable(null);
            var self = this;
            this.parent = instance;
            this.isValid = ko.pureComputed(function () {
                return self.parent.formIsValid(self);
            });
        }
        ViewModel.prototype.deleteItem = function () {
            this.parent.deleteListItem(this);
        };
        ViewModel.prototype.cancel = function () {
            window.location.href = this.parent.sourceUrl != null ? this.parent.sourceUrl : this.parent.rootUrl;
        };
        ViewModel.prototype.print = function () {
            window.print();
        };
        ViewModel.prototype.deleteAttachment = function (obj, event) {
            this.parent.deleteAttachment(obj);
            return false;
        };
        ViewModel.prototype.save = function (model, btn) {
            this.parent.saveListItem(model, false);
        };
        ViewModel.prototype.submit = function (model, btn) {
            this.parent.saveListItem(model, true);
        };
        ViewModel.createdByKey = 'CreatedByName';
        ViewModel.createdByEmailKey = 'CreatedByEmail';
        ViewModel.modifiedByKey = 'ModifiedByName';
        ViewModel.modifiedByEmailKey = 'ModifiedByEmail';
        ViewModel.createdKey = 'Created';
        ViewModel.modifiedKey = 'Modified';
        ViewModel.historyKey = 'history';
        ViewModel.historyDescriptionKey = 'description';
        ViewModel.historyDateKey = 'date';
        return ViewModel;
    })();
    Shockout.ViewModel = ViewModel;
})(Shockout || (Shockout = {}));
//# sourceMappingURL=ShockoutForms-0.0.1.js.map