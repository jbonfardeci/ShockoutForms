/// <reference path="../typings/knockout.d.ts" />
/// <reference path="../typings/jquery.d.ts" />
/// <reference path="../typings/jquery.ui.datetimepicker.d.ts" />
/// <reference path="../typings/jqueryui.d.ts" />
/// <reference path="../typings/accounting.d.ts" />

'use strict';

module Shockout {
    
    // This method for finding specific nodes in the returned XML was developed by Steve Workman. See his blog post
    // http://www.steveworkman.com/html5-2/javascript/2011/improving-javascript-xml-node-finding-performance-by-2000/
    // for performance details.
    jQuery.fn.SPFilterNode = function (name) {
        return this.find('*').filter(function () {
            return this.nodeName === name;
        });
    }; // End $.fn.SPFilterNode

    export class ShockoutForm {

        public $createdInfo;
        public $dialog;
        public $form;
        public $formAction;
        public $formStatus;

        public allowDelete: boolean = false;
        public allowPrint: boolean  = true;
        public allowSave: boolean = false;
        public allowedExtensions: Array<string>;
        public attachmentMessage: string = 'An attachment is required.';
        public currentUser: ICurrentUser;
        public confirmationUrl: string = '/SitePages/Confirmation.aspx';
        public debug: boolean = false;
        public editableFields: Array<string> = [];
        public enableErrorLog: boolean = true;
        public errorLogListName: string = 'Error Log';
        public fileHandlerUrl: string = '/_layouts/webster/SPFormFileHandler.ashx';
        public fileUploaderSettings: IFileUploaderSettings;
        public fileUploader: any = null;
        public form: HTMLElement = null;
        public hasAttachments: boolean = true;
        public itemId: number = null;
        public isSubmittedKey: string;
        public listId: string = null;
        public listItem: ISpItem;
        public listName: string = null;
        public preRender: Function;
        public postRender: Function;
        public preSave: Function;
        public requireAttachments: boolean = false;
        public rootUrl: string = '//' + window.location.hostname;
        public siteUrl: string = '/';
        public includeUserProfiles: boolean = true;
        public includeWorkflowHistory: boolean = true;
        public sourceUrl: string;
        public version: number = 1.0;
        public viewModelIsBound: boolean = false;
        public viewModel: IViewModel;   
        public workflowHistoryListName: string = 'Workflow History';
        
        private asyncFns: Array<any>;

        constructor(options: {}) {
            var self = this;

            if (!(this instanceof ShockoutForm)) {
                var error = "You must declare an instance of this class with 'new'.";
                alert(error);
                throw error;
                return;
            }

            if (!!this.getQueryParam("id")) {
                this.itemId = parseInt(this.getQueryParam("id"));
            }
            
            if (!!this.getQueryParam("formid")) {
                this.itemId = parseInt(this.getQueryParam("formid"));
            }

            this.sourceUrl = this.getQueryParam("source"); //if accessing the form from a SP list, take user back to the list on close

            if (!!this.sourceUrl) {
                this.sourceUrl = decodeURIComponent(this.sourceUrl);
            }

            // override default instance variables with key-value pairs from args
            if (options && options.constructor === Object) {
                for (var p in options) {
                    this[p] = options[p];
                }
            }
            else {
                error = "Missing required parameters.";
                alert(error);
                throw error;
                return;
            }

            // get the form container element
            this.form = <HTMLFormElement>(arguments['form'].constructor == String
                ? document.getElementById(arguments['form'])
                : arguments['form']);

            this.$form = $(this.form);

            this.viewModel = new ViewModel(this);

            //Cascading Asynchronous Function Execution (CAFE) Array
            this.asyncFns = [
                function () {
                    if (self.preRender) {
                        self.preRender(self);
                    }
                    self.nextAsync(true);
                }
                , self.getCurrentUserAsync
                , self.getUsersGroupsAsync
                , self.restrictSpGroupElementsAsync
                , self.initFormAsync
                , self.getListItemAsync
                , self.getAttachmentsAsync
                , self.getHistoryAsync
                , function () {
                    if (self.postRender) {
                        self.postRender(self);
                    }
                    self.nextAsync(true);
                }
            ];

            //start CAFE
            this.nextAsync();         
        }

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
            this.asyncFns.shift()(this, args);
        }

        initFormAsync(self: ShockoutForm, args: any = undefined) {            
            try {
                self.updateStatus("Initializing dynamic form features...");

                self.$form.prepend(Templates.BRANDING);

                self.$createdInfo = this.$form.find(".created-info");

                self.$formStatus = $('<div>', { 'class': 'form-status' }).appendTo(this.$form);

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

                // append action buttons
                self.$formAction = $(Templates.getFormAction(self.allowSave, self.allowDelete, self.allowPrint)).appendTo(self.$form);
                
                //append Created/Modified info to predefined section or append to form
                if (!!self.itemId) {
                    self.$createdInfo.html(Templates.getCreatedModifiedInfo().innerHTML);

                    //append Workflow history section
                    if (self.includeWorkflowHistory) {
                        self.$form.append(Templates.getHistoryTemplate());
                    }
                }

                if (this.editableFields.length == 0) {
                    //make array of SP field names and those that are editable from elements w/ data-bind attribute
                    self.$form.find("[data-bind]").each(function (i: number, e: HTMLElement) {
                        var key = self.observableNameFromControl(e);

                        //skip observable keys that have already been added or begins with an underscore '_' or dollar sign '$'
                        if (!!!key || self.editableFields.indexOf(key) > -1 || key.match(/^(_|\$)/) != null) { return; }

                        if (e.tagName == "INPUT" || e.tagName == "SELECT" || e.tagName == "TEXTAREA" || $(e).attr("contenteditable") == "true") {
                            self.editableFields.push(key);
                        }
                    });

                    self.editableFields.sort();
                }

                self.fileUploaderSettings = {
                    element: null,
                    action: self.siteUrl + self.fileHandlerUrl,
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

                //setup HTML fields
                // deprecated
                self.$form.find("textarea.rte").each(function (i: number, el: HTMLElement) {
                    var key = self.observableNameFromControl(el);
                    if (!!!key) { return; }

                    var $rte = $("<div>", {
                        "data-bind": "htmlValue: " + key,
                        "class": "content-editable",
                        "contenteditable": true
                    });

                    if ($(el).attr("required") != null || $(el).hasClass("required")) {
                        $rte.attr("required", "");
                        $rte.addClass("required");
                    }

                    $rte.insertBefore(el);

                    if (!self.debug) {
                        el.style.display = "none";
                    }
                });

                self.$form.find('[required]').addClass('required');

                self.nextAsync(true, "Form initialized.");

            }
            catch (e) {
                self.logError("initForm: " + e);
                self.nextAsync(false, "Failed to initialize form. " + e);
            }
        }

        getCurrentUserAsync(self: ShockoutForm, args: any = undefined): void {
            try {
                var currentUser: ICurrentUser;
                var query = '<Where><Eq><FieldRef Name="ID" /><Value Type="Counter"><UserID /></Value></Eq></Where>';
                var viewFields = '<FieldRef Name="ID" /><FieldRef Name="Name" /><FieldRef Name="EMail" /><FieldRef Name="Department" /><FieldRef Name="JobTitle" /><FieldRef Name="UserName" /><FieldRef Name="Office" />';

                self.getListItemsSoap(self.siteUrl, 'User Information List', viewFields, query, function (xData, Sstatus) {
                    
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

                    var $res: any = $(xData.responseXML);

                    $res.SPFilterNode("z:row").each(function (i: number, node: any) {
                        user.id = parseInt($(node).attr("ows_ID"));
                        user.title = $(node).attr("ows_Name");
                        user.login = $(node).attr("ows_UserName");
                        user.email = $(node).attr("ows_EMail");
                        user.account = user.id + ';#' + user.login;
                        user.jobtitle = $(node).attr("ows_JobTitle");
                        user.department = $(node).attr("ows_Department");
                    });

                    self.currentUser = user;
                    self.viewModel.currentUser(user);
                    self.nextAsync(true, 'Retrieved your account.');
                });
            }
            catch (e) {
                self.logError("getCurrentUserAsync:" + e);
                self.nextAsync(false, "Failed to retrieve your account.");
            }
        }

        getUsersGroupsAsync(self: ShockoutForm, args: any = undefined): void {
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
                });

                $jqXhr.fail(function (xData, status) {
                    var msg = "Failed to retrieve your groups: " + status;
                    self.logError(msg);
                    self.nextAsync(false, msg);
                });

                self.updateStatus("Retrieving your groups...");
            }
            catch (e) {
                self.logError("getUsersGroupsAsync: " + e);
                self.nextAsync(false, "Failed to retrieve your groups.");
            }
        }

        restrictSpGroupElementsAsync(self: ShockoutForm, args: any = undefined): void {
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
            }
            catch (e) {
                self.logError("restrictSpGroupElementsAsync: " + e);
                self.nextAsync(true, "Failed to retrieve your permissions.");
            }
        }

        getListItemAsync(self: ShockoutForm, args: any = undefined): void {
            var model: IViewModel = self.viewModel;

            self.updateStatus("Retrieving form values...");

            if (!!!self.itemId) {
                self.nextAsync(true, "This is a New form.");
                return;
            }

            var uri = self.rootUrl + self.siteUrl + '/_vti_bin/listdata.svc/' + self.listName.replace(/\s/g, '') + '(' + self.itemId + ')';
            // get the list item data
            self.getListItemsRest(uri, bindValues, fail);

            function bindValues(data: ISpWrapper<ISpItem>, status: string, jqXhr: any) {
                self.bindListItemValues(self, model, data.d);
                self.nextAsync(true, "Retrieved form data.");
            }

            function fail(obj: any, status: string, jqXhr: any) {
                if (obj.status && obj.status == '404') {
                    var msg = obj.statusText + ". The form may have been deleted by another user."
                }
                else {
                    msg = status + ' ' + jqXhr;
                }
                self.showDialog(msg);
                self.nextAsync(false, msg);
            }
        }

        getHistoryAsync(self: ShockoutForm, args: any = undefined): void {
            try {
                if (!!!self.itemId) {
                    self.nextAsync(true);
                    return;
                }
                var historyItems: Array<any> = [];
                var uri = self.rootUrl + self.siteUrl + "/_vti_bin/listdata.svc/" + self.workflowHistoryListName.replace(/\s/g, '') +
                    "?$filter=ListID eq '" + self.listId + "' and PrimaryItemID eq " + self.itemId + "&$select=Description,DateOccurred&$orderby=DateOccurred asc";

                self.getListItemsRest(uri, function (data: ISpCollectionWrapper<any>, status: string, jqXhr: any) {
                    $(data.d).each(function (i: number, item: any) {
                        historyItems.push(new HistoryItem(item.Description, self.parseJsonDate(item.DateOccurred)));
                    });
                    self.viewModel.history(historyItems);
                    self.nextAsync(true, "Retrieved workflow history.");
                });
            }
            catch (ex) {
                var wfUrl = self.rootUrl + self.siteUrl + '/Lists/' + self.workflowHistoryListName.replace(/\s/g, '%20');
                self.logError('The Workflow History list may be full at <a href="{url}">{url}</a>. Failed to retrieve workflow history in method, getHistoryAsync(). Error: '.replace(/\{url\}/g, wfUrl) + JSON.stringify(ex));
                self.nextAsync(true, 'Failed to retrieve workflow history.');
            }
        }

        bindListItemValues(self: ShockoutForm, model: IViewModel, item: ISpItem): void {
            self.listItem = self.clone(item, self); //store copy of the original SharePoint list item

            // Exclude these read-only metadata fields from the Knockout view model.
            var rxExclude = new RegExp("^(__metadata|ContentTypeID|ContentType|CreatedBy|ModifiedBy|Owshiddenversion|Version|Attachments|Path)");

            for (var key in item) {

                if (rxExclude.test(key) || !!model[key]) { continue; }

                // Object types will have a corresponding key name plus the suffix `Value` or `Id` for lookups.
                // For example: `SupervisorApproval` is an object container for `__deferred` that corresponds to `SupervisorApprovalValue` 
                // which is an ID or string value.
                if (item[key].constructor === Object && item[key]['__deferred']) {
                    if (item[key + 'Value']) {
                        model[key] = ko.observable(item[key + 'Value']);
                    } else if (item[key + 'Id']) {
                        model[key] = ko.observable(item[key + 'Id']);
                    }
                }
                else if (self.isJsonDate(item[key])) {
                    // parse JSON dates
                    model[key] = ko.observable(self.parseJsonDate(item[key]));
                }
                else {
                    // if there is a boolean field for storing the state of a form's submission status 
                    if (/submitted/i.test(key)) {
                        self.allowSave = true;
                        self.$formAction.find('.btn.save').show();
                        self.isSubmittedKey = key;
                    }
                    model[key] = ko.observable(item[key]);
                }
            } 
                
            // apply Knockout bindings
            ko.applyBindings(model, self.form);
            self.viewModelIsBound = true;  

            // get CreatedBy profile
            self.getListItemsRest(item.CreatedBy.__deferred.uri, function (data: ISpWrapper<ISpPerson>, status: string, jqXhr: any) {
                var person: ISpPerson = data.d;
                model.CreatedBy(person);
                model.isAuthor(self.currentUser.id == person.Id);
                if (self.includeUserProfiles) {
                    self.$createdInfo.find('.create-mod-info').prepend(Templates.getUserProfileTemplate(person, "Created By"));
                }
            });

            // get ModifiedBy profile
            self.getListItemsRest(item.ModifiedBy.__deferred.uri, function (data: ISpWrapper<ISpPerson>, status: string, jqXhr: any) {
                var person: ISpPerson = data.d;
                model.ModifiedBy(person);
                if (self.includeUserProfiles) {
                    self.$createdInfo.find('.create-mod-info').append(Templates.getUserProfileTemplate(person, "Last Modified By"));
                }
            });
        }

        // http://blog.vgrem.com/2014/03/22/list-items-manipulation-via-rest-api-in-sharepoint-2010/
        deleteListItem(model: IViewModel) {
            var self: ShockoutForm = model.parent;
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

        // http://blog.vgrem.com/2014/03/22/list-items-manipulation-via-rest-api-in-sharepoint-2010/
        saveListItem(model: IViewModel, isSubmit: boolean = true, refresh: boolean = true, customMsg: string = undefined): void {

            var self: ShockoutForm = model.parent,
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
            if (isSubmit && !self.formIsValid(model)) {
                return;
            }
            
            // prepare data to post
            $.each(this.editableFields, function (i: number, key: string): void {
                postData[key] = model[key]();
            });

            //Only update IsSubmitted if it's != true -- if it was already submitted.
            //Otherwise pressing Save would set it from true back to false - breaking any workflow logic in place!
            if (typeof(model[self.isSubmittedKey]) != "undefined" && (model[self.isSubmittedKey]() == null || model[self.isSubmittedKey]() == false)) {
                postData[self.isSubmittedKey] = isSubmit;
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
                var listItem: ISpItem = data.d;
                self.itemId = listItem.Id;

                if (isSubmit && !self.debug) {//submitting form
                    self.showDialog("<p>Your form has been submitted. You will be redirected in " + timeout / 1000 + " seconds.</p>", "Form Submission Successful");
                    setTimeout(function () {
                        window.location.href = self.sourceUrl != null ? self.sourceUrl : self.confirmationUrl;
                    }, timeout);
                }
                else {//saving form
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
                        self.bindListItemValues(self, model, listItem);
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

        getAttachmentsAsync(self: ShockoutForm = undefined): void {
            self = self || this;
            self.getListItemsRest(self.listItem.Attachments.__deferred.uri, function (data: ISpCollectionWrapper<ISpAttachment>, status: string, jqXhr: any) {
                $.each(data.d.results, function (i: number, att: ISpAttachment) {
                    self.viewModel.attachments().push(new Attachment(att));
                });
            });
        }

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
                    url: self.rootUrl + siteUrl + '/_vti_bin/lists.asmx',
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
        * Extract the Knockout observable name from a field with `data-bind` attribute
        * @param control: HTMLElement
        * @return string
        */
        observableNameFromControl(control: HTMLElement): string {
            var attr: string = $(control).attr("data-bind");
            if (!!!attr) { return null; }
            var rx: RegExp = new RegExp("\\b:(\\s+|)\\w*\\b");
            var exec: Array<string> = rx.exec(attr);
            var result: string = !!exec ? exec[0].replace(/:(\s+|)/gi, "") : null;
            return result; 
        }

        logError(msg: string): void {
            var self = this;

            //a dictionary lookup for known error messages from server
            var errors = [
                {
                    "message": "An error occurred. Invalid data has been used to update the list item. The field you are trying to update may be read only.",
                    "definition": "An Employee Account Name field contains an invalid company employee account name/ID. Please inspect each field for a valid account name/ID.",
                    "action": function (o) {
                        var labels = [];

                        //display labels of fields to correct
                        $("input.people-picker-control", self.form).each(function (el) {
                            var $parent = $(this).parent();
                            var label = $parent.first().html();
                            labels.push(label);
                        });

                        return o.definition += '<div><strong>' + labels.join('<br />') + '</strong></div>';
                    }
                }
            ];

            //lookup [from errors] and display a friendly error message for known issues to interpret canned server responses            
            for (var i = 0; i < errors.length; i++) {
                var rx = new RegExp(errors[i].message, "i");
                if (rx.test(msg)) {
                    if ("action" in errors[i]) {
                        msg = errors[i].action(errors[i]);
                    }
                    else {
                        msg = errors[i].definition;
                    }
                    break;
                }
            }

            if (this.debug) {
                this.log(msg);
                return;
            }

            var loc = window.location.href;
            var errorMsg = '<p>An error occurred at <a href="' + loc + '" target="_blank">' + loc + '</a></p>' +
                '<p>List Site URL: ' + self.rootUrl + self.siteUrl + '<br />' +
                'List Name: ' + self.listName + '<br />' +
                'Message: ' + msg + '</p>';

            if (!this.enableErrorLog) { return; }

            $.ajax({
                url: self.rootUrl + "/_vti_bin/listdata.svc/" + self.errorLogListName.replace(/\s/g, ''),
                type: "POST",
                processData: false,
                contentType: "application/json;odata=verbose",
                data: JSON.stringify({ "Title": "Web Form Error: " + this.listName, "Error": errorMsg }),
                headers: {
                    "Accept": "application/json;odata=verbose"
                },
                success: function (data) {
                    self.showDialog('<p>An error has occurred and the web administrator has been notified. They will be in touch with you soon.</p><p>Error Details: <pre>' + msg + '</pre></p>');
                },
                error: function (data) {
                    throw data.responseJSON.error;
                }
            });

        }

        log(msg) {
            if (this.debug) {
                console.log(msg);
            }
        }

        updateStatus(msg: string, success: boolean = undefined): void {
            success = success || true;
            this.$formStatus
                .html(msg)
                .css('color', (success ? "#ff0" : "$f00"))
                .slideup();
        }

        showDialog(msg: string, title: string = undefined, timeout: number = undefined) {
            title = title || "Form Dialog";
            msg = (msg).toString().match(/<\w>\w*/) == null ? '<p>' + msg + '</p>' : msg; //wrap non-html in <p>
            this.$dialog.html(msg).dialog('open');
            if (timeout) {
                setTimeout(function () { this.$dialog.dialog.close(); }, timeout);
            }
        }

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

        parseJsonDate(d: string): Date {
            if (!this.isJsonDate(d)) { return null; }
            return new Date(parseInt(d.replace(/\d/g, '')));
        }

        isJsonDate(val: any): boolean {
            return /\/Date\(\d+\)\//.test(val.toString());
        }

        getQueryParam(p) {
            var escape: Function = window["escape"], unescape: Function = window["unescape"];
            p = escape(unescape(p));
            var regex = new RegExp("[?&]" + p + "(?:=([^&]*))?", "i");
            var match = regex.exec(window.location.search);
            return match != null ? match[1] : null;
        }

        // https://developer.mozilla.org/en-US/docs/Web/Guide/API/DOM/The_structured_clone_algorithm
        clone(objectToBeCloned, self: ShockoutForm = undefined) {

            self = self || this;

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
        }

        /**
        * Validate the View Model's required fields
        * @returns: bool
        */
        formIsValid(model: IViewModel): boolean {
            var self: ShockoutForm = model.parent,
                labels: Array<string> = [],
                errorCount: number = 0,
                invalidCount: number = 0,
                invalidLabels: Array<string> = []
            ;

            try {

                self.$form.find('.required, [required]').each(function checkRequired(i: number, n: any): void {
                    var p = self.observableNameFromControl(n);
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
                self.$form.find(".invalid").each(function (i: number, el: HTMLElement): void {
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
                    self.showDialog('<p class="warning">The following are required:</p><p class="error"><strong>' + labels.join('<br/>') + '</strong></p>');
                    return false;
                }
                return true;
            }
            catch (e) {
                self.logError("Form validation error at formIsValid(): " + JSON.stringify(e));
                return false;
            }
        }

    }

}