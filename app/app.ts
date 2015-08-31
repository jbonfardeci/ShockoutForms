/// <reference path="../typings/knockout.d.ts" />
/// <reference path="../typings/jquery.d.ts" />
/// <reference path="../typings/jquery.ui.datetimepicker.d.ts" />
/// <reference path="../typings/jqueryui.d.ts" />

'use strict';

module Shockout {

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
        public fieldNames: Array<string>;
        public fileHandlerUrl: string = '/_layouts/webster/SPFormFileHandler.ashx';
        public fileUploaderSettings: IFileUploaderSettings;
        public fileUploader: any;
        public form: HTMLElement;
        public formId: string;
        public hasAttachments: boolean = true;
        public itemId: number;
        public isSubmittedKey: string;
        public listId: string;
        public listItem: ISpItem;
        public listName: string;
        public preRender: Function;
        public postRender: Function;
        public preSave: Function;
        public requireAttachments: boolean = false;
        public rootUrl: string = window.location.protocol + '//' + window.location.hostname + (!!window.location.port ? ':' + window.location.port : '');
        public siteUrl: string = '/';
        public includeUserProfiles: boolean = true;
        public includeWorkflowHistory: boolean = true;
        public sourceUrl: string;
        public viewModelIsBound: boolean = false;
        public viewModel: IViewModel;   
        public workflowHistoryListName: string = 'Workflow History';

        public static errorLogListName: string;
        
        private asyncFns: Array<any>;
        private version: string = '0.0.1';

        constructor(options: Object) {
            var self = this;

            if (!(this instanceof ShockoutForm)) {
                var error = "You must declare an instance of this class with 'new'.";
                alert(error);
                throw error;
                return;
            }

            if (!!Utils.getQueryParam("id")) {
                this.itemId = parseInt(Utils.getQueryParam("id"));
            }
            
            if (!!Utils.getQueryParam("formid")) {
                this.itemId = parseInt(Utils.getQueryParam("formid"));
            }

            this.sourceUrl = Utils.getQueryParam("source"); //if accessing the form from a SP list, take user back to the list on close

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
            this.form = <HTMLFormElement>(options['formId'].constructor == String
                ? document.getElementById(options['formId'])
                : options['formId']);

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

            ShockoutForm.errorLogListName = this.errorLogListName;

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
            this.nextAsync(true, 'Begin initialization...');         
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
            this.asyncFns.shift()(self, args);
        }

        initFormAsync(self: ShockoutForm, args: any = undefined) {            
            try {
                self.updateStatus("Initializing dynamic form features...");

                self.$createdInfo = self.$form.find(".created-info");

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

                if (self.editableFields.length == 0) {
                    //make array of SP field names and those that are editable from elements w/ data-bind attribute
                    self.$form.find('[data-bind]').each(function (i: number, e: HTMLElement) {
                        var key = Utils.observableNameFromControl(e);

                        //skip observable keys that have already been added or begins with an underscore '_' or dollar sign '$'
                        if (!!!key || self.editableFields.indexOf(key) > -1 || key.match(/^(_|\$)/) != null) { return; }

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
        }

        getCurrentUserAsync(self: ShockoutForm, args: any = undefined): void {
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
                });
            }
            catch (e) {
                if (self.debug) {
                    console.warn(e);
                }
                self.logError('getCurrentUserAsync():' + e);
                self.nextAsync(false, 'Failed to retrieve your account.');
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
                if (self.debug) {
                    console.warn(e);
                }
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
                if (self.debug) {
                    console.warn(e);
                }
                self.logError("restrictSpGroupElementsAsync: " + e);
                self.nextAsync(true, "Failed to retrieve your permissions.");
            }
        }

        getListItemAsync(self: ShockoutForm, args: any = undefined): void {
            try {
                var model: IViewModel = self.viewModel;

                self.updateStatus("Retrieving form values...");

                if (!!!self.itemId) {
                    self.nextAsync(true, "This is a New form.");
                    return;
                }

                var uri = self.rootUrl + self.siteUrl + '/_vti_bin/listdata.svc/' + self.listName.replace(/\s/g, '') + '(' + self.itemId + ')';
                // get the list item data
                self.getListItemsRest(uri, bindValues, fail);
            }
            catch (e) {
                if (self.debug) {
                    console.warn(e);
                }
            }

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

        bindListItemValues(self: ShockoutForm, model: IViewModel, item: ISpItem): void {
            self.listItem = Utils.clone(item); //store copy of the original SharePoint list item

            // Exclude these read-only metadata fields from the Knockout view model.
            var rxExclude = new RegExp("^(__metadata|ContentTypeID|ContentType|CreatedBy|ModifiedBy|Owshiddenversion|Version|Attachments|Path)");

            for (var key in item) {

                if (rxExclude.test(key) || !!model[key]) { continue; }

                // Object types will have a corresponding key name plus the suffix `Value` or `Id` for lookups.
                // For example: `SupervisorApproval` is an object container for `__deferred` that corresponds to `SupervisorApprovalValue` 
                // which is an ID or string value.
                if (item[key] != null && item[key].constructor === Object && item[key]['__deferred']) {
                    if (item[key + 'Value']) {
                        model[key] = ko.observable(item[key + 'Value']);
                    } else if (item[key + 'Id']) {
                        model[key] = ko.observable(item[key + 'Id']);
                    }
                }
                else if (item[key] != null && Utils.isJsonDate(item[key])) {
                    // parse JSON dates
                    model[key] = ko.observable(Utils.parseJsonDate(item[key]));
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
            $.each(self.editableFields, function (i: number, key: string): void {
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
            if (self.listItem == undefined) {
                self.nextAsync(true);
                return;
            }
            try {
                self.getListItemsRest(self.listItem.Attachments.__deferred.uri, function (data: ISpCollectionWrapper<ISpAttachment>, status: string, jqXhr: any) {
                    $.each(data.d.results, function (i: number, att: ISpAttachment) {
                        self.viewModel.attachments().push(new Attachment(att));
                    });
                    self.nextAsync(true, 'Retrieved attachments.');
                });
            }
            catch (e) {
                if (self.debug) {
                    console.warn(e);
                }
            }
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
                .slideUp();
        }

        showDialog(msg: string, title: string = undefined, timeout: number = undefined) {
            var self: ShockoutForm = this;
            title = title || "Form Dialog";
            msg = (msg).toString().match(/<\w>\w*/) == null ? '<p>' + msg + '</p>' : msg; //wrap non-html in <p>
            self.$dialog.html(msg).dialog('open');
            if (timeout) {
                setTimeout(function () { self.$dialog.dialog.close(); }, timeout);
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
                    var p = Utils.observableNameFromControl(n);
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

        getVersion(): string {
            return this.version;
        }

        logError(msg: string, self: ShockoutForm = undefined): void {
            self = self || this;
            self.showDialog('<p>An error has occurred and the web administrator has been notified. They will be in touch with you soon.</p><p>Error Details: <pre>' + msg + '</pre></p>');
            Utils.logError(msg, self.errorLogListName, self.rootUrl, self.debug);
        }
    }

}