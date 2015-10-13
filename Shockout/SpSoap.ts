module Shockout {

    export class SpSoap {

        public static getCurrentUser(callback: Function): void {

            var user: ICurrentUser = <ICurrentUser>{};
            var query = '<Query><Where><Eq><FieldRef Name="ID" /><Value Type="Counter"><UserID /></Value></Eq></Where></Query>';
            var viewFields = '<ViewFields><FieldRef Name="ID" /><FieldRef Name="Name" /><FieldRef Name="EMail" /><FieldRef Name="Department" /><FieldRef Name="JobTitle" /><FieldRef Name="UserName" /><FieldRef Name="Office" /></ViewFields>';

            SpSoap.getListItems('', 'User Information List', viewFields, query, function (xmlDoc: any, status: string, jqXhr: any) {
                    
                $(xmlDoc).find('*').filter(function () {
                    return this.nodeName == 'z:row';

                }).each(function (i: number, node: any) {
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
        }

        public static getUsersGroups(loginName: string, callback: Function) {
            var packet = '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
                '<soap:Body>' +
                '<GetGroupCollectionFromUser xmlns="http://schemas.microsoft.com/sharepoint/soap/directory/">' +
                '<userLoginName>' + loginName + '</userLoginName>' +
                '</GetGroupCollectionFromUser>' +
                '</soap:Body>' +
                '</soap:Envelope>';

            var $jqXhr: JQueryXHR = $.ajax({
                url: '/_vti_bin/usergroup.asmx',
                type: 'POST',
                dataType: 'xml',
                data: packet,
                contentType: 'text/xml; charset="utf-8"'
            });

            $jqXhr.done(cb);
            $jqXhr.fail(cb);

            function cb(xmlDoc: XMLDocument, status: string, jqXhr: JQueryXHR) {

                var $errorText = $(xmlDoc).find('errorstring');

                // catch and handle returned error
                if (!!$errorText && $errorText.text() != "") {
                    callback(null, $errorText.text());
                    return;
                }

                var groups: Array<any> = [];

                $(xmlDoc).find("Group").each(function (i: number, el: HTMLElement) {
                    groups.push({
                        id: parseInt($(el).attr("ID")),
                        name: $(el).attr("Name")
                    });
                });

                callback(groups);
            }

        }

        public static getListItems(siteUrl: string, listName: string, viewFields: string, query: string, callback: Function, rowLimit: number = 25): void {

            siteUrl = Utils.formatSubsiteUrl(siteUrl);

            if (!!!listName) {
                Utils.logError("Parameter `listName` is null or undefined in method SpSoap.getListItems()", SPForm.errorLogListName, SPForm.errorLogSiteUrl);
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

            var $jqXhr: JQueryXHR = $.ajax({
                url: siteUrl + '_vti_bin/lists.asmx',
                type: 'POST',
                dataType: 'xml',
                data: packet,
                headers: {
                    "SOAPAction": "http://schemas.microsoft.com/sharepoint/soap/GetListItems",
                    "Content-Type": "text/xml; charset=utf-8"
                }
            });

            $jqXhr.done(function (xmlDoc: XMLDocument, status: string, error: string) {
                callback(xmlDoc);
            })

            $jqXhr.fail(function (jqXhr: JQueryXHR, status: string, error: string) {
                callback(null, status + ': ' + error);
            });

        }

        public static getList(siteUrl: string, listName: string, callback: Function): void {

            siteUrl = Utils.formatSubsiteUrl(siteUrl);

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

            $jqXhr.done(function (xmlDoc: XMLDocument, status: string, jqXhr: JQueryXHR) {
                callback(xmlDoc);
            });

            $jqXhr.fail(function (jqXhr: JQueryXHR, status: string, error: string) {
                callback(null, status + ': ' + error);
            });

        }

        public static checkInFile(pageUrl: string, checkinType: string, comment: string = '') {
            var action = 'http://schemas.microsoft.com/sharepoint/soap/CheckInFile';
            var params = [pageUrl, comment, checkinType];
            var packet = '<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap:Body><CheckInFile xmlns="http://schemas.microsoft.com/sharepoint/soap/"><pageUrl>{0}</pageUrl><comment>{1}</comment><CheckinType>{2}</CheckinType></CheckInFile></soap:Body></soap:Envelope>';

            return this.executeSoapRequest(action, packet, params);
        }

        public static checkOutFile(pageUrl: string, checkoutToLocal: string, lastmodified: string) {
            var action = 'http://schemas.microsoft.com/sharepoint/soap/CheckOutFile';
            var params = [pageUrl, checkoutToLocal, lastmodified];
            var packet = '<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"><soap:Body><CheckOutFile xmlns="http://schemas.microsoft.com/sharepoint/soap/"><pageUrl>{0}</pageUrl><checkoutToLocal>{1}</checkoutToLocal><lastmodified>{2}</lastmodified></CheckOutFile></soap:Body></soap:Envelope>';

            return this.executeSoapRequest(action, packet, params);
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
        public static executeSoapRequest = function (action: string, packet: string, params: Array<any>, siteUrl: string = '/', callback: Function = undefined): void {

            siteUrl = Utils.formatSubsiteUrl(siteUrl);

            try {
                var serviceUrl: string = siteUrl + '_vti_bin/lists.asmx';

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

                $jqXhr.fail(function (jqXhr: any, status: string, error: string) {
                    var msg = 'Error in SpSoap.executeSoapRequest. ' + status + ': ' + error + ' ';
                    Utils.logError(msg, SPForm.errorLogListName);
                    console.warn(msg);
                });
            }
            catch (e) {
                Utils.logError('Error in SpSoap.executeSoapRequest.', JSON.stringify(e), SPForm.errorLogListName);
                console.warn(e);
            }
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
        public static updateListItem = function (itemId: number, listName: string, fields: Array<Array<any>>, isNew: boolean = true, siteUrl: string = '/', callback: Function = undefined): void {

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
                soapEnvelope += "<Field Name='ID'>" + itemId + "</Field>";
            }
            soapEnvelope += "</Method></Batch>";

            params.push(soapEnvelope);

            SpSoap.executeSoapRequest(action, packet, params, siteUrl, callback);
        }
    }

}