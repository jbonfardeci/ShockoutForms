
module Shockout {

    export class SpApi15 {

        /**
         * Get the current user.
         * @param {Function} callback
         * @param {boolean = false} expandGroups
         * @returns JQueryXHR
         */
        public static getCurrentUser(callback: Function, expandGroups: boolean = true, siteUrl: string = ''): JQueryXHR  {

            let url = Utils.formatSubsiteUrl(siteUrl);

            let $jqXhr: JQueryXHR = $.ajax({
                url: url + '_api/Web/CurrentUser' + (expandGroups ? '?$expand=Groups' : ''),
                type: 'GET',
                cache: true,
                dataType: 'json',
                contentType: 'application/json; charset=utf-8',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                }
            });

            $jqXhr.done(function (data: ISpWrapper<ISpApiPerson>, status: string, jqXhr: JQueryXHR) {
                let user: ISpApiPerson = data.d;
                let account = user.LoginName.replace(/^i\:0\#\.w\|/, ''); //remove SP 2013 prefix
                let currentUser: ICurrentUser = <ICurrentUser>{
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
                    let groups: any = data.d.Groups;
                    $(groups.results).each(function (i: number, group: any) {
                        currentUser.groups.push({id: group.Id, name: group.Title});
                    });
                }

                if(!!callback){
                    callback(currentUser);
                }
            });

            $jqXhr.fail(function (jqXhr: JQueryXHR, status: string, error: string) {
                if(!!callback){
                    callback(null, jqXhr.status); // '404'
                }
            });

            return $jqXhr;
        }

        /**
         * Get user's groups.
         * @param {number} userId
         * @param {JQueryPromiseCallback<any>} callback
         * @returns JQueryXHR
         */
        public static getUsersGroups(userId: number, callback: Function = undefined, siteUrl: string = ''): JQueryXHR {

            let url = `${Utils.formatSubsiteUrl(siteUrl)}_api/Web/GetUserById(${userId})/Groups`;

            let $jqXhr: JQueryXHR = $.ajax({
                url: url,
                type: 'GET',
                cache: true,
                dataType: 'json',
                contentType: 'application/json; charset=utf-8',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                }
            });

            $jqXhr.done(function (data: ISpCollectionWrapper<ISpApiUserGroup>, status: string, jqXhr: any) {
                let groups: Array<any> = [];
                for (var i = 0; i < data.d.results.length; i++) {
                    let group: ISpApiUserGroup = data.d.results[i];
                    groups.push({ id: group.Id, name: group.Title });
                }
                if(!!callback){
                    callback(groups);
                }
            });

            $jqXhr.fail(function (jqXhr: JQueryXHR, status: string, error: string) {
                if(!!callback){
                    callback(null, error);
                }
            });

            return $jqXhr;
        }

        /**
         * Send an email using the SharePoint Utilities library.
         * @param siteUrl
         * @param from 
         * @param to 
         * @param body 
         * @param subject 
         * @returns JQueryPromise<any>
         */
        public static SendEmail(siteUrl: string, from: string, to: string, body: string, subject: string): JQueryPromise<any> {
            let methodUrl = Utils.formatSubsiteUrl(siteUrl) + "_api/SP.Utilities.Utility.SendEmail";
            let data = {
                properties: {
                    __metadata: {
                        type: 'SP.Utilities.EmailProperties'
                    },
                    From: from,
                    To: {
                        results: [to]
                    },
                    Body: body,
                    Subject: subject
                }
            };
    
            return SpApi15.Post(siteUrl, methodUrl, data);
        }
    
        /**
         * Submit generic POST requests to SharePoint REST API.
         * @param siteUrl 
         * @param methodUrl 
         * @param data 
         * @returns JQueryPromise<any> 
         */
        public static Post(siteUrl: string, methodUrl: string, data: any, headers: any = undefined): JQueryPromise<any> {
            const deferred = $.Deferred(), self = this;
            const json = !!data ? JSON.stringify(data) : null;
    
            SpApi15.GetFormDigest(siteUrl).then(function(digest){

                let hdrs = {
                    'Accept': 'application/json;odata=verbose',
                    'Content-Type': 'application/json;odata=verbose',
                    'X-RequestDigest': digest.d.GetContextWebInformation.FormDigestValue,
                    'X-HTTP-Method': 'POST'
                };

                if(headers){
                    for(var p in headers){
                        hdrs[p] = headers[p];
                    }
                }

                $.ajax({
                    contentType: 'application/json;odata=verbose',
                    url: methodUrl,
                    type: 'POST',
                    data: json,
                    headers: hdrs,
                    success: function(data, status, xhr) {
                        deferred.resolve(data.d || data);
                    },
                    error: function(xhr, status, error) {
                        deferred.reject({xhr: xhr, status: status, error: error});
                        console.warn(error);
                    }
                });
    
            });
    
            return deferred.promise();
        }

        public static Get(url, cache: boolean = false): JQueryPromise<any> {
            const deferred = $.Deferred(), self = this;

            let $jqXhr: JQueryXHR = $.ajax({
                url: url,
                type: 'GET',
                cache: cache,
                dataType: 'json',
                contentType: 'application/json; charset=utf-8',
                headers: {
                    'Content-Type': 'application/json;odata=verbose',
                    'Accept': 'application/json;odata=verbose'
                }
            });

            $jqXhr.done(function (data: any, status: string, jqXhr: JQueryXHR) {
                deferred.resolve(data.d || data);
            });

            $jqXhr.fail(function (jqXhr: JQueryXHR, status: string, error: string) {
                if (!!status && parseInt(status) == 404) {
                    var msg = status + ". The data may have been deleted by another user."
                }
                else {
                    msg = status + ' ' + error;
                }
                deferred.reject(msg);
            });

            return deferred.promise();
        }

        public static GetListItem(listName: string, itemId: number, siteUrl: string = '/', cache: boolean = false, expand: string = undefined): JQueryPromise<any> {
            const methodUrl = `${Utils.formatSubsiteUrl(siteUrl)}_api/web/lists/GetByTitle('${listName}')/items(${itemId})${(!!expand ? '?$expand=' + expand : '')}`;
            return SpApi15.Get(methodUrl, cache);
        }
    
        /**
         * Get a SharePoint digest token to proceed with POST requests.
         * @param siteUrl 
         * @returns JQueryXHR
         */
        public static GetFormDigest(siteUrl: string): JQueryXHR {
            var opts: any = {
                'url': Utils.formatSubsiteUrl(siteUrl) + '_api/contextinfo',
                'method': 'POST',
                'headers': { 'Accept': 'application/json;odata=verbose' }
            };
            return $.ajax(opts);    
        }

        /**
         * Add or update a list item. If updating, include the item ID argument else leave undefined.
         * @param siteUrl 
         * @param listName 
         * @param data 
         * @param itemId
         * @returns JQueryPromise<any>
         */
        public static SaveListItem(siteUrl: string, listName: string, listItemType: string, data: any, itemId: number = undefined): JQueryPromise<any> {
            const methodUrl = `${Utils.formatSubsiteUrl(siteUrl)}_api/web/lists/GetByTitle('${listName}')/items${(!!itemId ? '(' + itemId + ')' : '')}`;
            
            let headers = undefined;
            // If updating a list item.
            if(!!itemId){
                headers = { 
                    'X-HTTP-Method': 'MERGE',
                    'If-Match': '*'
                };
            }

            let payload = {
                __metadata: {
                    type: formatListItemType(listItemType)
                }
            };

            for(const p in data){
                let o = data[p];
                if(o == null || o == undefined){
                    payload[p] = null;
                }
                else if(o.constructor === Array){
                    payload[p] = getSpRestArrayTemplate(o);
                }
                else{
                    payload[p] = o;
                }
            }

            return SpApi15.Post(siteUrl, methodUrl, payload, headers);

            function getSpRestArrayTemplate(array){
                return {
                    "__metadata": {
                        "type": "Collection(Edm.String)"
                    },
                    "results": array
                };
            }

            function formatListItemType(type: string){
                let t = type.replace(/\s/g, '_x0020_');
                let ptrn = `SP.Data.${t}ListItem`;
                let rx = new RegExp(ptrn);
                if(rx.test(t)){
                    return t;
                }
                return ptrn;
            }
        }

    }

}