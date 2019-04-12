
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
                    success: function(data, status, jqXhr) {
                        if(!!data){
                            deferred.resolve(data.d || data);
                        }
                        else{
                            deferred.resolve({status: status, jqXhr: jqXhr});
                        }
                    },
                    error: function(jqXhr, status, error) {
                        deferred.reject({jqXhr: jqXhr, status: status, error: error});
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
                if(!!data){
                    deferred.resolve(data.d || data);
                }
                else{
                    deferred.resolve({status: status, jqXhr: jqXhr});
                }
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
        public static AddListItem(siteUrl: string, listName: string, listItemType: string, data: any): JQueryPromise<any> {
            const methodUrl = `${Utils.formatSubsiteUrl(siteUrl)}_api/web/lists/GetByTitle('${listName}')/items`;

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

            return SpApi15.Post(siteUrl, methodUrl, payload);

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

        public static UpdateListItem(siteUrl: string, metadata: ISpMetadata, data: any): JQueryPromise<any> {
            const methodUrl = metadata.uri;
            const deferred = $.Deferred(), self = this;

            let payload = {
                __metadata: {
                    type: metadata.type
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
           
            const json = JSON.stringify(payload);
    
            SpApi15.GetFormDigest(siteUrl).then(function(digest){

                let hdrs = {
                    'Accept': 'application/json;odata=verbose',
                    'X-RequestDigest': digest.d.GetContextWebInformation.FormDigestValue,
                    'X-HTTP-Method': 'MERGE',
                    'If-Match': metadata.etag
                };

                $.ajax({
                    contentType: 'application/json;odata=verbose',
                    url: methodUrl,
                    type: 'POST',
                    data: json,
                    headers: hdrs,
                    success: function(data, status, xhr) {
                        if(!!data){
                            deferred.resolve(data.d || data);
                        }
                        else{
                            deferred.resolve({success: true});
                        }
                    },
                    error: function(xhr, status, error) {
                        deferred.reject({xhr: xhr, status: status, error: error});
                        console.warn(error);
                    }
                });

            });

            return deferred.promise();

            function getSpRestArrayTemplate(array){
                return {
                    "__metadata": {
                        "type": "Collection(Edm.String)"
                    },
                    "results": array
                };
            }
        }

        public static GetUserById(siteUrl: string, userId: number): JQueryPromise<ISpPerson> {
            let url = `${Utils.formatSubsiteUrl(siteUrl)}_vti_bin/ListData.svc/UserInformationList(${userId})`;
            return SpApi15.Get(url, false);
        }

        public static GetGenericPersonPng(): string{
            return "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAHkAAAB4CAYAAADWpl3sAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAkmSURBVHhe7Z33bxNLEMc3oXcIvTeBaEIEJJqQQIi/ml8AAYpoEj0gmui995r3PivPk19EbOf2bm/Gno90SuLYvvV8d8ru3a77hoaGRp4+fRomTJgQnO7i9+/fYenSpaG/8bfTxbjIPYCL3AO4yD2Ai9wDuMg9gIvcA7jIPYCL3AO4yD2Ai9wDuMg9gIvcA3SdyCMjI+HPnz/xCszPnz/Djx8/wvfv38O3b9/iwe889uvXr/gcnstrupmuuNSISCIux7Rp08LMmTPDwMBAmD17dpg8eXKYPn16fO6nT5+iyO/evQtv374NX758ieL39/dHG/T19cXndQNyqdG0yCIuH2bSpEnxA61ZsybMmzcvTJw4sfGs1uDRL168CPfv3w/Pnz+P78Vru0Fs8yIjLgLhpevWrQsbN26MQqeAV9+8eTMKTkSw7tmmRZawvHr16rBt27YwZcqUxn/KgZB+8eLF8OzZs+jVhHKLiMimWi/ei+H37t0bdu3aVbrAQD4/cOBAGBwc/C8dWMaMyCLwnDlzwuHDh8OyZcsa/6mO9evXh4MHD8aOxLmtYkJkEXjBggXh0KFD/1XKOaBCR2i826rQJkTGuPPnz48htI7aYcaMGfHcdC6LoVu9yBgVI+/bt6/W4hCB9+zZE9tA0WcJ1SLLMGb37t2VFFjjhfH3jh07YrtIIVZQK7JUtZs2bYp5UQurVq2KEy6kECtCqxUZgcnDTHJog7E5KcRK2FYpMh7CTBPG1AizbJs3b47ttODNKkXGixkHM2TSCrNtc+fONeHN6kQWL96wYUPjEb2QSix4szqR8eKFCxeqKrbGgnlhJklc5HEgXkH1agGGdytWrIgdU7PQ6kTmgv+iRYsaj+iH2qHTa9d1oUpkihjCNNWrFSi+CNmaCzA1IuPFGGrJkiWNR+xADeEidwh3dsyaNavxlx2Y7pR6QiOqPNmqyHKzoFZUiTx16lRT+Vig3XRQ9+Q2YCDmgy3CFTLNFbaqnGzRiwWGfu7JbcBAGMoqeLKL3AGW73FOvee7SlSJ7FSDKpGZA7YKi+u0okZkQjULz6zCIjqt6UaVJ7Os1Cp0UBe5Daw3Yg2SReicmm+8V+XJ5DWLIZvVkLTdPbkNGIi8ZtGbP3z44J7cKVTX7ABgDXYsAPfkNmAg8jKr/S3BLBc7FdS5hKcdqjwZoWUfDyu8efMmtlerF4M6kSm88AwrPH78OKYZF7lDMBRh7969e41HdEOhiMiaQzWoEhnIyxRfFnIzAmsP1aBOZAxGMXPr1q3GIzohRN++fTt2She5AIS/ly9fqvbmu3fvxvGxhZ2BVLZQPOPy5csqr0zJfl90Ru1eDGq7IR6Cp1y/fr3xiB4uXbqk+qrTaNSKjAG5pYbczGZyWiAPP3nyxIwXg+qEghHx6AsXLkSvrhs629WrV00JDOqrBkQmNA4NDYWvX782Hs0PM3Hnzp2Lv1sotpox0VrC9ufPn8PJkyfjz9y8fv06nD59OhaB1gQGMy1GaC5DHjt2LBo9Fw8fPgynTp2K14u1z2yNhaluidAYG4/OMVnCEO7s2bNxcsaqwGAu9oixGcacOHGikuvPXCA5evRo7Eh0LIshuhk1+13jLXLI34CBqWRHV7P8X67+sKUDm7Swg24KpIEbN27Efa45r5x7LJrbLO2VtsrrWr2+arBP7ZuaYxgWb4tY7F/J4jHWRPE7oZmh08ePH1uKza03/J91wuw3wnYUneyky2vJ84jKbvVEBd4HW7QTh3ZzsMsAy3tk8zZGAFwu5Sc3+PFeY7W9amoTWXo9wrC0hF33li9fHlfrIwwGGQ3FD/lRjPY3YzV3GN4DoyMAh3Qc/s95EYGOQ6XOwWubxWgFz+UctJ0NVcfa34RhHzcUPHr0KHYizknozyl2dpGbxWWh+dq1a+OGZ51ujIpnME4lX2LgVoaSc3EgrDwGYuTRRyfwHkQX0sL+/fs7XmqLsR88eBDnu4kcufJ8VpExDuLiqVu3bo2bkBaFuWzyJsJ04nllIB2Gz0A62LlzZ2GRKOaGh4djZxHProosImMY8SQKoy1btpTyoV69ehXOnz8fvaKdV6ci4hLu2QZ55cqVjf8Uh9BN+wnjVba/cpHFOIRm9qtmK6Qy4QOQp+VWIdpfprFoP+cAIs/27dtLXyRPRLp27VpsexXhu1KRRWBOwI7zVXoaxc2VK1fiTQZirJTzibj8pKgivVS5BSS2P3PmTDxf2Y5WmcgiMIUVuSsXhD7yNaIjtIjdieC0WdrN6xCXzdRz7dLLxQ+mTjl/mUJXJjJDB7YipPqsA/L1nTt3omdTkTcL3owIiyH47KQVhnJU/HVsUPP+/ftw/Pjx2CbaWwaViMybYqAjR46U1tCiUL1iOO4TY5KDW3YwICA4lS11AqEYj61D2NFw9+e/epRWjJUusngF36HEBIdTDOYCGFMjdCoicmnuhufwxZgucBp8JYPMzpVFKSLTIGauKFacNJgHpy7AC8uiFJFpEA1j+0EnHSIi6VNqiFSSRaYhNIiGOeXARRWKwbJCdrLINIQ8TMOc8pCQXYY3J4ksFXXKBQfn73DplfRXu8hAqa/5+5usgsDso63Ck5lE0DCR0I0sXry4lCo7SWTyMbfcONXAHDozh6ne7CIrhmK2jLxcWGQ5sVfV1UG9w900tXoyk/yd3qPlFIPhaep4OUlkJkHKvlvC+T+kw9pEJoQQTlzkaiEdpk5xJofr1EuUTmsovHCkWkQWT3aqRUROIUlknwTJA3auzZM7WW/kpJM6jEoS2YdPeeAmw5QK28O1AWoJ15yQqpqiwKkeImbKHHZhT0ZkhlBO9VBdp9i6sCdzUh9C5QFbc7gndzE4U3ZPBk7qs115wM4pK1IKvZJy3ivrvNRSePkYOS8pw6hCInMyy19obZHs4RqRfUozLylfp184J7vIdhi3yPQmSnofI+cluydT0rvIeUkZrhb2ZL/tJy9FvRgKe7LPduWFIWu2cM2JELiMPS2czkmZlygkss925SdruOZkPhFii0Ii+xjZFoVE9jtCbDEukRGYgsvDtS36maLkYLFzu4PnIbJPhNTD3zRpdYi2fcPDwyNsOtrpkIirIYODgz4Zkhm2kOQbdMbjYETegYGB8A/hTaV31Q8zwAAAAABJRU5ErkJggg==";
        }

    }

}