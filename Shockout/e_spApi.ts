module Shockout {

    export class SpApi{

        /**
         * Search the User Information list.
         * @param {string} term
         * @param {Function} callback
         * @param {number = 10} take
         */
        public static peopleSearch(term: string, callback: Function, take: number = 10): void {

            var filter: string  = "startswith(Name,'{0}')".replace(/\{0\}/g, term);
            var select: string  = null;
            var orderby: string = "Name";
            var top: number = 10;
            var cache: boolean = true;

            SpApi.getListItems('UserInformationList', fn, '/', filter, select, orderby, top, cache);

            function fn(data: Array<ISpPersonSearchResult>, error: string) {
                if (!!error) {
                    callback(null, error);
                    return;
                }
                callback(data, error);
            };

        }

        /**
         * Get a person by their ID from the User Information list.
         * @param {number} id
         * @param {Function} callback
         */
        public static getPersonById(id: number, callback: Function): void {
            SpApi.getListItem('UserInformationList', id, function (data: ISpPerson, error: string) {
                if (!!error) {
                    callback(null, error);
                }
                callback(data);
            }, '/', true);
        }

        /**
         * General REST request method.
         * @param {string} url
         * @param {JQueryPromiseCallback<any>} callback
         * @param {boolean = false} cache
         * @param {string = 'GET'} type
         */
        public static executeRestRequest(url: string, callback: JQueryPromiseCallback<any>, cache: boolean = false, type: string = 'GET'): void {

            var $jqXhr: JQueryXHR = $.ajax({
                url: url,
                type: type,
                cache: cache,
                dataType: 'json',
                contentType: 'application/json; charset=utf-8',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                }
            });

            $jqXhr.done(function (data: any, status: string, jqXhr: JQueryXHR) {
                callback(data);
            });

            $jqXhr.fail(function (jqXhr: JQueryXHR, status: string, error: string) {
                if (!!status && status == '404') {
                    var msg = status + ". The data may have been deleted by another user."
                }
                else {
                    msg = status + ' ' + error;
                }
                callback(null, msg);
            });

        }

        /**
         * Get list item via REST services.
         * @param {string} listName
         * @param {number} itemId
         * @param {Function} callback
         * @param {string = '/'} siteUrl
         * @param {boolean = false} cache
         * @param {string = null} expand
         */
        public static getListItem(listName: string, itemId: number, callback: Function, siteUrl: string = '/', cache: boolean = false, expand: string = null): void {

            siteUrl = Utils.formatSubsiteUrl(siteUrl);
            var url: string = siteUrl + '_vti_bin/listdata.svc/' + Utils.toCamelCase(listName) + '(' + itemId + ')?$expand=CreatedBy,ModifiedBy' + (!!expand ? ',' + expand : '');

            SpApi.executeRestRequest(url, fn, cache, 'GET');

            function fn(data: any, error) {
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
            };
        }

        /**
         * Get list item via REST services.
         * @param {string} listName
         * @param {Function} callback
         * @param {string = '/'} siteUrl
         * @param {string = null} filter
         * @param {string = null} select
         * @param {string = null} orderby
         * @param {number = 10} top
         * @param {boolean = false} cache
         */
        public static getListItems(listName: string, callback: Function, siteUrl: string = '/', filter: string = null, select: string = null, orderby: string = null, top: number = 10, cache: boolean = false): void {

            siteUrl = Utils.formatSubsiteUrl(siteUrl);

            var url: Array<string> = [siteUrl + '_vti_bin/listdata.svc/' + Utils.toCamelCase(listName)];

            if (!!filter) { url.push('$filter='+filter); }

            if (!!select) { url.push('$select=' + select); }

            if (!!orderby) { url.push('$orderby=' + orderby); }

            url.push('$top=' + top);

            SpApi.executeRestRequest(url.join('&').replace(/\&/, '\?'), fn, cache, 'GET');

            function fn(data: any, error) {
                var data = !!data && 'd' in data ? data.d : data;
                var results: any = null;
                if (!!data) {
                    results = 'results' in data ? data.results : data;
                }
                callback(results, error);
            };

        }

        /**
         * Insert a list item with REST service.
         * @param {string} url
         * @param {Function} callback
         * @param {any = undefined} data
         */
        public static insertListItem(url: string, callback: Function, data: any = undefined): void {

            var $jqXhr: JQueryXHR = $.ajax({
                url: url,
                type: 'POST',
                processData: false,
                contentType: 'application/json',
                data: !!data ? JSON.stringify(data) : null,
                headers: {
                    'Accept': 'application/json;odata=verbose'
                }
            });

            $jqXhr.done(function (data: any, status: string, jqXhr: JQueryXHR) {
                callback(data);
            });

            $jqXhr.fail(function (jqXhr: JQueryXHR, status: string, error: string) {
                callback(null, status + ': ' + error);
            });
        }

        /**
         * Update a list item with REST service.
         * @param {ISpItem} item
         * @param {Function} callback
         * @param {any = undefined} data
         */
        public static updateListItem(item: ISpItem, callback: Function, data: any = undefined): void {

            var $jqXhr: JQueryXHR = $.ajax({
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

            $jqXhr.done(function (data: any, status: string, jqXhr: JQueryXHR) {
                callback(data);
            });

            $jqXhr.fail(function (jqXhr: JQueryXHR, status: string, error: string) {
                callback(null, status + ': ' + error);
            });
        }

        /**
         * Delete the list item with REST service.
         * @param {ISpItem} item
         * @param {JQueryPromiseCallback<any>} callback
         */
        public static deleteListItem(item: ISpItem, callback: JQueryPromiseCallback<any>): void {

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
                callback(data);
            });

            $jqXhr.fail(function (jqXhr: JQueryXHR, status: string, error: string): void {
                callback(null, error);
            });
        }

        /**
         * Delete an attachment with REST service.
         * @param {ISpAttachment} att
         * @param {Function} callback
         */
        public static deleteAttachment(att: ISpAttachment, callback: Function): void {

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

            $jqXhr.done(function (data: any, status: string, jqXhr: JQueryXHR) {
                callback(data);
            });

            $jqXhr.fail(function (jqXhr: JQueryXHR, status: string, error: string) {
                callback(null, status + ': ' + error);
            });
        }

    }

}