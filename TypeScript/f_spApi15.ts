
module Shockout {

    export class SpApi15 {

        /**
         * Get the current user.
         * @param {Function} callback
         * @param {boolean = false} expandGroups
         */
        public static getCurrentUser(callback: Function, expandGroups: boolean = false): void {

            var $jqXhr: JQueryXHR = $.ajax({
                url: '/_api/Web/CurrentUser' + (expandGroups ? '?$expand=Groups' : ''),
                type: 'GET',
                cache: true,
                dataType: 'json',
                contentType: 'application/json; charset=utf-8',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                }
            });

            $jqXhr.done(function (data: ISpWrapper<ISpApiPerson>, status: string, jqXhr: JQueryXHR) {
                var user: ISpApiPerson = data.d;
                var currentUser: ICurrentUser = <ICurrentUser>{
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
                    var groups: any = data.d.Groups;
                    $(groups.results).each(function (i: number, group: any) {
                        currentUser.groups.push({id: group.Id, name: group.Title});
                    });
                }

                callback(currentUser);
            });

            $jqXhr.fail(function (jqXhr: JQueryXHR, status: string, error: string) {

                callback(null, jqXhr.status); // '404'
            });
        }

        /**
         * Get user's groups.
         * @param {number} userId
         * @param {JQueryPromiseCallback<any>} callback
         */
        public static getUsersGroups(userId: number, callback: JQueryPromiseCallback<any>): void {

            var $jqXhr: JQueryXHR = $.ajax({
                url: '/_api/Web/GetUserById('+userId+')/Groups',
                type: 'GET',
                cache: true,
                dataType: 'json',
                contentType: 'application/json; charset=utf-8',
                headers: {
                    'Accept': 'application/json;odata=verbose'
                }
            });

            $jqXhr.done(function (data: ISpCollectionWrapper<ISpApiUserGroup>, status: string, jqXhr: any) {
                var groups: Array<any> = [];
                for (var i = 0; i < data.d.results.length; i++) {
                    var group: ISpApiUserGroup = data.d.results[i];
                    groups.push({ id: group.Id, name: group.Title });
                }
                callback(groups);
            });

            $jqXhr.fail(function (jqXhr: JQueryXHR, status: string, error: string) {
                callback(null, error);
            });
        }

    }

}