module Shockout {

    export interface IShockoutObservable<T> extends KnockoutObservable<T> {
        _koName: string;
        _displayName: string;
        _name: string;
        _format: string;
        _required: boolean;
        _readOnly: boolean;
        _description: string;
        _type: string;
        _choices: Array<any>;
        _options: Array<any>;
        _isFillInChoice: boolean;
        _multiChoice: boolean;
    }

    export interface IHistoryItem {
        _description: string;
        _dateOccurred: Date;
    }

    export class HistoryItem implements IHistoryItem {
        public _description: string;
        public _dateOccurred: Date;
        constructor(d: string, date: Date) {
            this._description = d || null;
            this._dateOccurred = date || null;
        }
    }

    export interface IFileUploaderSettings {
        element: HTMLElement;
        action: string;
        debug: boolean;
        multiple: boolean;
        maxConnections: number;
        allowedExtensions: Array<string>;
        params: any;
        onSubmit(id, fileName);
        onComplete(id, fileName, json);
        template: string;
    }

    export interface ISpGroup {
        id: number;
        name: string;
    }

    export interface ICurrentUser {
        id: number;
        title: string;
        login: string;
        email: string;
        account: string;
        jobtitle: string;
        department: string;
        isAdmin: boolean;
        groups: Array<ISpGroup>;
    }


    /**
    * JSON "d" wrapper returned from SharePoint /_vti_bin/listdata.svc
    *   Prevents malicious scripts from executing
    */
    export interface ISpCollectionWrapper<T> {
        d: ISpCollection<T>;
    }

    export interface ISpWrapper<T> {
        d: T;
    }

    export interface ISpDeferred {
        uri: string;
    }

    export interface ISpDeferred {
        __deferred: ISpDeferred;
    }

    export interface ISpMetadata {
        uri: string;
        etag: string;
        type: string;
    }

    export interface ISpAttachmentMetadata {
        uri: string;
        type: string;
        edit_media: string;
        media_src: string;
        content_type: string;
        media_etag: string;
    }

    export interface ISpCollection<T> {
        results: Array<T>;
    }

    export interface ISpPersonSearchResult {
        __metadata: ISpMetadata;
        Id: number;
        Account: string;
        Name: string;
        WorkEMail: string;
    }

    export interface ISpPerson {
        __metadata: ISpMetadata;
        ContentTypeID: string;
        Name: string;
        Account: string;
        WorkEMail: string;
        EMail: string; //SP 2013 compat
        AboutMe: string;
        SIPAddress: string;
        IsSiteAdmin: boolean;
        Deleted: boolean;
        Picture: string;
        Department: string;
        Title: string;
        JobTitle: string; //SP 2013 compat
        MobilePhone: string;
        FirstName: string;
        LastName: string;
        WorkPhone: string;
        UserName: string;
        WebSite: string;
        AskMeAbout: string;
        Office: string;
        Id: number;
        ContentType: string;
        Modified: string;
        Created: string;
        CreatedBy: ISpDeferred;
        CreatedById: number;
        ModifiedById: number;
        Owshiddenversion: number;
        Version: string;
        Attachments: ISpDeferred;
        Path: string;
    }

    export interface ISpAttachment {
        __metadata: ISpAttachmentMetadata;
        EntitySet: string;
        ItemId: number;
        Name: string;
    }

    // recreate the SP REST object for an attachment
    export class SpAttachment implements ISpAttachment {
        __metadata: ISpAttachmentMetadata;
        EntitySet: string;
        ItemId: number;
        Name: string;

        constructor(rootUrl: string, siteUrl: string, listName: string, itemId: number, fileName: string) {
            var entitySet: string = listName.replace(/\s/g, '');
            siteUrl = Utils.formatSubsiteUrl(siteUrl);
            var uri = rootUrl + siteUrl + "_vti_bin/listdata.svc/Attachments(EntitySet='{0}',ItemId={1},Name='{2}')";
            uri = uri.replace(/\{0\}/, entitySet).replace(/\{1\}/, itemId + '').replace(/\{2\}/, fileName);

            this.__metadata = {
                uri: uri,
                content_type: "application/octetstream",
                edit_media: uri + "/$value",
                media_etag: null, // this property is unused for our purposes, so `null` is fine for now
                media_src: rootUrl + siteUrl + "/Lists/" + listName + "/Attachments/" + itemId + "/" + fileName,
                type: "Microsoft.SharePoint.DataService.AttachmentsItem"
            };
            this.EntitySet = entitySet;
            this.ItemId = itemId;
            this.Name = fileName;
        }
    }

    export interface ISpItem {
        __metadata: ISpMetadata;
        Title: string;
        ContentTypeID: string;
        Id: number;
        ContentType: string;
        Modified: any;
        Created: any;
        CreatedBy: ISpPerson;
        CreatedById: number;
        ModifiedBy: ISpPerson;
        ModifiedById: number;
        Owshiddenversion: number;
        Version: string;
        Attachments: ISpDeferred;
        Path: string;
    }

    export class SpItem implements ISpItem {
        __metadata: ISpMetadata;
        Title: string;
        ContentTypeID: string;
        Id: number;
        ContentType: string;
        Modified: any;
        Created: any;
        CreatedBy: ISpPerson;
        CreatedById: number;
        ModifiedBy: ISpPerson;
        ModifiedById: number;
        Owshiddenversion: number;
        Version: string;
        Attachments: ISpDeferred;
        Path: string;
        constructor() { }
    }

    export interface ISpMultichoiceValue {
        __metadata: ISpMetadata;
        Value: any;
    }

    export interface IPrincipalInfo {
        AccountName: string;
        UserInfoID: number;
        DisplayName: string;
        Email: string;
        Title: string; //job title
        IsResolved: boolean;
        PrincipalType: string;
    }
}