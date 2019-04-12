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

    export interface ISpCollection<T> {
        results: Array<T>;
    }

    export interface ISpPersonSearchResult {
        __metadata: ISpMetadata;
        Id: number;
        Account: string;
        Name: string;
        WorkEmail: string;
    }

    export interface ISpPerson {
        __metadata: ISpMetadata;
        ContentTypeID: string;
        Name: string;
        Account: string;
        WorkEmail: string;
        AboutMe: string;
        SIPAddress: string;
        IsSiteAdmin: boolean;
        Deleted: boolean;
        Picture: string;
        Department: string;
        Title: string;
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

    export interface ISp15User {
        __metadata: any,
        Groups: ISpDeferred,
        Id: number;
        LoginName: string;
        Title: string;
        PrincipalType: number;
        Email: string;
        IsSiteAdmin: boolean;
        UserId: any;
    }

    export interface ISpAttachmentMetadata {
        uri?: string;
        type?: string;
        id?: string;
    }

    export interface ISpAttachment {
        __metadata?: ISpAttachmentMetadata;
        FileName?: string;
        ServerRelativeUrl?: string;
    }

    export interface ISpAttachmentFilesWrapper {
        results: ISpAttachment[];
    }

    export interface ISpItem {
        __metadata: ISpMetadata;
        Title: string;
        ContentTypeID: string;
        Id: number;
        ContentType: string;
        Modified?: any;
        Created?: any;
        CreatedBy?: ISpPerson;
        CreatedById?: number;
        ModifiedBy?: ISpPerson;
        ModifiedById?: number;
        AuthorId?: number;
        EditorId?: number;
        Author?: any;
        Editor?: any;
        Owshiddenversion: number;
        Version: string;
        Attachments?: boolean;
        AttachmentFiles?: ISpAttachmentFilesWrapper;
        Path: string;
    }

    export class SpItem implements ISpItem {
        __metadata: ISpMetadata;
        Title: string;
        ContentTypeID: string;
        Id: number;
        ContentType: string;
        Modified?: any;
        Created?: any;
        CreatedBy?: ISpPerson;
        CreatedById?: number;
        ModifiedBy?: ISpPerson;
        ModifiedById?: number;
        AuthorId?: number;
        EditorId?: number;
        Author?: any;
        Editor?: any;
        Owshiddenversion: number;
        Version: string;
        AttachmentFiles: ISpAttachmentFilesWrapper;
        Attachments: boolean;
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