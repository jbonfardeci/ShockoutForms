module Shockout {

    export interface IHistoryItem {
        description: string;
        date: Date;
    }

    export class HistoryItem implements IHistoryItem {
        public description: string;
        public date: Date;
        constructor(description: string, date: Date) {
            this.description = description;
            this.date = date;
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

    export interface IAttachment {
        title: string;
        href: string;
        ext: string;
    }

    export class Attachment implements IAttachment {
        public title: string;
        public href: string;
        public ext: string;
        constructor(att: ISpAttachment) {
            this.title = att.name;
            this.href = att.__metadata.media_src;
            this.ext = att.name.match(/./) != null ? att.name.substring(att.name.lastIndexOf('.') + 1, att.name.length) : '';
        }
    }

    export interface ICurrentUser {
        id: number;
        title: string;
        login: string;
        email: string;
        account: string;
        jobtitle: string;
        department: string;
        groups: Array<any>;
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

    export interface __deferredUri {
        uri: string;
    }

    export interface ISpDeferred {
        __deferred: __deferredUri;
    }

    export interface ISpItemMetadata {
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
        __metadata: ISpItemMetadata;
        Id: number;
        Account: string;
        Name: string;
        WorkEMail: string;
    }

    export interface ISpPerson {
        __metadata: ISpItemMetadata;
        ContentTypeID: string;
        Name: string;
        Account: string;
        WorkEMail: string;
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

    export interface ISpAttachment {
        __metadata: ISpAttachmentMetadata;
        EntitySet: string;
        ItemId: number;
        name: string;
    }

    export interface ISpItem {
        __metadata: ISpItemMetadata;
        Title: string;
        ContentTypeID: string;
        Id: number;
        ContentType: string;
        Modified: any;
        Created: any;
        CreatedBy: ISpDeferred;
        CreatedById: number;
        ModifiedBy: ISpDeferred;
        ModifiedById: number;
        Owshiddenversion: number;
        Version: string;
        Attachments: ISpDeferred;
        Path: string;
    }

    export class SpItem implements ISpItem {
        __metadata: ISpItemMetadata;
        Title: string;
        ContentTypeID: string;
        Id: number;
        ContentType: string;
        Modified: any;
        Created: any;
        CreatedBy: ISpDeferred;
        CreatedById: number;
        ModifiedBy: ISpDeferred;
        ModifiedById: number;
        Owshiddenversion: number;
        Version: string;
        Attachments: ISpDeferred;
        Path: string;
        constructor() { }
    }
}