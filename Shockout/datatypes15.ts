module Shockout {

    /**
    * SharePoint 2013 API Data Types
    */

    export interface ISpApiMetadata {
        id: string;
        uri: string;
        type: string;
    }

    export interface ISpApiDeferred {
        uri: string;
    }

    export interface ISpApiUserIdMetadata {
        type: string;
    }

    export interface ISpApiUserId {
        __metadata: ISpApiUserIdMetadata;
        NameId: string;
        NameIdIssuer: string;
    }

    export interface ISpApiPerson {
        __metadata: ISpApiMetadata;
        Groups: ISpApiDeferred;
        Id: number;
        IsHiddenInUI: boolean;
        LoginName: string;
        Title: string;
        PrincipalType: number;
        Email: string;
        IsSiteAdmin: boolean;
        UserId: ISpApiUserId;
    }

    export interface ISpApiUserGroup {
        __metadata: ISpApiMetadata;
        Owner: ISpApiDeferred;
        Users: ISpApiDeferred;
        Id: number;
        IsHiddenInUI: boolean;
        LoginName: string;
        Title: string;
        PrincipalType: number;
        AllowMembersEditMembership: boolean;
        AllowRequestToJoinLeave: boolean;
        AutoAcceptRequestToJoinLeave: boolean;
        Description: string;
        OnlyAllowMembersViewMembership: boolean;
        OwnerTitle: string;
        RequestToJoinLeaveEmailSetting: string;
    }

}