
module Shockout {

    export interface IViewModelAttachments extends KnockoutObservable<Array<ISpAttachment>> {
        getViewModel: Function;
        getSpForm: Function;
    }
    
    export interface IViewModel {
        // SP List Item Fields
        Id: KnockoutObservable<number>;
        Created: KnockoutObservable<Date>;
        CreatedBy: KnockoutObservable<ISpPerson>;
        Modified: KnockoutObservable<Date>;
        ModifiedBy: KnockoutObservable<ISpPerson>;

        // non-list item properties
        parent: Shockout.SPForm;
        allowSave: KnockoutObservable<boolean>;
        allowPrint: KnockoutObservable<boolean>;
        allowDelete: KnockoutObservable<boolean>;
        attachments: IViewModelAttachments;
        currentUser: KnockoutObservable<any>;
        historyItems: KnockoutObservableArray<IHistoryItem>;       
        isValid: KnockoutComputed<boolean>;
        showUserProfiles: KnockoutObservable<boolean>;
        navMenuItems: KnockoutObservableArray<any>;
        isMember: KnockoutComputed<boolean>;

        // methods
        isAuthor(): boolean;
        deleteItem(): void;
        cancel(): void;
        print(): void;
        deleteAttachment(obj: ISpAttachment, event: any): boolean;
        getAttachments(self: SPForm, callback: Function): void;
        save(model: ViewModel, btn: HTMLElement): void;
        submit(model: ViewModel, btn: HTMLElement): void;
    }

    export class ViewModel implements IViewModel {

        // static properties
        public static isSubmittedKey: string;
        public static parent: SPForm;

        // SP List Item Fields
        public Id: KnockoutObservable<number> = ko.observable(null);
        public Created: KnockoutObservable<Date> = ko.observable(null);
        public CreatedBy: KnockoutObservable<ISpPerson> = ko.observable(null);
        public Modified: KnockoutObservable<Date> = ko.observable(null);
        public ModifiedBy: KnockoutObservable<ISpPerson> = ko.observable(null);

        // non-list item fields
        public parent: Shockout.SPForm;
        public allowSave: KnockoutObservable<boolean> = ko.observable(false);
        public allowPrint: KnockoutObservable<boolean> = ko.observable(false);
        public allowDelete: KnockoutObservable<boolean> = ko.observable(false);
        public attachments: IViewModelAttachments = <any>ko.observableArray();
        public currentUser: KnockoutObservable<ICurrentUser>;
        public historyItems: KnockoutObservableArray<any> = ko.observableArray();
        public isValid: KnockoutComputed<boolean>;
        public showUserProfiles: KnockoutObservable<boolean> = ko.observable(false);
        public navMenuItems: KnockoutObservableArray<any> = ko.observableArray();
        public isMember: KnockoutComputed<boolean>;

        public deleteAttachment;
        public getAttachments;

        constructor(instance: Shockout.SPForm) {
            var self = this;
            this.parent = instance;
            ViewModel.parent = instance;

            this.isValid = ko.pureComputed(function (): boolean {
                return self.parent.formIsValid(self);
            });

            this.deleteAttachment = instance.deleteAttachment;
            this.getAttachments = instance.getAttachments;
            this.currentUser = ko.observable(instance.getCurrentUser());
            this.attachments.getViewModel = function(){
                return self;
            };
            this.attachments.getSpForm = function(){
                return self.parent;
            }
            this.isMember = <KnockoutComputed<boolean>>ko.pureComputed(<any>{
                read: function (): boolean {
                    return false;
                },
                write: function (groups: string): boolean {
                    return self.parent.currentUserIsMemberOfGroups(groups);
                },
                owner: this
            });
        }

        public isAuthor(): boolean {
            if(!!!this.CreatedBy()){ return true; }
            return this.currentUser().id == this.CreatedBy().Id;
        }

        public deleteItem(): void {
            this.parent.deleteListItem(this);
        }

        public cancel(): void {
            var src: string = this.parent.getSourceUrl();
            window.location.href = !!src ? src : this.parent.getRootUrl();
        }

        public print(): void {
            window.print();
        }

        public save(model: ViewModel, btn: HTMLElement): void {
            this.parent.saveListItem(model, false);
        }

        public submit(model: ViewModel, btn: HTMLElement): void {
            this.parent.saveListItem(model, true);
        }

    }

}