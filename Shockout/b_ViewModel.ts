module Shockout {
    
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
        attachments: KnockoutObservable<Array<any>>;
        currentUser: KnockoutObservable<any>;
        historyItems: KnockoutObservable<Array<IHistoryItem>>;       
        isValid: KnockoutComputed<boolean>;
        showUserProfiles: KnockoutObservable<boolean>;

        // methods
        isAuthor(): boolean;
        deleteItem(): void;
        cancel(): void;
        print(): void;
        deleteAttachment(obj: any, event: any): boolean;
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
        public attachments: KnockoutObservableArray<any> = ko.observableArray();
        public currentUser: KnockoutObservable<ICurrentUser>;
        public historyItems: KnockoutObservable<Array<any>> = ko.observableArray();
        public isValid: KnockoutComputed<boolean>;
        public showUserProfiles: KnockoutObservable<boolean> = ko.observable(false);

        public deleteAttachment;

        constructor(instance: Shockout.SPForm) {
            var self = this;
            this.parent = instance;
            ViewModel.parent = instance;

            this.isValid = ko.pureComputed(function (): boolean {
                return self.parent.formIsValid(self);
            });

            this.deleteAttachment = instance.deleteAttachment;
            this.currentUser = ko.observable(instance.getCurrentUser());
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