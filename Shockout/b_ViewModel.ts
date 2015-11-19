module Shockout {
    
    export interface IViewModel {
        Id: KnockoutObservable<number>;
        Created: KnockoutObservable<Date>;
        CreatedBy: KnockoutObservable<ISpPerson>;
        Modified: KnockoutObservable<Date>;
        ModifiedBy: KnockoutObservable<ISpPerson>;

        parent: Shockout.SPForm;
        historyItems: KnockoutObservable<Array<IHistoryItem>>;
        attachments: KnockoutObservable<Array<any>>;
        currentUser: KnockoutObservable<any>;
        isValid: KnockoutComputed<boolean>;
        showUserProfiles: KnockoutObservable<boolean>;

        isAuthor(): boolean;
        deleteItem(): void;
        cancel(): void;
        print(): void;
        deleteAttachment(obj: any, event: any): boolean;
        save(model: ViewModel, btn: HTMLElement): void;
        submit(model: ViewModel, btn: HTMLElement): void;
    }

    export class ViewModel implements IViewModel {

        public static historyKey: string = 'history';
        public static historyDescriptionKey: string = 'description';
        public static historyDateKey: string = 'date';
        public static isSubmittedKey: string;
        public static parent: SPForm;

        public Id: KnockoutObservable<number> = ko.observable(null);
        public Created: KnockoutObservable<Date> = ko.observable(null);
        public CreatedBy: KnockoutObservable<ISpPerson> = ko.observable(null);
        public Modified: KnockoutObservable<Date> = ko.observable(null);
        public ModifiedBy: KnockoutObservable<ISpPerson> = ko.observable(null);
        public showUserProfiles: KnockoutObservable<boolean> = ko.observable(false);

        public parent: Shockout.SPForm;
        public historyItems: KnockoutObservable<Array<any>> = ko.observableArray();
        public attachments: KnockoutObservableArray<any> = ko.observableArray();
        public currentUser: KnockoutObservable<ICurrentUser>;
        public isValid: KnockoutComputed<boolean>;
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