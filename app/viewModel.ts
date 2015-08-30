module Shockout {
    
    export interface IViewModel {
        Title: KnockoutObservable<string>;
        CreatedBy: KnockoutObservable<ISpPerson>;
        ModifiedBy: KnockoutObservable<ISpPerson>;

        parent: any;
        history: KnockoutObservable<Array<any>>;
        attachments: KnockoutObservable<Array<any>>;
        isAuthor: KnockoutObservable<boolean>
        currentUser: KnockoutObservable<any>;

        deleteItem(): void;
        cancel(): void;
        print(): void;
        deleteAttachment(obj: any, event: any): boolean;
        save(model: ViewModel, btn: HTMLElement): void;
        submit(model: ViewModel, btn: HTMLElement): void;
    }

    export class ViewModel implements IViewModel {
        public Title: KnockoutObservable<string> = ko.observable(null);
        public CreatedBy: KnockoutObservable<ISpPerson> = ko.observable(null);
        public ModifiedBy: KnockoutObservable<ISpPerson> = ko.observable(null);

        public parent: ShockoutForm;
        public history: KnockoutObservable<Array<any>> = ko.observableArray([]);
        public attachments: KnockoutObservableArray<IAttachment> = ko.observableArray([]);
        public isAuthor: KnockoutObservable<boolean> = ko.observable(false);
        public isValid: KnockoutObservable<boolean> = ko.observable(false);
        public currentUser: KnockoutObservable<ICurrentUser> = ko.observable(null);
        
        constructor(instance: ShockoutForm) {
            this.parent = instance;
        }

        public deleteItem(): void {
            this.parent.deleteListItem(this);
        }

        public cancel(): void {
            window.location.href = this.parent.sourceUrl != null ? this.parent.sourceUrl : this.parent.rootUrl;
        }

        public print(): void {
            window.print();
        }

        public deleteAttachment(obj: any, event: any): boolean {
            this.parent.deleteAttachment(obj);
            return false;
        }

        public save(model: ViewModel, btn: HTMLElement): void {
            this.parent.saveListItem(model, false);
        }

        public submit(model: ViewModel, btn: HTMLElement): void {
            this.parent.saveListItem(model, true);
        }

    }

}