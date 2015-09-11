module Shockout {
    
    export interface IViewModel {
        Id: KnockoutObservable<number>;
        CreatedBy: KnockoutObservable<ISpPerson>;
        CreatedByName: KnockoutObservable<string>;
        CreatedByEmail: KnockoutObservable<string>;
        ModifiedBy: KnockoutObservable<ISpPerson>;
        ModifiedByName: KnockoutObservable<string>;
        ModifiedByEmail: KnockoutObservable<string>;

        parent: any;
        history: KnockoutObservable<Array<any>>;
        attachments: KnockoutObservable<Array<any>>;
        isAuthor: KnockoutObservable<boolean>
        currentUser: KnockoutObservable<any>;
        isValid: KnockoutComputed<boolean>;

        deleteItem(): void;
        cancel(): void;
        print(): void;
        deleteAttachment(obj: any, event: any): boolean;
        save(model: ViewModel, btn: HTMLElement): void;
        submit(model: ViewModel, btn: HTMLElement): void;
    }

    export class ViewModel implements IViewModel {

        public static createdByKey: string = 'CreatedByName';
        public static createdByEmailKey: string = 'CreatedByEmail';
        public static modifiedByKey: string = 'ModifiedByName';
        public static modifiedByEmailKey: string = 'ModifiedByEmail';
        public static createdKey: string = 'Created';
        public static modifiedKey: string = 'Modified';
        public static historyKey: string = 'history';
        public static historyDescriptionKey: string = 'description';
        public static historyDateKey: string = 'date';
        public static isSubmittedKey: string;
        public static parent: SPForm;

        public Id: KnockoutObservable<number> = ko.observable(null);
        public CreatedBy: KnockoutObservable<ISpPerson> = ko.observable(null);
        public CreatedByName: KnockoutObservable<string> = ko.observable(null);
        public CreatedByEmail: KnockoutObservable<string> = ko.observable(null);
        public ModifiedBy: KnockoutObservable<ISpPerson> = ko.observable(null);
        public ModifiedByName: KnockoutObservable<string> = ko.observable(null);
        public ModifiedByEmail: KnockoutObservable<string> = ko.observable(null);

        public parent: Shockout.SPForm;
        public history: KnockoutObservable<Array<any>> = ko.observableArray([]);
        public attachments: KnockoutObservableArray<ISpAttachment> = ko.observableArray([]);
        public isAuthor: KnockoutObservable<boolean> = ko.observable(false);
        public currentUser: KnockoutObservable<ICurrentUser> = ko.observable(null);
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