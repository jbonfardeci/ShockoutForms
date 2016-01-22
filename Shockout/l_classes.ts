module Shockout {

    export interface ICafe {
        asyncFns: Array<Function>;
        complete(fn: Function);
        fail(fn: Function);
        finally(fn: Function);
        next(success?: boolean, msg?: string, args?: any): void;
    }

    /**
     * CAFE - Cascading Asynchronous Function Execution. 
     * A class to control the sequential execution of asynchronous functions.
     * by John Bonfardeci <john.bonfardeci@gmail.com> 2014
     * @param {Array<Function>} asyncFns
     * @returns
     */
    export class Cafe {

        private _complete: Function;
        private _fail: Function;
        private _finally: Function;

        public asyncFns: Array<Function>;

        constructor(asyncFns: Array<Function> = undefined) {
            if (asyncFns) {
                this.asyncFns = asyncFns;
            }
            return this;
        }

        public complete(fn: Function): ICafe {
            this._complete = fn;
            return this;
        };

        public fail(fn: Function): ICafe {
            this._fail = fn;
            return this;
        }

        public finally(fn: Function): ICafe {
            this._finally = fn;
            return this;
        }

        public next(success: boolean = true, msg: string = undefined, args: any = undefined): void {

            if (!this.asyncFns) {
                throw "Error in Cafe: The required parameter `asyncFns` of type (Array<Function>) is undefined. Don't forget to instantiate Cafe with this parameter or set its value after instantiation.";
            }

            if (!success) {
                if (this._fail) {
                    this._fail(arguments);
                }
                return;
            }
            
            if (this._complete) {
                this._complete(arguments);
            }

            if (this.asyncFns.length == 0) {
                if (this._finally) {
                    this._finally(arguments);
                }
                return;
            }

            // execute the next function in the array
            this.asyncFns.shift()(this, args);
        }

    }

    /**
     * IFileUpload Interface
     * Interface for upload progress indicator for a Knockout observable array. 
     * @param {string} fileName
     * @param {number} bytes
     */
    export interface IFileUpload {
        label: KnockoutObservable<string>;
        progress: KnockoutObservable<number>;
        fileName: KnockoutObservable<string>;
        kb: KnockoutObservable<number>;
        className: KnockoutObservable<string>;
        getProgress: KnockoutComputed<string>;
    }

    /**
     * FileUpload Class
     * Creates an upload progress indicator for a Knockout observable array. 
     * @param {string} fileName
     * @param {number} bytes
     */
    export class FileUpload implements IFileUpload {

        public label: KnockoutObservable<string>;
        public progress: KnockoutObservable<number>;
        public fileName: KnockoutObservable<string>;
        public kb: KnockoutObservable<number>;
        public className: KnockoutObservable<string>;
        public getProgress: KnockoutComputed<string>;

        constructor(fileName: string, bytes: number) {
            var self = this;
            this.label = ko.observable(null);
            this.progress = ko.observable(0);
            this.fileName = ko.observable(fileName);
            this.kb = ko.observable((bytes / 1024));
            this.className = ko.observable('progress-bar progress-bar-success progress-bar-striped active');
            this.getProgress = ko.pureComputed(function () {
                return self.fileName() + ' ' + self.progress() + '%';
            }, this);
        }
    }

}