module Shockout {

    // recreate the SP REST object for an attachment
    export class SpAttachment implements ISpAttachment {
        __metadata: ISpAttachmentMetadata;
        EntitySet: string;
        ItemId: number;
        Name: string;

        constructor(rootUrl: string, siteUrl: string, listName: string, itemId: number, fileName: string) {
            var entitySet: string = listName.replace(/\s/g, '');
            siteUrl = Utils.formatSubsiteUrl(siteUrl);
            var uri = `${rootUrl + siteUrl}_vti_bin/listdata.svc/Attachments(EntitySet='${entitySet}',ItemId=${itemId},Name='${fileName}')`;

            this.__metadata = {
                uri: uri,
                content_type: "application/octetstream",
                edit_media: uri + "/$value",
                media_etag: null, // this property is unused for our purposes, so `null` is fine for now
                media_src: `${rootUrl + siteUrl}/Lists/${listName}/Attachments/${itemId}/${fileName}`,
                type: "Microsoft.SharePoint.DataService.AttachmentsItem"
            };
            this.EntitySet = entitySet;
            this.ItemId = itemId;
            this.Name = fileName;
        }
    }

    export interface ICafe {
        asyncFns: Array<Function>;
        start(msg?: string): void;
        complete(fn: Function): ICafe;
        fail(fn: Function): ICafe;
        finally(fn: Function): ICafe;
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

        public start(msg: string = undefined): void {
            this.next(true, msg);
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

            if (this._complete) {
                this._complete(msg, success, args);
            }

            if (!success) {
                if (this._fail) {
                    this._fail(msg, success, args);
                }
                return;
            }

            if (this.asyncFns.length == 0) {
                if (this._finally) {
                    this._finally(msg, success, args);
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
            this.className = ko.observable('progress-bar progress-bar-info progress-bar-striped active');
            this.getProgress = ko.pureComputed(function () {
                return self.fileName() + ' ' + self.progress() + '%';
            }, this);
        }
    }

    /**
     * Date Time Model Interface
     */
    export interface IDateTimeModel{
        $element: JQuery;
        $parent: JQuery;
        $hh: JQuery;
        $mm: JQuery;
        $tt: JQuery;
        $display: JQuery;
        $error: JQuery;
        $reset: JQuery;
        required: Boolean;
        koName: string;
        setModelValue: Function;
        setDisplayValue: Function;
        toString: Function;
    }

    /**
     * Date Time Model Class
     */
    export class DateTimeModel implements IDateTimeModel {
        public $element: JQuery;
        public $parent: JQuery;
        public $hh: JQuery;
        public $mm: JQuery;
        public $tt: JQuery;
        public $display: JQuery;
        public $error: JQuery;
        public $reset: JQuery;
        public required: Boolean;
        public koName: string;

        constructor(element: HTMLElement, modelValue: KnockoutObservable<Date>) {
            var self = this;
            var date = ko.unwrap(modelValue);
            this.$element = $(element);
            this.$parent = this.$element.parent();

            if (Utils.isJsonDateTicks(date)) {
                modelValue(Utils.parseJsonDate(date));
            }

            this.required = this.$element.hasClass('required') || this.$element.attr('required') != null;
            this.$element.attr({
                'placeholder': 'MM/DD/YYYY',
                'maxlength': 10,
                'class': 'datepicker med form-control'
            }).css('display', 'inline-block')
                .on('change', onChange)
                .datepicker({
                    changeMonth: true,
                    changeYear: true
                })
            ;

            if (this.required) {
                this.$element.attr('required', '');
            }

            this.$element.after(Templates.getTimeControlsHtml());
            this.$display = this.$parent.find('.so-datetime-display');
            this.$error = this.$parent.find('.error');
            this.$hh = this.$parent.find('.so-select-hours').val('12').on('change', onChange);
            this.$mm = this.$parent.find('.so-select-minutes').val('0').on('change', onChange);
            this.$tt = this.$parent.find('.so-select-tt').val('AM').on('change', onChange);
            this.$reset = this.$parent.find('.btn.reset')
                .on('click', function () {
                    try {
                        self.$element.val('');
                        self.$hh.val('12');
                        self.$mm.val('0');
                        self.$tt.val('AM');
                        self.$display.html('');
                        modelValue(null);
                    }
                    catch (e) {
                        console.warn(e);
                    }
                    return false;
                });

            function onChange() {
                self.setModelValue(modelValue);
            }
        }

        public setModelValue(modelValue: KnockoutObservable<Date>): void {
            try {

                var val = this.$element.val();
                if (!!!$.trim(this.$element.val())) {
                    return;
                }

                var hrs: number = parseInt(this.$hh.val());
                var min: number = parseInt(this.$mm.val());
                var tt: string = this.$tt.val();

                if (tt == 'PM' && hrs < 12) {
                    hrs += 12;
                }
                else if (tt == 'AM' && hrs > 11) {
                    hrs -= 12;
                }

                // SP saves date/time in UTC
                var curDateTime: Date = new Date(val);
                curDateTime.setUTCHours(hrs, min, 0, 0);
                modelValue(curDateTime);
            }
            catch (e) {
                if (SPForm.DEBUG) {
                    throw e;
                }
            }
        }

        public setDisplayValue(modelValue: KnockoutObservable<Date>): void {
            try {
                var date = ko.unwrap(modelValue);

                if (!!!date || date.constructor != Date) { return; }

                var date = ko.unwrap(modelValue);
                var hrs: number = date.getUTCHours(); // converts UTC hours to locale hours
                var min: number = date.getUTCMinutes();

                this.$element.val((date.getUTCMonth() + 1) + '/' + date.getUTCDate() + '/' + date.getUTCFullYear());

                // set TT based on military hours
                if (hrs > 12) {
                    hrs -= 12;
                    this.$tt.val('PM');
                }
                else if (hrs == 0) {
                    hrs = 12;
                    this.$tt.val('AM');
                }
                else if (hrs == 12) {
                    this.$tt.val('PM');
                }
                else {
                    this.$tt.val('AM');
                }

                this.$hh.val(hrs.toString());
                this.$mm.val(min.toString());
                this.$display.html(this.toString(modelValue));
            }
            catch (e) {
                if (SPForm.DEBUG) {
                    throw e;
                }
            }
        }

        public toString(modelValue: KnockoutObservable<Date>): string {
            return DateTimeModel.toString(modelValue);
        }

        public static toString(modelValue: KnockoutObservable<Date>): string {
            // convert from UTC to locale
            try {
                var date = ko.unwrap(modelValue);
                if (date == null || date.constructor != Date) { return; }
                var dateTimeStr: string = Utils.toDateTimeLocaleString(date);
                // add time zone
                var timeZone = /\b\s\(\w+\s\w+\s\w+\)/i.exec(date.toString());
                if (!!timeZone) {
                    // e.g. convert '(Central Daylight Time)' to '(CDT)'
                    dateTimeStr += ' ' + timeZone[0].replace(/\b\w+/g, function (x) {
                            return x[0];
                        }).replace(/\s/g, '');
                }
                return dateTimeStr;
            }
            catch (e) {
                if (SPForm.DEBUG) {
                    throw e;
                }
            }
            return null;
        }
    }

}