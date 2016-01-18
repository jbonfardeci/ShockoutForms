module Shockout {

    export class KoHandlers {
        public static bindKoHandlers() {
            bindKoHandlers(ko);
        }
    }

    interface IDateTimeModel{
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

    class DateTimeModel implements IDateTimeModel {
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

            var hrsOpts = [];
            for (var i = 1; i <= 12; i++) {
                hrsOpts.push('<option value="' + i + '">' + (i < 10 ? '0' + i : i) + '</option>');
            }

            var mmOpts = [];
            for (var i = 0; i < 60; i++) {
                mmOpts.push('<option value="' + i + '">' + (i < 10 ? '0' + i : i) + '</option>');
            }

            var timeHtml: string = `<span class="glyphicon glyphicon-calendar"></span>
                <select class="form-control so-select-hours" style="margin-left:1em; max-width:5em; display:inline-block;">${hrsOpts.join('')}</select>
                <span> : </span>
                <select class="form-control so-select-minutes" style="width:5em; display:inline-block;">${mmOpts.join('')}</select>
                <select class="form-control so-select-tt" style="margin-left:1em; max-width:5em; display:inline-block;"><option value="AM">AM</option><option value="PM">PM</option></select>
                <button class="btn btn-sm btn-default reset" style="margin-left:1em;">Reset</button>
                <span class="error no-print" style="display:none;">Invalid Date-time</span>
                <span class="so-datetime-display no-print" style="margin-left:1em;"></span>`;

            this.$element.after(timeHtml);
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

                if (!!!$.trim(this.$element.val())) {
                    return;
                }

                var date: Date = Utils.parseDate(this.$element.val());
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
                var curDateTime: Date = new Date();
                curDateTime.setUTCFullYear(date.getFullYear());
                curDateTime.setUTCMonth(date.getMonth());
                curDateTime.setUTCDate(date.getDate());
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

    /* Knockout Custom handlers */
    function bindKoHandlers(ko) {

        ko.bindingHandlers['spHtmlEditor'] = {
            init: function (element: HTMLElement, valueAccessor: Function, allBindings: Function, vm: IViewModel) {
                var koName: string = Utils.observableNameFromControl(element);

                $(element)
                    .blur(update)
                    .change(update)
                    .keydown(update);

                function update(): void {
                    vm[koName]($(this).html());                 
                }
            }
            , update: function (element: HTMLElement, valueAccessor: Function, allBindings: Function, vm: IViewModel) {
                var value = ko.utils.unwrapObservable(valueAccessor()) || "";
                if (element.innerHTML !== value) {
                    element.innerHTML = value;
                }
            }
        }

        /* SharePoint People Picker */
        ko.bindingHandlers['spPerson'] = {
            init: function (element, valueAccessor, allBindings, bindingContext) {
                try {
                    // stop if not an editable field 
                    if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') { return; }

                    // This will be called when the binding is first applied to an element
                    // Set up any initial state, event handlers, etc. here
                    var viewModel = bindingContext.$data
                        , modelValue = valueAccessor()
                        , person = ko.unwrap(modelValue)
                        ;

                    var $element = $(element)
                        .addClass('people-picker-control')
                        .attr('placeholder', 'Employee Account Name');

                    //create wrapper for control
                    var $parent = $(element).parent();

                    var $spError = $('<div>', { 'class': 'sp-validation person' });
                    $element.after($spError);

                    var $desc = $('<div>', {
                        'class': 'no-print'
                        , 'html': '<em>Enter the employee name. The auto-suggest menu will appear below the field. Select the account name.</em>'
                    });
                    $spError.after($desc);

                    //controls
                    var $spValidate = $('<button>', {
                        'html': '<span class="glyphicon glyphicon-user"></span>',
                        'class': 'btn btn-sm btn-default no-print',
                        'title': 'Validate the employee account name.'
                    }).on('click', function () {
                        if ($.trim($element.val()) == '') {
                            $element.removeClass('invalid').removeClass('valid');
                            return false;
                        }

                        if (!Utils.validateSpPerson(modelValue())) {
                            $spError.text('Invalid').addClass('error').show();
                            $element.addClass('invalid').removeClass('valid');
                        }
                        else {
                            $spError.text('Valid').removeClass('error');
                            $element.removeClass('invalid').addClass('valid').show();
                        }
                        return false;
                    }).insertAfter($element);

                    var $reset = $('<button>', { 'class': 'btn btn-sm btn-default reset', 'html': 'Reset' })
                        .on('click', function () {
                            modelValue(null);
                            return false;
                        })
                        .insertAfter($spValidate);
                    
                    var autoCompleteOpts: any = {
                        source: function (request, response) {

                            // Use People.asmx instead of REST services against the User Information List, 
                            // which allows you to search users that haven't logged into SharePoint yet.
                            // Thanks to John Kerski from Definitive Logic for the suggestion.
                            SpSoap.searchPrincipals(request.term, function (data: Array<IPrincipalInfo>) {
                                response($.map(data, function (item: IPrincipalInfo) {
                                    return {
                                        label: item.DisplayName + ' (' + item.Email + ')',
                                        value: item.UserInfoID + ';#' + item.AccountName
                                    }
                                }));
                            }, 10, 'User');
                        },
                        minLength: 3,
                        select: function (event, ui) {
                            modelValue(ui.item.value);
                        }
                    };

                    $(element).autocomplete(autoCompleteOpts);
                    $(element).on('focus', function () { $(this).removeClass('valid'); })
                        .on('blur', function () { onChangeSpPersonEvent(this, modelValue); })
                        .on('mouseout', function () { onChangeSpPersonEvent(this, modelValue); });
                }
                catch (e) {
                    if (SPForm.DEBUG) {
                        console.info('Error in Knockout handler spPerson init()');
                        console.info(e);
                    }
                }

                function onChangeSpPersonEvent(self, modelValue) {
                    var value = $.trim($(self).val());
                    if (value == '') {
                        modelValue(null);
                        $(self).removeClass('valid').removeClass('invalid');
                        return;
                    }

                    if (Utils.validateSpPerson(modelValue())) {
                        $(self).val(modelValue().split('#')[1]);
                        $(self).addClass('valid').removeClass('invalid');
                    }
                    else {
                        $(self).removeClass('valid').addClass('invalid');
                    }
                };
            },
            update: function (element, valueAccessor, allBindings, bindingContext) {
                // This will be called once when the binding is first applied to an element,
                // and again whenever any observables/computeds that are accessed change
                // Update the DOM element based on the supplied values here.
                try {
                    var viewModel = bindingContext.$data;

                    // First get the latest data that we're bound to
                    var modelValue = valueAccessor();

                    // Next, whether or not the supplied model property is observable, get its current value
                    var person = ko.unwrap(modelValue);

                    // Now manipulate the DOM element
                    var displayName = "";
                    if (Utils.validateSpPerson(person)) {
                        displayName = person.split('#')[1];
                        $(element).addClass("valid");
                    }

                    if ('value' in element) {
                        $(element).val(displayName);
                    } else {
                        $(element).text(displayName);
                    }
                }
                catch (e) {
                    if (SPForm.DEBUG) {
                        console.info('Error in Knockout handler spPerson update()');
                        console.info(e);
                    }
                }
            }
        };

        ko.bindingHandlers['spMoney'] = {
            'init': function (element, valueAccessor, allBindings, viewModel: IViewModel, bindingContext) {

                /* stop if not an editable field */
                if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') { return; }

                viewModel = bindingContext.$data;
                var value = valueAccessor();
                var valueUnwrapped = ko.unwrap(value);

                $(element).on('blur', onChange).on('change', onChange);

                function onChange() {
                    var val = this.value.toString().replace(/[^\d\.\-]/g, '');
                    val = val == '' ? null : (val - 0);
                    value(val);
                };
            },
            'update': function (element, valueAccessor, allBindings, viewModel, bindingContext) {
                viewModel = bindingContext.$data;
                var value = valueAccessor();
                var valueUnwrapped = ko.unwrap(value);

                if (valueUnwrapped != null) {
                    if (valueUnwrapped < 0) {
                        $(element).addClass('negative');
                    } else {
                        $(element).removeClass('negative');
                    }
                } else {
                    valueUnwrapped = 0;
                }

                var formattedValue = Utils.formatMoney(valueUnwrapped);
                Utils.updateKoField(element, formattedValue);
            }
        };

        ko.bindingHandlers['spDecimal'] = {
            'init': function (element, valueAccessor, allBindings, viewModel, bindingContext) {

                // stop if not an editable field 
                if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') { return; }

                viewModel = bindingContext.$data;
                var value = valueAccessor();
                var valueUnwrapped = ko.unwrap(value);

                $(element).on('blur', onChange).on('change', onChange);

                function onChange() {
                    var val = this.value.toString().replace(/[^\d\-\.]/g, '');
                    val = val == '' ? null : (val - 0);
                    value(val);
                };
            },
            'update': function (element, valueAccessor, allBindings, viewModel, bindingContext) {
                viewModel = bindingContext.$data;
                var value = valueAccessor();
                var valueUnwrapped = ko.unwrap(value);
                var precision = allBindings.get('precision') || 2;

                var formattedValue = Utils.toFixed(valueUnwrapped, precision);

                if (valueUnwrapped != null) {
                    if (valueUnwrapped < 0) {
                        $(element).addClass('negative');
                    } else {
                        $(element).removeClass('negative');
                    }
                } else {
                    valueUnwrapped = 0;
                }

                Utils.updateKoField(element, formattedValue);
            }
        };

        ko.bindingHandlers['spNumber'] = {
            /* executes on load */
            init: function (element, valueAccessor, allBindings, viewModel, bindingContext) {

                /* stop if not an editable field */
                if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') { return; }

                viewModel = bindingContext.$data;
                var value = valueAccessor();
                var valueUnwrapped = ko.unwrap(value);

                $(element).on('blur', onChange).on('change', onChange);

                function onChange() {
                    var val = this.value.toString().replace(/[^\d\-]/g, '');
                    val = val == '' ? null : (val - 0);
                    value(val);
                };
            },
            /* executes on load and on change */
            update: function (element, valueAccessor, allBindings, viewModel, bindingContext) {
                viewModel = bindingContext.$data;
                var value = valueAccessor();
                var valueUnwrapped = ko.unwrap(value);

                valueUnwrapped = valueUnwrapped == null ? 0 : valueUnwrapped;
                valueUnwrapped = valueUnwrapped.constructor == String ? valueUnwrapped = valueUnwrapped.replace(/\D/g) - 0 : valueUnwrapped;

                Utils.updateKoField(element, valueUnwrapped);

                if (value.constructor == Function) {
                    value(valueUnwrapped);
                }
            }
        };

        ko.bindingHandlers['spDate'] = {
            after: ['attr'],
            init: function (element, valueAccessor, allBindings, viewModel: IViewModel, bindingContext) {
                var modelValue: KnockoutObservable<Date> = valueAccessor();

                if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') { return; }// stop if not an editable field

                $(element)
                    .css('display', 'inline-block')
                    .addClass('datepicker med')
                    .attr('placeholder', 'MM/DD/YYYY')
                    .on('blur', onDateChange)
                    .on('change', onDateChange)
                    .after('<span class="glyphicon glyphicon-calendar"></span>');

                $(element).datepicker({
                    changeMonth: true,
                    changeYear: true
                });

                function onDateChange() {
                    modelValue(Utils.parseDate(this.value));
                };
            },
            update: function (element, valueAccessor, allBindings, viewModel: IViewModel, bindingContext) {
                var modelValue: KnockoutObservable<Date> = valueAccessor();
                var date: Date = Utils.parseDate(ko.unwrap(modelValue));
                var dateStr = '';

                if (!!date && date != null) {
                    dateStr = Utils.dateToLocaleString(date);
                }

                if ('value' in element) {
                    $(element).val(dateStr);
                } else {
                    $(element).text(dateStr);
                }
            }
        };

        // 1. REST returns UTC
        // 2. getUTCHours converts UTC to Locale
        ko.bindingHandlers['spDateTime'] = {
            after: ['attr'],
            init: function (element, valueAccessor, allBindings, viewModel: IViewModel, bindingContext) {
                // stop if not an editable field
                if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden'){ return; }

                try {
                    var modelValue: KnockoutObservable<Date> = valueAccessor();
                    var model = new DateTimeModel(element, modelValue); 
                    element['$$model'] = model;           
                }
                catch (e) {
                    if (SPForm.DEBUG) {
                        console.warn('Error in Knockout handler spDateTime init()...');
                        throw e;
                    }
                }
            },
            update: function (element, valueAccessor, allBindings, viewModel: IViewModel, bindingContext) {
                try {
                    var modelValue: KnockoutObservable<Date> = valueAccessor();
                    var date: Date = Utils.parseDate(ko.unwrap(modelValue));
                    if (element.tagName.toLowerCase() == 'input') {
                        var model: IDateTimeModel = element['$$model'];
                        if (!!model && model.constructor == DateTimeModel) {
                            model.setDisplayValue(modelValue);
                        }
                    }
                    else {
                        $(element).text(DateTimeModel.toString(modelValue));
                    }   
                }
                catch (e) {
                    if (SPForm.DEBUG) {
                        console.warn('Error in Knockout handler spDateTime update()...s');
                        throw e;
                    }
                }
            }
        };
    }
}