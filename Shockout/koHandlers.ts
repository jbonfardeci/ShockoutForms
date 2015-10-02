module Shockout {

    export class KoHandlers {

        public static bindKoHandlers() {
            bindKoHandlers(ko);
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
                    if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') { return; }/*stop if not an editable field */

                    // This will be called when the binding is first applied to an element
                    // Set up any initial state, event handlers, etc. here
                    var viewModel = bindingContext.$data
                        , modelValue = valueAccessor()
                        , person = ko.unwrap(modelValue)
                        ;

                    var $element = $(element);
                    $element.addClass('people-picker-control');
                    $element.attr('placeholder', 'Employee Account Name'); //.addClass('people-picker-control');

                    //create wrapper for control
                    var $parent = $(element).parent();

                    var $spError = $('<div>', { 'class': 'sp-validation person' }).appendTo($parent);

                    var $desc = $('<div>', {
                        'class': 'no-print'
                        , 'html': '<em>Enter the employee name. The auto-suggest menu will appear below the field. Select the account name.</em>'
                    }).appendTo($parent);

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
                    
                    var autoCompleteOpts: any = {
                        source: function (request, response) {
                            SpApi.peopleSearch(request.term, function (data: Array<ISpPersonSearchResult>) {
                                response($.map(data, function (item) {
                                    var email: string = item['EMail'] || item['WorkEMail']; // SP 2013 vs SP 2010 Email key name.
                                    var name: string = item['Name'] || item['Account'];
                                    return {
                                        label: item.Name + ' (' + email + ')',
                                        value: item.Id + ';#' + name
                                    }
                                }));
                            });
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

        ko.bindingHandlers['spDate'] = {
            init: function (element, valueAccessor, allBindings, viewModel: IViewModel, bindingContext) {
                var modelValue: KnockoutObservable<Date> = valueAccessor();

                if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') { return; }// stop if not an editable field

                $(element)
                    .datepicker()
                    .addClass('datepicker med')
                    .attr('placeholder', 'MM/DD/YYYY')
                    .on('blur', onDateChange)
                    .on('change', onDateChange);

                function onDateChange() {
                    modelValue(Utils.parseDate(this.value));
                };
            },
            update: function (element, valueAccessor, allBindings, viewModel: IViewModel, bindingContext) {
                var modelValue: KnockoutObservable<Date> = valueAccessor();
                var date: Date = Utils.parseDate( ko.unwrap(modelValue) );
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

        // 1. REST returns UTC
        // 2. getUTCHours converts UTC to Locale
        ko.bindingHandlers['spDateTime'] = {
            init: function (element, valueAccessor, allBindings, viewModel: IViewModel, bindingContext) {

                if (element.tagName.toLowerCase() != 'input' || $(element).attr('type') == 'hidden') { return; }// stop if not an editable field

                var modelValue: KnockoutObservable<Date> = valueAccessor()
                    , required
                    , $hh: JQuery
                    , $mm: JQuery
                    , $tt: JQuery
                    , $display
                    , $error
                    , $element: JQuery = $(element)
                    , $parent: JQuery = $element.parent()
                    ;

                try {

                    var currentVal: Date = Utils.parseDate(modelValue());
                    modelValue(currentVal); // just in case it was a string date
                    var koName = Utils.koNameFromControl(element);

                    $display = $('<span>', { 'class': 'no-print', 'style': 'margin-left:1em;' }).insertAfter($element);
                    $error = $('<span>', { 'class': 'error', 'html': 'Invalid Date-time', 'style': 'display:none;' }).insertAfter($element);
                    element.$display = $display;
                    element.$error = $error;

                    required = $element.hasClass('required') || $element.attr('required') != null;

                    $element.attr({
                        'placeholder': 'MM/DD/YYYY',
                        'maxlength': 10,
                        'class': 'datepicker med form-control'
                    }).css('display', 'inline-block').datepicker().on('change', function () {
                        try {
                            $error.hide();
                            var date: Date = Utils.parseDate(this.value);
                            modelValue(date);
                            $display.html(Utils.toDateTimeLocaleString(date));
                        }
                        catch (e) {
                            $error.show();
                        }
                    });

                    if (required) {
                        $element.attr('required', 'required');
                    }

                    var timeHtml: Array<string> = ['<span class="glyphicon glyphicon-calendar" style="margin-left:.2em;"></span>'];

                    // Hours 
                    timeHtml.push('<select class="form-control select-hours" style="margin-left:1em;width:5em;display:inline-block;">');
                    for (var i = 1; i <= 12; i++) {
                        timeHtml.push('<option value="' + i + '">' + (i < 10 ? '0' + i : i) + '</option>');
                    }
                    timeHtml.push('</select>');

                    timeHtml.push('<span> : </span>');
                    
                    // Minutes     
                    timeHtml.push('<select class="form-control select-minutes" style="width:5em;display:inline-block;">');
                    for (var i = 0; i < 60; i++) {
                        timeHtml.push('<option value="' + i +'">' + (i < 10 ? '0'+i : i) + '</option>');
                    }
                    timeHtml.push('</select>');
                         
                    // TT: AM/PM
                    timeHtml.push('<select class="form-control select-tt" style="margin-left:1em;width:5em;display:inline-block;"><option value="AM">AM</option><option value="PM">PM</option></select>');

                    $element.after(timeHtml.join(''));

                    $hh = $parent.find('.select-hours');
                    $mm = $parent.find('.select-minutes');
                    $tt = $parent.find('.select-tt');

                    $hh.on('change', setDateTime);
                    $mm.on('change', setDateTime);
                    $tt.on('change', setDateTime);

                    element.$hh = $hh;
                    element.$mm = $mm;
                    element.$tt = $tt;

                    // set default time
                    if (!!currentVal) {
                        setDateTime();
                    }
                    else {
                        $element.val('');
                        $hh.val('12');
                        $mm.val('0');
                        $tt.val('AM');
                    }

                }
                catch (e) {
                    if (SPForm.DEBUG) {
                        console.warn('Error in Knockout handler spDateTime init()...s');
                        console.warn(e);
                    }
                }

                // must conver user's locale date/time to UTC for SP
                function setDateTime(): void {
                    try {
                        var date: Date = Utils.parseDate($element.val());
                        if (!!!date) {
                            date = new Date();
                        }
                        var hrs: number = parseInt($hh.val());
                        var min: number = parseInt($mm.val());
                        var tt: string = $tt.val();

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
                            console.warn(e);
                        }
                    }
                }

            },
            update: function (element, valueAccessor, allBindings, viewModel: IViewModel, bindingContext) {

                try {
                    var modelValue: KnockoutObservable<Date> = valueAccessor();
                    var date: Date = Utils.parseDate(ko.unwrap(modelValue));

                    if (typeof modelValue == 'function') {
                        modelValue(date); // just in case it was a string date 
                    }

                    if (!!date) {
                        var dateTimeStr: string = Utils.toDateTimeLocaleString(date); // convert from UTC to locale
                        // add time zone
                        dateTimeStr += /\b\s\(\w+\s\w+\s\w+\)/i.exec(date.toString())[0];
                        
                        if (element.tagName.toLowerCase() == 'input') {
                            element.value = (date.getUTCMonth()+1) + '/' + date.getUTCDate() + '/' + date.getUTCFullYear();
                            var hrs: number = date.getUTCHours(); // converts UTC hours to locale hours
                            var min: number = date.getUTCMinutes(); 

                            // set TT based on military hours
                            if (hrs > 12) {
                                hrs -= 12;
                                element.$tt.val('PM');
                            }
                            else if (hrs == 0) {
                                hrs = 12;
                                element.$tt.val('AM');
                            }
                            else if (hrs == 12) {
                                element.$tt.val('PM');
                            }
                            else {
                                element.$tt.val('AM');
                            }

                            element.$hh.val(hrs);
                            element.$mm.val(min);
                            element.$display.html(dateTimeStr);
                        }
                        else {
                            element.innerHTML = dateTimeStr;
                        }
                    }
                }
                catch (e) {
                    if (SPForm.DEBUG) {
                        console.warn('Error in Knockout handler spDateTime update()...s');
                        console.warn(e);
                    }
                }
            }
        };
    }
}