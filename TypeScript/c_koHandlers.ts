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
                        'html': Templates.personIcon,
                        'class': Templates.buttonDefault,
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

                    var $reset = $('<button>', { 'class': Templates.resetButton, 'html': 'Reset' })
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
                    .after(Templates.calendarIcon);

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
                        console.warn('Error in Knockout handler spDateTime update()...');
                        throw e;
                    }
                }
            }
        };
    }
}