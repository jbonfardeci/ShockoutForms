module Shockout {

    /* Knockout Custom handlers */
    (function bindKoHandlers(ko) {

        //http://stackoverflow.com/questions/7904522/knockout-content-editable-custom-binding?lq=1
        ko.bindingHandlers['htmlValue'] = {
            init: function (element, valueAccessor, allBindingsAccessor) {
                ko.utils.registerEventHandler(element, "blur", update);
                ko.utils.registerEventHandler(element, "keydown", update);
                ko.utils.registerEventHandler(element, "change", update);
                ko.utils.registerEventHandler(element, "mousedown", update);
                function update() {
                    var modelValue = valueAccessor();
                    var elementValue = element.innerHTML;
                    if (ko.isWriteableObservable(modelValue)) {
                        modelValue(elementValue);
                    }
                    else { //handle non-observable one-way binding
                        var allBindings = allBindingsAccessor();
                        if (allBindings['_ko_property_writers'] && allBindings['_ko_property_writers'].htmlValue) {
                            allBindings['_ko_property_writers'].htmlValue(elementValue);
                        }
                    }
                }
            },
            update: function (element, valueAccessor) {
                var value = ko.utils.unwrapObservable(valueAccessor()) || "";
                if (element.innerHTML !== value) {
                    element.innerHTML = value;
                }
            }
        };

        /* SharePoint People Picker */
        ko.bindingHandlers['spPerson'] = {
            init: function (element, valueAccessor, allBindings, bindingContext) {
                try {
                    if (element.tagName.toLowerCase() != "input" || $(element).attr("type") == "hidden") { return; }/*stop if not an editable field */

                    // This will be called when the binding is first applied to an element
                    // Set up any initial state, event handlers, etc. here
                    var viewModel = bindingContext.$data
                        , modelValue = valueAccessor()
                        , person = ko.unwrap(modelValue)
                        ;

                    $(element).attr("placeholder", "Employee Account Name").addClass("people-picker-control");

                    //create wrapper for control
                    var $parent = $(element).parent();

                    //controls
                    var $spValidate = $("<button>", { "html": "<span>Validate</span>", "class": "sp-validate-person", "title": "Validate the employee account name." }).on("click", function () {
                        if ($.trim($(element).val()) == "") {
                            $(element).removeClass("invalid").removeClass("valid");
                            return false;
                        }

                        if (!validateSpPerson(modelValue())) {
                            $spError.text("Invalid").addClass("error");
                            $(element).addClass("invalid").removeClass("valid");
                        }
                        else {
                            $spError.text("Valid").removeClass("error");
                            $(element).removeClass("invalid").addClass("valid");
                        }
                        return false;
                    });
                    $parent.append($spValidate);

                    /*var $spLookup = $("<button>", { "html": "<span>Lookup</span>", "class": "sp-lookup-person" }).on("click", function () {
                        return false;
                    });
                    $parent.append($spLookup);
                    */

                    var $spError = $("<span>", { "class": "sp-validation person" });
                    $parent.append($spError);

                    var $desc = $('<div>', { 'class': 'no-print', 'html': '<em>Enter the employee name. The auto-suggest menu will appear below the field. Select the account name.</em>' });
                    $parent.append($desc);

                    $(element).autocomplete({
                        source: function (request, response) {
                            $.ajax({
                                url: "/_layouts/webster/SPUserAutoComplete.ashx",
                                dataType: "json",
                                data: { term: request.term },
                                success: function (data) {
                                    response($.map(data, function (item) {
                                        return {
                                            label: item.name,
                                            value: item.id + ';#' + item.name
                                        }
                                    }));
                                }
                            });
                        },
                        minLength: 3,
                        select: function (event, ui) {
                            modelValue(ui.item.value);
                        }
                    })
                        .on("focus", function () { $(this).removeClass("valid"); })
                        .on("blur", function () { onChangeSpPersonEvent(this, modelValue); })
                        .on("mouseout", function () { onChangeSpPersonEvent(this, modelValue); })
                    ;
                }
                catch (e) {

                }

                function onChangeSpPersonEvent(self, modelValue) {
                    var value = $.trim($(self).val());
                    if (value == "") {
                        modelValue(null);
                        $(self).removeClass("valid").removeClass("invalid");
                        return;
                    }

                    if (validateSpPerson(modelValue())) {
                        $(self).val(modelValue().split('#')[1]);
                        $(self).addClass("valid").removeClass("invalid");
                    }
                    else {
                        $(self).removeClass("valid").addClass("invalid");
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
                    if (validateSpPerson(person)) {
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

                }
            }
        };

        ko.bindingHandlers['spDate'] = {
            init: function (element, valueAccessor, allBindings, bindingContext) {
                var modelValue = valueAccessor();
                if (element.tagName.toLowerCase() != "input" || $(element).attr("type") == "hidden") { return; }/*stop if not an editable field */
                $(element).datepicker().addClass("datepicker med").on("blur", onDateChange).on("change", onDateChange);
                $(element).attr("placeholder", "MM/DD/YYYY");

                function onDateChange() {
                    try {
                        if ($.trim(this.value) == "") { modelValue(null); return; }
                        modelValue(new Date(this.value));
                    } catch (e) { modelValue(null); this.value = ""; }
                };
            },
            update: function (element, valueAccessor, allBindings, bindingContext) {
                var viewModel = bindingContext.$data;
                var modelValue = valueAccessor();
                var value = ko.unwrap(modelValue);
                var dateStr = '';

                if (value && value != null) {
                    var d = new Date(value);
                    dateStr = dateToLocaleString(d);
                }

                if ('value' in element) {
                    $(element).val(dateStr);
                } else {
                    $(element).text(dateStr);
                }
            }
        };

        ko.bindingHandlers['spDateTime'] = {
            init: function (element, valueAccessor, allBindings, bindingContext) {
                if (element.tagName.toLowerCase() != "input" || $(element).attr("type") == "hidden") { return; }/*stop if not an editable field */

                var viewModel = bindingContext.$data
                    , modelValue = valueAccessor()
                    , value = ko.unwrap(modelValue)
                    , required
                    , $time
                    , $display
                    , $error
                    , $element = $(element)
                    ;

                try {

                    $display = $('<span>', { 'class': 'no-print' }).insertAfter($element);
                    $error = $('<span>', { 'class': 'error', 'html': 'Invalid Date-time', 'style': 'display:none;' }).insertAfter($element);
                    element.$display = $display;
                    element.$error = $error;

                    required = $element.hasClass("required") || $element.attr("required") != null;

                    $element.attr({
                        'placeholder': 'MM/DD/YYYY',
                        'maxlength': 10,
                        'class': 'datepicker med'
                    }).datepicker().on('change', function () {
                        try {
                            $error.hide();
                            if (!isDate(this.value)) {
                                $error.show();
                                return;
                            }
                            var val = this.value;
                            val = val.split('/');
                            var m = val[0] - 1;
                            var d = val[1] - 0;
                            var y = val[2] - 0;
                            var date = modelValue() == null ? new Date(y, m, d) : modelValue();
                            date.setMonth(m);
                            date.setDate(d);
                            date.setYear(y);
                            modelValue(date);
                            $display.html(toDateTimeLocaleString(date));
                        }
                        catch (e) {
                            $error.show();
                        }
                    });

                    $time = $("<input>", {
                        'type': 'text',
                        'maxlength': 8,
                        'style': 'width:6em;',
                        'class': (required ? 'required' : ''),
                        'placeholder': 'HH:MM PM'
                    })
                        .insertAfter($element)
                        .on('change', function () {
                            try {
                                $error.hide();
                                var time = this.value.toString().toUpperCase().replace(/[^\d\:AMP\s]/g, '');
                                this.value = time;

                                if (modelValue() == null) {
                                    return;
                                }

                                if (!isTime(time)) {
                                    $error.show();
                                    return;
                                }

                                var d = modelValue();
                                var tt = time.replace(/[^AMP]/g, ''); // AM/PM
                                var t = time.replace(/[^\d\:]/g, '').split(':');
                                var h = t[0] - 0; //hours
                                var m = t[1] - 0; //minutes

                                if (tt == 'PM' && h < 12) {
                                    h += 12; //convert to military time
                                }
                                else if (tt == 'AM' && h == 12) {
                                    h = 0; //convert to military midnight
                                }

                                d.setHours(h);
                                d.setMinutes(m);
                                modelValue(d);

                                $display.html(toDateTimeLocaleString(d));
                                $error.hide();
                            }
                            catch (e) {
                                $display.html(e);
                                $error.show();
                            }
                        });

                    $time.before('<span> Time: </span>').after('<span class="no-print"> (HH:MM PM) </span>');

                    element.$time = $time;

                    if (modelValue() == null) {
                        $element.val('');
                        $time.val('');
                    }

                }
                catch (e) {

                }
            },
            update: function (element, valueAccessor, allBindings, bindingContext) {
                var viewModel = bindingContext.$data
                    , modelValue = valueAccessor()
                    , value = ko.unwrap(modelValue)
                    ;

                try {
                    if (value && value != null) {
                        var d = new Date(value);
                        var dateStr = dateToLocaleString(d);
                        var timeStr = toTimeLocaleString(d);

                        if (element.tagName.toLowerCase() == "input") {
                            element.value = dateStr;
                            element.$time.val(timeStr);
                            element.$display.html(dateStr + ' ' + timeStr);
                        } else {
                            element.innerHTML = dateStr + ' ' + timeStr;
                        }
                    }
                }
                catch (e) { }
            }
        };

        ko.bindingHandlers['spMoney'] = {
            'init': function (element, valueAccessor, allBindings, viewModel, bindingContext) {

                /* stop if not an editable field */
                if (element.tagName.toLowerCase() != "input" || $(element).attr("type") == "hidden") { return; }

                viewModel = bindingContext.$data;
                var value = valueAccessor();
                var valueUnwrapped = ko.unwrap(value);

                $(element).on("blur", onChange).on("change", onChange);

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
                        $(element).addClass("negative");
                    } else {
                        $(element).removeClass("negative");
                    }
                } else {
                    valueUnwrapped = 0;
                }

                var formattedValue = accounting.formatMoney(valueUnwrapped);
                updateKoField(element, formattedValue);
            }
        };

        ko.bindingHandlers['spDecimal'] = {
            'init': function (element, valueAccessor, allBindings, viewModel, bindingContext) {

                // stop if not an editable field 
                if (element.tagName.toLowerCase() != "input" || $(element).attr("type") == "hidden") { return; }

                viewModel = bindingContext.$data;
                var value = valueAccessor();
                var valueUnwrapped = ko.unwrap(value);

                $(element).on("blur", onChange).on("change", onChange);

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

                var formattedValue = accounting.toFixed(valueUnwrapped, precision);

                if (valueUnwrapped != null) {
                    if (valueUnwrapped < 0) {
                        $(element).addClass("negative");
                    } else {
                        $(element).removeClass("negative");
                    }
                } else {
                    valueUnwrapped = 0;
                }

                updateKoField(element, formattedValue);
            }
        };

        ko.bindingHandlers['spNumber'] = {
            /* executes on load */
            init: function (element, valueAccessor, allBindings, viewModel, bindingContext) {

                /* stop if not an editable field */
                if (element.tagName.toLowerCase() != "input" || $(element).attr("type") == "hidden") { return; }

                viewModel = bindingContext.$data;
                var value = valueAccessor();
                var valueUnwrapped = ko.unwrap(value);

                $(element).on("blur", onChange).on("change", onChange);

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

                updateKoField(element, valueUnwrapped);

                if (typeof (value) == "function") {
                    value(valueUnwrapped);
                }
            }
        };

    })(ko);


    /* update a KO observable whether it's an update or text field */
    function updateKoField(el, val) {
        if (el.tagName.toLowerCase() == "input") {
            $(el).val(val);
        } else {
            $(el).html(val);
        }
    }

    //validate format ID;#UserName
    function validateSpPerson(person) {
        return person != null && person.toString().match(/^\d*;#/) != null;
    }

    function isTime(val) {
        var rx = new RegExp("\\d{1,2}:\\d{2}\\s{0,1}(AM|PM)");
        return rx.test(val);
    }

    function isDate(val) {
        var rx = new RegExp("\\d{1,2}\/\\d{1,2}\/\\d{4}");
        return rx.test(val.toString());
    }

    function dateToLocaleString(d) {
        try {
            var dd = d.getDate();
            dd = dd < 10 ? "0" + dd : dd;
            var mo = d.getMonth() + 1;
            mo = mo < 10 ? "0" + mo : mo;
            return mo + '/' + dd + '/' + d.getFullYear();
        }
        catch (e) {
            return 'Invalid Date';
        }
    }

    function toTimeLocaleObject(d) {
        var hours: number = 0;
        var minutes: any;
        var tt: string;

        hours = d.getHours();
        minutes = d.getMinutes();
        tt = hours > 11 ? 'PM' : 'AM';

        if (minutes < 10) {
            minutes = '0' + minutes;
        }

        if (hours > 12) {
            hours -= 12;
        }

        return {
            hours: hours,
            minutes: minutes,
            tt: tt
        };
    }

    function toTimeLocaleString(d) {
        var str = '12:00 AM';
        var hours = d.getHours();
        var minutes = d.getMinutes();
        var tt = hours > 11 ? 'PM' : 'AM';

        if (minutes < 10) {
            minutes = '0' + minutes;
        }

        if (hours > 12) {
            hours -= 12;
        }
        else if (hours == 0) {
            hours = 12;
        }

        return hours + ':' + minutes + ' ' + tt;
    }

    function toDateTimeLocaleString(d) {
        var time = toTimeLocaleString(d);
        return dateToLocaleString(d) + ' ' + time;
    }
}