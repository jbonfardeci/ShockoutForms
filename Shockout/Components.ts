module Shockout {
    
    export class KoComponents {

        public static registerKoComponents() {

            ko.components.register('so-text-field', {
                viewModel: soFieldModel,
                template: KoComponents.soTextFieldTemplate
            });

            ko.components.register('so-html-field', {
                viewModel: soFieldModel,
                template: KoComponents.soHtmlFieldTemplate
            });

            ko.components.register('so-person-field', {
                viewModel: soFieldModel,
                template: KoComponents.soTextFieldTemplate.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spPerson: modelValue')
            });

            ko.components.register('so-date-field', {
                viewModel: soFieldModel,
                template: KoComponents.soTextFieldTemplate.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spDate: modelValue')
            });

            ko.components.register('so-datetime-field', {
                viewModel: soFieldModel,
                template: KoComponents.soTextFieldTemplate.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spDateTime: modelValue')
            });

            ko.components.register('so-money-field', {
                viewModel: soFieldModel,
                template: KoComponents.soTextFieldTemplate.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spMoney: modelValue')
            });

            ko.components.register('so-number-field', {
                viewModel: soFieldModel,
                template: KoComponents.soTextFieldTemplate.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spNumber: modelValue')
            });

            ko.components.register('so-decimal-field', {
                viewModel: soFieldModel,
                template: KoComponents.soTextFieldTemplate.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spDecimal: modelValue')
            });

            ko.components.register('so-checkbox-field', {
                viewModel: soFieldModel,
                template: KoComponents.soCheckboxFieldTemplate
            });

            ko.components.register('so-select-field', {
                viewModel: soFieldModel,
                template: KoComponents.soSelectFieldTemplate
            });

            ko.components.register('so-checkbox-group', {
                viewModel: soFieldModel,
                template: KoComponents.soCheckboxGroupTemplate
            });

            ko.components.register('so-radio-group', {
                viewModel: soFieldModel,
                template: KoComponents.soRadioGroupTemplate
            });

            ko.components.register('so-usermulti-group', {
                viewModel: soUsermultiModel,
                template: KoComponents.soUsermultiFieldTemplate
            });

            ko.components.register('so-static-field', {
                viewModel: soStaticModel,
                template: KoComponents.soStaticFieldTemplate
            });

            ko.components.register('so-static-person', {
                viewModel: soStaticModel,
                template: KoComponents.soStaticFieldTemplate.replace(/data-bind="text: modelValue/g, 'data-bind="spPerson: modelValue')
            });

            ko.components.register('so-static-date', {
                viewModel: soStaticModel,
                template: KoComponents.soStaticFieldTemplate.replace(/data-bind="text: modelValue/g, 'data-bind="spDate: modelValue')
            });

            ko.components.register('so-static-datetime', {
                viewModel: soStaticModel,
                template: KoComponents.soStaticFieldTemplate.replace(/data-bind="text: modelValue/g, 'data-bind="spDateTime: modelValue')
            });

            ko.components.register('so-static-money', {
                viewModel: soStaticModel,
                template: KoComponents.soStaticFieldTemplate.replace(/data-bind="text: modelValue/g, 'data-bind="spMoney: modelValue')
            });

            ko.components.register('so-static-number', {
                viewModel: soStaticModel,
                template: KoComponents.soStaticFieldTemplate.replace(/data-bind="text: modelValue/g, 'data-bind="spNumber: modelValue')
            });

            ko.components.register('so-static-decimal', {
                viewModel: soStaticModel,
                template: KoComponents.soStaticFieldTemplate.replace(/data-bind="text: modelValue/g, 'data-bind="spDecimal: modelValue')
            });

            ko.components.register('so-static-html', {
                viewModel: soStaticModel,
                template: KoComponents.soStaticFieldTemplate.replace(/data-bind="text: modelValue/g, 'data-bind="html: modelValue')
            });

            ko.components.register('so-attachments', {
                viewModel: function (params) {

                    if (!params) {
                        throw 'params is undefined in so-attachments';
                        return;
                    }

                    if (!params.val) {
                        throw "Parameter `val` for so-attachments is required!";
                    }

                    this.attachments = params.val; 
                    this.title = params.title || 'Attachments';
                    this.id = params.id;
                },
                template: '<section>' +
                    '<h4 data-bind="text: title">Attachments (<span data-bind="text: attachments().length"></span>)</h4>' +
                    '<div data-bind="attr:{id: id}"></div>'+
                    '<div data-bind="foreach: attachments">' +
                        '<div>' +
                            '<a href="" data-bind="attr: {href: __metadata.media_src}"><span class="glyphicon glyphicon-paperclip"></span> <span data-bind="text: Name"></span></a>&nbsp;' + 
                        '</div>' +
                    '</div>' +
                '</section>'
            });

            function soStaticModel(params) {
                if (!params) {
                    throw 'params is undefined in so-static-field';
                    return;
                }

                var koObj: IShockoutObservable<string> = params.val || params.modelValue;

                if (!koObj) {
                    throw "Parameter `val` or `modelValue` for so-static-field is required!";
                }

                this.modelValue = koObj;
                this.id = params.id || koObj._koName;
                this.label = params.label || koObj._displayName;
                this.description = params.description || koObj._description;
            };

            function soFieldModel(params) {

                if (!params) {
                    throw 'params is undefined in soFieldModel';
                    return;
                }

                var koObj: IShockoutObservable<string> = params.val || params.modelValue;

                if (!koObj) {
                    throw "Parameter `val` or `modelValue` for so-text-field is required!";
                }

                this.modelValue = koObj;
                this.id = params.id || koObj._koName;
                this.name = params.name || koObj._koName || params.id;
                this.label = params.label || koObj._displayName;
                this.title = params.title;
                this.caption = params.caption;
                this.maxlength = params.maxlength || 255;
                this.placeholder = params.placeholder || koObj._displayName;
                this.description = params.description || koObj._description;
                this.valueUpdate = params.valueUpdate;
                this.editable = !!koObj._koName; // if `_koName` is a prop of our KO var, it's a field we can update in theSharePoint list.
                this.koName = koObj._koName; // include the name of the KO var in case we need to reference it.
                this.options = params.options || koObj._options;
                this.required = params.required || koObj._required;
                this.readOnly = params.readOnly || false;
                this.inline = params.inline || false;

            };

            function soUsermultiModel(params) {

                if (!params) {
                    throw 'params is undefined in soFieldModel';
                    return;
                }

                var self = this;
                var koObj: IShockoutObservable<string> = params.val || params.modelValue;

                if (!koObj) {
                    throw "Parameter `val` or `modelValue` for so-text-field is required!";
                }

                this.modelValue = koObj;
                this.id = params.id || koObj._koName;
                this.name = params.name || koObj._koName || params.id;
                this.label = params.label || koObj._displayName;
                this.title = params.title;
                this.caption = params.caption;
                this.required = params.required;
                this.maxlength = params.maxlength || 255;
                this.placeholder = params.placeholder || koObj._displayName;
                this.description = params.description || koObj._description;
                this.editable = !!koObj._koName; // if `_koName` is a prop of our KO var, it's a field we can update in theSharePoint list.
                this.koName = koObj._koName; // include the name of the KO var in case we need to reference it.
                this.person = ko.observable(null);
                this.readOnly = params.readOnly || false;

                // add a person to KO object People
                this.addPerson = function (model, ctrl) {
                    if (self.modelValue() == null) {
                        self.modelValue([]);
                    }

                    self.modelValue().push(self.person());
                    self.modelValue.valueHasMutated();
                    self.person(null);
                    return false;
                };

                // remove a person from KO object People
                this.removePerson = function (person, event) {
                    self.modelValue.remove(person);
                    return false;
                }
            };

        };

        //&& !!required && !readOnly
        private static hasErrorCssDiv: string = '<div class="form-group" data-bind="css: {\'has-error\': !!!modelValue() && !!required, \'has-success has-feedback\': !!modelValue() && !!required}">';

        private static requiredFeedbackSpan: string = '<span class="glyphicon glyphicon-ok form-control-feedback" aria-hidden="true"></span>';

        public static soStaticFieldTemplate: string =
        '<div class="form-group">'+
            '<div class="row">' +            
                // field label
                '<!-- ko if: label -->' +
                    '<div class="col-sm-3"><label data-bind="html: label"></label></div>'+
                '<!-- /ko -->' +
                // field
                '<div class="col-sm-9" data-bind="text: modelValue"></div>' +
            '</div>' +
            // description
            '<!-- ko if: description -->' +
            '<div class="row"><div class="col-sm-3">&nbsp;</div><div class="col-sm-9" data-bind="html: description"></div></div>' +
            '<!-- /ko -->' +
        '</div>';

        public static soTextFieldTemplate: string =
        KoComponents.hasErrorCssDiv +
            '<div class="row">' + 
                   
                // field label
                '<!-- ko if: label -->' +
                    '<div class="col-sm-3"><label data-bind="html: label, attr: {for: id}"></label></div>' +
                '<!-- /ko -->' +
            
                // field
                '<div class="col-sm-9">' +
                    '<!-- ko if: readOnly -->' +
                        '<div data-bind="text: modelValue"></div>' +
                    '<!-- /ko -->' +

                    '<!-- ko ifnot: readOnly -->' +
                        '<input type="text" data-bind="value: modelValue, css: {\'so-editable\': editable}, attr: {id: id, placeholder: placeholder, title: title, required: required, \'ko-name\': koName }" class="form-control" />' +
                        '<!-- ko if: !!required -->' +
                            KoComponents.requiredFeedbackSpan +
                        '<!-- /ko -->' +
                    '<!-- /ko -->' +

                '</div>'+
            '</div>' +

            // description
            '<!-- ko if: description -->' +
            '<div class="row"><div class="col-sm-3">&nbsp;</div><div class="col-sm-9" data-bind="html: description"></div></div>' +
            '<!-- /ko -->' +

        '</div>';

        //'<div data-bind="spHtmlEditor: modelValue" contenteditable="true" class="form-control content-editable"></div>'+
        //'<textarea data-bind="value: modelValue, css: {\'so-editable\': editable}, attr: {id: id, required: required, \'ko-name\': koName }" data-sp-html="" style="display:none;"></textarea>' +
        public static soHtmlFieldTemplate: string =
        KoComponents.hasErrorCssDiv +
            '<div class="row">' + 
                   
                // field label
                '<!-- ko if: label -->' +
                    '<div class="col-sm-3"><label data-bind="html: label, attr: {for: id}"></label></div>' +
                '<!-- /ko -->' +
            
                // field
                '<div class="col-sm-9">' +
                    '<!-- ko if: readOnly -->' +
                        '<div data-bind="html: modelValue"></div>' +
                    '<!-- /ko -->' +

                    '<!-- ko ifnot: readOnly -->' +
                        '<div data-bind="spHtmlEditor: modelValue" contenteditable="true" class="form-control content-editable"></div>' +
                        '<textarea data-bind="value: modelValue, css: {\'so-editable\': editable}, attr: {id: id, required: required, \'ko-name\': koName }" data-sp-html="" style="display:none;"></textarea>' +
                        '<!-- ko if: !!required -->' +
                            KoComponents.requiredFeedbackSpan +
                        '<!-- /ko -->' +
                    '<!-- /ko -->' +

                '</div>' +
            '</div>' +

            // description
            '<!-- ko if: description -->' +
            '<div class="row"><div class="col-sm-3">&nbsp;</div><div class="col-sm-9" data-bind="html: description"></div></div>' +
            '<!-- /ko -->' +

        '</div>';

        public static soCheckboxFieldTemplate: string =
        '<div class="form-group">' +

            '<div class="row">' + 
                   
                // field label
                '<!-- ko if: label -->' +
                    '<div class="col-sm-3"><label data-bind="html: label"></label></div>' +
                '<!-- /ko -->' +
            
                // field
                '<div class="col-sm-9">' +
                    '<!-- ko if: readOnly -->' +
                        '<div data-bind="text: !!modelValue() ? \'Yes\' : \'No\'"></div>' +
                    '<!-- /ko -->' +

                    '<!-- ko ifnot: readOnly -->' +
                        '<label class="checkbox">' +
                            '<input type="checkbox" data-bind="checked: modelValue, css: {\'so-editable\': editable}, attr: {id: id, \'ko-name\': koName}, valueUpdate: valueUpdate" />' +
                            '<span data-bind="html: label"></span>' +
                        '</label>' +
                    '<!-- /ko -->' +

                '</div>' +
            '</div>' +

            // description
            '<!-- ko if: description -->' +
            '<div class="row"><div class="col-sm-3">&nbsp;</div><div class="col-sm-9" data-bind="html: description"></div></div>' +
            '<!-- /ko -->' +

        '</div>';

        public static soSelectFieldTemplate: string =
        KoComponents.hasErrorCssDiv +
            '<div class="row">' + 
                   
                // field label
                '<!-- ko if: label -->' +
                '<div class="col-sm-3"><label data-bind="html: label, attr: {for: id}"></label></div>' +
                '<!-- /ko -->' +
            
                // field
                '<div class="col-sm-9">' +
                    '<!-- ko if: readOnly -->' +
                    '<div data-bind="text: modelValue"></div>' +
                    '<!-- /ko -->' +

                    '<!-- ko ifnot: readOnly -->' +
                        '<select data-bind="value: modelValue, options: options, optionsCaption: caption, css: {\'so-editable\': editable}, attr: {id: id, title: title, required: required, \'ko-name\': koName}" class="form-control"></select>' +
                        '<!-- ko if: !!required -->' +
                            KoComponents.requiredFeedbackSpan +
                        '<!-- /ko -->' +
                    '<!-- /ko -->' +

                '</div>' +
            '</div>' +

            // description
            '<!-- ko if: description -->' +
            '<div class="row"><div class="col-sm-3">&nbsp;</div><div class="col-sm-9" data-bind="html: description"></div></div>' +
            '<!-- /ko -->' +

        '</div>';

        public static soCheckboxGroupTemplate: string =
        '<div class="form-group">' +
            // description
            '<!-- ko if: description -->' +
            '<div class="row"><div class="col-sm-3">&nbsp;</div><div class="col-sm-9" data-bind="html: description"></div></div>' +
        '<!-- /ko -->' +

            '<div class="row">'+

                // field label
                '<!-- ko if: label -->' +
                '<div><label data-bind="html: label"></label></div>' +
                '<!-- /ko -->' +

                '<div>'+
                    // show static elements if inline
                    '<!-- ko if: readOnly -->' +

                        // show static unordered list if !inline
                        '<!-- ko ifnot: inline -->'+
                            '<ul class="list-group">' +

                                '<!-- ko foreach: modelValue -->' +
                                    '<li data-bind="text: $data" class="list-group-item"></li>' +
                                '<!-- /ko -->' +

                                '<!-- ko if: modelValue().length == 0 -->' +
                                    '<li class="list-group-item">--None--</li>'+
                                '<!-- /ko -->' +

                            '</ul>' +
                        '<!-- /ko -->' +

                        // show static inline elements if inline
                        '<!-- ko if: inline -->' +

                            '<!-- ko foreach: modelValue -->' +
                                '<span data-bind="text: $data"></span>' +
                                '<!-- ko if: $index() < $parent.modelValue().length-1 -->,&nbsp;<!-- /ko -->' +
                            '<!-- /ko -->' +

                            '<!-- ko if: modelValue().length == 0 -->' +
                            '<span>--None--</span>' +
                            '<!-- /ko -->' +

                        '<!-- /ko -->' +

                    '<!-- /ko -->' +

                    // show input field if not readOnly
                    '<!-- ko ifnot: readOnly -->' +
                        '<!-- ko foreach: options -->' +
                            '<label data-bind="css:{\'checkbox\': !$parent.inline, \'checkbox-inline\': $parent.inline}">' +
                                '<input type="checkbox" data-bind="checked: $parent.modelValue, css: {\'so-editable\': $parent.editable}, attr: {\'ko-name\': $parent.koName, \'value\': $data}" />' +
                                '<span data-bind="text: $data"></span>' +
                            '</label>' +
                        '<!-- /ko -->' +
                    '<!-- /ko -->' +

                '</div>'+
            '</div>';

        public static soRadioGroupTemplate: string =
        '<div class="form-group">' +
            // description
            '<!-- ko if: description -->' +
            '<div class="row"><div class="col-sm-3">&nbsp;</div><div class="col-sm-9" data-bind="html: description"></div></div>' +
            '<!-- /ko -->' +

            '<div class="row">' +

                // field label
                '<!-- ko if: label -->'+
                    '<div class="col-sm-3"><label data-bind="html: label"></label></div>' +
                '<!-- /ko -->'+

                '<div class="col-sm-9">'+
                    // show static field if readOnly
                    '<!-- ko if: readOnly -->' +
                        '<div data-bind="text: modelValue"></div>' +
                    '<!-- /ko -->' +

                    // show input field if not readOnly
                    '<!-- ko ifnot: readOnly -->' +
                        '<!-- ko foreach: options -->' +  
                            '<label data-bind="css:{\'radio\': !$parent.inline, \'radio-inline\': $parent.inline}">' +
                                '<input type="radio" data-bind="checked: $parent.modelValue, attr:{value: $data, name: $parent.name, \'ko-name\': $parent.koName}, css:{\'so-editable\': $parent.editable}" />' +
                                '<span data-bind="text: $data"></span>' +
                            '</label>' +
                        '<!-- /ko -->' +
                    '<!-- /ko -->' +
                '</div>'+
            '</div>' +
            
        '</div>';

        public static soUsermultiFieldTemplate: string =
        '<div class="form-group">' +

            // show input field if not readOnly
            '<!-- ko ifnot: readOnly -->' +
                '<input type="hidden" data-bind="value: modelValue, css: {\'so-editable\': editable}, attr: {id: id, \'ko-name\': koName, required: required}" />' +
                '<div class="row">' +
                    '<div class="col-md-6 col-xs-6">' +
                        '<input type="text" data-bind="spPerson: person, attr: {placeholder: placeholder}" />' +
                        '<button class="btn btn-success" data-bind="click: addPerson, attr: {\'disabled\': person() == null}"><span>Add</span></button>' +
                    '</div>' +

                    '<!-- ko if: required && modelValue() == null && !readOnly -->' +
                        '<div class="col-md-6 col-xs-6">' +
                            '<p class="error">This field is required.</p>' +
                        '</div>' +
                    '<!-- /ko -->' +

                '</div>' +
            '<!-- /ko -->' +

            '<!-- ko foreach: modelValue -->' +
                '<div class="row">' +
                    '<div class="col-md-10 col-xs-10" data-bind="spPerson: $data"></div>' +

                    '<!-- ko ifnot: readOnly -->'+
                        '<div class="col-md-2 col-xs-2">' +
                            '<button class="btn btn-xs btn-danger" data-bind="click: $parent.removePerson"><span class="glyphicon glyphicon-trash"></span></button>' +
                        '</div>' +
                    '<!-- /ko -->'+

                '</div>' +
            '<!-- /ko -->' +

            '<!-- ko if: description -->' +
                '<p data-bind="html: description"></p>' +
            '<!-- /ko -->' +

        '</div>';
    }

}