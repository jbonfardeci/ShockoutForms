 module Shockout {

    export class KoComponents {

        public static registerKoComponents() {

            var uniqueId = (function () {
                var i = 0;
                return function () {
                    return i++;
                };
            })();

            ko.components.register('so-text-field', {
                viewModel: soFieldModel,
                template: Templates.soTextField
            });

            ko.components.register('so-html-field', {
                viewModel: soFieldModel,
                template: Templates.soHtmlFieldTemplate
            });

            ko.components.register('so-person-field', {
                viewModel: soFieldModel,
                template: Templates.soTextField.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spPerson: modelValue')
            });

            ko.components.register('so-date-field', {
                viewModel: soFieldModel,
                template: Templates.soTextField.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spDate: modelValue')
            });

            ko.components.register('so-datetime-field', {
                viewModel: soFieldModel,
                template: Templates.soTextField.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spDateTime: modelValue')
            });

            ko.components.register('so-money-field', {
                viewModel: soFieldModel,
                template: Templates.soTextField.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spMoney: modelValue')
            });

            ko.components.register('so-number-field', {
                viewModel: soFieldModel,
                template: Templates.soTextField.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spNumber: modelValue')
            });

            ko.components.register('so-decimal-field', {
                viewModel: soFieldModel,
                template: Templates.soTextField.replace(/data-bind="(value|text): modelValue/g, 'data-bind="spDecimal: modelValue')
            });

            ko.components.register('so-checkbox-field', {
                viewModel: soFieldModel,
                template: Templates.soCheckboxField
            });

            ko.components.register('so-select-field', {
                viewModel: soFieldModel,
                template: Templates.soSelectField
            });

            ko.components.register('so-checkbox-group', {
                viewModel: soFieldModel,
                template: Templates.soCheckboxGroup
            });

            ko.components.register('so-radio-group', {
                viewModel: soFieldModel,
                template: Templates.soRadioGroup
            });

            ko.components.register('so-usermulti-group', {
                viewModel: soUsermultiModel,
                template: Templates.soUsermultiField
            });

            ko.components.register('so-static-field', {
                viewModel: soStaticModel,
                template: Templates.soStaticField
            });

            ko.components.register('so-static-person', {
                viewModel: soStaticModel,
                template: Templates.soStaticField.replace(/data-bind="text: modelValue/g, 'data-bind="spPerson: modelValue')
            });

            ko.components.register('so-static-date', {
                viewModel: soStaticModel,
                template: Templates.soStaticField.replace(/data-bind="text: modelValue/g, 'data-bind="spDate: modelValue')
            });

            ko.components.register('so-static-datetime', {
                viewModel: soStaticModel,
                template: Templates.soStaticField.replace(/data-bind="text: modelValue/g, 'data-bind="spDateTime: modelValue')
            });

            ko.components.register('so-static-money', {
                viewModel: soStaticModel,
                template: Templates.soStaticField.replace(/data-bind="text: modelValue/g, 'data-bind="spMoney: modelValue')
            });

            ko.components.register('so-static-number', {
                viewModel: soStaticModel,
                template: Templates.soStaticField.replace(/data-bind="text: modelValue/g, 'data-bind="spNumber: modelValue')
            });

            ko.components.register('so-static-decimal', {
                viewModel: soStaticModel,
                template: Templates.soStaticField.replace(/data-bind="text: modelValue/g, 'data-bind="spDecimal: modelValue')
            });

            ko.components.register('so-static-html', {
                viewModel: soStaticModel,
                template: Templates.soStaticField.replace(/data-bind="text: modelValue/g, 'data-bind="html: modelValue')
            });

            ko.components.register('so-attachments', {
                viewModel: soAttachmentsModel,
                template: Templates.soAttachments
            });

            ko.components.register('so-created-modified-info', {
                viewModel: soCreatedModifiedInfoModel,
                template: Templates.soCreatedModifiedInfo                 
            });

            ko.components.register('so-nav-menu', {
                viewModel: soNavMenuModel,
                template: Templates.soNavMenu
            });

            ko.components.register('so-workflow-history', {
                viewModel: function (params) {
                    this.historyItems = <Array<IHistoryItem>>(params.val || params.historyItems);
                },
                template: Templates.soWorkflowHistory
            });

            function soCreatedModifiedInfoModel(params) {
                this.CreatedBy = <KnockoutObservable<ISpPerson>>params.createdBy
                this.ModifiedBy = <KnockoutObservable<ISpPerson>>params.modifiedBy;
                this.profiles = ko.observableArray([
                    { header: 'Created By', profile: this.CreatedBy },
                    { header: 'Modified By', profile: this.ModifiedBy }
                ]);
                this.Created = params.created;
                this.Modified = params.modified;
                this.showUserProfiles = params.showUserProfiles;
            };

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

                var labelX: number = parseInt(params.labelColWidth || 3); // Bootstrap label column width 1-12
                var fieldX: number = parseInt(params.fieldColWidth || (12 - (labelX - 0))); // Bootstrap field column width 1-12
                this.labelColWidth = 'col-sm-' + labelX;
                this.fieldColWidth = 'col-sm-' + fieldX;
            };

            function soFieldModel(params): void {

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
                this.label = params.label = null || params.label == '' ? undefined : params.label || koObj._displayName;
                this.title = params.title;
                this.caption = params.caption;
                this.maxlength = params.maxlength || 255;
                this.placeholder = params.placeholder || params.label || koObj._displayName;
                this.description = (typeof params.description != 'undefined') ? (params.description == null ? undefined : params.description) : koObj._description;
                this.valueUpdate = params.valueUpdate;
                this.editable = !!koObj._koName; // if `_koName` is a prop of our KO var, it's a field we can update in theSharePoint list.
                this.koName = koObj._koName; // include the name of the KO var in case we need to reference it.
                this.options = params.options || koObj._options;
                this.required = (typeof params.required == 'function') ? params.required : ko.observable(!!params.required || false);
                this.inline = params.inline || false;
                this.multiline = params.multiline || false;

                var labelX: number = parseInt(params.labelColWidth || 3); // Bootstrap label column width 1-12
                var fieldX: number = parseInt(params.fieldColWidth || (12 - (labelX - 0))); // Bootstrap field column width 1-12
                this.labelColWidth = 'col-sm-' + labelX;
                this.fieldColWidth = 'col-sm-' + fieldX;

                // allow for static bool or ko obs
                this.readOnly = (typeof params.readOnly == 'function') ? params.readOnly : ko.observable(!!params.readOnly || false);
            };

            function soAttachmentsModel(params) {
                var self = this;
                var w: any = window;
                this.errorMsg = ko.observable(null);
                if (!!!params) {
                    this.errorMsg('`params` is undefined in component so-attachments');
                    throw this.errorMsg();
                    return;
                }
                if (!!!params.val) {
                    this.errorMsg('Parameter `val` for component so-attachments is required!');
                    throw this.errorMsg();
                    return;
                }

                var spForm: SPForm = params.val.getSpForm();
                var vm: IViewModel = spForm.getViewModel();
                var allowedExtensions: Array<string> = params.allowedExtensions || spForm.allowedExtensions;
                var reader: FileReader;
                // CAFE - Cascading Asynchronous Function Exectuion; 
                // Required to let SharePoint only write one file at a time, otherwise you'll get a 'changes conflict with another user's changes...' when attempting to write multiple files at once
                var cafe: ICafe;
                var asyncFns: Array<Function>;

                this.attachments = <IViewModelAttachments>params.val;
                this.label = params.label || 'Attach Files';
                this.drop = params.drop || true;
                this.dropLabel = params.dropLabel || '...or Drag and Drop Files Here';
                this.className = params.className || 'btn btn-primary';
                this.title = params.title || 'Attachments';
                this.description = params.description;
                this.readOnly = (typeof params.readOnly == 'function') ? params.readOnly : ko.observable(params.readOnly || false); // allow for static bool or ko observable
                this.length = ko.pureComputed(function () { return self.attachments().length; });
                this.fileUploads = <KnockoutObservableArray<any>>ko.observableArray();

                //check for compatibility
                this.hasFileReader = ko.observable(w.File && w.FileReader && w.FileList && w.Blob);

                if (!this.hasFileReader) {
                    this.errorMsg('This browser does not support the FileReader class required for uplaoding files. You may be using IE 9 or another unsupported browser.');
                }

                this.id = params.id || 'so_fileUploader_' + uniqueId();

                this.deleteAttachment = function (att, event) {
                    if (!confirm('Are you sure you want to delete ' + att.Name + '? This can\'t be undone.')) {
                        return;
                    }
                    Shockout.SpApi.deleteAttachment(att, function (data, error) {
                        if (!!error) {
                            alert("Failed to delete attachment: " + error);
                            return;
                        }
                        self.attachments.remove(att);
                    });
                };

                // event handler for input[type='file']
                this.fileHandler = function (e) {
                    var files: Array<File> = document.getElementById(self.id)['files'];
                    readFiles(files);
                };

                // event handler for Attach button
                this.onSelect = function (e) {
                    cancel(e);
                    //trigger click on the input file control
                    document.getElementById(self.id).click();
                };

                // event handler for Drag adn Drop Zone
                this.onDrop = function (localViewModel: any, e: any) {
                    cancel(e);

                    if (spForm.debug) {
                        console.info('dropped files over dropzone, arguments are...');
                        console.info(arguments);
                    }

                    var dt = (e.originalEvent || e).dataTransfer;
                    var files: Array<File> = dt.files;
                    if (!!!files) {
                        console.warn('Error in so-attachments - event.dataTransfer.files is ' + typeof files);
                        return false;
                    }
                    else {
                        readFiles(files);
                    }
                };

                // read files array
                function readFiles(files: Array<File>): void {
                    asyncFns = [];

                    // build the cascading function execution array
                    var fileArray: Array<File> = Array.prototype.slice.call(files, 0);
                    fileArray.map(function (file: File, i: number) {
                        asyncFns.push(function () {
                            readFile(file);
                        });
                    });

                    cafe = new Cafe(asyncFns);

                    // If this is a new form, save it first; you can't attach a file unless the list item already exists.
                    if (vm.Id() == null) {
                        spForm.saveListItem(vm, false, undefined, function (itemId: number) {
                            // catch-all if for some reason vm.Id is still null or lost reference of vm and we're referencing a local copy of the actual view model?
                            if (vm.Id() == null && !!itemId && itemId.toFixed) {
                                vm.Id(itemId);
                            }
                            setTimeout(function () {
                                cafe.next(true); //start the async function exectuion cascade
                            }, 1000);
                        });
                    } else {
                        cafe.next(true); //start the async function exectuion cascade
                    }
                };

                // upload a File object
                function readFile(file: File): void {

                    if (spForm.debug) {
                        console.info('uploading file...');
                        console.info(file);
                    }

                    var fileName: string = file.name.replace(/[^a-zA-Z0-9_\-\.]/g, ''); // clean the filename
                    var ext: string = /\.\w{2,4}$/.exec(fileName)[0]; //extract extension from filename, e.g. '.docx'
                    var rootName = fileName.replace(new RegExp(ext + '$'), ''); // e.g. 'test.docx' becomes 'test'
                    // Is the extension of the fileName in the array of allowed extensions? 
                    var allowedExtension: boolean = new RegExp("^(\\.|)(" + allowedExtensions.join('|') + ")$", "i").test(ext);
                    if (!allowedExtension) {
                        self.errorMsg('Only files with the extensions: ' + allowedExtensions.join(', ') + ' are allowed.');
                        return;
                    }

                    // Check for duplicate filename. If found, append a number.
                    for (var i = 0; i < self.attachments().length; i++) {
                        if (new RegExp(fileName, 'i').test(self.attachments()[i].Name)) {
                            fileName = rootName + '-1' + ext;
                            break;
                        }
                    }

                    var fileUpload: IFileUpload = new FileUpload(fileName, file.size);
                    self.fileUploads().push(fileUpload);
                    self.fileUploads.valueHasMutated();

                    reader = new FileReader();
                    reader.onerror = function errorHandler(e) {
                        var evt: any = e;
                        var className = fileUpload.className();
                        fileUpload.className(className.replace('-success', '-danger'));
                        switch (evt.target.error.code) {
                            case evt.target.error.NOT_FOUND_ERR:
                                self.errorMsg = 'File Not Found!';
                                break;
                            case evt.target.error.NOT_READABLE_ERR:
                                self.errorMsg = 'File is not readable.';
                                break;
                            case evt.target.error.ABORT_ERR:
                                break; // noop
                            default:
                                self.errorMsg = 'An error occurred reading this file.';
                        };
                    };
                    reader.onprogress = function (e) {
                        updateProgress(e, fileUpload);
                    };
                    reader.onabort = function (e) {
                        self.errorMsg('File read cancelled');
                    };
                    reader.onloadstart = function (e) {
                        fileUpload.progress(0);
                    };
                    reader.onload = function (e) {
                        var event: any = e;
                        // Ensure that the progress bar displays 100% at the end.
                        fileUpload.progress(100);
                        // Send the base64 string to the AddAttachment service for upload.
                        Shockout.SpSoap.addAttachment(event.target.result, fileName, spForm.listName, spForm.viewModel.Id(), spForm.siteUrl, callback);
                    }
                    reader.onloadend = function (loadend) {
                        /*loadend = { 
                            target: FileReader, 
                            isTrusted: true, 
                            lengthComputable: true, 
                            loaded: 1972, 
                            total: 1972, 
                            eventPhase: 0, 
                            bubbles: false, 
                            cancelable: false, 
                            defaultPrevented: false, 
                            timeStamp: 1453336901529000, 
                            originalTarget: FileReader 
                        }*/
                        //console.info('loaded ' +  + (loadend.loaded/1024).toFixed(2) + ' KB.');
                    };

                    // read as base64 string
                    reader.readAsDataURL(file);

                    function callback() {

                        // on error: jqXhr: JQueryXHR, status: string, error: string
                        // success: xmlDoc: any, status: string, jqXhr: JQueryXHR
                        var status: string = arguments[1];

                        if (spForm.debug) {
                            console.info('so-html5-attachments.onFileUploadComplete()...');
                            console.info(arguments);
                        }

                        /* error XML: 
                        <?xml version="1.0" encoding="utf-8"?>
                        <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
                            <soap:Body>
                                <soap:Fault>
                                    <faultcode>soap:Server</faultcode>
                                    <faultstring>Exception of type 'Microsoft.SharePoint.SoapServer.SoapServerException' was thrown.</faultstring>
                                    <detail>
                                        <errorstring xmlns="http://schemas.microsoft.com/sharepoint/soap/">Parameter listItemID is missing or invalid.</errorstring>
                                        <errorcode xmlns="http://schemas.microsoft.com/sharepoint/soap/">0x82000001</errorcode>
                                    </detail>
                                </soap:Fault>
                            </soap:Body>
                        </soap:Envelope>                       
                        */
                        if (!!!status && status == 'error') {
                            var jqXhr: JQueryXHR = arguments[0];
                            var responseXml: Document = jqXhr.responseXML;
                            var errorString = $(jqXhr.responseXML).find('errorstring').text();
                            if (!!errorString) {
                                spForm.$dialog.html('Error on file upload. Message from server: ' + (errorString || jqXhr.statusText)).dialog('open');
                            }
                            fileUpload.className(fileUpload.className().replace('-success', '-danger'));
                            cafe.next(false); // will cause Cafe to stop execution of all async functions
                        }
                        else if (status == 'success') {
                            // push a new SP attachment instance to the view model's `attachments` collection
                            var att: ISpAttachment = new SpAttachment(spForm.getRootUrl(), spForm.siteUrl, spForm.listName, spForm.getItemId(), fileName);
                            self.attachments().push(att);
                            self.attachments.valueHasMutated();
                            cafe.next(true); //execute the next file read
                        }

                        setTimeout(function () {
                            self.fileUploads.remove(fileUpload);
                        }, 1000);
                    }
                };

                this.onDragenter = cancel;
                this.onDragover = cancel;

                function updateProgress(e, fileUpload: IFileUpload): void {
                    // e is a ProgressEvent.
                    if (e.lengthComputable) {
                        var percentLoaded = Math.round((e.loaded / e.total) * 100);
                        // Increase the progress bar length.
                        if (percentLoaded < 100) {
                            fileUpload.progress(percentLoaded);
                        }
                    }
                };

                function cancel(e: Event): void {
                    if (e.preventDefault) {
                        e.preventDefault();
                    }
                    if (e.stopPropagation) {
                        e.stopPropagation();
                    }
                };

                if (!spForm.enableAttachments) {
                    this.errorMsg('Attachments are disabled for this form or SharePoint list.');
                    this.readOnly(true);
                }
            }

            function soUsermultiModel(params) {
                if (!params) {
                    throw 'params is undefined in soFieldModel';
                    return;
                }
                var self = this;
                var koObj = params.val || params.modelValue;
                if (!koObj) {
                    throw "Parameter `val` or `modelValue` for so-text-field is required!";
                }
                this.modelValue = koObj;
                this.placeholder = "";
                this.id = params.id || koObj._koName;
                this.name = params.name || koObj._koName || params.id;
                this.label = params.label || koObj._displayName;
                this.title = params.title;
                this.required = params.required;
                this.description = params.description || koObj._description;
                this.editable = !!koObj._koName; // if `_koName` is a prop of our KO var, it's a field we can update in theSharePoint list.
                this.koName = koObj._koName; // include the name of the KO var in case we need to reference it.
                this.person = ko.observable(null);
                // allow for static bool or ko obs
                this.readOnly = (typeof params.readOnly == 'function') ? params.readOnly : ko.observable(!!params.readOnly || false);
                // add a person to KO object People
                this.addPerson = function (model, ctrl) {
                    //if array is null, create it...
                    if (self.modelValue() == null) {
                        self.modelValue([]);
                    }
                    
                    //if the person is already in the list...don't add.
                    var isAlreadyInArray = false;
                    ko.utils.arrayForEach(self.modelValue(), function (item) {
                        if (item == self.person()) {
                            isAlreadyInArray = true;
                        }
                        return;
                    });

                    if (!isAlreadyInArray) {
                        self.modelValue().push(self.person());
                        self.modelValue.valueHasMutated();
                        self.person(null);
                    } else {
                        this.shake(ctrl);
                    }
                    return false;
                };
                // remove a person from KO object People
                this.removePerson = function (person, event) {
                    try {
                        self.modelValue.remove(person);
                    } catch (err) {
                        var index = self.modelValue().indexOf(person);
                        if (index > -1) {  //did not find the item...so don't remove it.
                            self.modelValue().splice(self.modelValue().indexOf(person), 1);
                            self.modelValue.valueHasMutated();
                        }
                    }
                    return false;
                };

                this.showRequiredText = ko.pureComputed(function () {
                    if (self.required) {
                        if (!!self.modelValue()) {
                            return self.modelValue().length < 1;
                        }
                        return true;  //the field is required, but there are no entries in the array, so show the required text.
                    }
                    return false;  //the field is not required, so do not show required text.
                });
                
                //shake behaviour using jQuery animate:
                this.shake = function (element) {
                    var $el = $('button[id=' + element.currentTarget.id + ']');
                    var shakes = 3;
                    var distance = 5;
                    var duration = 200; //total shake animation in miliseconds
					
                    $el.css("position", "relative");
                    for (var x = 1; x <= shakes; x++) {
                        $el.removeClass("btn-success")
                            .addClass("btn-danger")
                            .animate({ left: (distance * -1) }, (((duration / shakes) / 4)))
                            .animate({ left: distance }, ((duration / shakes) / 2))
                            .animate({ left: 0 }, (((duration / shakes) / 4)));
                        setTimeout(function () {
                            $el.removeClass("btn-danger btn-warning").addClass("btn-success");
                        }, 1000);
                    }
                };
            };

            function soNavMenuModel(params) {
                this.navMenuItems = <KnockoutObservableArray<any>>params.val;
                this.title = params.title || 'Navigation';
            };

        };

    }

}