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

                    this.attachments = <IViewModelAttachments>params.val;
                    this.label = params.label || 'Attach Files';
                    this.drop = params.drop || true;
                    this.dropLabel = params.dropLabel || 'OR Drag and Drop Files Here';
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

                    // dropped support for IE9 uploader
                    //if (!this.hasFileReader()) {
                    //    // instantiate the qq file uploader instance
                    //    this.qqFileUploaderId = 'so_qq_fileUploader_' + uniqueId();
                    //    var settings: IFileUploaderSettings = new FileUploaderSettings(spForm, this.qqFileUploaderId, spForm.allowedExtensions);
                    //    var uploader = new Shockout.qq.FileUploader(settings);
                    //}

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

                    // HTML 5 methods
                    this.fileHandler = function (e) { 
                        // If this is a new form, save it first; you can't attach a file unless the list item already exists.
                        if (vm.Id() == null) {
                            spForm.saveListItem(vm, false, undefined, function (itemId) {
                                setTimeout(function () {
                                    readFiles();
                                }, 1000);
                            });
                            return;
                        }          
                        readFiles();             
                    };

                    this.onSelect = function (e) {
                        cancel(e);
                        //trigger click on the input file control
                        document.getElementById(self.id).click();
                    };

                    // WIP
                    this.onDrop = function (localViewModel, e) {
                        cancel(e);

                        if (spForm.debug) {
                            console.info('dropped files over dropzone, arguments are...');
                            console.info(arguments);
                        }

                        var dt = (e.originalEvent || e).dataTransfer;
                        var files = dt.files;
                        if (!!!files) {
                            console.warn('Error in so-attachments - event.dataTransfer.files is ' + typeof files);
                            return false;
                        }
                        else{
                            Array.prototype.slice.call(files, 0).map(readFile);
                        }
                    };

                    function readFiles() {
                        Array.prototype.slice.call(document.getElementById(self.id)['files'], 0).map(readFile);
                    };

                    function readFile(file: File) {       

                        if (spForm.debug) {
                            console.info('uploading file...');
                            console.info(file);
                        }
                                           
                        var fileName: string = file.name.replace(/[^a-zA-Z0-9_\-\.]/g, ''); // clean the filename
                        var allowedExtension: boolean = new RegExp("\\b.(" + allowedExtensions.join('|') + ")$").test(fileName);
                        if (!allowedExtension) {
                            self.errorMsg('Only files with the extensions: ' + allowedExtensions.join(', ') + ' are allowed.'); 
                            return;
                        }

                        //var extension: Array<string> = /\.\w{3,4}$/.exec(fileName); //extract extension from filename
                        // Check for duplicate filename. If found, append a number.
                        var duplicateCt: number = self.attachments().filter(function (file: ISpAttachment): boolean {
                            return fileName == file.Name;
                        }).length;

                        if (spForm.debug && duplicateCt > 0) {
                            console.warn(duplicateCt + ' duplicate files found in Attachments Array');
                        }

                        if (duplicateCt > 0) {
                            var ext = fileName.split('.').slice(-1); // e.g. 'txt'
                            var rootName = fileName.split('.').slice(0, -1).join('.'); //if they have more than one period in the filename - e.g. 'test.2016.01.20'
                            fileName = rootName + '-' + duplicateCt + '.' + ext;                      
                        }
                                                               
                        var fileUpload = new FileUpload(fileName, file.size);
                        self.fileUploads().push(fileUpload);
                        self.fileUploads.valueHasMutated();

                        reader = new FileReader();
                        reader.onerror = errorHandler;
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

                        function callback(xmlDoc: any, status: string, jqXhr: JQueryXHR) {
                            if (spForm.debug) {
                                console.info('so-html5-attachments.onFileUploadComplete()...');
                                console.info(arguments);
                            }

                            if (!!!status || status != 'success') {
                                spForm.$dialog.html('Error on file upload. Please ensure you upload unique filenames to this form.').dialog('open');
                            }
                            else if (status == 'success') {
                                // push a new SP attachment instance to the view model's `attachments` collection
                                var att: ISpAttachment = new SpAttachment(spForm.getRootUrl(), spForm.siteUrl, spForm.listName, spForm.getItemId(), fileName);
                                self.attachments.push(att);
                            }

                            setTimeout(function () {
                                self.fileUploads.remove(fileUpload);
                            }, 1000);
                        }
                    };

                    function FileUpload(filename: string, bytes: number) {
                        this.label = ko.observable();
                        this.progress = ko.observable(0);
                        this.filename = ko.observable(filename);
                        this.kb = ko.observable((bytes / 1024).toFixed(1));
                    };

                    this.onDragenter = cancel;
                    this.onDragover = cancel;

                    function errorHandler(evt) {
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

                    function updateProgress(e, fileUpload) {
                        // e is a ProgressEvent.
                        if (e.lengthComputable) {
                            var percentLoaded = Math.round((e.loaded / e.total) * 100);
                            // Increase the progress bar length.
                            if (percentLoaded < 100) {
                                fileUpload.label(fileUpload.filename() + ' ' + percentLoaded + '% Complete');
                                fileUpload.progess(percentLoaded);
                            }
                        }
                    };

                    function cancel(e) {
                        if (e.preventDefault) {
                            e.preventDefault();
                        }
                        if (e.stopPropagation) {
                            e.stopPropagation();
                        }
                        return false;
                    };

                    if (!spForm.enableAttachments) {
                        this.errorMsg('Attachments are disabled for this form or SharePoint list.');
                        this.readOnly(true);
                    }
                },
                template: 
                `<section>
                    <h4><span data-bind="text: title"></span><span data-bind="text: length" class="badge"></span></h4>
                    <div data-bind="visible: !!errorMsg()" class="alert alert-danger"><span class="glyphicon glyphicon-exclamation-sign"></span>&nbsp;<span data-bind="text: errorMsg"></span></div> 
                    <!-- ko ifnot: hasFileReader() -->
                    <div data-bind="visible: !!!readOnly(), attr: {id: this.qqFileUploaderId}"></div>
                    <!-- /ko -->
                    <!-- ko if: !readOnly() && hasFileReader() -->
                        <input type="file" data-bind="attr: {'id': id}, event: {'change': fileHandler}" multiple class="form-control" style="display:none;" /> 
                        <div data-bind="attr:{'class': className}, event: {'click': onSelect}"><span class="glyphicon glyphicon-paperclip"></span>&nbsp;<span data-bind="text: label"></span></div> 
                        <!-- ko if: drop -->
                            <div class="so-file-dropzone" data-bind="event: {'dragenter': onDragenter, 'dragover': onDragover, 'drop': onDrop}"><div data-bind="text: dropLabel"></div></div>
                        <!-- /ko -->
                        <!-- ko foreach: fileUploads -->
                            <div class="progress"> 
                                <div data-bind="text: label, attr: {'aria-valuenow': progress(), 'style': 'width:' + progress() + '%;' }" class="progress-bar progress-bar-success progress-bar-striped active" role="progressbar" aria-valuemin="0" aria-valuemax="100"></div>  
                            </div>
                        <!-- /ko -->
                    <!-- /ko -->
                    <div data-bind="foreach: attachments" style="margin:1em auto;">
                        <div>      
                            <a href="" data-bind="attr: {href: __metadata.media_src}"><span class="glyphicon glyphicon-paperclip"></span>&nbsp;<span data-bind="text: Name"></span></a>
                            <!-- ko ifnot: $parent.readOnly() -->
                            <button data-bind="event: {click: $parent.deleteAttachment}" class="btn btn-sm btn-danger" title="Delete Attachment"><span class="glyphicon glyphicon-remove"></span></button>
                            <!-- /ko -->
                        </div>
                    </div>
                    <!-- ko if: length() == 0 && readOnly() -->
                        <p>No attachments have been included.</p> 
                    <!-- /ko -->
                    <!-- ko if: description -->
                        <div data-bind="text: description"></div>
                    <!-- /ko -->
                </section>`
            });

            ko.components.register('so-created-modified-info', {
                viewModel: function (params) {
                    this.CreatedBy = <KnockoutObservable<ISpPerson>>params.createdBy
                    this.ModifiedBy = <KnockoutObservable<ISpPerson>>params.modifiedBy;
                    this.profiles = ko.observableArray([
                        { header: 'Created By', profile: this.CreatedBy },
                        { header: 'Modified By', profile: this.ModifiedBy }
                    ]);
                    this.Created = params.created;
                    this.Modified = params.modified;
                    this.showUserProfiles = params.showUserProfiles;
                },
                template:
                    `<!-- ko if: showUserProfiles() -->
                        <div class="create-mod-info no-print hidden-xs">
                            <!-- ko foreach: profiles -->
                                <div class="user-profile-card">
                                    <h4 data-bind="text: header"></h4>
                                    <!-- ko with: profile -->
                                        <img data-bind="attr: {src: Picture, alt: Name}" />
                                        <ul>
                                            <li><label>Name</label><span data-bind="text: Name"></span><li>
                                            <li><label>Job Title</label><span data-bind="text: JobTitle"></span></li>
                                            <li><label>Department</label><span data-bind="text: Department"></span></li>
                                            <li><label>Email</label><a data-bind="text: WorkEMail, attr: {href: (\'mailto:\' + WorkEMail)}"></a></li>
                                            <li><label>Phone</label><span data-bind="text: WorkPhone"></span></li>
                                            <li><label>Office</label><span data-bind="text: Office"></span></li>
                                        </ul>
                                    <!-- /ko -->
                                </div>
                            <!-- /ko -->             
                        </div>
                    <!-- /ko -->
                    <div class="row">
                        <!-- ko with: CreatedBy -->
                            <div class="col-md-3"><label>Created By</label> <a data-bind="text: Name, attr: {href: \'mailto:\' + WorkEMail}" class="email"> </a></div>
                        <!-- /ko -->
                        <div class="col-md-3"><label>Created</label> <span data-bind="spDateTime: Created"></span></div>                    
                        <!-- ko with: ModifiedBy -->
                            <div class="col-md-3"><label>Modified By</label> <a data-bind="text: Name, attr: {href: \'mailto:\' + WorkEMail}" class="email"></a></div>
                        <!-- /ko -->
                        <div class="col-md-3"><label>Modified</label> <span data-bind="spDateTime: Modified"></span></div>
                    </div>`
            });

            ko.components.register('so-workflow-history', {
                viewModel: function (params) {
                    this.historyItems = <Array<IHistoryItem>>(params.val || params.historyItems);
                },
                template:
                `<div class="row">
                    <div class="col-sm-8"><strong>Description</strong></div>
                    <div class="col-sm-4"><strong>Date</strong></div>
                </div>
                <!-- ko foreach: historyItems -->
                    <div class="row">
                        <div class="col-sm-8"><span data-bind="text: _description"></span></div>
                        <div class="col-sm-4"><span data-bind="spDateTime: _dateOccurred"></span></div>
                    </div>
                <!-- /ko -->`
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

                var labelX: number = parseInt(params.labelColWidth || 3); // Bootstrap label column width 1-12
                var fieldX: number = parseInt(params.fieldColWidth || (12 - (labelX - 0))); // Bootstrap field column width 1-12
                this.labelColWidth = 'col-sm-' + labelX;
                this.fieldColWidth = 'col-sm-' + fieldX;
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

        };

        //&& !!required && !readOnly
        private static hasErrorCssDiv: string = `<div class="form-group" data-bind="css: {\'has-error\': !!!modelValue() && !!required(), \'has-success has-feedback\': !!modelValue() && !!required()}">`;

        private static requiredFeedbackSpan: string = `<span class="glyphicon glyphicon-ok form-control-feedback" aria-hidden="true"></span>`;

        public static soStaticFieldTemplate: string =
        `<div class="form-group">
            <div class="row">            
                <!-- ko if: !!label -->
                    <div class="col-sm-3" data-bind="attr:{\'class\': labelColWidth}"><label data-bind="html: label"></label></div>
                <!-- /ko -->
                <div class="col-sm-9" data-bind="text: modelValue, attr:{\'class\': fieldColWidth}"></div>
            </div>
            <!-- ko if: description -->
            <div class="so-field-description"><p data-bind="html: description"></p></div>
            <!-- /ko -->
        </div>`;

        public static soTextFieldTemplate: string =
        `${KoComponents.hasErrorCssDiv}
        <div class="row">                 
	        <!-- ko if: !!label -->
		        <div class="col-sm-3" data-bind="attr:{\'class\': labelColWidth}"><label data-bind="html: label, attr: {for: id}"></label></div>
	        <!-- /ko -->          
	        <div class="col-sm-9" data-bind="attr:{\'class\': fieldColWidth}">
		        <!-- ko if: readOnly() -->
			        <div data-bind="text: modelValue"></div>
		        <!-- /ko -->
		        <!-- ko ifnot: readOnly() -->
			        <!-- ko if: multiline -->
				        <textarea data-bind="value: modelValue, css: {\'so-editable\': editable}, attr: {id: id, placeholder: placeholder, title: title, required: required, \'ko-name\': koName }" class="form-control"></textarea>
			        <!-- /ko -->
			        <!-- ko ifnot: multiline -->
				        <input type="text" data-bind="value: modelValue, css: {\'so-editable\': editable}, attr: {id: id, placeholder: placeholder, title: title, required: required, maxlength: maxlength, \'ko-name\': koName }" class="form-control" />
			        <!-- /ko -->
			        <!-- ko if: !!required() -->
				        ${KoComponents.requiredFeedbackSpan}
			        <!-- /ko -->
		        <!-- /ko -->
	        </div>
	        <!-- ko if: description -->
		    <div class="so-field-description"><p data-bind="html: description"></p></div>
	        <!-- /ko -->
        </div>`;

        public static soHtmlFieldTemplate: string =
        `${KoComponents.hasErrorCssDiv}
        <div class="row"> 
            <!-- ko if: !!label -->
                <div class="col-sm-3" data-bind="attr:{\'class\': labelColWidth}"><label data-bind="html: label, attr: {for: id}"></label></div>
            <!-- /ko -->
            <div class="col-sm-9" data-bind="attr:{\'class\': fieldColWidth}">
                <!-- ko if: readOnly() -->
                    <div data-bind="html: modelValue"></div>
                <!-- /ko -->
                <!-- ko ifnot: readOnly() -->
                    <div data-bind="spHtmlEditor: modelValue" contenteditable="true" class="form-control content-editable"></div>
                    <textarea data-bind="value: modelValue, css: {\'so-editable\': editable}, attr: {id: id, required: required, \'ko-name\': koName }" data-sp-html="" style="display:none;"></textarea>
                    <!-- ko if: !!required() -->
                        ${KoComponents.requiredFeedbackSpan}
                    <!-- /ko -->
                <!-- /ko -->
                </div>
            </div>
            <!-- ko if: description -->
                <div class="so-field-description"><p data-bind="html: description"></p></div>
            <!-- /ko -->
        </div>`;

        public static soCheckboxFieldTemplate: string =
        `<div class="form-group">
            <div class="row">
                <div class="col-sm-3" data-bind="attr:{\'class\': labelColWidth}"></div>
                <div class="col-sm-9" data-bind="attr:{\'class\': fieldColWidth}">
                    <!-- ko if: readOnly() -->
                        <div data-bind="text: !!modelValue() ? \'Yes\' : \'No\'"></div>
                    <!-- /ko -->
                    <!-- ko ifnot: readOnly() -->
                        <label class="checkbox">
                            <input type="checkbox" data-bind="checked: modelValue, css: {\'so-editable\': editable}, attr: {id: id, \'ko-name\': koName}, valueUpdate: valueUpdate" />
                            <span data-bind="html: label" style="margin-left:1em;"></span>
                        </label>
                    <!-- /ko -->
                </div>
            </div>
            <!-- ko if: description -->
                <div class="so-field-description"><p data-bind="html: description"></p></div>
            <!-- /ko -->
        </div>`;

        public static soSelectFieldTemplate: string =
            `${KoComponents.hasErrorCssDiv}<div class="row">
                <!-- ko if: !!label -->
                    <div class="col-sm-3" data-bind="attr:{\'class\': labelColWidth}"><label data-bind="html: label, attr: {for: id}"></label></div>
                <!-- /ko -->
                <div class="col-sm-9" data-bind="attr:{\'class\': fieldColWidth}">
                    <!-- ko if: readOnly() -->
                        <div data-bind="text: modelValue"></div>
                    <!-- /ko -->
                    <!-- ko ifnot: readOnly() -->
                        <select data-bind="value: modelValue, options: options, optionsCaption: caption, css: {\'so-editable\': editable}, attr: {id: id, title: title, required: required, \'ko-name\': koName}" class="form-control"></select>
                        <!-- ko if: !!required() -->
                            ${KoComponents.requiredFeedbackSpan}
                        <!-- /ko -->
                    <!-- /ko -->
                </div>
            </div>
            <!-- ko if: description -->
                <div class="so-field-description"><p data-bind="html: description"></p></div>
            <!-- /ko -->
        </div>`;

        public static soCheckboxGroupTemplate: string =
        `<div class="form-group">
            <!-- ko if: description -->
	            <div class="so-field-description"><p data-bind="html: description"></p></div>
            <!-- /ko -->
            <div class="row">
	            <!-- ko if: !!label -->
		            <div><label data-bind="html: label"></label></div>
	            <!-- /ko -->
	            <div>
		            <!-- ko if: readOnly() -->
			            <!-- ko ifnot: inline -->
				            <ul class="list-group">
					            <!-- ko foreach: modelValue -->
						            <li data-bind="text: $data" class="list-group-item"></li>
					            <!-- /ko -->
					            <!-- ko if: modelValue().length == 0 -->
						            <li class="list-group-item">--None--</li>
					            <!-- /ko -->
				            </ul>
			            <!-- /ko -->
			            <!-- ko if: inline -->
				            <!-- ko foreach: modelValue -->
					            <span data-bind="text: $data"></span>
					            <!-- ko if: $index() < $parent.modelValue().length-1 -->,&nbsp;<!-- /ko -->
				            <!-- /ko -->
				            <!-- ko if: modelValue().length == 0 -->
					            <span>--None--</span>
				            <!-- /ko -->
			            <!-- /ko -->
		            <!-- /ko -->
		            <!-- ko ifnot: readOnly() -->
			            <input type="hidden" data-bind="value: modelValue, attr:{required: !!required}" /><p data-bind="visible: !!required" class="req">(Required)</p>
			            <!-- ko foreach: options -->
				            <label data-bind="css:{\'checkbox\': !$parent.inline, \'checkbox-inline\': $parent.inline}">
					            <input type="checkbox" data-bind="checked: $parent.modelValue, css: {\'so-editable\': $parent.editable}, attr: {\'ko-name\': $parent.koName, \'value\': $data}" />
					            <span data-bind="text: $data"></span>
				            </label>
			            <!-- /ko -->
		            <!-- /ko -->
	            </div>
            </div>`;

        public static soRadioGroupTemplate: string =
        `<div class="form-group">
	        <!-- ko if: description -->
		        <div class="so-field-description"><p data-bind="html: description"></p></div>
	        <!-- /ko -->
	        <div class="row">
		        <!-- ko if: !!label -->
			        <div class="col-sm-3" data-bind="attr:{\'class\': labelColWidth}"><label data-bind="html: label"></label></div>
		        <!-- /ko -->
		        <div class="col-sm-9" data-bind="attr:{\'class\': fieldColWidth}">
			        <!-- ko if: readOnly() -->
				        <div data-bind="text: modelValue"></div>
			        <!-- /ko -->
			        <!-- ko ifnot: readOnly() -->
				        <!-- ko foreach: options -->  
					        <label data-bind="css:{\'radio\': !$parent.inline, \'radio-inline\': $parent.inline}">
						        <input type="radio" data-bind="checked: $parent.modelValue, attr:{value: $data, name: $parent.name, \'ko-name\': $parent.koName}, css:{\'so-editable\': $parent.editable}" />
						        <span data-bind="text: $data"></span>
					        </label>
				        <!-- /ko -->
			        <!-- /ko -->
		        </div>
	        </div>
        </div>`;

        public static soUsermultiFieldTemplate: string =
        `<div class="form-group">
	        <input type="hidden" data-bind="value: modelValue, css: {\'so-editable\': editable}, attr: {id: id, \'ko-name\': koName, required: required}" />
	        <div class="row">
		        <div class="col-md-3 col-xs-3">
			        <label data-bind="html: label"></label>
		        </div>
		        <div class="col-md-9 col-xs-9">
			        <!-- ko ifnot: readOnly -->
				        <input type="text" data-bind="spPerson: person, attr: {placeholder: placeholder}" />
				        <button class="btn btn-success" data-bind="click: addPerson, attr: {\'disabled\': person() == null, id: koName + \'_AddButton\' }"><span>Add</span></button>
				        <!-- ko if: showRequiredText -->
					        <div class="col-md-6 col-xs-6">
						        <p class="text-danger">At least one person must be added.</p>
					        </div>
				        <!-- /ko -->
			        <!-- /ko -->		
			        <!-- ko foreach: modelValue -->
				        <div class="row">
					        <div class="col-md-10 col-xs-10" data-bind="spPerson: $data"></div>
					        <!-- ko ifnot: $parent.readOnly() -->
						        <div class="col-md-2 col-xs-2">
							        <button class="btn btn-xs btn-danger" data-bind="click: $parent.removePerson"><span class="glyphicon glyphicon-trash"></span></button>
						        </div>
					        <!-- /ko -->
				        </div>
			        <!-- /ko -->
			        <!-- ko if: description -->
				        <div class="so-field-description"><p data-bind="html: description"></p></div>
			        <!-- /ko -->		
		        </div>
	        </div>
        </div>`;

        public static soCreatedModifiedTemplate =
        `<!-- ko if: !!CreatedBy && CreatedBy() != null -->
            <section>
                <so-created-modified-info params="created: Created, createdBy: CreatedBy, modified: Modified, modifiedBy: ModifiedBy, showUserProfiles: showUserProfiles"></so-created-modified-info>
            </section>
        <!-- /ko -->`;

        public static soWorkflowHistoryTemplate =
        `<!-- ko if: !!Id() -->
            <section id="workflowHistory" class="nav-section">
                <h4>Workflow History</h4>
                <so-workflow-history params="val: historyItems"></so-workflow-history>
            </section>
        <!-- /ko -->`;
    }

}