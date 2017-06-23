module Shockout {

    export class Templates {

        public static buttonDefault = 'btn btn-sm btn-default no-print';

        public static calendarIcon = '<span class="glyphicon glyphicon-calendar"></span>';
        
        public static personIcon = '<span class="glyphicon glyphicon-user"></span>';
        
        public static resetButton = 'btn btn-sm btn-default no-print reset';
        
        public static timeControlsHtml = `<span class="glyphicon glyphicon-calendar"></span>
            <select class="form-control so-select-hours" style="margin-left:1em; max-width:5em; display:inline-block;">{0}</select><span> : </span>
            <select class="form-control so-select-minutes" style="width:5em; display:inline-block;">{1}</select>
            <select class="form-control so-select-tt" style="margin-left:1em; max-width:5em; display:inline-block;"><option value="AM">AM</option><option value="PM">PM</option></select>
            <button class="btn btn-sm btn-default reset" style="margin-left:1em;">Reset</button>
            <span class="error no-print" style="display:none;">Invalid Date-time</span>
            <span class="so-datetime-display no-print" style="margin-left:1em;"></span>`;
        
        public static getTimeControlsHtml = (): string => {
            var hrsOpts = [];
            for (var i = 1; i <= 12; i++) {
                hrsOpts.push('<option value="' + i + '">' + (i < 10 ? '0' + i : i) + '</option>');
            }

            var mmOpts = [];
            for (var i = 0; i < 60; i++) {
                mmOpts.push('<option value="' + i + '">' + (i < 10 ? '0' + i : i) + '</option>');
            }
            
            return Templates.timeControlsHtml.replace('{0}', hrsOpts.join('')).replace('{1}', mmOpts.join(''));
            
        };
        
        public static soFormAction: string = 
        `<div class="row">
            <div class="col-sm-8 col-sm-offset-4 text-right">
                <label>Logged in as:</label><span data-bind="text: currentUser().title" class="current-user"></span>
                <button class="btn btn-default cancel" data-bind="event: { click: cancel }" title="Close"><span class="glyphicon glyphicon-remove"></span><span class="hidden-xs">Close</span></button>
                <!-- ko if: allowPrint() -->
                    <button class="btn btn-primary print" data-bind="event: {click: print}" title="Print"><span class="glyphicon glyphicon-print"></span><span class="hidden-xs">Print</span></button>
                <!-- /ko -->
                <!-- ko if: allowDelete() && Id() != null -->
                    <button class="btn btn-warning delete" data-bind="event: {click: deleteItem}" title="Delete"><span class="glyphicon glyphicon-remove"></span><span class="hidden-xs">Delete</span></button>
                <!-- /ko -->
                <!-- ko if: allowSave() -->
                    <button class="btn btn-success save" data-bind="event: { click: save }" title="Save your work."><span class="glyphicon glyphicon-floppy-disk"></span><span class="hidden-xs">Save</span></button>
                <!-- /ko -->
                <button class="btn btn-danger submit" data-bind="event: { click: submit }" title="Submit for routing."><span class="glyphicon glyphicon-floppy-open"></span><span class="hidden-xs">Submit</span></button>
            </div>
        </div>`;

        public static getFormAction(): HTMLDivElement {
            var div = document.createElement('div');
            div.className = 'form-action no-print';
            div.innerHTML = Templates.soFormAction;
            return div;
        }

        public static hasErrorCssDiv: string =
        '<div class="form-group" data-bind="css: {\'has-error\': !!!modelValue() && !!required(), \'has-success has-feedback\': !!modelValue() && !!required()}">';

        public static requiredFeedbackSpan: string = '<span class="glyphicon glyphicon-ok form-control-feedback" aria-hidden="true"></span>';

        public static soNavMenuControl: string = '<so-nav-menu params="val: navMenuItems"></so-nav-menu>';

        public static soNavMenu: string =
        `<nav class="navbar navbar-default no-print" id="TOP">
            <div class="navbar-header">
                <button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#bs-example-navbar-collapse-1" aria-expanded="false">
                <span class="sr-only">Toggle navigation</span>
                <span class="icon-bar"></span>
                <span class="icon-bar"></span>
                <span class="icon-bar"></span>
                </button>
            </div>
            <div class="collapse navbar-collapse" id="bs-example-navbar-collapse-1">
                <ul class="nav navbar-nav" data-bind="foreach: navMenuItems">
                    <li><a href="#" data-bind="html: title, attr: {'href': '#' + anchorName }"></a></li>
                </ul>
            </div>
        </nav>`;

        public static soStaticField: string =
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

        public static soTextField: string =
        `${Templates.hasErrorCssDiv}
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
				        ${Templates.requiredFeedbackSpan}
			        <!-- /ko -->
		        <!-- /ko -->
	        </div>
	        <!-- ko if: description -->
		    <div class="so-field-description"><p data-bind="html: description"></p></div>
	        <!-- /ko -->
        </div>`;

        public static soAttachments: string =
        `<section class="nav-section">
            <h4><span data-bind="text: title"></span> &ndash; <span data-bind="text: length" class="badge"></span></h4>
            <div data-bind="visible: !!errorMsg()" class="alert alert-danger alert-dismissable">
                <button type="button" class="close" data-dismiss="alert" aria-label="Close"><span aria-hidden="true">&times;</span></button>
                <span class="glyphicon glyphicon-exclamation-sign"></span>&nbsp;<span data-bind="text: errorMsg"></span>
            </div>
            <!-- ko ifnot: hasFileReader() -->
            <div data-bind="visible: !!!readOnly(), attr: {id: this.qqFileUploaderId}"></div>
            <!-- /ko -->
            <!-- ko if: !readOnly() && hasFileReader() -->
                <div class="row">
                    <div class="col-md-2 so-attach-files-btn">
                        <input type="file" data-bind="attr: {'id': id}, event: {'change': fileHandler}" multiple class="form-control" style="display:none;" />
                        <div data-bind="attr:{'class': className}, event: {'click': onSelect}"><span class="glyphicon glyphicon-paperclip"></span>&nbsp;<span data-bind="text: label"></span></div>
                    </div>
                    <div class="col-md-10">
                        <!-- ko if: drop -->
                            <div class="so-file-dropzone" data-bind="event: {'dragenter': onDragenter, 'dragover': onDragover, 'drop': onDrop}">
                                <div><span data-bind="text: dropLabel"></span> <span class="glyphicon glyphicon-upload"></span></div>
                            </div>
                        <!-- /ko -->
                    </div>
                </div>
                <!-- ko foreach: fileUploads -->
                    <div class="progress">
                        <div data-bind="attr: {'aria-valuenow': progress(), 'style': 'width:' + progress() + '%;', 'class': className() }" role="progressbar" aria-valuemin="0" aria-valuemax="100">
                            <span data-bind="text: fileName"></span>
                        </div>
                    </div>
                <!-- /ko -->
            <!-- /ko -->
            <div data-bind="foreach: attachments" style="margin:1em auto;">
                <div class="so-attachment">
                    <a href="" data-bind="attr: {href: __metadata.media_src}"><span class="glyphicon glyphicon-paperclip"></span>&nbsp;<span data-bind="text: Name"></span></a>
                    <!-- ko ifnot: $parent.readOnly() -->
                    <button data-bind="event: {click: $parent.deleteAttachment}" class="btn btn-sm btn-danger delete" title="Delete Attachment"><span class="glyphicon glyphicon-remove"></span></button>
                    <!-- /ko -->
                </div>
            </div>
            <!-- ko if: length() == 0 && readOnly() -->
                <p>No attachments have been included.</p>
            <!-- /ko -->
            <!-- ko if: description -->
                <div data-bind="text: description"></div>
            <!-- /ko -->
        </section>`;

        public static soHtmlFieldTemplate: string =
        `${Templates.hasErrorCssDiv}
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
                        ${Templates.requiredFeedbackSpan}
                    <!-- /ko -->
                <!-- /ko -->
                </div>
            </div>
            <!-- ko if: description -->
                <div class="so-field-description"><p data-bind="html: description"></p></div>
            <!-- /ko -->
        </div>`;

        public static soCheckboxField: string =
        `<div class="form-group">
            <div class="row">
                <div class="col-sm-3" data-bind="attr:{\'class\': labelColWidth}">
                    <!-- ko if: readOnly() -->
                    <label data-bind="html: label, attr: {for: id}"></label>
                    <!-- /ko -->
                </div>
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

        public static soSelectField: string =
        `${Templates.hasErrorCssDiv}<div class="row">
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
                            ${Templates.requiredFeedbackSpan}
                        <!-- /ko -->
                    <!-- /ko -->
                </div>
            </div>
            <!-- ko if: description -->
                <div class="so-field-description"><p data-bind="html: description"></p></div>
            <!-- /ko -->
        </div>`;

        public static soCheckboxGroup: string =
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

        public static soRadioGroup: string =
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

        public static soUsermultiField: string =
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

        public static soWorkflowHistoryControl =
        `<!-- ko if: !!Id() && historyItems().length > 0 -->
            <section id="workflowHistory" class="nav-section">
                <h4>Workflow History</h4>
                <so-workflow-history params="val: historyItems"></so-workflow-history>
            </section>
        <!-- /ko -->`;

        public static soWorkflowHistory: string =
        `<div class="row">
            <div class="col-sm-8"><strong>Description</strong></div>
            <div class="col-sm-4"><strong>Date</strong></div>
        </div>
        <!-- ko foreach: historyItems -->
            <div class="row">
                <div class="col-sm-8"><span data-bind="text: _description"></span></div>
                <div class="col-sm-4"><span data-bind="spDateTime: _dateOccurred"></span></div>
            </div>
        <!-- /ko -->`;

        public static soCreatedModifiedInfoControl =
        `<!-- ko if: !!CreatedBy && CreatedBy() != null -->
            <section class="nav-section">
                <h4>Created/Modified By</h4>
                <so-created-modified-info params="created: Created, createdBy: CreatedBy, modified: Modified, modifiedBy: ModifiedBy, showUserProfiles: showUserProfiles"></so-created-modified-info>
            </section>
        <!-- /ko -->`;

        public static soCreatedModifiedInfo: string =
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
        </div>`;
    }
   
}