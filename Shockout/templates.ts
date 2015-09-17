module Shockout {

    export class Templates {

        public static attachmentsTemplate = '<h4>Attachments (<span data-bind="text: attachments().length"></span>)</h4>' +
            '<div id="{0}"></div>' +
            '<div data-bind="foreach: attachments">' +
            '<div>' +
            '<a href="" data-bind="attr: {href: __metadata.media_src}"><span class="glyphicon glyphicon-paperclip"></span> <span data-bind="text: Name"></span></a>&nbsp;' +
            '<button data-bind="event: {click: $root.deleteAttachment}" class="btn btn-sm btn-danger" title="Delete Attachment" data-author-only><span class="glyphicon glyphicon-remove"></span></button>' +
            '</div>' +
            '</div>';

        public static fileuploadTemplate: string = '<div class="qq-uploader" data-author-only>' +
            '<div class="qq-upload-drop-area"><span>Drop files here to upload</span></div>' +
            '<div class="btn btn-primary qq-upload-button"><span class="glyphicon glyphicon-paperclip"></span> Attach File</div>' +
            '<ul class="qq-upload-list"></ul></div>';

        public static createdModifiedTemplate: string = '<div class="create-mod-info no-print hidden-xs"></div>' +
            '<div class="row">' +
            '<div class="col-md-3"><label>Created By</label> <a data-bind="text: {0}, attr:{href: \'mailto:\'+{1}()}" class="email" > </a></div>' +
            '<div class="col-md-3"><label>Created</label> <span data-bind="spDateTime: {2}"></span></div>' +
            '<div class="col-md-3"><label>Modified By</label> <a data-bind="text: {3}, attr:{href: \'mailto:\'+{4}()}" class="email"></a></div>' +
            '<div class="col-md-3"><label>Modified</label> <span data-bind="spDateTime: {5}"></span></div>' +
            '</div>';
        
        public static historyTemplate: string = '<h4>Workflow History</h4>' +
            '<div class="row">' +
            '<div class="col-md-8 col-xs-8"><strong>Description</strong></div>' +
            '<div class="col-md-4 col-xs-4"><strong>Date</strong></div>' +
            '</div>' +
            '<div class="row" data-bind="foreach: {0}">' +
            '<div data-bind="text: {1}" class="col-md-8 col-xs-8"></div>' +
            '<div data-bind="spDateTime: {2}" class="col-md-4 col-xs-4"></div>' +
            '</div>';

        public static userProfileTemplate = '<h4>{header}</h4>' +
            '<img src="{pictureurl}" alt="{name}" />' +
            '<ul>' +
            '<li><label>Name</label>{name}<li>' +
            '<li><label>Title</label>{jobtitle}</li>' +
            '<li><label>Department</label>{department}</li>' +
            '<li><label>Email</label><a href="mailto:{workemail}">{workemail}</a></li>' +
            '<li><label>Phone</label>{workphone}</li>' +
            '<li><label>Office</label>{office}</li>' +
            '</ul>';

        public static getFileUploadTemplate(): string {
            var $div = $('<div>').html(Templates.fileuploadTemplate);
            return $div.html();
        }

        public static getCreatedModifiedHtml(): string {
            var template: string = Templates.createdModifiedTemplate.replace(/\{0\}/g, ViewModel.createdByKey)
                    .replace(/\{1\}/g, ViewModel.createdByEmailKey)
                    .replace(/\{2\}/g, ViewModel.createdKey)
                    .replace(/\{3\}/g, ViewModel.modifiedByKey)
                    .replace(/\{4\}/g, ViewModel.modifiedByEmailKey)
                    .replace(/\{5\}/g, ViewModel.modifiedKey);

            var $div = $('<div>').html(template);
            return $div.html();          
        }

        public static getHistoryTemplate(): JQuery {
            var template: string = Templates.historyTemplate.replace(/\{0\}/g, ViewModel.historyKey)
                        .replace(/\{1\}/g, ViewModel.historyDescriptionKey)
                        .replace(/\{2\}/g, ViewModel.historyDateKey);

            var $div = $('<section>', {
                'html': template,
                'data-bind': 'visible: {0}().length > 0'.replace(/\{0\}/, ViewModel.historyKey)
            });

            return $div;
        }

        public static getFormAction(allowSave: boolean = true, allowDelete: boolean = true, allowPrint: boolean = true): JQuery {
            var template: Array<string> = [];
            //template.push('<div class="form-breadcrumbs"><a href="/">Home</a> &gt; eForms</div>');
            template.push('<button class="btn btn-default cancel" data-bind="event: { click: cancel }" title="Close"><span class="glyphicon glyphicon-remove"></span><span class="hidden-xs">Close</span></button>');

            if (allowPrint) {
                template.push('<button class="btn btn-primary print" data-bind="visible: Id() != null, event: {click: print}" title="Print"><span class="glyphicon glyphicon-print"></span><span class="hidden-xs">Print</span></button>');
            }
            if (allowDelete) {
                template.push('<button class="btn btn-warning delete" data-bind="visible: Id() != null, event: {click: deleteItem}" title="Delete"><span class="glyphicon glyphicon-remove"></span><span class="hidden-xs">Delete</span></button>');
            }
            template.push('<button class="btn btn-success save" data-bind="event: { click: save }" style="display:none;" title="Save your work."><span class="glyphicon glyphicon-floppy-disk"></span><span class="hidden-xs">Save</span></button>');

            template.push('<button class="btn btn-danger submit" data-bind="event: { click: submit }, disable: !isValid()" title="Submit for routing."><span class="glyphicon glyphicon-floppy-open"></span><span class="hidden-xs">Submit</span></button>');

            var $div = $('<div>', { 'class': 'form-action no-print', 'html': template.join('') });
            return $div;
        }

        public static getAttachmentsTemplate(fileuploaderId: string): JQuery {
            var template = Templates.attachmentsTemplate.replace(/\{0\}/, fileuploaderId);
            var $div = $('<div>', { 'html': template });
            return $div;
        }

        public static getUserProfileTemplate(profile: ISpPerson, headerTxt: string): JQuery{

            var pictureUrl = '/_layouts/images/person.gif';
            if (profile.Picture != null && profile.Picture.indexOf(',') > -1) {
                pictureUrl = profile.Picture.split(',')[0];
            }

            var template: string = Templates.userProfileTemplate.replace(/\{header\}/g, headerTxt)
                .replace(/\{pictureurl\}/g, pictureUrl)
                .replace(/\{name\}/g, (profile.Name || ''))
                .replace(/\{jobtitle\}/g, profile.Title || '')
                .replace(/\{department\}/g, profile.Department || '')
                .replace(/\{workemail\}/g, profile.WorkEMail || '')
                .replace(/\{workphone\}/g, profile.WorkPhone || '')
                .replace(/\{office\}/g, profile.Office || '');

            var $div = $('<div>', { 'class': 'user-profile-card', 'html': template });
            return $div;
        }
    }
   
}