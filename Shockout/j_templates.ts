module Shockout {

    export class Templates {

        public static attachmentsTemplate = '<h4>Attachments (<span data-bind="text: attachments().length"></span>)</h4>\
            <div id="{0}"></div>\
            <div data-bind="foreach: attachments">\
            <div>\
            <a href="" data-bind="attr: {href: __metadata.media_src}"><span class="glyphicon glyphicon-paperclip"></span> <span data-bind="text: Name"></span></a>&nbsp;\
            <button data-bind="event: {click: $root.deleteAttachment}" class="btn btn-sm btn-danger" title="Delete Attachment"><span class="glyphicon glyphicon-remove"></span></button>\
            </div>\
            </div>';

        public static fileuploadTemplate: string = '<div class="qq-uploader" data-author-only>\
            <div class="qq-upload-drop-area"><span>Drop files here to upload</span></div>\
            <div class="btn btn-primary qq-upload-button"><span class="glyphicon glyphicon-paperclip"></span> Attach File</div>\
            <ul class="qq-upload-list"></ul></div>';

        public static getFileUploadTemplate(): string {
            var $div = $('<div>').html(Templates.fileuploadTemplate);
            return $div.html();
        }

        public static getFormAction(allowSave: boolean = true, allowDelete: boolean = true, allowPrint: boolean = true): JQuery {
            var template: Array<string> = [];
            template.push('<label>Logged in as:</label><span data-bind="text: currentUser().title" class="current-user"></span>');
            template.push('<button class="btn btn-default cancel" data-bind="event: { click: cancel }" title="Close"><span class="glyphicon glyphicon-remove"></span><span class="hidden-xs">Close</span></button>');

            if (allowPrint) {
                template.push('<button class="btn btn-primary print" data-bind="visible: Id() != null, event: {click: print}" title="Print"><span class="glyphicon glyphicon-print"></span><span class="hidden-xs">Print</span></button>');
            }
            if (allowDelete) {
                template.push('<button class="btn btn-warning delete" data-bind="visible: Id() != null, event: {click: deleteItem}" title="Delete"><span class="glyphicon glyphicon-remove"></span><span class="hidden-xs">Delete</span></button>');
            }
            template.push('<button class="btn btn-success save" data-bind="event: { click: save }" style="display:none;" title="Save your work."><span class="glyphicon glyphicon-floppy-disk"></span><span class="hidden-xs">Save</span></button>');

            template.push('<button class="btn btn-danger submit" data-bind="event: { click: submit }" title="Submit for routing."><span class="glyphicon glyphicon-floppy-open"></span><span class="hidden-xs">Submit</span></button>');

            var $div = $('<div>', { 'class': 'form-action no-print', 'html': template.join('') });
            return $div;
        }

        public static getAttachmentsTemplate(fileuploaderId: string): JQuery {
            var template = Templates.attachmentsTemplate.replace(/\{0\}/, fileuploaderId);
            var $div = $('<div>', { 'html': template });
            return $div;
        }
    }
   
}