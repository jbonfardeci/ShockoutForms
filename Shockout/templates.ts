module Shockout {

    export class Templates {
        
        public static getFileUploadTemplate(): string {
            var $div = $('<div>').html('<div class="qq-uploader">' +
                        '<div class="qq-upload-drop-area"><span>Drop files here to upload</span></div>' +
                        '<div class="btn btn-primary qq-upload-button"><span class="glyphicon glyphicon-paperclip"></span> Attach File</div>' +
                        '<ul class="qq-upload-list"></ul></div>');
            return $div.html();
        }

        public static getCreatedModifiedHtml(): string {
            var template: string = '<div class="create-mod-info no-print"></div>' +
                '<div class="row">' +
                '<div class="col-md-3"><label>Created By</label> <a data-bind="text: {0}, attr:{href: \'mailto:\'+{1}()}" class="email" > </a></div>' +
                '<div class="col-md-3"><label>Created</label> <span data-bind="spDateTime: {2}"></span></div>' +
                '<div class="col-md-3"><label>Modified By</label> <a data-bind="text: {3}, attr:{href: \'mailto:\'+{4}()}" class="email"></a></div>' +
                '<div class="col-md-3"><label>Modified</label> <span data-bind="spDateTime: {5}"></span></div>' +
                '</div>';
            
            template = template.replace(/\{0\}/g, ViewModel.createdByKey)
                    .replace(/\{1\}/g, ViewModel.createdByEmailKey)
                    .replace(/\{2\}/g, ViewModel.createdKey)
                    .replace(/\{3\}/g, ViewModel.modifiedByKey)
                    .replace(/\{4\}/g, ViewModel.modifiedByEmailKey)
                    .replace(/\{5\}/g, ViewModel.modifiedKey);

            var $div = $('<div>').html(template);
            return $div.html();          
        }

        public static getHistoryTemplate(): JQuery {
            var template: string =
                '<h4>Workflow History</h4>' +
                '<table border="1" cellpadding="5" cellspacing="0" class="data-table" style="width:100%;border-collapse:collapse;">' +
                '<thead>' +
                '<tr><th>Description</th><th>Date</th></tr>' +
                '</thead>' +
                '<tbody data-bind="foreach: {0}">' +
                '<tr><td data-bind="text: {1}"></td><td data-bind="text: {2}"></td></tr>' +
                '</tbody>' +
                '</table>'.replace(/\{0\}/g, ViewModel.historyKey)
                    .replace(/\{1\}/g, ViewModel.historyDescriptionKey)
                    .replace(/\{2\}/g, ViewModel.historyDateKey);

            var $div = $('<div>', {
                'data-bind': 'visible: {0}().length > 0'.replace(/\{0\}/i, ViewModel.historyKey)
            });

            return $div;
        }

        public static getFormAction(allowSave: boolean = true, allowDelete: boolean = true, allowPrint: boolean = true): JQuery {
            var template: Array<string> = [];
            //template.push('<div class="form-breadcrumbs"><a href="/">Home</a> &gt; eForms</div>');
            template.push('<button class="btn btn-default cancel" data-bind="event: { click: cancel }"><span>Close</span></button>');

            if (allowPrint) {
                template.push('<button class="btn btn-primary print" data-bind="visible: Id() != null, event: {click: print}"><span class="glyphicon glyphicon-print"></span><span>Print</span></button>');
            }
            if (allowDelete) {
                template.push('<button class="btn btn-warning delete" data-bind="visible: Id() != null, event: {click: deleteItem}"><span class="glyphicon glyphicon-remove"></span><span>Delete</span></button>');
            }
            template.push('<button class="btn btn-success save" data-bind="event: { click: save }" style="display:none;"><span class="glyphicon glyphicon-floppy-disk"></span><span>Save</span></button>');

            template.push('<button class="btn btn-danger submit" data-bind="event: { click: submit }, disable: !isValid()"><span class="glyphicon glyphicon-floppy-open"></span><span>Submit</span></button>');

            var $div = $('<div>', { 'class': 'form-action no-print', 'html': template.join('') });
            return $div;
        }

        public static getAttachmentsTemplate(fileuploaderId: string): JQuery {
            var template =
                '<h4>Attachments (<span data-bind="text: attachments().length"></span>)</h4>' + 
                '<div id="' + fileuploaderId + '"></div>' + 
                '<div data-bind="visible: attachments().length > 0">' + 
                '<table class="attachments-table">' +
                '<tbody data-bind="foreach: attachments">' +
                '<tr>' +
                '<td><a href="" data-bind="text: title, attr: {href: href, \'class\': ext}"></a></td>' +
                '<td><button data-bind="event: {click: $root.deleteAttachment}" class="btn del" title="Delete"><span class="glyphicon glyphicon-remove"></span><span>Delete</span></button></td>' +
                '</tr>' +
                '</tbody>' +
                '</table>' + 
                '</div>';

            var $div = $('<div>', { 'html': template });
            return $div;
        }

        public static getUserProfileTemplate(profile: ISpPerson, headerTxt: string): JQuery{

            var template: string =
                '<h4>{header}</h4>' + 
                '<img src="{pictureurl}" alt="{name}" />' + 
                '<ul>' +  
                '<li><label>Name</label>{name}<li>' + 
                '<li><label>Title</label>{jobtitle}</li>' + 
                '<li><label>Department</label>{department}</li>' + 
                '<li><label>Email</label><a href="mailto:{workemail}">{workemail}</a></li>' + 
                '<li><label>Phone</label>{workphone}</li>' +
                '<li><label>Office</label>{office}</li>' +
                '</ul>';

            template = template.replace(/\{header\}/g, headerTxt)
                .replace(/\{pictureurl\}/g, (profile.Picture.indexOf(',') > 0 ? profile.Picture.split(',')[0] : profile.Picture))
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