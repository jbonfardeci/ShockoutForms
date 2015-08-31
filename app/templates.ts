module Shockout {

    export class Templates {
        
        public static createdByKey: string = 'CreatedBy';
        public static modifiedByKey: string = 'ModifiedBy';
        public static createdKey: string = 'Created';
        public static modifiedKey: string = 'Modified';
        public static historyKey: string = 'history';
        public static historyDescriptionKey: string = 'description';
        public static historyDateKey: string = 'date';

        public static getFileUploadTemplate(): string {
            return '<div class="qq-uploader">' +
                '<div class="qq-upload-drop-area"><span>Drop files here to upload</span></div>' +
                '<div class="btn qq-upload-button">Attach Files</div>' +
                '<ul class="qq-upload-list"></ul>' +
                '</div>';
        }

        public static getCreatedModifiedInfo(): HTMLElement {
            var template: string = '<h4>Created/Modified Information</h4>' +
                '<ul>' +
                '<li class="create-mod-info no-print"></li>' +
                '<li><label>Created By</label><a data-bind="text: {0}().Name, attr:{href: \'mailto:\'+{0}().WorkEMail}" class="email"></a></li>' +
                '<li><label>Created</label><span data-bind="spDateTime: {1}"></span></li>' +
                '<li><label>Modified By</label><a data-bind="text: {2}().Name, attr:{href: \'mailto:\'+{2}().WorkEMail}" class="email"></a></li>' +
                '<li><label>Modified</label><span data-bind="spDateTime: {3}"></span></li>' +
                '</ul>';

            var section: HTMLElement = document.createElement('section');
            section.className = 'created-mod-info';
            section.innerHTML = template
                .replace(/\{0\}/g, Templates.createdByKey)
                .replace(/\{1\}/g, Templates.createdKey)
                .replace(/\{2\}/g, Templates.modifiedByKey)
                .replace(/\{3\}/g, Templates.modifiedKey);
            return section;
        }

        public static getHistoryTemplate(): HTMLElement {
            var template: string = '<h4>Workflow History</h4>' +
                '<table border="1" cellpadding="5" cellspacing="0" class="data-table" style="width:100%;border-collapse:collapse;">' +
                '<thead>' +
                '<tr><th>Description</th><th>Date</th></tr>' +
                '</thead>' +
                '<tbody data-bind="foreach: {0}">' +
                '<tr><td data-bind="text: {1}"></td><td data-bind="text: {2}"></td></tr>' +
                '</tbody>' +
                '</table>';
            var section: HTMLElement = document.createElement('section');
            section.setAttribute('data-bind', 'visible: {0}.length > 0'.replace(/\{0\}/i, Templates.historyKey));
            section.innerHTML = template
                .replace(/\{0\}/g, Templates.historyKey)
                .replace(/\{1\}/g, Templates.historyDescriptionKey)
                .replace(/\{2\}/g, Templates.historyDateKey);
            return section;
        }

        public static getFormAction(allowSave: boolean = true, allowDelete: boolean = true, allowPrint: boolean = true): HTMLDivElement {
            var template: Array<string> = [];
            template.push('<div class="form-breadcrumbs"><a href="/">Home</a> &gt; eForms</div>');
            template.push('<button class="btn cancel" data-bind="event: { click: cancel }"><span>Close</span></button>');

            if (allowPrint) {
                template.push('<button class="btn print" data-bind="visible: Id() != null, event: {click: print}"><span>Print</span></button>');
            }
            if (allowDelete) {
                template.push('<button class="btn delete" data-bind="visible: Id() != null, event: {click: deleteItem}"><span>Delete</span></button>');
            }
            if (allowSave) {
                template.push('<button class="btn save" data-bind="event: { click: save }"><span>Save</span></button>');
            }

            template.push('<button class="btn submit" data-bind="event: { click: submit }"><span>Submit</span></button>');

            var div: HTMLDivElement = document.createElement('div');
            div.className = 'form-action';
            div.innerHTML = template.join('');
            return div;
        }

        public static getAttachmentsTemplate(fileuploaderId: string): HTMLDivElement {
            var template = '<h4>Attachments</h4>' +
                '<div id="{0}"></div>' +
                '<table class="attachments-table">' +
                '<tbody data-bind="foreach: attachments">' +
                '<tr>' +
                '<td><a href="" data-bind="text: title, attr: {href: href, \'class\': ext}"></a></td>' +
                '<td><button data-bind="event: {click: $root.deleteAttachment}" class="btn del" title="Delete"><span>Delete</span></button></td>' +
                '</tr>' +
                '</tbody>' +
                '</table>';
            var div: HTMLDivElement = document.createElement('div');
            div.innerHTML = template.replace(/\{0\}/, fileuploaderId);
            return div;
        }

        public static getUserProfileTemplate(profile: ISpPerson, headerTxt: string): HTMLDivElement {

            var template = '<h4>{header}</h4>' +
                '<img src="{pictureurl}" alt="{name}" />' +
                '<ul>' +
                '<li><label>Name</label>{name}</li>' +
                '<li><label>Title</label>{jobtitle}</li>' +
                '<li><label>Department</label>{department}</li>' +
                '<li><label>Email</label><a href="mailto:{workemail}">{workemail}</a></li>' +
                '<li><label>Phone</label>{workphone}</li>' +
                '<li><label>Office</label>{office}</li>' +
                '</ul>';

            var div: HTMLDivElement = document.createElement("div");
            div.className = "user-profile-card";
            div.innerHTML = template
                .replace(/\{header\}/g, headerTxt)
                .replace(/\{pictureurl\}/g, profile.Picture)
                .replace(/\{name\}/g, (profile.Name || ''))
                .replace(/\{jobtitle\}/g, profile.Title || '')
                .replace(/\{department\}/g, profile.Department || '')
                .replace(/\{workemail\}/g, profile.WorkEMail || '')
                .replace(/\{workphone\}/g, profile.WorkPhone || '')
                .replace(/\{office\}/g, profile.Office || '')
            ;
            return div;
        }
    }
   
}