module Shockout {

    export class Templates {

        // dropped support for qq fileuploader
        //public static fileuploadTemplate: string = `
        //<div class="qq-uploader" data-author-only>
        //    <div class="qq-upload-drop-area"><span>Drop files here to upload</span></div>
        //    <div class="btn btn-primary qq-upload-button"><span class="glyphicon glyphicon-paperclip"></span> Attach File</div>
        //    <ul class="qq-upload-list"></ul>
        //</div>`;

        public static actionTemplate: string = 
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
            div.innerHTML = Templates.actionTemplate;
            return div;
        }
    }
   
}