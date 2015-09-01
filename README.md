# ShockoutForms
### a Work in Progress
SharePoint + Knockout MVVM forms - an InfoPath killer

Forget the frustrations of legacy InfoPath and XSL SharePoint forms. Leverage the power of Knockout's databinding with this framework.

#### Usage
<pre>var spForm = new Shockout.SPForm(
	/*listName:*/ 'My SharePoint List Name', 
	/*formId:*/ 'my-form-ID', 
	/*options:*/ {
		debug: false, // default false
		preRender: function(spForm){}, // default undefined
        postRender: function(spForm){}, // default undefined
        preSave: function(spForm){}, // default undefined	
		allowDelete: false, // default false
        allowPrint: true, // default true
        allowSave: true, // default true
        allowedExtensions: []  // default is ['txt', 'rtf', 'zip', 'pdf', 'doc', 'docx', 'jpg', 'gif', 'png', 'ppt', 'tif', 'pptx', 'csv', 'pub', 'msg']
        attachmentMessage: 'An attachment is required.', // the default
        confirmationUrl: '/SitePages/Confirmation.aspx', // the default
        enableErrorLog: true, // default true
        errorLogListName: 'Error Log', // Designated SharePoint list for logging user and form errrors. default 'Error Log'
        fileHandlerUrl: string = '/_layouts/SPFormFileHandler.ashx',  // default    
        hasAttachments: true, // default true
        includeUserProfiles: true, // default true
        includeWorkflowHistory: true, // default true        
        requireAttachments: false, // default false
        siteUrl: '/companyForms', // default is ''        
        workflowHistoryListName: 'Workflow History' // the default
	});
</pre>

