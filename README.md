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
        enableAttachments: true, // default true
        includeUserProfiles: true, // default true
        includeWorkflowHistory: true, // default true        
        requireAttachments: false, // default false
        siteUrl: '/companyForms', // default is ''        
        workflowHistoryListName: 'Workflow History' // the default
	});
</pre>

### Custom Knockout binding handlers for SP list field types included:
	
####spContentEditor
`<textarea data-bind="spContentEditor: myVar"></textarea>`

####spPerson
`<input type="text" data-bind="spPerson: myVar" />`
	OR
`<div data-bind="spPerson: myVar" />`

####spDate
`<input type="text" data-bind="spDate: myVar" />`	
	OR
`<div data-bind="spDate: myVar" />`

####spDateTime
`<input type="text" data-bind="spDateTime: myVar" />`
	OR
`<div data-bind="spDateTime: myVar" />`

####spMoney
`<input type="text" data-bind="spMoney: myVar" />`
	OR
`<div data-bind="spMoney: myVar" />`

####spDecimal
`<input type="text" data-bind="spDecimal: myVar" />`
	OR
`<div data-bind="spDecimal: myVar" />`

####spNumber
`<input type="text" data-bind="spNumber: myVar" />`
	OR
`<div data-bind="spNumber: myVar" />`


Copyright (C) 2015  John T. Bonfardeci

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as
published by the Free Software Foundation, either version 3 of the
License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.

