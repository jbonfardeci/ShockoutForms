# ShockoutForms
### a Work in Progress
SharePoint + Knockout MVVM forms - an InfoPath killer

Forget the frustrations of legacy InfoPath and XSL SharePoint forms. Leverage the power of Knockout's databinding with this framework.

####Dependencies: 
jQuery 1.72+, jQuery UI<any>, KnockoutJS 3.2+
Looks best with Bootstrap CSS - http://getbootstrap.com or use the CDN (Content Delivery Network) links below.

You must be familiar with the Knockout JS MVVM framework syntax. Visit http://knockoutjs.com if you need an introduction or refresher.

#### Usage
```
// These are included in the sample Master page provided - Shockout.SpForms.master
<!-- Bootstrap CSS (in head)-->
<link href="//maxcdn.bootstrapcdn.com/bootstrap/3.3.5/css/bootstrap.min.css" rel="stylesheet" type="text/css" />

<!-- jQuery UI CSS (in head)-->
<link href="//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css" rel="stylesheet" type="text/css" />

<!-- It's recommended to place your scripts at the bottom of the page, before the ending </body> tag, for faster page loads. -->

<!-- jQuery -->
<script src="//code.jquery.com/jquery-1.11.3.min.js" type="text/javascript"></script>
<script src="//code.jquery.com/ui/1.11.4/jquery-ui.min.js" type="text/javascript"></script>

<!-- Bootstrap -->
<script src="//maxcdn.bootstrapcdn.com/bootstrap/3.3.5/js/bootstrap.min.js" type="text/javascript"></script>

<!-- Knockout JS -->
<script src="//cdnjs.cloudflare.com/ajax/libs/knockout/3.3.0/knockout-min.js" type="text/javascript"></script>

<!-- Shockout SPForms -->
<script src="ShockoutForms-0.0.1.min.js" type="text/javascript"></script>

<!-- Setup your form - this goes at the bottom of your form's page -->
<script type="text/javascript">
	var spForm = new Shockout.SPForm(
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
</script>
```

###Displaying SharePoint Choice Fields - Select Menu
How to display the choices from a SharePoint Choice Field in a select menu.
```
<div class="form-group">
	<label data-bind="text: MySpChoiceFieldName._displayName" class="control-label"></label>
	<select data-bind="value: MySpChoiceFieldName, options: MySpChoiceFieldName._choices, optionsValue: 'value', optionsCaption: '--SELECT--'" class="form-control"></select>
</div>
```

###Displaying SharePoint MultiChoice Fields - Checkboxes
How to display the choices from a SharePoint MultiChoice Field with checkboxes.
```
<div class="form-group">
    <label data-bind="text: MySpChoiceFieldName._displayName" class="control-label"></label>

	<!-- optional Field Description -->
	<p data-bind="text: MySpChoiceFieldName._description"></p>

    <!-- ko foreach: MySpChoiceFieldName._choices -->
    <label class="checkbox">
        <input type="checkbox" data-bind="checked: $root.MySpChoiceFieldName, attr: { value: $data.value, name: 'MySpChoiceFieldName_' + $index() }" />
        <span data-bind="text: $data.value"></span>
    </label>
    <!-- /ko --> 	            
</div>
```

###Displaying SharePoint MultiChoice Fields - Radio Buttons
How to display the choices from a SharePoint MultiChoice Field with radio buttons.
```
<div class="form-group">
    <label data-bind="text: MySpChoiceFieldName._displayName" class="control-label"></label>

	<!-- optional Field Description -->
	<p data-bind="text: MySpChoiceFieldName._description"></p>

    <!-- ko foreach: MySpChoiceFieldName._choices -->
    <label class="radio">
        <input type="radio" data-bind="checked: $root.MySpChoiceFieldName, attr: { value: $data.value }" name="MySpChoiceFieldName" />
        <span data-bind="text: $data.value"></span>
    </label>
    <!-- /ko -->             
</div>
```

### Custom Knockout binding handlers for SP list field types included:
	
####spHtmlEditor
`<textarea data-bind="value: Comments, spHtml: true"></textarea>` 

####spPerson
`<input type="text" data-bind="spPerson: myVar" />`
	OR
`<div data-bind="spPerson: myVar"></div>`

####spDate
`<input type="text" data-bind="spDate: myVar" />`	
	OR
`<div data-bind="spDate: myVar"></div>`

####spDateTime
`<input type="text" data-bind="spDateTime: myVar" />`
	OR
`<div data-bind="spDateTime: myVar"></div>`

####spMoney
`<input type="text" data-bind="spMoney: myVar" />`
	OR
`<div data-bind="spMoney: myVar"></div>`

####spDecimal
`<input type="text" data-bind="spDecimal: myVar" />`
	OR
`<div data-bind="spDecimal: myVar"></div>`

####spNumber
`<input type="text" data-bind="spNumber: myVar" />`
	OR
`<div data-bind="spNumber: myVar"></div>`

####Element Attributes

```
// Restricts element to authors only. Removes from DOM otherwise.
// Useful for restricting edit fields to the person that created the form.
<section data-author-only></section>
```

```
// Restricts element to non-authors of a form. Removes from DOM otherwise. 
// Useful for displaying read-only/non-edit sections to non-authors only.
<section data-non-authors></section>
```

```
// Restricts elements to forms with an ID in the querystring. Removes from DOM otherwise. 
// Useful for sections that require another person's input (approval sections) on an existing form.
<section data-edit-only></section>
```

```
// Restricts elements to forms with NO ID in the querystring. Removes from DOM otherwise. 
<section data-new-only></section>
```

```
// Control permissions to elements by SP group membership, such as manager approval sections/fields.
// Value is a comma delimitted list of user groups `<groupId>;#<groupName>`.
// Example:
<section data-sp-groups="1;#Administrators,2;#Managers"></section>
```

```
// For approval sections, you can combine these attributes:
<section data-edit-only data-sp-groups="1;#Administrators,2;#Managers"></section>
// This element will be shown to users who beleong to the SP user groups specified and only when there is an ID in the querystring of the form URL. 
```

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

