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
			allowedExtensions: ['txt', 'rtf', 'zip', 'pdf', 'doc', 'docx', 'jpg', 'gif', 'png', 'ppt', 'tif', 'pptx', 'csv', 'pub', 'msg']  // the default 
			attachmentMessage: 'An attachment is required.', // the default
			confirmationUrl: '/SitePages/Confirmation.aspx', // the default
			enableErrorLog: true, // default true
			errorLogListName: 'Error Log', // Designated SharePoint list for logging user and form errrors; default 'Error Log' on root site
			fileHandlerUrl: string = '/_layouts/SPFormFileHandler.ashx',  // the default    
			enableAttachments: true, // default true
			includeUserProfiles: true, // default true
			includeWorkflowHistory: true, // default true        
			requireAttachments: false, // default false
			siteUrl: '/companyForms', // default is ''        
			workflowHistoryListName: 'Workflow History' // the default
		});
</script>
```

###Attachments
To enable your form to allow attaching files, ensure the `enableAttachments` option is `true` (the default) and include an element with the class name "attachments". Shockout will place everything inside the element(s). 
Also ensure your SharePoint list has attachments enabled. Shockout will detect this setting and render attachments based on your SP list settings.
```
<section class="attachments"></section>
```

###Show the User Profiles for Created By and Modified By
To enable this feature, ensure that `includeUserProfiles` is `true` (the default) include an element with the class name "created-info". 
Shockout will query the User Information List or User Profile Service, if you have it, and display user profiles with: picture, full name, job title, email, phone, department, and office.
If this feature is disabled, Shockout will only show the Created By/Created and Modified By/Modified fields. 
```
<section class="created-info" data-edit-only></section>
```

###SharePoint Field Variable Names
Shockout relies on SharePoint REST Services and SP REST Services returns your list's field names in a specific format; basically the current display name, minus spaces and special characters, in "CamelCase."

The Shockout framework will map these camel case variable names to an instance of a Knockout view model. You'll use these variable names to create your form's fields.

To preview the formatting of your SharePoint list's field names, it's very helpful to use a REST client such as the Advanced REST Client for Google's Chrome browser - https://chrome.google.com/webstore/detail/advanced-rest-client/hgmloofddffdnphfgcellkdfbfbjeloo

Once this application is installed, you can preview the JSON returned by entering the following in the address bar. Your list name must NOT contain spaces and must be in CamelCase
```
http://<SiteUrl>/<Subsite>/_vti_bin/listdata.svc/<MyListName>(<ID>)

// example: https://mysite.com/forms/_vti_bin/listdata.svc/PurchaseRequisitions(1)
```
Choose the GET radio option and enter `Accept: application/json;odata=verbose` in the RAW field. This tells SP to return JSON, not XML!
Now that you know the variable names, you're ready to create your Shockout form.

###Displaying a SharePoint Text Field
```
<div class="form-group">
	<label data-bind="text: MySpFieldName._displayName" for="MySpFieldName" class="control-label"></label>
	
	<input type="text" data-bind="value: MySpFieldName, attr:{'placeholder': MySpFieldName._displayName}" maxlength="255" id="MySpFieldName" class="form-control" />
	
	<!-- optional Field Description -->
	<p data-bind="text: MySpFieldName._description"></p>
</div>
```

###Displaying a SharePoint Checkbox Field (Boolean)
```
<div class="form-group">
	<label class="checkbox">
        <input type="checkbox" data-bind="checked: MySpFieldName" />
        <span data-bind="text: MySpFieldName._displayName"></span>
    </label>

	<!-- optional Field Description -->
	<p data-bind="text: MySpFieldName._description"></p>
</div>
```

###Displaying SharePoint Choice Fields - Select Menu
How to display the choices from a SharePoint Choice Field in a select menu.
```
<div class="form-group">
	<label data-bind="text: MySpChoiceFieldName._displayName" class="control-label" for="MySpChoiceFieldName"></label>
	
	<select data-bind="value: MySpChoiceFieldName, options: MySpChoiceFieldName._choices, optionsValue: 'value', optionsCaption: '--SELECT--'" id="MySpChoiceFieldName" class="form-control"></select>

	<!-- optional Field Description -->
	<p data-bind="text: MySpChoiceFieldName._description"></p>
</div>
```

###Displaying SharePoint MultiChoice Fields - Checkboxes
How to display the choices from a SharePoint MultiChoice Field with checkboxes.
```
<div class="form-group">
    <label data-bind="text: MySpChoiceFieldName._displayName" class="control-label"></label>

    <!-- ko foreach: MySpChoiceFieldName._choices -->
    <label class="checkbox">
        <input type="checkbox" data-bind="checked: $root.MySpChoiceFieldName, attr: { value: $data.value, name: 'MySpChoiceFieldName_' + $index() }" />
        <span data-bind="text: $data.value"></span>
    </label>
    <!-- /ko --> 
	
	<!-- optional Field Description -->
	<p data-bind="text: MySpChoiceFieldName._description"></p>	            
</div>
```

###Displaying SharePoint MultiChoice Fields - Radio Buttons
How to display the choices from a SharePoint MultiChoice Field with radio buttons.
```
<div class="form-group">
    <label data-bind="text: MySpChoiceFieldName._displayName" class="control-label"></label>

    <!-- ko foreach: MySpChoiceFieldName._choices -->
    <label class="radio">
        <input type="radio" data-bind="checked: $root.MySpChoiceFieldName, attr: { value: $data.value }" name="MySpChoiceFieldName" />
        <span data-bind="text: $data.value"></span>
    </label>
    <!-- /ko -->   
	
	<!-- optional Field Description -->
	<p data-bind="text: MySpChoiceFieldName._description"></p>          
</div>
```

##Required Field Validation
Simply add the `required="required"` attribute to required fields. Shockout will do the rest!

###Knockout SharePoint Field Binding Handlers
	
####spHtml
`<textarea data-bind="value: Comments, spHtml: true"></textarea>` 

####spPerson
Displays an auto-complete field with user name and email address as you type a user's first or last name in the field. The minimum number of characters is 3 before it will query the User Information List.
SharePoint group names aren't supported at this time, but coming soon.
`<input type="text" data-bind="spPerson: myVar" />`
	OR
`<div data-bind="spPerson: myVar"></div>`

####spDate
Displays a date in MM/DD/YYYY format.
`<input type="text" data-bind="spDate: myVar" />`	
	OR
`<div data-bind="spDate: myVar"></div>`

####spDateTime
Displays a date/time in MM/DD/YYYY HH:MMM TT format.
`<input type="text" data-bind="spDateTime: myVar" />`
	OR
`<div data-bind="spDateTime: myVar"></div>`

####spMoney
Displays currency in USD. Negative values are displayed in red and enclosed in parenthesis `()`. Other currency symbols coming soon.
`<input type="text" data-bind="spMoney: myVar" />`
	OR
`<div data-bind="spMoney: myVar"></div>`

####spDecimal
Displays number with 2 decimal places. Negative values are displayed in red. You can change the decmial places with the precision attribute.
`<input type="text" data-bind="spDecimal: myVar, precision: 2" />`
	OR
`<div data-bind="spDecimal: myVar, precision: 2"></div>`

####spNumber
Displays integer/whole number. Negative values are displayed in red.
`<input type="text" data-bind="spNumber: myVar" />`
	OR
`<div data-bind="spNumber: myVar"></div>`

###Element Attributes

####data-author-only
Restricts element to authors only. Removes from DOM otherwise.
Useful for restricting edit fields to the person that created the form.
```
<section data-author-only></section>
```

####data-non-authors
Restricts element to non-authors of a form. Removes from DOM otherwise. 
Useful for displaying read-only/non-edit sections to non-authors only.
```
<section data-non-authors></section>
```

####data-edit-only
Restricts elements to forms with an ID in the querystring. Removes from DOM otherwise. 
Useful for sections that require another person's input (approval sections) on an existing form.
```
<section data-edit-only></section>
```

####data-new-only
Restricts elements to forms with NO ID in the querystring. Removes from DOM otherwise. 
```
<section data-new-only></section>
```

####data-sp-groups
Control permissions to elements by SP group membership, such as manager approval sections/fields.
// Value is a comma delimitted list of user groups `<groupId>;#<groupName>`.
```
// Example:
<section data-sp-groups="1;#Administrators,2;#Managers"></section>
```

For approval sections, you can combine these attributes:
```
<section data-edit-only data-sp-groups="1;#Administrators,2;#Managers"></section>
```
This element will be shown to users who beleong to the SP user groups specified and only when there is an ID in the querystring of the form URL. 

###Form Events
You may further customize your form by adding extra functionality within the appropriate event methods. 
You specify the code for these methods in the third parameter of the constructor - the options object.

####preRender()
```
preRender: function(spForm){
	// Run custom code here BEFORE the form is rendered and BEFORE the Knockout view model is bound.
	// Useful for adding custom markup and/or custom local Knockout variables to your form.
	// Shockout will know the difference between your variables and the ones that exist in your SharePoint list.
}
```

####postRender()
```
postRender: function(spForm){
	// Run custom code here AFTER the form is rendered and AFTER the Knockout view model is bound.
	// Useful for:
	//	- setting default values for your Knockout objects
	//	- using JSON.parse() to convert string data stored in a text field to JSON - think tables in InfoPath but with JSON instead of XML!
}
```

####preSave()
```
preSave: function(spForm){
	// Run code before the form is saved.
	// Useful for: 
	//	- implementing custom validation
	//	- converting JSON data to a string with JSON.stringify(), which is saved in a plain text field.
}	
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

