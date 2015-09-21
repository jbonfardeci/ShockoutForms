<%@ Page Language="C#" masterpagefile="../_catalogs/masterpage/Shockout.SpForms.master" title="Purchase Requisition" %>
<%@ Register tagprefix="SharePoint" namespace="Microsoft.SharePoint.WebControls" assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<asp:Content runat="server" ID="main1" ContentPlaceHolderID="main">
	
	<div class="panel panel-primary">
	
		<div class="panel-heading">
			<div class="panel-title">Purchase Requisition</div>
		</div>
		
		<div class="panel-body" id="ShockoutSpForm">
		
			<input type="hidden" data-bind="value: IsSubmitted" />
		
			<section class="created-info" data-edit-only></section>
			
			<section class="attachments"></section>
		
			<section>
				<h4>Vendor Information</h4>
				
				<div class="form-group" data-bind="with: Title._metadata">
				    <label data-bind="text: displayName, attr:{'for': name}" class="control-label"></label>
					<input type="text" data-bind="value: $parent, attr:{'placeholder': displayName, 'id': name}" maxlength="255" class="form-control" />
					<!-- optional Field Description -->
					<p data-bind="text: description"></p>
				</div>
				
				<div class="form-group" data-bind="with: VendorName._metadata">
				    <label data-bind="text: displayName, attr:{'for': name}" class="control-label"></label>
					<input type="text" data-bind="value: $parent, attr:{'placeholder': displayName, 'id': name}" maxlength="255" class="form-control" />
					<!-- optional Field Description -->
					<p data-bind="text: description"></p>
				</div>
				
				<div class="form-group">
				    <label data-bind="text: VendorAddress._displayName, attr:{'for': VendorAddress._name}" class="control-label"></label>
					<textarea data-bind="value: VendorAddress, attr:{'id': VendorAddress._name}" class="rte"></textarea>
					<!-- optional Field Description -->
					<p data-bind="text: VendorAddress._description"></p>
				</div>
				
			</section>

            <section>
                <h4>Routing</h4>
                <div class="form-group">
				    <label data-bind="text: Supervisor._displayName, attr: { 'for': Supervisor._name }" class="control-label"></label>
					<input type="text" data-bind="spPerson: Supervisor, attr: {'id': Supervisor._name }" maxlength="255" class="form-control" />s
					<!-- optional Field Description -->
					<p data-bind="text: Supervisor._description"></p>
				</div>
            </section>
			
			<section data-edit-only>
				<h4>Supervisor Approval Section</h4>
				<div class="form-group" data-bind="with: SupervisorApproval._metadata">
				    <label data-bind="text: displayName" class="control-label"></label>
				
				    <!-- ko foreach: choices -->
				    <label class="radio">
				        <input type="radio" data-bind="checked: $parent.$parent, attr: { value: $data.value, 'name': $parent.name }" />
				        <span data-bind="text: $data.value"></span>
				    </label>
				    <!-- /ko -->   
					
					<!-- optional Field Description -->
					<p data-bind="text: description"></p>          
				</div>
			</section>
			
			<section>
				<h4>Approvals</h4>
				<div>
					<label class="control-label">Supervisor</label>
					<span data-bind="text: SupervisorApproval"></span>
				</div>
			</section>
		
		</div>
	
	</div>

</asp:Content>

<asp:Content runat="server" ID="head1" ContentPlaceHolderID="head">

</asp:Content>

<asp:Content runat="server" ID="scripts1" ContentPlaceHolderID="scripts">

<script type="text/javascript">
(function(){
var spForm = new Shockout.SPForm(
    /*listName:*/ 'Purchase Requisitions', 
    /*formId:*/ 'ShockoutSpForm', 
    /*options:*/ {
        debug: false, // default false
        siteUrl: '', // default  
        confirmationUrl: '/SitePages/Confirmation.aspx', // the default
        preRender: function (spForm) {

        }, // default undefined
        postRender: function (spForm) {

        }, // default undefined
        preSave: function (spForm) {

        }, // default undefined   
        allowDelete: true, // default false
        allowPrint: true, // default true
        allowSave: true, // default true
        allowedExtensions: ['txt', 'rtf', 'zip', 'pdf', 'doc', 'docx', 'jpg', 'gif', 'png', 'ppt', 'tif', 'pptx', 'csv', 'pub', 'msg'],  // the default
        enableAttachments: true, // default true
        requireAttachments: false, // default false
        attachmentMessage: 'An attachment is required.', // the default 
        fileHandlerUrl: '/_layouts/SPFormFileHandler.ashx', // the default 
        enableErrorLog: false, // default true
        errorLogListName: 'Error Log', // Designated SharePoint list for logging user and form errors; Requires a custom SP list named 'Error Log' on root site with fields: 'Title' and 'Error'  
        includeUserProfiles: true, // default true   
        includeWorkflowHistory: false, // default true 
        workflowHistoryListName: 'Workflow History' // the default
    });
})();
</script>

</asp:Content>