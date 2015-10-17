<%@ Page Language="C#" masterpagefile="/_catalogs/masterpage/Shockout.SpForms.master" title="Purchase Requisition" inherits="Microsoft.SharePoint.WebPartPages.WebPartPage, Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" meta:progid="SharePoint.WebPartPage.Document" %>
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<asp:Content runat="server" ID="main1" ContentPlaceHolderID="main">
	
	<div class="panel panel-primary">
	
		<div class="panel-heading">
			<div class="panel-title">Purchase Requisition</div>
		</div>
		
		<div class="panel-body" id="ShockoutSpForm">
		
			<input type="hidden" data-bind="value: IsSubmitted" />
			<input type="hidden" data-bind="value: ItemData" />
			<input type="hidden" data-bind="value: TotalCost" />
		
			<section class="created-info" data-edit-only=""></section>
			
			<section class="attachments nav-section"></section>
		
			<section class="nav-section">
				<h4>Vendor Information</h4>
				
				<div class="form-group">
				    <label data-bind="text: Title._displayName" for="Title" class="control-label"></label>
					<input type="text" data-bind="value: Title, attr:{'placeholder': Title._displayName}" maxlength="255" id="Title" class="form-control" />
					<!-- optional Field Description -->
					<p data-bind="text: Title._description"></p>
				</div>
				
				<div class="form-group">
				    <label data-bind="text: VendorName._displayName" for="VendorName" class="control-label"></label>
					<input type="text" data-bind="value: VendorName, attr: { 'placeholder': VendorName._displayName }" maxlength="255" id="VendorName" class="form-control" />
					<!-- optional Field Description -->
					<p data-bind="text: VendorName._description"></p>
				</div>
				
				<div class="form-group">
				    <label data-bind="text: VendorAddress._displayName" class="control-label"></label>
					<textarea data-bind="value: VendorAddress" data-sp-html=""></textarea>
					<!-- optional Field Description -->
					<p data-bind="text: VendorAddress._description"></p>
				</div>
				
			</section>

            <!--Item List-->
            <section class="nav-section">
			    <h4>Item List</h4>

                <div>

                    <div class="right" data-author-only="">
                        <button data-bind="click: toggle_edit" class="btn btn-primary"><span data-bind="visible: !edit()">Edit</span><span data-bind="visible: edit()">Stop Editing</span></button>&nbsp;
					    <button data-bind="click: add_item" class="btn btn-success"><span>+ Add Item</span></button>
                    </div>

                    <div class="row">
                        <div class="col-md-3 col-xs-3"><strong>Specific Item(s) Requested</strong></div>
					    <div class="col-md-2 col-xs-2"><strong>Product/Code</strong></div>
					    <div class="col-md-2 col-xs-2"><strong>Quantity</strong></div>
					    <div class="col-md-2 col-xs-2"><strong>Per Unit Cost</strong></div>
					    <div class="col-md-2 col-xs-2"><strong>Item Total</strong></div>
					    <div class="col-md-1 col-xs-1">&nbsp;</div>
                    </div>

                    <!-- ko foreach: items -->
                    <div class="row item" data-bind="css: {'even': $index() % 2 != 0}">
                        <div class="col-md-3 col-xs-3">
                            <input type="text" data-bind="value: item, visible: $root.edit()" class="form-control" />
                            <span data-bind="text: item, visible: !$root.edit()"></span>
                        </div>
					    <div class="col-md-2 col-xs-2">
                            <input type="text" data-bind="value: code, visible: $root.edit()" class="form-control short" />
                            <span data-bind="text: code, visible: !$root.edit()"></span>
					    </div>
					    <div class="col-md-2 col-xs-2 right">
                            <input type="text" data-bind="spNumber: quantity, visible: $root.edit()" class="form-control short" />
                            <span data-bind="spNumber: quantity, visible: !$root.edit()"></span>
					    </div>
					    <div class="col-md-2 col-xs-2 right">
                            <input type="text" data-bind="spMoney: cost, visible: $root.edit()" class="form-control short" />
                            <span data-bind="spMoney: cost, visible: !$root.edit()"></span>
					    </div>
					    <div class="col-md-2 col-xs-2 right" data-bind="spMoney: (quantity() * cost())"></div>
					    <div class="col-md-1 col-xs-1 right">
                            <button class="btn btn-sm btn-danger" data-bind="click: $root.del_item, visible: $root.edit()" data-author-only=""><span class="glyphicon glyphicon-trash"></span></button>
					    </div>
                    </div>
                    <!-- /ko -->

				    <div class="row">
					    <div class="col-md-9 col-xs-9 right">Subtotal</div>
					    <div class="col-md-2 col-xs-2 right"><span data-bind="spMoney: subtotal"></span></div>
					    <div class="col-md-1 col-xs-1">&nbsp;</div>
				    </div>
				    <div class="row">
					    <div class="col-md-9 col-xs-9 right">Shipping Charge</div>
					    <div class="col-md-2 col-xs-2 right">
                            <input type="text" data-bind="spMoney: shipping, visible: edit()" class="form-control short" />
                            <span data-bind="spMoney: shipping, visible: !edit()"></span>
					    </div>
					    <div class="col-md-1 col-xs-1">&nbsp;</div>
				    </div>
				    <div class="row">
					    <div class="col-md-9 col-xs-9 right">Tax &ndash; Percentage: <strong>8.125%</strong> OR Flat Amount: <strong>#.##</strong></div>
					    <div class="col-md-2 col-xs-2 right">
                            <input type="text" data-bind="value: tax, visible: edit()" class="form-control short" />
                            <span data-bind="text: tax, visible: !edit()"></span>
					    </div>
					    <div class="col-md-1 col-xs-1">&nbsp;</div>
				    </div>
				    <div class="row">
					    <div class="col-md-9 col-xs-9 right">Total Tax</div>
					    <div class="col-md-2 col-xs-2 right"><span data-bind="spMoney: total_tax"></span></div>
					    <div class="col-md-1 col-xs-1">&nbsp;</div>
				    </div>
				    <div class="row">
					    <div class="col-md-9 col-xs-9 right"><strong>Total</strong></div>
					    <div class="col-md-2 col-xs-2 right"><strong data-bind="spMoney: total"></strong></div>
					    <div class="col-md-1 col-xs-1">&nbsp;</div>
				    </div>
					
				    <div class="right" data-author-only="">
					    <button data-bind="click: toggle_edit" class="btn btn-primary"><span data-bind="visible: !edit()">Edit</span><span data-bind="visible: edit()">Stop Editing</span></button>&nbsp;
					    <button data-bind="click: add_item" class="btn btn-success"><span>+ Add Item</span></button>				
				    </div>		

                </div>
		
		    </section>

            <section class="nav-section">
                <h4>Routing</h4>
                <div class="form-group">
				    <label for="Supervisor" class="control-label">Your Supervisor</label>
					<input type="text" data-bind="spPerson: YourSupervisor" maxlength="255" id="YourSupervisor" class="form-control" />
					<!-- optional Field Description -->
					<p data-bind="text: YourSupervisor._description"></p>
				</div>
            </section>
			
            <input type="hidden" data-bind="value: SupervisorApproval" />
			<section data-edit-only="" data-bind="visible: isSupervisor()">
				<h4>Supervisor Approval Section</h4>
				<div class="form-group">
				    <label data-bind="text: SupervisorApproval._displayName" class="control-label"></label>
				
				    <!-- ko foreach: SupervisorApproval._choices -->
				    <label class="radio">
				        <input type="radio" data-bind="checked: $root.SupervisorApproval, attr: { value: $data.value }" name="SupervisorApproval" />
				        <span data-bind="text: $data.value"></span>
				    </label>
				    <!-- /ko -->   
					
					<!-- optional Field Description -->
					<p data-bind="text: SupervisorApproval._description"></p>          
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
<script src="/Forms/ShockoutForms-0.0.2.min.js" type="text/javascript"></script>
<script type="text/javascript">
(function(){
var spForm = new Shockout.SPForm(
    /*listName:*/ 'Purchase Requisitions', 
    /*formId:*/ 'ShockoutSpForm', 
    /*options:*/ {
        debug: false, 
        siteUrl: '/media', 
        confirmationUrl: '/SitePages/Confirmation.aspx',
        preRender: function (spForm, vm) {

            try{
            
                // set up the KO variables and methods for updating the Item list table
                vm.edit = ko.observable(false);
                vm.shipping = ko.observable(0);
                vm.tax = ko.observable('0.000%');
                vm.items = ko.observableArray([]);

                vm.subtotal = ko.computed(function () {
                    var subtotal = 0;
                    for (var i = 0; i < vm.items().length; i++) {
                        var r = vm.items()[i];
                        subtotal += r.quantity() * r.cost();
                    }
                    return subtotal;
                }, vm);

                vm.total_tax = ko.computed(function () {
                    var tax = vm.tax().toString().replace(/[^\d\%\.]/g, '');
                    var subtotal = vm.subtotal();

                    if (/\%/.test(tax)) {
                        tax = (tax.replace(/\%/g, '') - 0) / 100;
                        tax = tax * subtotal;
                    } else {
                        tax -= 0;
                    }

                    return tax;
                }, vm);

                vm.total = ko.computed(function () {
                    var total = vm.subtotal() + vm.total_tax() + vm.shipping();
                    return total;
                }, vm);

                vm.add_item = function (model, btn) {
                    model.items().push(new Item());
                    model.items.valueHasMutated();
                    if (!model.edit()) { model.edit(true); }
                    return false;
                };

                vm.toggle_edit = function (model, btn) {
                    model.edit(!model.edit());
                    return false;
                };

                vm.del_item = function (row, btn) {
                    vm.items.remove(row);
                    return false;
                };

                vm.isSupervisor = function () {
                    return vm.YourSupervisor() == vm.currentUser().account && vm.IsSubmitted();
                };
            }
            catch (e) {
                spForm.logError(e);
            }

        }, // default undefined
        postRender: function (spForm, vm) {

            try{
                //convert Line Item JSON data to KO Observable Array to display on form	    		
                if (vm.ItemData() != null) {
                    var json = JSON.parse(vm.ItemData());

                    $.each(json.items, function (i, o) {
                        vm.items().push(new Item(o.item, o.code, o.quantity, o.cost));
                    });

                    vm.items.valueHasMutated();
                    vm.tax(json.tax);
                    vm.shipping(json.shipping - 0);
                }
            }
            catch (e) {
                spForm.logError(e);
            }

        }, // default undefined
        preSave: function (spForm, vm) {

            /* save JSON string to SP list item field just before Save */
            try {
                var json = {
                    items: [],
                    tax: vm.tax(),
                    shipping: vm.shipping()
                };

                $.each(vm.items(), function (i, o) {
                    var row = {};
                    for (var p in o) {
                        row[p] = o[p]();
                    }
                    json.items.push(row);
                });

                vm.ItemData(JSON.stringify(json));
                vm.TotalCost(vm.total());
                
            }
            catch (e) {
                spForm.logError(e);
            }

        }
    });

    /* KO Line Item Model */
    function Item(item, code, quantity, cost) {
        this.item = ko.observable(item || null);
        this.code = ko.observable(code || null);
        this.quantity = ko.observable(quantity || 1);
        this.cost = ko.observable(cost || 0);
    };

})();
</script>

</asp:Content>