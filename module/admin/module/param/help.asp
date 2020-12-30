<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../header_popup.asp" -->

<!-- SECTION_START:Overview -->
<h3>Module Parameters  Help</h3>

<p>
Module parameters are configuration settings that may be applied
to a module.  The administration area allows developers to define
new parameters that they need for their modules.  The parameters
are designed to be fully configurable and support many different
data types.
</p>

<p>
Once parameters have been defined for a module, a special icon will
appear in the module admin area that links to the module configuration.
The icon is a black circle with a screwdriver and a wrench that
appears just to the right of the tab navigation.  Clicking on this
icon will bring up the module configuration page.
</p>

<p>
Within your module pages, you refernece the module parameters using
the <kbd>steParam</kbd> library function.  Refer to the the site library
documentation for more information.
</p>

<!-- SECTION_END:Overview -->

<!-- SECTION_START:Param -->

<h3>Module Parameters Help</h3>

<p>
Define all of the parameters needed by your module in the Parameters
administration.  Unless you are a module developer, you shouldn't need
to modify the parameters for a module.  In fact, the whole parameters
configuration area shouldn't need to be touched at all.
</p>

<p>
After you have defined some parameters for your module, you may change
the order in which they appear on the module configuration page.  Do
this by clicking on the <i>up</i> and <i>down</i> action links which
appear next to the parameters in the list.
</p>

<h4>Module Parameter Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Parameter Name</td>
	<td class="formd">Variable name used for the parameter (must be unique to this module.)</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Parameter Value</td>
	<td class="formd">Default value for this parameter (parameter will have this value when the module is first installed)</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Data Type</td>
	<td class="formd">Type of data stored in this parameter (the sytem will enforce the data type chosen.)</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Label</td>
	<td class="formd">Label shown in the parameter configuration page next to the form input.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Minimum Value</td>
	<td class="formd">The minimum value for numeric data types such as <kbd>integer</kbd> and <kbd>float</kbd>.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Maximum Value</td>
	<td class="formd">The maximum value for numeric data types such as <kbd>integer</kbd> and <kbd>float</kbd>.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Is Required?</td>
	<td class="formd">Is this parameter required to have a value? (value cannot be empty)</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Help Text</td>
	<td class="formd">Any instructions regarding the value that should be entered for this parameter.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Param -->

<!-- SECTION_START:Options -->

<h3>Module Options Help</h3>

<p>
Module options are used to define a list of options for a parameter
that has a data type of "drop list".  The module configuration page
will display a drop-down select list populated with the associated
options you define in this area.
</p>

<p>
Options consist of a label and a value.  The label is shown to the
user in the drop-list.  After they choose an option, the parameter
will take on the associated value (not the label).  It is this
value that will be returned by the <kbd>steParam</kbd> library
function.
</p>

<p>
The <i>up</i> and <i>down</i> action links for the options allows
you to change the sort order of the options within the drop-list.
The <i>Active</i> checkbox allows you to mark options as active
(when checked show in the drop-list) or inactive (when unchecked
don't show in the drop-list.)
</p>

<p>
Check the <i>Valid</i> checkbox to indicate that the option is a
valid selection in the list.  You would create invalid selections
when you want to make headings in the drop-list that serve to mark
groups of options but can't be selected as a valid choice in the list.
</p>

<p>
You may also define drop-list options for a parameter type instead
of assigning them to an individual parameter.  If you do this, any
module configuration that uses your special data type will render
the input as a drop-list.
</p>

<h4>Module Options Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Option Value</td>
	<td class="formd">This is the value that will be assigned to the parameter when this option is selected.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Option Label</td>
	<td class="formd">The label that is displayed in the drop-list on the module configuration page.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Options -->

<!-- SECTION_START:Types -->

<h3>Module Types Help</h3>

<p>
The module types administration allows you to configure the module
data types that are valid for the module parameters. Unless you are
a module developer, you really shouldn't touch this administration
area.
</p>

<p>
ASP Nuke comes pre-populated with the basic data types used in most
modules such as: "Integer", "String", "Float", "Date" and "Time".  You
may add additional types to the list as needed.  Be aware that
everything besided drop-lists will need to have supporting code
written in the backend to do the data validation.
</p>

<p>
Your most common addition to this list will probably be custom
drop-lists.  Even then, if your module only uses a drop-list for one
paramter, then you should just associate the options with the
parameter and not with a new type.
</p>

<p>
It is best to keep the list of types fairly short to avoid confusing
when defining module parameters.
</p>

<h4>Module Types Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Type Code</td>
	<td class="formd">Short character code to identify the parameter data type.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Type Name</td>
	<td class="formd">Human readable data type name that is used in the drop-list to choose the data type.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>ASP Convert Function</td>
	<td class="formd">ASP data conversion function for the type (mainly used for the built-in conversion functions such as <kbd>CInt</kbd> and <kbd>CDbl</kbd>)</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>HTML Input Type</td>
	<td class="formd">If your parameter type is not a complex data type (like dates) then you can indicate what type of HTML form input is used to input the parameter value.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>RegEx Validation</td>
	<td class="formd">A regular expression which is used to validate that parameter values entered.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Label Position</td>
	<td class="formd">Allows you to place the label on <kbd>TOP</kbd> of the input control instead of to the left when more space is needed for the input. (mainly used for large TEXTAREA inputs.)</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Has Options?</td>
	<td class="formd">Does this input type have options?  Unless you choose <kbd>yes</kbd> here, the parameters assigned this type will not be rendered as a drop-list.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Minimum Value</td>
	<td class="formd">For numeric data types, defines the minimum acceptable value this parameter may have.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Maximum Value</td>
	<td class="formd">For numeric data types, defines the maximum acceptable value this parameter may have.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Is Numeric?</td>
	<td class="formd">Does this parameter type use a numeric value?</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>DB Quote Character</td>
	<td class="formd">Defines any special quoting caracters that are needed when inserting or updating the parameter value in the database.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Types -->

<!-- #include file="../../../../footer_popup.asp" -->