<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../header_popup.asp" -->

<!-- SECTION_START:Overview -->
<h3>Application Help</h3>

<p>
The application administration allows you to configure the ASP Nuke
application itself.  The configuration properties are the same ones
you used if you ran the <kbd>setup</kbd> wizard when you first installed
ASP Nuke.
</p>

<p>
You should never change the application variable properties for the default
set of application variables that comes pre-defined with the ASP Nuke install.
The one exception to this is that you can change a varible's value.  Bear in
mind that some changes (such as path information), has the ability to render
your whole site inoperable.
</p>

<p>
If you do make a mistake and your ASP Nuke installation is broken, you should
be able to go back and run the <kbd>setup</kbd> script to repair the variables
that broke the application.  The best advice for non-technical types is not to
touch the application configuration unless absolutely necessary.
</p>

<!-- SECTION_END:Overview -->

<!-- SECTION_START:Configure -->

<h3>Application Variable Help</h3>

<p>
The Application Variables configure the entire ASP Nuke application and
are mainly used to setup the database connection, web site location and
the various folders used to hold data.
</p>

<p>
Unless you are an experienced developer, it is best to never change any
of the variable setting with the exception of the values.  Even then,
you need to be careful that you don't configure your ASP Nuke application
incorrectly.
</p>

<h4> Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Tab Group</td>
	<td class="formd">Indicate the major tab (for the setup wizard) where the application variable will appear.  We define a different tab for each functional group in the application.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Name</td>
	<td class="formd">Variable name for the application variable.  References to this name are hard-coded in the ASP Nuke code, so changing the name is strongly discouraged.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Value</td>
	<td class="formd">This is the value for the application variable.  You can make changes to these values to change the setup values for ASP Nuke.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Data Type</td>
	<td class="formd">Indicate the type of data which should be entered for the variable.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Is Required?</td>
	<td class="formd">Does this variable require a value? (cannot be empty)</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Help Text</td>
	<td class="formd">Include instructions on the use of this variable and what type of value should be entered.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Configure -->

<!-- SECTION_START:Tabs -->

<h3>Application Tabs Help</h3>

<p>
The application tabs are used to group application variables (settings)
into functional groups.  Each tab is represented in the setup wizard
as a different step in configuring the ASP Nuke application.
</p>

<p>
</p>

<h4>Application Tabs Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Name</td>
	<td class="formd">A short name that is used to label this application setting tab.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Title</td>
	<td class="formd">A long title for the application group that is shown as the main heading for the setup page.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Introduction</td>
	<td class="formd">Provide short introductory text that will appear before the application settings on the setup page.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Summary</td>
	<td class="formd">Provide short summary content to be displayed at the end of the application settings page.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Tabs -->

<!-- SECTION_START:Types -->

<h3>Application Types Help</h3>

<p>
The Application Types administration is used to configure the basic
data types which an application variable may be assigned.  It includes
basic data types such as "Integer", "Float", "Date" and "String" along
with a special type named "Drop List".
</p>

<p>
You can define as many data types as you need for your ASP Nuke
configuration settings.  We strongly recommend that you stick with
the default set that is included with the ASP Nuke install package.
This will help you in the future should you upgrade your application
to a newer version.
</p>

<p>
Each of the different data types has it's own validation rule to ensure
that the correct data is being entered for the application setting.
</p>

<h4>Application Types Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Type Code</td>
	<td class="formd">Short character code to identify the application variable data type.</td>
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
	<td class="formd">If your application variable type is not a complex data type (like dates) then you can indicate what type of HTML form input is used to input the application variable.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>RegEx Validation</td>
	<td class="formd">A regular expression which is used to validate that variable values entered.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Label Position</td>
	<td class="formd">Allows you to place the label on <kbd>TOP</kbd> of the input control instead of to the left when more space is needed for the input. (mainly used for large TEXTAREA inputs.)</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Has Options?</td>
	<td class="formd">Does this input type have options?  Unless you choose <kbd>yes</kbd> here, the application variables assigned this type will not be rendered as a drop-list.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Minimum Value</td>
	<td class="formd">For numeric data types, defines the minimum acceptable value this application variable may have.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Maximum Value</td>
	<td class="formd">For numeric data types, defines the maximum acceptable value this application variable may have.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Is Numeric?</td>
	<td class="formd">Does this application variable type use a numeric value?</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>DB Quote Character</td>
	<td class="formd">Defines any special quoting caracters that are needed when inserting or updating the variable value in the database.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Types -->

<!-- #include file="../../../../footer_popup.asp" -->