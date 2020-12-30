<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../header_popup.asp" -->

<!-- SECTION_START:Overview -->
<h3>Module Admin Help</h3>

<p>
The module administration is used mainly for the development of modules.
Specifically, it is for doing a manual setup for a new module that you are
creating.  One exception to this rule is placing the module capsules in
the three column layout.
</p>

<p>
If you go to the <b>group</b> tab, you can define the layout of the modules
on the page. DO NOT edit or delete the existing module groups ("LEFT", "RIGHT"
and "CENTER").  These must exist for the ASP Nuke layout to render properly.
</p>

<p>
Unless you are developing ASP Nuke modules, we strongly suggest that you only
use the <b>group</b> tab to change the layout of modules on the page.  All of
the standard modules that come in the ASP Nuke installation come pre-configured
to work out-of-the-box.
</p>

<!-- SECTION_END:Overview -->

<!-- SECTION_START:Module -->

<h3>Module Administration Help</h3>

<p>
The Module Administration is intended for module developers and
advanced site operators only.  You should not modify any of the
module configurations unless you know exactly what you are doing.
There is nothing really interesting you can do by
configuration of a module.
</p>

<p>
The module administration is primarily used for module developers
and allows you to define new modules in the database.  Once defined,
you can incorporate the modules into your site layout.
</p>

<p>
The process for developing a new ASP Nuke module generally goes as
follows:
</p>

<ul>
<li>Create a folder for the new module under the appropriate
	category folder (<kbd>/module/categoryname/foldername</kbd>)
<li>Create a configuration for the module using this module
	administration.
<li>Create the database tables necessary for your new module
<li>Develop the ASP pages for the capsule and/or the full-size
	module.
<li>Develop the module admistration pages under the directory:
	<kbd>/module/categoryname/foldername/admin</kbd>
<li>Configure the module group to display your module.
<li>Configure the <i>Access</I> permissions in the site administration
	to include your new module administration.
<li>Test and debug your code.
</ul>

<h4>Module Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Category</td>
	<td class="formd">Indicate the category where the module will be found (the module package must be installed under this category's folder)</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Title</td>
	<td class="formd">The title for this module as it appears in the ASP Nuke site administration pages.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Synopsis</td>
	<td class="formd">A short summary (couple of sentences) of the module and what it's used for.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Description</td>
	<td class="formd">Long description of the module and how it works.  Type one carriage return to create a line break and two consecutive carriage returns to create a paragraph break.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Folder Name</td>
	<td class="formd">Indicate the folder name where the module files can be found (underneath the category folder)</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Version No</td>
	<td class="formd">Indicate the latest version number for the module.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Small Capsule Module</td>
	<td class="formd">Script name for the file containing the ASP page that renders the small capsule for this module (suitable for placing in the narrow LEFT or RIGHT module groups.)</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Full Module</td>
	<td class="formd">Script name for the file containing the ASP page that renders the large content area (suitable for the CENTER module group)</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Update URL</td>
	<td class="formd">Internet URL to check for updates for the module (this feature is not implemented yet.)</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Do Update Check?</td>
	<td class="formd">Should we perform an automatic check for updates to this module? (this feature is not implemented yet.)</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Check Days</td>
	<td class="formd">Number of days to wait between automatic checks for updates to this module (not implemented yet.)</td>
</tr>
</table>
</p>

<!-- SECTION_END:Module -->
<!-- SECTION_START:Group -->

<h3>Module Group Help</h3>

<p>
The module groups are the ASP Nuke page layout components that control
how modules are arranged on the page.  For the most part, the module
groups are pre-defined and should never be modified.  The one exception
to this is for an experienced developer who wants to design a custom
layout for their Nuke site.
</p>

<p>
The standard module groups are "LEFT", "MID" and "RIGHT".  The "LEFT"
and "RIGHT" groups contain the narrow columns with the small module capsules
that appear on the left and right-hand side of the page.  The "MID"
group is the central content area which is only used on the home page
of the ASP Nuke site and holds the full-size modules.
</p>

<p>
The ordering of the columns (controlled with the <i>up</i> and <i>down</i>
action links) is not really used at this point.  If you click on the
<i>layout</i> action link, you will be able to configure what modules
will appear within each column.
</p>

<p>
More groups may be added in the future to incorporate a dynamic (drop-down)
menu that appears above the three columns and also a similar bar right under
the three main columns.
</p>

<h4>Module Group Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Group Code</td>
	<td class="formd">Identifier referenced in the ASP Nuke header and footer templates to build the layout.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Group Name</td>
	<td class="formd">Human readable name for the group to display in the administration area.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Has 140 (Pixel) Size Modules?</td>
	<td class="formd">Does this layout group support the small module capsules (narrow columns on the LEFT and RIGHT.)</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Has Full Size Modules?</td>
	<td class="formd">Does this layout support full-size module capsules (this should represent the main content area.)</td>
</tr>
</table>
</p>

<!-- SECTION_END:Group -->
<!-- SECTION_START:Category -->

<h3>Module Category Help</h3>

<p>
The module categories exist to organize the various modules available for the
ASP Nuke content management system.  They define a logical and physical grouping
of the module packages through a category name and a folder under the
<kbd>module</kbd> folder in the web site structure.
</p>

<p>
The module category list should remain static.  For consistancy sake and the
future addition of a centralized update service, we need to keep the list of
categories consistant.  If you are not concerned with compatibility with the
main ASP Nuke project, then feel free to create your own categories.
</p>

<h4>Module Category Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Name</td>
	<td class="formd">The title for this module category a visitor will see when browsing the module directory.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Folder Name</td>
	<td class="formd">This is the folder name directly under <kbd>/module</kbd> where the modules will be stored.  No path information of any kind should be included in this property (like: <kbd>articles</kbd>)</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Comments</td>
	<td class="formd">Include any additional comments about the module category that a visitor should see when they browse to the category.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Category -->

<!-- #include file="../../../../footer_popup.asp" -->