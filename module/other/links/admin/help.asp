<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../header_popup.asp" -->

<!-- SECTION_START:Overview -->
<h3>Link Help</h3>

<p>
The link module allows you define links to related web sites, services
or other interesting links.  You can define as many links as needed and
divide them up into groups for displaying in a capsule.
</p>

<p>
This could be used to create a menu for your site although you might want
to consider using the Dynamic Menu module.  Or you might use the links
module to define links to related sites or provide very simple advertising
for your affiliate web sites.
</p>

<!-- SECTION_END:Overview -->

<!-- SECTION_START:Link -->

<h3>Link Administration Help</h3>

<p>
The link administration will display a listing of all of the links which are
currently active.  This list includes all links regardless of what category
group they are placed in.  In fact, the Category name is actually a part of
the listing.
</p>

<p>
When you add a new link, you simply select the Category where the new
link should be created along with its properties.  If you don't have any
Categories configured, you should take care of that first by clicking on
the Categories tab and adding a new category.
</p>

<h4>Link Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Category</td>
	<td class="formd">Category group where the link should be posted.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>URL</td>
	<td class="formd">Defines the destination (target) for the link which is usually something like: <kbd>http://www.google.com/</kbd>)</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Label</td>
	<td class="formd">Defines the label for the link which is the hyperlinked text which appears within a category group.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Link -->


<!-- SECTION_START:Category -->

<h3>Link Category Help</h3>

<p>
The link categories serve to group related links together in the link module.  
Individual groups are separted and the category group heading is displayed above
each sub-list.  You can define what the groups are and the order in which they
are displayed using the link category administration.
</p>

<p>
You can change the order of category groups by using the <b>up</b> and <b>down</b>
links in the category listing.  The changes you make to the category order take
effect immediately.
</p>

<p>
Be careful about deleting categories that have links assigned to it.  Without a group
the associated links will be inaccessible through the module admin or the public side
of the site.  Only a manual repair of the data will recover the lost links and 
assign them to a new category group.
</p>

<h4>Category Properties</h4>

<p>
<table border=0 cellpadding=4 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Input Label</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Category Name</td>
	<td class="formd">Title or label for the category used as a heading for the list of links.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Category -->


<!-- #include file="../../../../footer_popup.asp" -->