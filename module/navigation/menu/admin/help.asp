<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' help.asp
'	Display the help information for the menu items.
'	THIS FILE WILL AUTOMATICALLY BE PARSED INTO THE DATABASE
'	CHANGES TO THE FILE WILL BE DETECTED AND RELOADED AS NEEDED
'
' Copyright (C) 2002 Orvado Technologies (http://www.orvado.com)
'
' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'--------------------------------------------------------------------
%>
<!-- #include file="../../../../header_popup.asp" -->

<!-- SECTION_START:Overview -->

<h3>Menu Items Help</h3>

<p>
The main menu is a horizontal list of menu choices.  When the user moves
the cursor over the main menu item, a list of choices drops down.  From
this list the user may select one option from the menu.  Each choice is
usually an action to take which will involve going to a new page on the
web site.
</p>

<p>
The dynamic menu system is designed to work just like the menus used in
your operating system.  They require Javascript in order to work properly
so you should make sure your users support Javascript if you want to use
this module.
</p>

<p>
The menu items are arranged hierarchically meaning that they use a parent-child
relationship between each of the items.  Although you could define menu items
as many levels deep as you like, the dynamic menu system only supports two
levels.  The top level are the main menu choices displayed horizontally.  The
second level are the choiced displayed in the drop-lists.
</p>

<h4>Menu Items Properties</h4>

<p>
<table border=0 cellpadding=2 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Property</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Parent Menu</td>
	<td class="formd">Define the menu option under which this menu option should appear.  (Leave this blank if you want the item to appear at the top level.)</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Menu Name</td>
	<td class="formd">This is the label that will be used when displaying the option in the menu.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>URL</td>
	<td class="formd">Define the path and script to link the menu option to (such as <kbd>/docs/intro.asp</kbd>.)  Doesn't apply to top-level menu items which don't link to any web page.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>HTML Content</td>
	<td class="formd">As an alternative to linking to a new page, you can simply define all of the content for the target page inside the textarea.  This eliminates the need to create a separate web page for the content.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Overview -->

<!-- SECTION_START:Menu Items -->

<p>
The dynamic menu system is designed to work just like the menus used in
your operating system.  They require Javascript in order to work properly
so you should make sure your users support Javascript if you want to use
this module.
</p>

<p>
The menu module will usually appear at the top of the screen right below
your main site logo.  You can add, edit and delete menu items in real time.
Each of the menu items will take the user to a specific URL (which may be
located on your ASP Nuke site or on an external site.  You also have the
option of creating HTML comment directly for the menu item.
</p>

<p>
When you create the HTML for the menu item, the content will be stored in
your database.  Whenever you need to edit the item, you will need to come
back to the "menu" administration.
</p>

<p>
<table border=0 cellpadding=2 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Property</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>Parent Menu</td>
	<td class="formd">Define the menu option under which this menu option should appear.  (Leave this blank if you want the item to appear at the top level.)</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>Menu Name</td>
	<td class="formd">This is the label that will be used when displaying the option in the menu.</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top" nowrap>URL</td>
	<td class="formd">Define the path and script to link the menu option to (such as <kbd>/docs/intro.asp</kbd>.)  Doesn't apply to top-level menu items which don't link to any web page.</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top" nowrap>HTML Content</td>
	<td class="formd">As an alternative to linking to a new page, you can simply define all of the content for the target page inside the textarea.  This eliminates the need to create a separate web page for the content.</td>
</tr>
</table>
</p>

<!-- SECTION_END:Menu Items -->

<!-- #include file="../../../../footer_popup.asp" -->