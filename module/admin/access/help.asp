<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' help.asp
'	Display the help information for the discussion forums.
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

<h3>Access Help</h3>

<p>
The access administration area is where you define permissions for users
who will access the ASP nuke administration.  Each module is typically
associated with a single access right.  This right may be assigned to
your administrative users by using the <i>users</i> action in the list.
</p>

<p>
Basically, the access rights will control what appears in the administrator
menu that appears at the top of the main content area after the admin user
logs into the site.  The order of the access rights determine how the menu
choices appear in the menu.
</p>

<p>
In the future we will integrate a system that automatically sets up these
rights for you when modules are installed or removed from the system.  For
now, you will need to set these up manually as you install or develop new
modules.
</p>

<h4>How Rights are Used</h4>

<p>
Access rights are generally used in two different ways.  First, they are
used to manager the administrator menu items that show up in the web site
administration nav bar after a valid user logs in.  Secondly, they are used as
named permissions which can be referenced in your code.
</p>

<p>
To simplify the process of creating and managing permissions.  Four "pseudo-permissions"
were created to indicate the common actions: "add", "edit", "delete" and "view". 
You can configure your access rights to use these or not for each right in the
hierarchy.
</p>

<!-- SECTION_END:Overview -->

<!-- SECTION_START:Rights -->

<h3>Access Rights</h3>

Use the list of access rights to define the permissions that are used to secure
areas of the site.  You cannot just create an access right and expect it to work.
There has to be some code behind the access right in order for the security check
to be done.

<h4>Access Right Properties</h4>

<p>
<table border=0 cellpadding=2 cellspacing=0 class="list">
<tr>
	<Td class="listhead">Property</td>
	<td class="listhead">Description</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top">Parent Right</td>
	<td class="formd">Defines the parent access right where this right falls under.  If the item you are creating should be a "top-level" item than you can leave this field blank</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top">Right Name</td>
	<td class="formd">Name for the access right as it will appear in the administration menu displayed at the top of the main content area</td>
</tr>
<tr class="list0">
	<Td class="forml" valign="top">Menu Name</td>
	<td class="formd">Defines the menu label to display in the user's navigation bar.  If you leave this value blank, the access right WILL NOT be displayed in the user's menu. </td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top">Hyperlink</td>
	<td class="formd">Path and script entry point for the module admin.  This will typically look something like <kbd>/module/forum/discuss/admin/topic_list.asp</kbd></td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top">Has Add?</td>
	<td class="formd">Should a sub-permission of "Add" be created under this access right?</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top">Has Edit?</td>
	<td class="formd">Should a sub-permission of "Edit" be created under this access right?</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top">Has Delete?</td>
	<td class="formd">Should a sub-permission of "Delete" be created under this access right?</td>
</tr>
<tr class="list1">
	<Td class="forml" valign="top">Has View?</td>
	<td class="formd">Should a sub-permission of "View" be created under this access right?</td>
</tr>
</table>
</p>

<!-- SECTION_END:Rights -->

<!-- #include file="../../../../footer_popup.asp" -->