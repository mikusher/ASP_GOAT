﻿<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/tree_lib.asp" -->
<%
'--------------------------------------------------------------------
' country_list.asp
'	Displays a list of the registered members for the site
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

Dim sStat
Dim rsList
Dim nArchive
Dim nSelected

nArchive = steNForm("Archive")
nSelected = steNForm("FolderID")
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Categories" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3>Categories</H3>

<% treTreeAdmin "tblDocFolder", "FolderID", "ParentFolderID", "FolderName", "Category Name", "", _
		nArchive, nSelected %>

<P ALIGN="Center">
	<A HREF="country_add.asp">Add New Category</A>
</P>

<!-- #include file="../../../../footer.asp" -->
