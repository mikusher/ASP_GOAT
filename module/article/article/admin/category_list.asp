﻿<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' category_list.asp
'	Displays a list of the article categories for the site
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
Dim rsCat

sStat = "SELECT	CategoryID, CategoryName, Comments, Created, Modified " &_
		"FROM	tblArticleCategory " &_
		"ORDER BY CategoryName"
Set rsCat = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Category" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "Article Category List" %></H3>

<P>
<% steTxt "Shown below are all of the current article categories defined in the database." %>
</P>

<% If Not rsCat.EOF Then %>

<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 CLASS="list">
<TR>
	<TD CLASS="listhead"><% steTxt "Category Name" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Created" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Modified" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<%
I = 0
Do Until rsCat.EOF %>
<TR CLASS="list<%= I mod 2 %>">
	<TD><%= rsCat.Fields("CategoryName").Value %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsCat.Fields("Created").Value, vbShortDate) %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsCat.Fields("Modified").Value, vbShortDate) %></TD>
	<TD>
		<A HREF="category_edit.asp?categoryid=<%= rsCat.Fields("CategoryID").Value %>" CLASS="actionlink"><% steTxt "edit" %></A> .
		<A HREF="category_delete.asp?categoryid=<%= rsCat.Fields("CategoryID").Value %>" CLASS="actionlink"><% steTxt "delete" %></A>
	</TD>
</TR>
<%	rsCat.MoveNext
	I = I + 1
   Loop %>
</TABLE>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No article categories exist in the database" %></B></P>

<% End If %>

<P ALIGN="center">
	<A HREF="category_add.asp" CLASS="adminlink"><% steTxt "Add New Article Category" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->