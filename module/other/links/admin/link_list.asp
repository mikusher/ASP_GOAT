<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' link_list.asp
'	Displays a list of the current links for the site
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
Dim rsLink

sStat = "SELECT	tblLink.LinkID, tblLink.URL, tblLink.Label, tblLinkCategory.CategoryName, " &_
		"		tblLink.Created, tblLink.Modified " &_
		"FROM	tblLink " &_
		"INNER JOIN	tblLinkCategory ON tblLink.CategoryID = tblLinkCategory.CategoryID " &_
		"ORDER BY tblLink.Created DESC"
Set rsLink = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Link" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "Link List" %></H3>

<P>
<% steTxt "Shown below are all of the current links defined in the database." %>
</P>

<% If Not rsLink.EOF Then %>

<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 class="list">
<TR>
	<TD class="listhead"><% steTxt "Label" %></TD>
	<TD class="listhead"><% steTxt "URL" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Category" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Modified" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<%	I = 0
	Do Until rsLink.EOF %>
<TR class="list<%= I mod 2 %>">
	<TD><%= rsLink.Fields("Label").Value %></TD>
	<TD><a href="<%= rsLink.Fields("URL").Value %>" target="_new"><%= rsLink.Fields("URL").Value %></A></TD>
	<TD ALIGN="right"><%= rsLink.Fields("CategoryName").Value %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsLink.Fields("Modified").Value, vbShortDate) %></TD>
	<TD>
		<A HREF="link_edit.asp?LinkID=<%= rsLink.Fields("LinkID").Value %>" class="actionlink"><% steTxt "edit" %></A> .
		<A HREF="link_delete.asp?LinkID=<%= rsLink.Fields("LinkID").Value %>" class="actionlink"><% steTxt "delete" %></A>
	</TD>
</TR>
<%	rsLink.MoveNext
	I = I + 1
   Loop %>
</TABLE>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No links exist in the database" %></B></P>

<% End If %>

<P ALIGN="center">
	<A HREF="link_add.asp" class="adminlink"><% steTxt "Add New Link" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->