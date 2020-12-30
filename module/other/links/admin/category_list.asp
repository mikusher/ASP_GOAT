<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' category_list.asp
'	Displays a list of the current admins for the site
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
Dim nCategoryID

nCategoryID = steNForm("CategoryID")

Select Case LCase(steForm("action"))
	Case "moveup"
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblLinkCategory " &_
					"SET	OrderNo = OrderNo + 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & steNForm("OrderNo") - 1
			Call adoExecute(sStat)

			sStat = "UPDATE	tblLinkCategory " &_
					"SET	OrderNo = " & (steNForm("OrderNo") - 1) & ", Modified = " & adoGetDate & " " &_
					"WHERE	CategoryID = " & steNForm("CategoryID")
			Call adoExecute(sStat)
	Case "movedown"
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblLinkCategory " &_
					"SET	OrderNo = OrderNo - 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & steNForm("OrderNo") + 1
			Call adoExecute(sStat)

			sStat = "UPDATE	tblLinkCategory " &_
					"SET	OrderNo = " & (steNForm("OrderNo") + 1) & ", Modified = " & adoGetDate & " " &_
					"WHERE	CategoryID = " & steNForm("CategoryID")
			Call adoExecute(sStat)
End Select

sStat = "SELECT	OrderNo, CategoryID, CategoryName, Modified " &_
		"FROM	tblLinkCategory " &_
		"ORDER BY OrderNo"
Set rsCat = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Category" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "Link Category List" %></H3>

<P>
<% steTxt "Shown below are all of the link categories defined in the database." %>
</P>

<% If Not rsCat.EOF Then %>

<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 class="list">
<TR>
	<TD class="listhead"><% steTxt "Order" %></TD>
	<TD class="listhead"><% steTxt "Category Name" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Modified" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<%	I = 0
	Do Until rsCat.EOF %>
<TR class="<%= I mod 2 %>">
	<TD><%= rsCat.Fields("OrderNo").Value %></TD>
	<TD><%= rsCat.Fields("CategoryName").Value %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsCat.Fields("Modified").Value, vbShortDate) %></TD>
	<TD>
		<A HREF="category_list.asp?CategoryID=<%= rsCat.Fields("CategoryID").Value %>&orderno=<%= rsCat.Fields("OrderNo").Value %>&action=moveup" class="actionlink"><% steTxt "up" %></A> .
		<A HREF="category_list.asp?CategoryID=<%= rsCat.Fields("CategoryID").Value %>&orderno=<%= rsCat.Fields("OrderNo").Value %>&action=movedown" class="actionlink"><% steTxt "down" %></A> .
		<A HREF="category_edit.asp?CategoryID=<%= rsCat.Fields("CategoryID").Value %>" class="actionlink"><% steTxt "edit" %></A> .
		<A HREF="category_delete.asp?CategoryID=<%= rsCat.Fields("CategoryID").Value %>" class="actionlink"><% steTxt "delete" %></A>
	</TD>
</TR>
<%	rsCat.MoveNext
   Loop %>
</TABLE>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No categories exist in the database" %></B></P>

<% End If %>

<P ALIGN="center">
	<A HREF="category_add.asp" class="adminlink"><% steTxt "Add New Link Category" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->