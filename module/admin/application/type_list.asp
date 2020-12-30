<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' type_list.asp
'	Displays a list of the current variable types defined in the database
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
Dim rsType
Dim sAction

sAction = LCase(steForm("Action"))

Select Case sAction
	Case "moveup"
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblApplicationVarType " &_
					"SET	OrderNo = OrderNo + 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & steNForm("OrderNo") - 1
			Call adoExecute(sStat)

			sStat = "UPDATE	tblApplicationVarType " &_
					"SET	OrderNo = " & (steNForm("OrderNo") - 1) & ", Modified = " & adoGetDate & " " &_
					"WHERE	TypeID = " & steNForm("TypeID")
			Call adoExecute(sStat)
			modRefresh True
	Case "movedown"
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblApplicationVarType " &_
					"SET	OrderNo = OrderNo - 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & steNForm("OrderNo") + 1
			Call adoExecute(sStat)

			sStat = "UPDATE	tblApplicationVarType " &_
					"SET	OrderNo = " & (steNForm("OrderNo") + 1) & ", Modified = " & adoGetDate & " " &_
					"WHERE	TypeID = " & steNForm("TypeID")
			Call adoExecute(sStat)
			modRefresh True
End Select

sStat = "SELECT	OrderNo, TypeID, TypeCode, TypeName, OrderNo, HTMLInputType, HasOptions, Modified " &_
		"FROM	tblApplicationVarType " &_
		"ORDER BY OrderNo"
Set rsType = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Types" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "Application Variable Type List" %></H3>

<P>
<% steTxt "The following application variable types are defined in the database." %>&nbsp;
<% steTxt "These control the main areas of your content layout." %>&nbsp;
<% steTxt "You cannot just create a new variable type and have it show up on the site, special ASP script must be written to handle new types." %>
</P>

<% If Not rsType.EOF Then %>

<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 class="list">
<TR>
	<TD CLASS="listhead"><% steTxt "Order" %></TD>
	<TD CLASS="listhead"><% steTxt "Code" %></TD>
	<TD CLASS="listhead"><% steTxt "Type Name" %></TD>
	<TD CLASS="listhead"><% steTxt "Input" %></TD>
	<TD CLASS="listhead"><% steTxt "Opt" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Modified" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<% I = 0
Do Until rsType.EOF %>
<TR CLASS="list<%= I mod 2 %>">
	<TD><%= rsType.Fields("OrderNo").Value %></TD>
	<TD><%= rsType.Fields("TypeCode").Value %></TD>
	<TD><%= rsType.Fields("TypeName").Value %></TD>
	<TD><%= rsType.Fields("HTMLInputType").Value %></TD>
	<TD><% If steRecordBoolValue(rsType, "HasOptions") Then Response.Write "Y" Else Response.Write "N" %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsType.Fields("Modified").Value, vbShortDate) %></TD>
	<TD>
		<A HREF="type_list.asp?TypeID=<%= rsType.Fields("TypeID").Value %>&orderno=<%= rsType.Fields("OrderNo").Value %>&action=moveup" class="actionlink"><% steTxt "up" %></A> .
		<A HREF="type_list.asp?TypeID=<%= rsType.Fields("TypeID").Value %>&orderno=<%= rsType.Fields("OrderNo").Value %>&action=movedown" class="actionlink"><% steTxt "down" %></A> .
		<A HREF="option_list.asp?TypeID=<%= rsType.Fields("TypeID").Value %>" class="actionlink"><% steTxt "options" %></A> .
		<A HREF="type_edit.asp?TypeID=<%= rsType.Fields("TypeID").Value %>" class="actionlink"><% steTxt "edit" %></A> .
		<A HREF="type_delete.asp?TypeID=<%= rsType.Fields("TypeID").Value %>" class="actionlink"><% steTxt "delete" %></A>
	</TD>
</TR>
<%	rsType.MoveNext
	I = I + 1
   Loop %>
</TABLE>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No variable types exist in the database" %></B></P>

<% End If %>

<P ALIGN="center">
	<A HREF="type_add.asp" class="adminlink"><% steTxt "Add New Type" %></A>
</P>

<!-- #include file="../../../footer.asp" -->