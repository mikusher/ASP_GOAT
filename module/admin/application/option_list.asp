<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' option_list.asp
'	Displays a list of the current variable options defined in the database
'	May be related to a variable type (tblApplicationVarType) or a variable
'	(tblApplicationVar)
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
Dim rsOption
Dim sAction
Dim rsType
Dim rsVar
Dim sEditName
Dim sWhere

sAction = LCase(steForm("Action"))

Select Case sAction
	Case "activ"
			sStat = "UPDATE	tblApplicationVarOption " &_
					"SET	Archive = 0, Modified = " & adoGetDate & " " &_
					"WHERE	OptionID = " & steNForm("OptionID")
			Call adoExecute(sStat)
	Case "deactiv"
			sStat = "UPDATE	tblApplicationVarOption " &_
					"SET	Archive = 1, Modified = " & adoGetDate & " " &_
					"WHERE	OptionID = " & steNForm("OptionID")
			Call adoExecute(sStat)
	Case "valid"
			sStat = "UPDATE	tblApplicationVarOption " &_
					"SET	IsValid = 1, Modified = " & adoGetDated & " " &_
					"WHERE	OptionID = " & steNForm("OptionID")
			Call adoExecute(sStat)
	Case "invalid"
			sStat = "UPDATE	tblApplicationVarOption " &_
					"SET	IsValid = 0, Modified = " & adoGetDate & " " &_
					"WHERE	OptionID = " & steNForm("OptionID")
			Call adoExecute(sStat)
	Case "moveup"
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblApplicationVarOption " &_
					"SET	OrderNo = OrderNo + 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & steNForm("OrderNo") - 1
			Call adoExecute(sStat)

			sStat = "UPDATE	tblApplicationVarOption " &_
					"SET	OrderNo = " & (steNForm("OrderNo") - 1) & ", Modified = " & adoGetDate & " " &_
					"WHERE	OptionID = " & steNForm("OptionID")
			Call adoExecute(sStat)
	Case "movedown"
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblApplicationVarOption " &_
					"SET	OrderNo = OrderNo - 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & steNForm("OrderNo") + 1
			Call adoExecute(sStat)

			sStat = "UPDATE	tblApplicationVarOption " &_
					"SET	OrderNo = " & (steNForm("OrderNo") + 1) & ", Modified = " & adoGetDate & " " &_
					"WHERE	OptionID = " & steNForm("OptionID")
			Call adoExecute(sStat)
End Select

' retrieve the name of the type we are working with
If steNForm("TypeID") > 0 Then
	sStat = "SELECT TypeName FROM tblApplicationVarType WHERE TypeID = " & steNForm("TypeID")
	Set rsType = adoOpenRecordset(sStat)
	If Not rsType.EOF Then sEditName = rsType.Fields("TypeName").Value
	rsType.Close
	Set rsType = Nothing
Else
	sStat = "SELECT VarName FROM tblApplicationVar WHERE VarID = " & steNForm("VarID")
	Set rsVar = adoOpenRecordset(sStat)
	If Not rsVar.EOF Then sEditName = rsVar.Fields("VarName").Value
	rsVar.Close
	Set rsVar = Nothing
End If

' retrieve the list of options for the type or variable
If steNForm("TypeID") > 0 Then
	sWhere = "TypeID = " & steNForm("TypeID")
End If
If steNForm("VarID") > 0 Then
	sWhere = "VarID = " & steNForm("VarID")
End If
sStat = "SELECT	OrderNo, OptionID, OptionValue, OptionLabel, OrderNo, IsValid, Archive, Modified " &_
		"FROM	tblApplicationVarOption " &_
		"WHERE " & sWhere & " " &_
		"ORDER BY OrderNo"
Set rsOption = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Options" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><%= sEditName %> <% steTxt "Option List" %></H3>

<P>
<% steTxt "The following application variable options are defined in the database for the specified item." %>&nbsp;
<% steTxt "The user may only choose one of the options from the list below." %>
</P>

<% If Not rsOption.EOF Then %>

<form method="post" action="#">
<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 class="list">
<TR>
	<TD CLASS="listhead"><% steTxt "Order" %></TD>
	<TD CLASS="listhead"><% steTxt "Value" %></TD>
	<TD CLASS="listhead"><% steTxt "Label" %></TD>
	<TD CLASS="listhead"><% steTxt "Valid" %></TD>
	<TD CLASS="listhead"><% steTxt "Active" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Modified" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<% I = 0
Do Until rsOption.EOF %>
<TR CLASS="list<%= I mod 2 %>">
	<TD><%= rsOption.Fields("OrderNo").Value %></TD>
	<TD><%= rsOption.Fields("OptionValue").Value %></TD>
	<TD><%= rsOption.Fields("OptionLabel").Value %></TD>
	<TD><% If steRecordBoolValue(rsOption, "IsValid") Then %>
		<input name="isvalid" type="checkbox" checked onclick="location.href='option_list.asp?typeid=<%= steNForm("TypeID") %>&varid=<%= steNForm("varid") %>&optionid=<%= rsOption.Fields("OptionID").Value %>&action=invalid'">
		<% Else %>
		<input name="isvalid" type="checkbox" onclick="location.href='option_list.asp?typeid=<%= steNForm("TypeID") %>&varid=<%= steNForm("varid") %>&optionid=<%= rsOption.Fields("OptionID").Value %>&action=valid'">
		<% End If %>
	</TD>
	<TD><% If Not steRecordBoolValue(rsOption, "Archive") Then %>
		<input name="archive" type="checkbox" checked onclick="location.href='option_list.asp?typeid=<%= steNForm("TypeID") %>&varid=<%= steNForm("varid") %>&optionid=<%= rsOption.Fields("OptionID").Value %>&action=deactiv'">
		<% Else %>
		<input name="archive" type="checkbox" onclick="location.href='option_list.asp?typeid=<%= steNForm("TypeID") %>&varid=<%= steNForm("varid") %>&optionid=<%= rsOption.Fields("OptionID").Value %>&action=activ'">
		<% End If %>
	</TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsOption.Fields("Modified").Value, vbShortDate) %></TD>
	<TD>
		<A HREF="option_list.asp?typeid=<%= steNForm("TypeID") %>&varid=<%= steNForm("varid") %>&OptionID=<%= rsOption.Fields("OptionID").Value %>&orderno=<%= rsOption.Fields("OrderNo").Value %>&action=moveup" class="actionlink"><% steTxt "up" %></A> .
		<A HREF="option_list.asp?typeid=<%= steNForm("TypeID") %>&varid=<%= steNForm("varid") %>&OptionID=<%= rsOption.Fields("OptionID").Value %>&orderno=<%= rsOption.Fields("OrderNo").Value %>&action=movedown" class="actionlink"><% steTxt "down" %></A> .
		<A HREF="option_edit.asp?typeid=<%= steNForm("TypeID") %>&varid=<%= steNForm("varid") %>&OptionID=<%= rsOption.Fields("OptionID").Value %>" class="actionlink"><% steTxt "edit" %></A> .
		<A HREF="option_delete.asp?typeid=<%= steNForm("TypeID") %>&varid=<%= steNForm("varid") %>&OptionID=<%= rsOption.Fields("OptionID").Value %>" class="actionlink"><% steTxt "delete" %></A>
	</TD>
</TR>
<%	rsOption.MoveNext
	I = I + 1
   Loop %>
</TABLE>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No parameter options exist in the database" %></B></P>

<% End If %>

<P ALIGN="center">
<% If steNForm("TypeID") > 0 Then %>
	<A HREF="type_list.asp" class="adminlink"><% steTxt "Type List" %></a> &nbsp;
<% Else %>
	<A HREF="variable_list.asp" class="adminlink"><% steTxt "Variable List" %></a> &nbsp;
<% End If %>
	<A HREF="option_add.asp?TypeID=<%= steNForm("TypeID") %>&varid=<%= steNForm("varid") %>" class="adminlink"><% steTxt "Add New Option" %></A>
</P>

<!-- #include file="../../../footer.asp" -->