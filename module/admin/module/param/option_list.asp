<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' option_list.asp
'	Displays a list of the current param options defined in the database
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
Dim rsParam
Dim sEditName
Dim sWhere

sAction = LCase(steForm("Action"))

Select Case sAction
	Case "activ"
			sStat = "UPDATE	tblModuleParamOption " &_
					"SET	Archive = 0, Modified = " & adoGetDate & " " &_
					"WHERE	OptionID = " & steNForm("OptionID")
			Call adoExecute(sStat)
	Case "deactiv"
			sStat = "UPDATE	tblModuleParamOption " &_
					"SET	Archive = 1, Modified = " & adoGetDate & " " &_
					"WHERE	OptionID = " & steNForm("OptionID")
			Call adoExecute(sStat)
	Case "valid"
			sStat = "UPDATE	tblModuleParamOption " &_
					"SET	IsValid = 1, Modified = " & adoGetDate & " " &_
					"WHERE	OptionID = " & steNForm("OptionID")
			Call adoExecute(sStat)
	Case "invalid"
			sStat = "UPDATE	tblModuleParamOption " &_
					"SET	IsValid = 0, Modified = " & adoGetDate & " " &_
					"WHERE	OptionID = " & steNForm("OptionID")
			Call adoExecute(sStat)
	Case "moveup"
		Dim rsPrev, sPrevOrder

		' retrieve the previous order no
		sStat = "SELECT	OrderNo " &_
				"FROM	tblModuleParamOption " &_
				"WHERE	OrderNo < " & steNForm("OrderNo") & " " &_
				"ORDER BY OrderNo DESC"
		Set rsPrev = adoOpenRecordset(sStat)
		If Not rsPrev.EOF Then
			sPrevOrder = rsPrev.Fields("OrderNo").Value
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblModuleParamOption " &_
					"SET	OrderNo = OrderNo + 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & sPrevOrder
			Call adoExecute(sStat)

			sStat = "UPDATE	tblModuleParamOption " &_
					"SET	OrderNo = " & sPrevOrder & ", Modified = " & adoGetDate & " " &_
					"WHERE	OptionID = " & steNForm("OptionID")
			Call adoExecute(sStat)
		End If
		rsPrev.Close
		Set rsPrev = Nothing
	Case "movedown"
		Dim rsNext, sNextOrder

		' retrieve the next order no
		sStat = "SELECT	OrderNo " &_
				"FROM	tblModuleParamOption " &_
				"WHERE	OrderNo > " & steNForm("OrderNo") & " " &_
				"ORDER BY OrderNo"
		Set rsNext = adoOpenRecordset(sStat)
		If Not rsNext.EOF Then
			sNextOrder = rsNext.Fields("OrderNo").Value
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblModuleParamOption " &_
					"SET	OrderNo = OrderNo - 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & sNextOrder
			Call adoExecute(sStat)

			sStat = "UPDATE	tblModuleParamOption " &_
					"SET	OrderNo = " & sNextOrder & ", Modified = " & adoGetDate & " " &_
					"WHERE	OptionID = " & steNForm("OptionID")
			Call adoExecute(sStat)
		End If
		rsNext.Close
		Set rsNext = Nothing
End Select

' retrieve the name of the type we are working with
If steNForm("TypeID") > 0 Then
	sStat = "SELECT TypeName FROM tblModuleParamType WHERE TypeID = " & steNForm("TypeID")
	Set rsType = adoOpenRecordset(sStat)
	If Not rsType.EOF Then sEditName = rsType.Fields("TypeName").Value
	rsType.Close
	Set rsType = Nothing
Else
	sStat = "SELECT ParamName FROM tblModuleParam WHERE ParamID = " & steNForm("ParamID")
	Set rsParam = adoOpenRecordset(sStat)
	If Not rsParam.EOF Then sEditName = rsParam.Fields("ParamName").Value
	rsParam.Close
	Set rsParam = Nothing
End If

' retrieve the list of options for the type or parameter
If steNForm("TypeID") > 0 Then
	sWhere = "TypeID = " & steNForm("TypeID")
End If
If steNForm("ParamID") > 0 Then
	sWhere = "ParamID = " & steNForm("ParamID")
End If
If sWhere <> "" Then
	sStat = "SELECT	OrderNo, OptionID, OptionValue, OptionLabel, OrderNo, IsValid, Archive, Modified " &_
			"FROM	tblModuleParamOption " &_
			"WHERE " & sWhere & " " &_
			"ORDER BY OrderNo"
	Set rsOption = adoOpenRecordset(sStat)
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->
<!-- #include file="../../../../lib/class/form/listinput.asp" -->

<% sCurrentTab = "Options" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><%= sEditName %>&nbsp;<% steTxt "Option List" %></H3>

<% If Not IsObject(rsOption) Then %>

<p>
<% steTxt "Please choose a paramter type or parameter to work with from the lists below:" %>
</p>

<p>
<table border=0 cellpadding=4 cellspacing=0>
<tr>
	<Td class="forml"><% steTxt "Param Type" %></td><td>&nbsp;&nbsp;</td>
	<td class="formd">
<% ' display the drop-list to choose the type or param to work with
	Dim oList
	Set oList = New clsListInput
 	oList.TreeListInput "TypeID", "tblModuleParamType", "TypeID", "", "Archive = 0", _
		"TypeName", "TypeID", "TypeName", steNForm("TypeID"), "", True %>
	</td>
	<Td class="forml"><% steTxt "Param Name" %></td><td></td>
	<td class="formd">
<%	' next build a drop-list for the params themselves
 	oList.TreeListInput "ParamID", "tblModuleParam", "ParamID", "", "Archive = 0", _
		"ParamName", "ParamID", "ParamName", steNForm("ParamID"), "", True %>
	</td>
</tr>
</table>
</p>

<% Else %>

<P>
<% steTxt "The following module param options are defined in the database for the specified item." %>&nbsp;
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
		<input name="isvalid" type="checkbox" checked onclick="location.href='option_list.asp?typeid=<%= steNForm("TypeID") %>&paramid=<%= steNForm("paramid") %>&optionid=<%= rsOption.Fields("OptionID").Value %>&action=invalid'">
		<% Else %>
		<input name="isvalid" type="checkbox" onclick="location.href='option_list.asp?typeid=<%= steNForm("TypeID") %>&paramid=<%= steNForm("paramid") %>&optionid=<%= rsOption.Fields("OptionID").Value %>&action=valid'">
		<% End If %>
	</TD>
	<TD><% If Not steRecordBoolValue(rsOption, "Archive") Then %>
		<input name="archive" type="checkbox" checked onclick="location.href='option_list.asp?typeid=<%= steNForm("TypeID") %>&paramid=<%= steNForm("paramid") %>&optionid=<%= rsOption.Fields("OptionID").Value %>&action=deactiv'">
		<% Else %>
		<input name="archive" type="checkbox" onclick="location.href='option_list.asp?typeid=<%= steNForm("TypeID") %>&paramid=<%= steNForm("paramid") %>&optionid=<%= rsOption.Fields("OptionID").Value %>&action=activ'">
		<% End If %>
	</TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsOption.Fields("Modified").Value, vbShortDate) %></TD>
	<TD>
		<A HREF="option_list.asp?typeid=<%= steNForm("TypeID") %>&paramid=<%= steNForm("paramid") %>&OptionID=<%= rsOption.Fields("OptionID").Value %>&orderno=<%= rsOption.Fields("OrderNo").Value %>&action=moveup" class="actionlink"><% steTxt "up" %></A> .
		<A HREF="option_list.asp?typeid=<%= steNForm("TypeID") %>&paramid=<%= steNForm("paramid") %>&OptionID=<%= rsOption.Fields("OptionID").Value %>&orderno=<%= rsOption.Fields("OrderNo").Value %>&action=movedown" class="actionlink"><% steTxt "down" %></A> .
		<A HREF="option_edit.asp?typeid=<%= steNForm("TypeID") %>&paramid=<%= steNForm("paramid") %>&OptionID=<%= rsOption.Fields("OptionID").Value %>" class="actionlink"><% steTxt "edit" %></A> .
		<A HREF="option_delete.asp?typeid=<%= steNForm("TypeID") %>&paramid=<%= steNForm("paramid") %>&OptionID=<%= rsOption.Fields("OptionID").Value %>" class="actionlink"><% steTxt "delete" %></A>
	</TD>
</TR>
<%	rsOption.MoveNext
	I = I + 1
   Loop %>
</TABLE>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No parameter options exist in the database" %></B></P>

<% End If %>

<% End If %>

<P ALIGN="center">
	<A HREF="option_add.asp?ModuleID=<%= steNForm("ModuleID") %>&paramid=<%= steNForm("paramid") %>" class="adminlink"><% steTxt "Add New Option" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->