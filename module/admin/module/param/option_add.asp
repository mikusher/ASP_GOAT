﻿<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' option_add.asp
'	Add a new module parameter option to the database
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

Dim sErrorMsg
Dim sStat
Dim rsCat
Dim rsType
Dim rsParam
Dim sEditName
Dim rsOrder
Dim nModuleID
Dim nOrderNo
Dim sWhere

nModuleID = steNForm("ModuleID")
If steNForm("TypeID") > 0 Then
	sWhere = "TypeID = " & steNForm("TypeID")
Else
	sWhere = "ParamID = " & steNForm("ParamID")
End If

If steForm("action") = "add" Then
	' make sure the required fields are present
	If Trim(steForm("OptionValue")) = "" Then
		sErrorMsg = steGetText("Please enter the Option Value for this parameter option")
	ElseIf Trim(steForm("OptionLabel")) = ""	Then
		sErrorMsg = steGetText("Please enter the Option Label for this parameter option")
	ElseIf steNForm("TypeID") = 0 And steNForm("ParamID") = 0 Then
		sErrorMsg = steGetText("Param Type ID or Parameter ID missing, unable to continue")
	Else
		' retrieve the new order no
		sStat = "SELECT Coalesce(Max(OrderNo) + 1, 1) As OrderNo FROM tblModuleParamOption WHERE " & sWhere
		Set rsOrder = adoOpenRecordset(sStat)
		If Not rsOrder.EOF Then nOrderNo = rsOrder.Fields("OrderNo").Value Else nOrderNo = 1
		rsOrder.Close
		Set rsOrder = Nothing

		' create the new module option in the database
		sStat = "INSERT INTO tblModuleParamOption (" &_
				"	OrderNo, TypeID, ParamID, OptionValue, OptionLabel, Modified " &_
				") VALUES (" &_
				nOrderNo & "," &_
				steNForm("TypeID") & "," &_
				steNForm("ParamID") & "," &_
				steQForm("OptionValue") & "," &_
				steQForm("OptionLabel") & "," & adoGetDate &_
				")"
		Call adoExecute(sStat)
	End If
End If

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
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Options" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><%= sEditName %> - <% steTxt "Add Option" %></H3>

<P>
<% steTxt "Please enter the properties for the new module option using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="option_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">
<input type="hidden" name="moduleid" value="<%= nModuleID %>">
<input type="hidden" name="typeid" value="<%= steNForm("TypeID") %>">
<input type="hidden" name="paramid" value="<%= steNForm("ParamID") %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Option Value" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="OptionValue" VALUE="<%= steEncForm("OptionValue") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Option Label" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="OptionLabel" VALUE="<%= steEncForm("OptionLabel") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Option" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><%= sEditName %> - <% steTxt "Option Added" %></H3>

<P>
<% steTxt "The new module parameter option has been added to the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
<% If steNForm("TypeID") > 0 Then %>
	<A HREF="type_list.asp" class="adminlink"><% steTxt "Type List" %></A> &nbsp;
<% ElseIf steNForm("ParamID") > 0 Then %>
	<A HREF="param_list.asp?ModuleID=<%= nModuleID %>&paramid=<%= steNForm("paramid") %>" class="adminlink"><% steTxt "Param List" %></A> &nbsp;
<% End If %>
	<A HREF="option_list.asp?ModuleID=<%= nModuleID %>&typeid=<%= steNForm("typeid") %>&paramid=<%= steNForm("paramid") %>" class="adminlink"><% steTxt "Option List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->
