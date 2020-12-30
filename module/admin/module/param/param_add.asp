<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' param_add.asp
'	Add a new module param to the database
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
Dim rsOrder
Dim rsMod
Dim sModuleName
Dim nOrderNo

If steForm("action") = "add" Then
	' make sure the required fields are present
	If Trim(steForm("ParamName")) = "" Then
		sErrorMsg = steGetText("Please enter the Parameter Name for this param")
	ElseIf Trim(steForm("Label")) = ""	Then
		sErrorMsg = steGetText("Please enter the Label for this param")
	ElseIf steNForm("TypeID") = 0 Then
		sErrorMsg = steGetText("Please select a Data Type for the param")
	Else
		' retrieve a new order no
		sStat = "SELECT Coalesce(Max(OrderNo) + 1, 1) As OrderNo " &_
				"FROM tblModuleParam " &_
				"WHERE	ModuleID = " & steNForm("ModuleID")
		Set rsOrder = adoOpenRecordset(sStat)
		If Not rsOrder.EOF Then nOrderNo = rsOrder.Fields("OrderNo").Value Else nOrderNo = 1
		' create the new module param in the database
		sStat = "INSERT INTO tblModuleParam (" &_
				"	OrderNo, ModuleID, ParamName, ParamValue, Label, TypeID, " &_
				"	MinValue, MaxValue, IsRequired, HelpText, Modified" &_
				") VALUES (" &_
				nOrderNo & "," &_
				steNForm("ModuleID") & "," &_
				steQForm("ParamName") & "," &_
				steQForm("ParamValue") & "," &_
				steQForm("Label") & "," &_
				steNForm("TypeID") & "," &_
				steQForm("MinValue") & "," &_
				steQForm("MaxValue") & "," &_
				steNForm("IsRequired") & "," &_
				steQForm("HelpText") & "," & adoGetDate &_
				")"
		Call adoExecute(sStat)
		' update the parameter cache
		Call modParamCache(steNForm("ModuleID"), "")
	End If
End If

' retrieve the module name
sStat = "SELECT Title FROM tblModule WHERE ModuleID = " & steNForm("ModuleID")
Set rsMod = adoOpenRecordset(sStat)
If Not rsMod.EOF Then sModuleName = rsMod.Fields("Title").Value Else sModuleName = "* Unknown *"
rsMod.Close
Set rsMod = Nothing

' retrieve the list of module parameter types
sStat = "SELECT	TypeID, TypeName " &_
		"FROM	tblModuleParamType " &_
		"WHERE	Archive = 0 " &_
		"ORDER BY OrderNo"
Set rsType = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Param" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><%= sModuleName %>&nbsp;<% steTxt "Module - Add New Parameter" %></H3>

<P>
<% steTxt "Please enter the properties for the new module param using the form below." %>&nbsp;
<% steTxt "Parameters are used to configure ASP Nuke modules and ASP code must be written to support the parameters you create." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="param_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">
<input type="hidden" name="ModuleID" value="<%= steNForm("ModuleID") %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Parameter Name" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="ParamName" VALUE="<%= steEncForm("ParamName") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Parameter Value" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="ParamValue" VALUE="<%= steEncForm("ParamValue") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Data Type" %></TD><TD></TD>
	<TD>
		<SELECT NAME="TypeID" class="form">
		<option value=""> -- Choose --
		<% Do Until rsType.EOF %>
		<option value="<%= rsType.Fields("TypeID").Value %>"<% If steEncForm("TypeID") = CStr(rsType.Fields("TypeID").Value) Then Response.Write " SELECTED" %>> <%= rsType.Fields("TypeName").Value %>
		<%	rsType.MoveNext
		   Loop
		rsType.Close
		Set rsType = Nothing %>
		</SELECT>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Label" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Label" VALUE="<%= steEncForm("Label") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Minimum Value" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="MinValue" VALUE="<%= steEncForm("MinValue") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Maximum Value" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="MaxValue" VALUE="<%= steEncForm("MaxValue") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Is Required?" %></TD><TD></TD>
	<TD class="formd">
		<INPUT TYPE="radio" NAME="IsRequired" VALUE="1"<% If steNForm("IsRequired") <> 0 Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="IsRequired" VALUE="0"<% If steNForm("IsRequired") = 0 Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Help Text" %></TD><TD></TD>
	<TD class="formd">
		<TEXTAREA name="HelpText" cols="42" rows="8" class="form"><%= steEncForm("HelpText") %></TEXTAREA>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Parameter" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><%= sModuleName %>&nbsp;<% steTxt "Module - Param Added" %></H3>

<P>
<% steTxt "The new module param has been added to the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="group_list.asp" class="adminlink"><% steTxt "Group List" %></A> &nbsp;
	<A HREF="module_list.asp" class="adminlink"><% steTxt "Module List" %></A> &nbsp;
	<A HREF="param_list.asp?moduleid=<%= steNForm("ModuleID") %>" class="adminlink"><% steTxt "Param List" %></A> &nbsp;
	<A HREF="param_add.asp?moduleid=<%= steNForm("ModuleID") %>" class="adminlink"><% steTxt "Add Another" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->
