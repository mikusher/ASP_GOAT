<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' param_edit.asp
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

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("ParamName")) = "" Then
		sErrorMsg = steGetText("Please enter the Parameter Name for this param")
	ElseIf Trim(steForm("Label")) = ""	Then
		sErrorMsg = steGetText("Please enter the Label for this param")
	ElseIf steNForm("TypeID") = 0 Then
		sErrorMsg = steGetText("Please select a Data Type for the param")
	Else
		' create the new module param in the database
		sStat = "UPDATE tblModuleParam SET " &_
				"	ParamName = " & steQForm("ParamName") & "," &_
				"	ParamValue = " & steQForm("ParamValue") & "," &_
				"	Label = " & steQForm("Label") & "," &_
				"	TypeID = " & steNForm("TypeID") & "," &_
				"	MinValue = " & steQForm("MinValue") & "," &_
				"	MaxValue = " & steQForm("MaxValue") & "," &_
				"	IsRequired = " & steNForm("IsRequired") & "," &_
				"	HelpText = " & steQForm("HelpText") & " " &_
				"WHERE ParamID = " & steNForm("ParamID")
		Call adoExecute(sStat)
		' update the parameter cache
		Call modParamCache(steNForm("ModuleID"), "")
	End If
End If

' retrieve the list of module parameter types
sStat = "SELECT	TypeID, TypeName " &_
		"FROM	tblModuleParamType " &_
		"WHERE	Archive = 0 " &_
		"ORDER BY OrderNo"
Set rsType = adoOpenRecordset(sStat)

' retrieve the record to edit
sStat = "SELECT * FROM tblModuleParam WHERE ParamID = " & steNForm("ParamID")
Set rsParam = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Param" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Module Param" %></H3>

<P>
<% steTxt "Please make your changes to the module param using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="param_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<input type="hidden" name="ModuleID" value="<%= steNForm("ModuleID") %>">
<input type="hidden" name="ParamID" value="<%= steNForm("ParamID") %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Parameter Name" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="ParamName" VALUE="<%= steRecordEncValue(rsParam, "ParamName") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Parameter Value" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="ParamValue" VALUE="<%= steRecordEncValue(rsParam, "ParamValue") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Data Type" %></TD><TD></TD>
	<TD>
		<SELECT NAME="TypeID" class="form">
		<% Do Until rsType.EOF %>
		<option value="<%= rsType.Fields("TypeID").Value %>"<% If steRecordEncValue(rsParam, "TypeID") = CStr(rsType.Fields("TypeID").Value) Then Response.Write " SELECTED" %>> <%= rsType.Fields("TypeName").Value %>
		<%	rsType.MoveNext
		   Loop
		rsType.Close
		Set rsType = Nothing %>
		</SELECT>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Label" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Label" VALUE="<%= steRecordEncValue(rsParam, "Label") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Minimum Value" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="MinValue" VALUE="<%= steRecordEncValue(rsParam, "MinValue") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Maximum Value" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="MaxValue" VALUE="<%= steRecordEncValue(rsParam, "MaxValue") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Is Required?" %></TD><TD></TD>
	<TD class="formd">
		<INPUT TYPE="radio" NAME="IsRequired" VALUE="1"<% If steRecordBoolValue(rsParam, "IsRequired") Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="IsRequired" VALUE="0"<% If Not steRecordBoolValue(rsParam, "IsRequired") Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Help Text" %></TD><TD></TD>
	<TD class="formd">
		<TEXTAREA name="HelpText" cols="42" rows="8" class="form"><%= steRecordEncValue(rsParam, "HelpText") %></TEXTAREA>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update Param" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Module Param Updated" %></H3>

<P>
<% steTxt "The module param has been updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="param_list.asp?moduleid=<%= steNForm("ModuleID") %>" class="adminlink"><% steTxt "Param List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->
