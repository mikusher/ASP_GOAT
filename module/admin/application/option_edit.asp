<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' option_edit.asp
'	Edit an existing variable option to the database
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
Dim sWhere

If steNForm("TypeID") > 0 Then
	sWhere = "TypeID = " & steNForm("TypeID")
Else
	sWhere = "VarID = " & steNForm("VarID")
End If

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("OptionValue")) = "" Then
		sErrorMsg = steGetText("Please enter the Option Value for this variable option")
	ElseIf Trim(steForm("OptionLabel")) = ""	Then
		sErrorMsg = steGetText("Please enter the Option Label for this variable option")
	ElseIf steNForm("TypeID") = 0 And steNForm("VarID") = 0 Then
		sErrorMsg = steGetText("Type ID or Variable ID missing, unable to continue")
	Else
		' create the new variable option in the database
		sStat = "UPDATE tblApplicationVarOption SET " &_
				"OptionValue = " & steQForm("OptionValue") & "," &_
				"OptionLabel = " & steQForm("OptionLabel") &_
				"WHERE " & sWhere & " " &_
				"AND	OptionID = " & steNForm("OptionID")
		Call adoExecute(sStat)
	End If
End If

' retrieve the option to edit here
sStat = "SELECT * FROM tblApplicationVarOption WHERE OptionID = " & steNForm("OptionID")
Set rsEdit = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Options" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Variable Option" %></H3>

<P>
<% steTxt "Please enter your changes for the application variable option using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="option_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<input type="hidden" name="typeid" value="<%= steNForm("TypeID") %>">
<input type="hidden" name="varid" value="<%= steNForm("VarID") %>">
<input type="hidden" name="optionid" value="<%= steNForm("OptionID") %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Option Value" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="OptionValue" VALUE="<%= steRecordEncValue(rsEdit, "OptionValue") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Option Label" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="OptionLabel" VALUE="<%= steRecordEncValue(rsEdit, "OptionLabel") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update Option" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3>New Variable Option Updated</H3>

<P>
<% steTxt "The application variable option was successfully updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
<% If steNForm("TypeID") > 0 Then %>
	<A HREF="type_list.asp" class="adminlink"><% steTxt "Type List" %></A> &nbsp;
<% ElseIf steNForm("VarID") > 0 Then %>
	<A HREF="variable_list.asp?varid=<%= steNForm("varid") %>" class="adminlink"><% steTxt "Variable List" %></A> &nbsp;
<% End If %>
	<A HREF="option_list.asp?typeid=<%= steNForm("typeid") %>&varid=<%= steNForm("varid") %>" class="adminlink"><% steTxt "Option List" %></A>
</P>
<!-- #include file="../../../footer.asp" -->
