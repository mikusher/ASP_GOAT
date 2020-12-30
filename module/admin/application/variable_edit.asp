<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' variable_edit.asp
'	Update existing application variable in the database
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
Dim rsVar
Dim rsTab
Dim aTab
Dim nVarID
Dim I

nVarID = steNForm("VarID")

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If steNForm("TabID") = 0 Then
		sErrorMsg = steGetText("Please select a Tab Group for the variable")
	ElseIf Trim(steForm("VarName")) = ""	Then
		sErrorMsg = steGetText("Please enter the Name for this variable")
	ElseIf Trim(steForm("VarValue")) = ""	Then
		sErrorMsg = steGetText("Please enter the Value for this variable")
	ElseIf steNForm("TypeID") = 0 Then
		sErrorMsg = steGetText("Please select a Data Type for this variable")
	ElseIf steForm("IsRequired") = "" Then
		sErrorMsg = steGetText("Please indicate whether this variable is required")
	Else
		' create the new application var in the database
		sStat = "UPDATE tblApplicationVar SET " &_
				"	TabID = " & steNForm("TabID") & "," &_
				"	VarName = " & steQForm("VarName") & "," &_
				"	VarValue = " & steQForm("VarValue") & "," &_
				"	TypeID = " & steNForm("TypeID") & "," &_
				"	IsRequired = " & steNForm("IsRequired") & "," &_
				"	HelpText = " & steQForm("HelpText") & " " &_
				"WHERE	VarID = " & nVarID
		Call adoExecute(sStat)

		' update the application global variable immediately
		Application(steForm("VarName")) = steForm("VarValue")
	End If
End If

' retrieve the list of tabs to choose from
sStat = "SELECT TabID, TabName FROM tblApplicationVarTab WHERE Archive = 0 ORDER BY OrderNo"
Set rsTab = adoOpenRecordset(sStat)
If Not rsTab.EOF Then aTab = rsTab.GetRows
rsTab.Close
Set rsTab = Nothing

sStat = "SELECT TypeID, TypeName FROM tblApplicationVarType WHERE Archive = 0 ORDER BY OrderNo"
Set rsType = adoOpenRecordset(sStat)
If Not rsType.EOF Then aType = rsType.GetRows
rsType.Close
Set rsType = Nothing

sStat = "SELECT * FROM tblApplicationVar WHERE VarID = " & nVarID
Set rsVar = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Configure" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Application Variable" %></H3>

<P>
<% steTxt "Please make your changes to the application variable using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="variable_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="VarID" VALUE="<%= nVarID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Tab Group" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><SELECT NAME="TabID" class="form">
		<OPTION VALUE=""> -- Choose --
		<% If IsArray(aTab) Then
			For I = 0 To UBound(aTab, 2) %>
		<OPTION VALUE="<%= aTab(0, I) %>"<% If steRecordEncValue(rsVar, "TabID") = CStr(aTab(0, I)) Then Response.Write " SELECTED" %>> <%= Server.HTMLEncode(aTab(1, I)) %>
		<%	Next
		   End If %>
		</SELECT>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="VarName" VALUE="<%= steRecordEncValue(rsVar, "VarName") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Value" %></TD><TD></TD>
	<TD><TEXTAREA NAME="VarValue" COLS="52" ROWS="10" WRAP="virtual" class="form"><%= steRecordEncValue(rsVar, "VarValue") %></TEXTAREA></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Data Type" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><SELECT NAME="TypeID" class="form">
		<OPTION VALUE=""> -- Choose --
		<% If IsArray(aType) Then
			For I = 0 To UBound(aType, 2) %>
		<OPTION VALUE="<%= aType(0, I) %>"<% If steRecordEncValue(rsVar, "TypeID") = CStr(aType(0, I)) Then Response.Write " SELECTED" %>> <%= Server.HTMLEncode(aType(1, I)) %>
		<%	Next
		   End If %>
		</SELECT>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Is Required?" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="IsRequired" VALUE="1" class="formradio"<% If steRecordBoolValue(rsVar, "IsRequired") Then Response.Write " CHECKED" %>> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="IsRequired" VALUE="0" class="formradio"<% If Not steRecordBoolValue(rsVar, "IsRequired") Then Response.Write " CHECKED" %>> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top">Help Text</TD><TD></TD>
	<TD><TEXTAREA NAME="HelpText" COLS="52" ROWS="8" WRAP="virtual" class="form"><%= steRecordEncValue(rsVar, "HelpText") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update Variable" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Application Variable Updated" %></H3>

<P>
<% steTxt "The application variable was successfully updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="variable_list.asp" class="adminlink"><% steTxt "Variable List" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
