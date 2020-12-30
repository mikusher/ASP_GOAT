<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' variable_add.asp
'	Add a new application variable to the database
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
Dim rsTab
Dim aTab
Dim rsType
Dim aType
Dim rsVar

If steForm("action") = "add" Then
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
		' create the new application variable in the database
		sStat = "INSERT INTO tblApplicationVar (" &_
				"	TabID, VarName, VarValue, TypeID, IsRequired, HelpText, Created " &_
				") VALUES (" &_
				steNForm("TabID") & "," &_
				steQForm("VarName") & "," &_
				steQForm("VarValue") & "," &_
				steNForm("TypeID") & "," &_
				steNForm("IsRequired") & "," &_
				steQForm("HelpText") & "," &_
				adoGetDate &_
				")"
		Call adoExecute(sStat)

		' create the application global variable immediately
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
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Configure" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New Application Variable" %></H3>

<P>
<% steTxt "Please enter the new properties for the new application variable using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="variable_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Tab Group" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><SELECT NAME="TabID" class="form">
		<OPTION VALUE=""> -- Choose --
		<% If IsArray(aTab) Then
			For I = 0 To UBound(aTab, 2) %>
		<OPTION VALUE="<%= aTab(0, I) %>"<% If steEncForm("TabID") = CStr(aTab(0, I)) Then Response.Write " SELECTED" %>> <%= Server.HTMLEncode(aTab(1, I)) %>
		<%	Next
		   End If %>
		</SELECT>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Name" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="VarName" VALUE="<%= steEncForm("VarName") %>" SIZE="32" MAXLENGTH="40" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Value" %></TD><TD></TD>
	<TD><TEXTAREA NAME="VarValue" COLS="52" ROWS="8" WRAP="virtual" class="form"><%= steEncForm("VarValue") %></TEXTAREA></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Data Type" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><SELECT NAME="TypeID" class="form">
		<OPTION VALUE=""> -- Choose --
		<% If IsArray(aType) Then
			For I = 0 To UBound(aType, 2) %>
		<OPTION VALUE="<%= aType(0, I) %>"<% If steEncForm("TypeID") = CStr(aType(0, I)) Then Response.Write " SELECTED" %>> <%= Server.HTMLEncode(aType(1, I)) %>
		<%	Next
		   End If %>
		</SELECT>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Is Required?" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="IsRequired" VALUE="1" class="formradio"<% If steNForm("IsRequired") = 1 Then Response.Write " CHECKED" %>> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="IsRequired" VALUE="0" class="formradio"<% If steNForm("IsRequired") = 0 Then Response.Write " CHECKED" %>> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Help Text" %></TD><TD></TD>
	<TD><TEXTAREA NAME="HelpText" COLS="52" ROWS="8" WRAP="virtual" class="form"><%= steEncForm("HelpText") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Variable" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "New Application Variable Added" %></H3>

<P>
<% steTxt "The new application variable has been added to the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="variable_list.asp" class="adminlink"><% steTxt "Variable List" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
