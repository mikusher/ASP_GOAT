<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/class/form/listinput.asp" -->
<%
'--------------------------------------------------------------------
' userright_edit.asp
'	Update existing user right in the database
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
Dim rsRight
Dim nRightID

nRightID = steNForm("rightid")

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("RightName")) = ""	Then
		sErrorMsg = steGetText("Please enter the Right Name for this user right")
	' ElseIf Trim(steForm("Hyperlink")) = ""	Then
	'	sErrorMsg = steGetText("Please enter the Hyperlink for this user right")
	Else
		' create the new user right in the database
		sStat = "UPDATE tblUserRight SET " &_
				"	ParentRightID = " & steNForm("ParentRightID") & "," &_
				"	RightName = " & steQForm("RightName") & "," &_
				"	AdminMenuName = " & steQForm("AdminMenuName") & "," &_
				"	Hyperlink = " & steQForm("Hyperlink") & "," &_
				"	AccessKey = '" & navNewKey & "'," &_
				"	HasAdd = " & steNForm("HasAdd") & "," &_
				"	HasEdit = " & steNForm("HasEdit") & "," &_
				"	HasDelete = " & steNForm("HasDelete") & "," &_
				"	HasView = " & steNForm("HasView") & "," &_
				"	Modified = " & adoGetDate & " " &_
				"WHERE	RightID = " & nRightID
		Call adoExecute(sStat)

		' force a refresh of the navigation items
		navRefreshNav True

		' re-login the user to get the new access key
		Call navLogin(Request.Cookies("AdminUserID"))
	End If
End If

sStat = "SELECT * FROM tblUserRight WHERE RightID = " & nRightID
Set rsRight = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Rights" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit User Right" %></H3>

<P>
Please make your changes to the user right using the form below.
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="userright_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="rightid" VALUE="<%= nRightID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Parent Right" %></TD><TD></TD>
	<TD>
<%
Dim oList
Set oList = New clsListInput
oList.ChooseOptionLabel = "TOP-LEVEL RIGHT"
Call oList.TreeListInput("ParentRightID", "tblUserRight", "RightID", "ParentRightID", "RightID <> " & nRightID, _
		"OrderNo", "RightID", "RightName", steRecordValue(rsRight, "ParentRightID"), "", False)
%>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Right Name" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="RightName" VALUE="<%= steRecordEncValue(rsRight, "RightName") %>" SIZE="32" MAXLENGTH="40" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Menu Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="AdminMenuName" VALUE="<%= steRecordEncValue(rsRight, "AdminMenuName") %>" SIZE="32" MAXLENGTH="40" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Hyperlink" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Hyperlink" VALUE="<%= steRecordEncValue(rsRight, "Hyperlink") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has Add?" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="HasAdd" VALUE="1" class="form"<% If CStr(steRecordValue(rsRight, "HasAdd")) = "1" Then Response.Write " CHECKED" %>> Yes
		<INPUT TYPE="radio" NAME="HasAdd" VALUE="0" class="form"<% If CStr(steRecordValue(rsRight, "HasAdd")) = "0" Then Response.Write " CHECKED" %>> No
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has Edit?" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="HasEdit" VALUE="1" class="form"<% If steRecordValue(rsRight, "HasEdit") = "1" Then Response.Write " CHECKED" %>> Yes
		<INPUT TYPE="radio" NAME="HasEdit" VALUE="0" class="form"<% If steRecordValue(rsRight, "HasEdit") = "0" Then Response.Write " CHECKED" %>> No
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has Delete?" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="HasDelete" VALUE="1" class="form"<% If steRecordValue(rsRight, "HasDelete") = "1" Then Response.Write " CHECKED" %>> Yes
		<INPUT TYPE="radio" NAME="HasDelete" VALUE="0" class="form"<% If steRecordValue(rsRight, "HasDelete") = "0" Then Response.Write " CHECKED" %>> No
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has View?" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="HasView" VALUE="1" class="form"<% If steRecordValue(rsRight, "HasView") = "1" Then Response.Write " CHECKED" %>> Yes
		<INPUT TYPE="radio" NAME="HasView" VALUE="0" class="form"<% If steRecordValue(rsRight, "HasView") = "0" Then Response.Write " CHECKED" %>> No
	</TD>
</TR><TR>
	<TD ALIGN="right" COLSPAN=3><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update User Right" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "User Right Updated" %></H3>

<P>
The user right was successfully updated in the database.  Please proceed administering
the site by using the menu shown at the top of the screen.
</P>

<% End If %>

<p align="center">
	<a href="userright_list.asp" class="adminlink"><% steTxt "User Right List" %></A>
</p>

<!-- #include file="../../../footer.asp" -->
