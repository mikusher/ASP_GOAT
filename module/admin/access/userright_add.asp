<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/class/form/listinput.asp" -->
<%
'--------------------------------------------------------------------
' userright_add.asp
'	Add a new user right to the database
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

If steForm("action") = "add" Then
	' make sure the required fields are present
	If Trim(steForm("RightName")) = ""	Then
		sErrorMsg = steGetText("Please enter the Right Name for this user right")
	' ElseIf Trim(steForm("Hyperlink")) = ""	Then
	'	sErrorMsg = steGetText("Please enter the Hyperlink for this user right")
	Else
		' determine the order no
		sStat = "SELECT Coalesce(Max(OrderNo) + 1, 1) As OrderNo FROM tblUserRight"
		Set rsOrder = adoOpenRecordset(sStat)
		If Not rsOrder.EOF Then nOrderNo = rsOrder.Fields("OrderNo").Value Else nOrderNo = 1
		rsOrder.Close
		Set rsOrder = Nothing

		' create the new user right in the database
		sStat = "INSERT INTO tblUserRight (" &_
				"	ParentRightID, OrderNo, RightName, AdminMenuName, Hyperlink, AccessKey, " &_
				"	HasAdd, HasEdit, HasDelete, HasView, Created " &_
				") VALUES (" &_
				steNForm("ParentRightID") & "," &_
				nOrderNo & "," &_
				steQForm("RightName") & "," &_
				steQForm("AdminMenuName") & "," &_
				steQForm("Hyperlink") & "," &_
				"'" & navNewKey & "'," &_
				steNForm("HasAdd") & "," &_
				steNForm("HasEdit") & "," &_
				steNForm("HasDelete") & "," &_
				steNForm("HasView") & "," &_
				adoGetDate &_
				")"
		Call adoExecute(sStat)

		' force a refresh of the navigation items
		navRefreshNav True
	End If
End If
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Rights" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New User Right" %></H3>

<P>
Please enter the new properties for the new user right using the form below.
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="userright_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Parent Right" %></TD><TD></TD>
	<TD>
<%
Dim oList
Set oList = New clsListInput
oList.ChooseOptionLabel = "TOP-LEVEL RIGHT"
Call oList.TreeListInput("ParentRightID", "tblUserRight", "RightID", "ParentRightID", "", _
		"OrderNo", "RightID", "RightName", steEncForm("ParentRightID"), "", False)
%>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Right Name" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="RightName" VALUE="<%= steEncForm("RightName") %>" SIZE="32" MAXLENGTH="40" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Menu Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="AdminMenuName" VALUE="<%= steEncForm("AdminMenuName") %>" SIZE="32" MAXLENGTH="40" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Hyperlink" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Hyperlink" VALUE="<%= steEncForm("Hyperlink") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has Add?" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="HasAdd" VALUE="1" class="form"<% If CStr(steEncForm("HasAdd")) = "1" Then Response.Write " CHECKED" %>> Yes
		<INPUT TYPE="radio" NAME="HasAdd" VALUE="0" class="form"<% If CStr(steEncForm("HasAdd")) = "0" Then Response.Write " CHECKED" %>> No
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has Edit?" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="HasEdit" VALUE="1" class="form"<% If steEncForm("HasEdit") = "1" Then Response.Write " CHECKED" %>> Yes
		<INPUT TYPE="radio" NAME="HasEdit" VALUE="0" class="form"<% If steEncForm("HasEdit") = "0" Then Response.Write " CHECKED" %>> No
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has Delete?" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="HasDelete" VALUE="1" class="form"<% If steEncForm("HasDelete") = "1" Then Response.Write " CHECKED" %>> Yes
		<INPUT TYPE="radio" NAME="HasDelete" VALUE="0" class="form"<% If steEncForm("HasDelete") = "0" Then Response.Write " CHECKED" %>> No
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has View?" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="HasView" VALUE="1" class="form"<% If steEncForm("HasView") = "1" Then Response.Write " CHECKED" %>> Yes
		<INPUT TYPE="radio" NAME="HasView" VALUE="0" class="form"<% If steEncForm("HasView") = "0" Then Response.Write " CHECKED" %>> No
	</TD>
</TR><TR>
	<TD ALIGN="right" COLSPAN=3><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add User Right" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "New User Right Added" %></H3>

<P>
The new user right has been added to the database.  Please proceed administering
the site by using the menu shown at the top of the screen.
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="userright_list.asp" class="adminlink"><% steTxt "User Right List" %></A>
</P>

<!-- #include file="../../../footer.asp" -->