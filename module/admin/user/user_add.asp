<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->
<%
'--------------------------------------------------------------------
' user_add.asp
'	Add a new admin user to the database
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
Dim rsUser

If steForm("action") = "add" Then
	' make sure the required fields are present
	If Trim(steForm("username")) = "" Then
		sErrorMsg = steGetText("Please enter a username for this new admin user")
	ElseIf Trim(steForm("password")) = "" Then
		sErrorMsg = steGetText("Please enter a password for this new admin user")
	ElseIf Trim(steForm("cpassword")) = "" Then
		sErrorMsg = steGetText("Please enter the confirmation password for this new admin user")
	ElseIf steForm("password") <> steForm("cpassword") Then
		sErrorMsg = steGetText("Password and Confirmation do not match, please try again")
	ElseIf Trim(steForm("firstname")) = ""	Then
		sErrorMsg = steGetText("Please enter the first name for this new admin user")
	ElseIf Trim(steForm("lastname")) = "" Then
		sErrorMsg = steGetText("Please enter the last name for this new admin user")
	Else
		' create the new user in the database
		sStat = "INSERT INTO tblUser (" &_
				"	Username, Password, FirstName, MiddleName, LastName, EmailAddress, " &_
				"	DayPhone, EvePhone, Created" &_
				") VALUES (" &_
				steQForm("username") & ",'" & SHA256(steForm("password")) & "'," &_
				steQForm("FirstName") & "," & steQForm("MiddleName") & "," &_
				steQForm("lastName") & "," & steQForm("EmailAddress") & "," &_
				steQForm("DayPhone") & "," & steQForm("EvePhone") & "," &_
				adoGetDate &_
				")"
		Call adoExecute(sStat)
	End If
End If
%>

<% sCurrentTab = "Users" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New Admin User" %></H3>

<P>
<% steTxt "Please enter the new properties for the new admin using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="user_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD class="forml"><% steTxt "Username" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="Username" VALUE="<%= steEncForm("Username") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Password" %></TD><TD></TD>
	<TD><INPUT TYPE="password" NAME="Password" VALUE="<%= steEncForm("Password") %>" SIZE="32" MAXLENGTH="64" class="form"></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Confirm Password" %></TD><TD></TD>
	<TD><INPUT TYPE="password" NAME="CPassword" VALUE="<%= steEncForm("CPassword") %>" SIZE="32" MAXLENGTH="64" class="form"></TD>
</TR><TR>
	<TD class="forml"><% steTxt "First Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="FirstName" VALUE="<%= steEncForm("FirstName") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Middle Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="MiddleName" VALUE="<%= steEncForm("MiddleName") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Last Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="LastName" VALUE="<%= steEncForm("LastName") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Email Address" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="EmailAddress" VALUE="<%= steEncForm("EmailAddress") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Daytime Phone" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="DayPhone" VALUE="<%= steEncForm("DayPhone") %>" SIZE="32" MAXLENGTH="16" class="form"></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Evening Phone" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="EvePhone" VALUE="<%= steEncForm("EvePhone") %>" SIZE="32" MAXLENGTH="16" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add User" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "New Admin User Added" %></H3>

<P>
<% steTxt "The new admin user has been added to the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<p align="center">
	<a href="user_list.asp" class="adminlink"><% steTxt "User List" %></a> &nbsp;
	<a href="user_add.asp" class="adminlink"><% steTxt "Add Another" %></a>
</p>

<% End If %>

<!-- #include file="../../../footer.asp" -->
