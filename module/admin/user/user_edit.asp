<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' user_edit.asp
'	Modify an existing admin user from the database
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
Dim sPassword		' SQL to modify password
Dim rsUser

If steForm("action") = "update" Then
	' make sure the required fields are present
	If Trim(steForm("username")) = "" Then
		sErrorMsg = steGetText("Please enter a username for this admin user")
	ElseIf Trim(steForm("password")) = "" And Trim(steForm("cpassword")) <> "" Then
		sErrorMsg = steGetText("Please enter the password for this admin user")
	ElseIf Trim(steForm("cpassword")) = "" And Trim(steForm("password")) <> "" Then
		sErrorMsg = steGetText("Please enter the confirmation password for this admin user")
	ElseIf steForm("password") <> steForm("cpassword") Then
		sErrorMsg = steGetText("Password and Confirmation do not match, please try again")
	ElseIf Trim(steForm("firstname")) = ""	Then
		sErrorMsg = steGetText("Please enter the first name for this new admin user")
	ElseIf Trim(steForm("lastname")) = "" Then
		sErrorMsg = steGetText("Please enter the last name for this new admin user")
	Else
		' create the new user in the database
		If Trim(steForm("password")) <> "" Then
			sPassword = "		Password = '" & SHA256(steForm("password")) & "',"
		End If
		sStat = "UPDATE tblUser SET " &_
				"		Username = " & steQForm("username") & "," &_
				sPassword &_
				"		FirstName = " & steQForm("FirstName") & "," &_
				"		MiddleName = " & steQForm("MiddleName") & "," &_
				"		LastName = " & steQForm("lastName") & "," &_
				"		EmailAddress = " & steQForm("EmailAddress") & "," &_
				"		DayPhone = " & steQForm("DayPhone") & "," &_
				"		EvePhone = " & steQForm("EvePhone") & "," &_
				"		Modified = " & adoGetDate & " " &_
				"WHERE	UserID = " & steForm("UserID")
		Call adoExecute(sStat)
	End If
End If

' retrieve the user information to edit (if nec)
If steForm("action") <> "update" Or sErrorMsg <> "" Then
	sStat = "SELECT * FROM tblUser WHERE UserID = " & steForm("UserID")
	Set rsUser = adoOpenRecordset(sStat)
End If
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Users" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "update" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Admin User" %></H3>

<P>
<% steTxt "Please make your changes to the user record in the form below." %>&nbsp;
<% steTxt "If you don't wish to change this user's password, you should leave the <I>Password</I> and <I>Confirm Password</I> fields blank." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="user_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="update">
<INPUT TYPE="hidden" NAME="UserID" VALUE="<%= steForm("UserID") %>">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD CLASS="forml"><% steTxt "Username" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="Username" VALUE="<%= steRecordEncValue(rsUser, "Username") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Password" %></TD><TD></TD>
	<TD><INPUT TYPE="password" NAME="Password" VALUE="" SIZE="32" MAXLENGTH="64" class="form"></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Confirm Password" %></TD><TD></TD>
	<TD><INPUT TYPE="password" NAME="CPassword" VALUE="" SIZE="32" MAXLENGTH="64" class="form"></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "First Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="FirstName" VALUE="<%= steRecordEncValue(rsUser, "FirstName") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Middle Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="MiddleName" VALUE="<%= steRecordEncValue(rsUser, "MiddleName") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Last Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="LastName" VALUE="<%= steRecordEncValue(rsUser, "LastName") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Email Address" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="EmailAddress" VALUE="<%= steRecordEncValue(rsUser, "EmailAddress") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Daytime Phone" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="DayPhone" VALUE="<%= steRecordEncValue(rsUser, "DayPhone") %>" SIZE="32" MAXLENGTH="16" class="form"></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Evening Phone" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="EvePhone" VALUE="<%= steRecordEncValue(rsUser, "EvePhone") %>" SIZE="32" MAXLENGTH="16" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR>
		<INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Modify User" %> " class="form">
	</TD>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Admin User Modified" %></H3>

<P>
<% steTxt "The admin user has been modified according to the changes you requested." %>
</P>

<% End If %>

<!-- #include file="../../../footer.asp" -->
