<!-- #include file="../../lib/site_lib.asp" -->
<!-- #include file="../../lib/sha256.asp" -->
<%
'--------------------------------------------------------------------
' login.asp
'	Stand-alone login  script for the nuke administration.
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

Dim sStat
Dim rsUser
Dim sErrorMsg
Dim bLoginSuccess	' user is logged in successfully

bLoginSuccess = false

If steForm("action") <> "" Then
	' check the form variables
	If Trim(steForm("username")) = "" Then
		sErrorMsg = steGetText("Please enter your username")
	ElseIf Trim(steForm("password")) = "" Then
		sErrorMsg = steGetText("Plesae enter your password")
	Else
		' retrieve the user information here
		sStat = "SELECT	UserID, FirstName, LastName, Username " & _
				"FROM	tblUser " &_
				"WHERE	Username = " & steQForm("username") & " " &_
				"AND	Password = '" & Replace(SHA256(steForm("password")), "'", "''") & "'"
		Set rsUser = adoOpenRecordset(sStat)
		If Not rsUser.EOF Then
			' login the user
			Response.Cookies("AdminUserID") = rsUser.Fields("UserID").Value
			Response.Cookies("AdminUsername") = rsUser.Fields("Username").Value
			Response.Cookies("AdminFullname") = rsUser.Fields("Firstname").Value & " " & rsUser.Fields("LastName").Value
			bLoginSuccess = True

			' redirect to the login here (if nec)
			If steForm("goto") <> "" Then
				Response.Redirect(steForm("goto"))
			End If
		Else
			' display error message to user here
			sErrorMsg = steGetText("Username and Password are invalid")
		End If
	End If
End If
%>
<!-- #include file="../../header.asp" -->

<% If Not bLoginSuccess Then %>

<H3><% steTxt "Admin Login" %></H3>

<P>
<% steGetText "Please enter your username and login below to administer the site." %>
<% steGetText "Only users with the proper access permissions may access the admin area." %>
</P>

<FORM METHOD="post" ACTION="login.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="login">
<INPUT TYPE="hidden" NAME="goto" VALUE="<%= steEncForm("goto") %>">
<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD class="forml"><% steTxt "Username" %><BR>
	<INPUT TYPE="text" NAME="username" VALUE="<%= steEncForm("username") %>" SIZE="16" MAXLENGTH="16" class="form">
	</TD>
</TR><TR>
	<TD class="forml"><% steTxt "Password" %><BR>
	<INPUT TYPE="password" NAME="password" VALUE="" SIZE="16" MAXLENGTH="16" class="form">
	</TD>
</TR><TR>
	<TD ALIGN="center"><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Login" %> " class="form">
	</TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Login Successful" %></H3>

<P>
<% steTxt "You have successfully logged onto the site." %>&nbsp;
<% steTxt "You now have access to the site administration tools made available to you." %>
</P>

<% End If %>

<!-- #include file="../../footer.asp" -->