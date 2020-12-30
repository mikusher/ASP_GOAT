<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/sha256.asp" -->
<%
'--------------------------------------------------------------------
' login.asp
'	Perform a login of a member to our site.
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

' adoDebug = True
Dim sStat
Dim rsMember
Dim sErrorMsg
Dim sStatusMsg
Dim bLoginSuccess	' user is logged in successfully
Dim bLogoffSuccess	' user has logged off successfully

bLoginSuccess = false
sStatusMsg = Request.QueryString("error")

If steForm("action") = "login" Then
	' check the form variables
	If Trim(steForm("username")) = "" Then
		sErrorMsg = steGetText("Please enter your username") & "<BR>"
	ElseIf Trim(steForm("password")) = "" Then
		sErrorMsg = steGetText("Please enter your password") & "<BR>"
	Else
		' retrieve the user information here
		sStat = "SELECT	MemberID, FirstName, LastName, Username " & _
				"FROM	tblMember " &_
				"WHERE	Username = " & steQForm("username") & " " &_
				"AND	Password = '" & SHA256(steForm("password")) & "'"
		Set rsMember = adoOpenRecordset(sStat)
		If Not rsMember.EOF Then
			' login the user
			Response.Cookies("MemberID") = rsMember.Fields("MemberID").Value
			Response.Cookies("FullName") = rsMember.Fields("FirstName").Value & " " & rsMember.Fields("LastName").Value
			Response.Cookies("Username") = rsMember.Fields("Username").Value
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
ElseIf steForm("action") = "logoff" Then
	Response.Cookies("MemberID").Expires = Now()
	Response.Cookies("Username").Expires = Now()
	' logoff admin also (if defined)
	Response.Cookies("AdminUserID").Expires = Now()
	Response.Cookies("AdminUsername").Expires = Now()
	Response.Cookies("AdminFullname").Expires = Now()
	bLogoffSuccess = True
End If
%>
<!-- #include file="../../../header.asp" -->

<% If bLogoffSuccess Then %>

<H3><% steTxt "Member Logoff" %></H3>

<P>
You have successfully logged off of <%= Application("CompanyName") %>.
For added security, you should close all your browser windows from
which you accessed our web site or shut down your computer.  Thank
you for visiting us today.
</P>

<% ElseIf Not bLoginSuccess Then %>

<H3><% steTxt "Member Login" %></H3>

<P>
Welcome to the member login page.  This will allow you to gain access to
the member-only ares of <%= Application("CompanyName") %>.
Please enter your username and login below.
</P>

<% If sErrorMsg <> "" Then %>
	<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% ElseIf sStatusMsg <> "" Then %>
	<P><B><%= sStatusMsg %></B></P>
<% End If %>

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
<% steTxt "You have successfully logged onto the site.  You now have access to all of the member areas for our site." %>
</P>

<% End If %>

<!-- #include file="../../../footer.asp" -->