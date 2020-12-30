<!-- #include file="../../../lib/site_lib.asp"-->
<!-- #include file="../../../lib/sha256.asp" -->
<%
'--------------------------------------------------------------------
' change_password.asp
'	Change a member's password.
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

Dim query
Dim sAction
Dim sUsername
Dim sPassword
Dim sConfirm
Dim sCode
Dim sErrorMsg
Dim sStatusMsg
Dim sEmail

sAction = steForm("Action")
sUsername = steForm("Username")
sPassword = steForm("Password")
sConfirm = steForm("Confirm")
sEmail = steForm("Email")
sCode = steForm("Code")

If sAction = "SEND" Then
	' check for required fields
	If Trim(sPassword) = "" Then
		sErrorMsg = sErrorMsg & steGetText("Please Enter your New Password") & "<BR>"
	End If
	If Trim(sConfirm) = "" Then
		sErrorMsg = sErrorMsg & steGetText("Please Enter your Confirm Password") & "<BR>"
	End If
	If sPassword <> sConfirm Then
		sErrorMsg = sErrorMsg & steGetText("Your Password and Confirmation do not match") & "<BR>"
	End If
	If sErrorMsg = "" Then
		' check to see if the username is already in use
		query = "SELECT MemberID, FirstName, LastName, Username, Password " &_
				"FROM	tblMember " &_
				"WHERE	Username = " & steQForm("Username") & " " &_
				"AND	AuthCode = '" & sCode & "' " &_
				"AND	Active <> 0 " &_
				"AND	Archive = 0"	
		Set rsMember = adoOpenRecordset(query)
		If rsMember.EOF Then
			sErrorMsg = sErrorMsg & steGetText("The password change request could not be located") & "<BR>"
		Else
			' no errors occurred - reset the member's password
			query = "UPDATE	tblMember " &_
					"SET	Password = '" & SHA256(steForm("password")) & "' " &_
					"WHERE	Username = " & steQForm("Username") & " " &_
					"AND	AuthCode = '" & sCode & "' " &_
					"AND	Active <> 0 " &_
					"AND	Archive = 0"
			Call adoExecute(query)

			' login the member to the site
			Response.Cookies("MemberID") = rsMember.Fields("MemberID").Value
			Response.Cookies("FullName") = rsMember.Fields("FirstName").Value & " " & rsMember.Fields("LastName").Value
			Response.Cookies("Username") = rsMember.Fields("Username").Value

			sStatusMsg = steGetText("Now logged in as member") & " " & sUsername & "<BR>"

		End If
		rsMember.Close
		rsMember = Empty
	End IF
End If
%>
<!-- #include file="../../../header.asp"-->
<% If sAction <> "SEND" Or sErrorMsg <> "" Then %>
	<H3><% steTxt "Change Password" %></H3>

	<P>
	Please enter your e-mail registered with <%= Application("CompanyName") %>
	and we will send you your password.
	</P>

	<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
	<% End If %>

	<FORM METHOD="post" ACTION="change_password.asp">
	<input type="hidden" name="username" value="<%= steEncForm("Username") %>">
	<input type="hidden" name="code" value="<%= Server.HTMLEncode(sCode) %>">

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
	<TR>
		<TD class="forml">
		<% steTxt "Username" %><BR>
		<font class="formd"><%= sUsername %></font>
		</TD>
	</TR><TR>
		<TD class="forml">
		<% steTxt "New Password" %> (req)<BR>
		<INPUT TYPE="password" NAME="password" VALUE="<%= steEncForm("Password") %>" SIZE="22" MAXLENGTH="32" class="form">
		</TD>
	</TR><TR>
		<TD class="forml">
		<% steTxt "Confirm Password" %> (req)<BR>
		<INPUT TYPE="password" NAME="confirm" VALUE="<%= steEncForm("Confirm") %>" SIZE="22" MAXLENGTH="32" class="form">
		</TD>
	</TR>
	<TR>
		<TD COLSPAN=2 ALIGN="center">
		<INPUT TYPE="hidden" NAME="Action" VALUE="SEND">
		<INPUT TYPE="submit" NAME="_dummy" VALUE=" <% steTxt "Change Password" %> " class="form">
		</TD>
	</TR>
	</TABLE>
	</P>

	</FORM>

	<P>
	<FONT CLASS="small">(req) - <% steTxt "Indicates that a form field is required" %></FONT>
	</P>
<% Else %>
	<H3><% steTxt "Password Changed!" %></H3>

	<% If sStatusMsg <> "" Then %>
	<p><B><%= sStatusMsg %></B></p>
	<% End If %>
	<P>
	Your password has been reset for your member account.
	Please write down your password and keep it in a safe place.  This is
	the username you will use to login to <%= Application("CompanyName") %>
	from now on.
	</P>
<% End If %>

<!-- #include file="../../../footer.asp"-->