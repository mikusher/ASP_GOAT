<!-- #include file="../../../lib/site_lib.asp"-->
<!-- #include file="../../../lib/sha256.asp" -->
<!-- #include file="../../../lib/class/nukemail.asp"-->
<%
'--------------------------------------------------------------------
' forgot_password.asp
'	Send an e-mail reminder of a member's password.
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
Dim sErrorMsg
Dim sCode
Dim sEmail

sAction = steForm("Action")
sEmail = steForm("Email")

If sAction = "SEND" Then
	' check for required fields
	If (Trim(sEmail) = "") Then
		sErrorMsg = sErrorMsg & steGetText("Please Enter your E-Mail Address") & "<BR>"
	ElseIf Not IsValidEmail(Trim(sEmail)) Then
		sErrorMsg = sErrorMsg & steGetText("The E-Mail you entered is not a valid address") & "<BR>"
	End If
	If sErrorMsg = "" Then
		' check to see if the username is already in use
		query = "SELECT MemberID, FirstName, LastName, Username, Password " &_
				"FROM	tblMember " &_
				"WHERE	EmailAddress = " & steQForm("Email")
		Set rsMember = adoOpenRecordset(query)
		If rsMember.EOF Then
			sErrorMsg = sErrorMsg & steGetText("The e-mail account specified does not exist") & "<BR>"
		Else
			Dim sURL
			' no errors occurred - resend the member's password
			sCode = locAuthorizationCode
			query = "UPDATE tblMember " &_
					"SET AuthCode = '" & sCode & "' " &_
					"WHERE	MemberID = " & rsMember.Fields("MemberID").Value & " " &_
					"AND	Active <> 0 " &_
					"AND	Archive = 0"
			Call adoExecute(query)

			' send the e-mail to activate the account
			sURL = Application("ASPNukeURL") & "module/account/register/change_password.asp" & "?username=" & Server.URLEncode(rsMember.Fields("Username").Value) & "&code=" & sCode
			sBody = "<P>Dear " & rsMember.Fields("FirstName").Value & " " & rsMember.Fields("LastName").Value & ",</P>" & vbCrLf &_
				"<P>As you requested through our <i>forgot password</i> feature, here are the " &_
				"instructions you will need to reset your password registered with " & Application("CompanyName") & ".</P>" & vbCrLf &_
				"<P>Click or enter the web address below and enter your new password in the web form displayed " &_
				"and your account will be updated and your account will be logged-in.</P>" & vbCrLf &_
				"<p><A HREF=""" & sURL & """>" & sURL & "</A></p>" & vbCrLf &_
				"<P>Thank you for visiting " & Application("CompanyName") & ".</P>"

			Set oMail = New NukeMail
			oMail.FromAddress = Application("SupportEmail")
			oMail.FromName = Application("CompanyName") & " Support"
			oMail.ToAddress = sEmail
			oMail.ToName = rsMember.Fields("FirstName").Value & " " & rsMember.Fields("LastName").Value
			oMail.Subject = Application("CompanyName") & " Info Request"
			oMail.HTMLBody = sBody
			If Not oMail.Send Then
				Response.Write "<p><b class=""error"">" & oMail.ErrorMsg & "</b></p>"
			End If
		End If
	End IF
End If
%>
<!-- #include file="../../../header.asp"-->
<% If sAction <> "SEND" Or sErrorMsg <> "" Then %>
	<H3><% steTxt "Password Change Request" %></H3>

	<P>
	Please enter your e-mail registered with <%= Application("CompanyName") %>
	and we will send you instructions on how you may reset your password.
	It is not possible to retrieve your password because it is stored
	encrypted in our database.
	</P>

	<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
	<% End If %>

	<FORM METHOD="post" ACTION="forgot_password.asp">

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
	<TR>
		<TD class="forml">
		<% steTxt "E-Mail" %> (req)<BR>
		<INPUT TYPE="text" NAME="email" VALUE="<%= sEmail %>" SIZE="22" MAXLENGTH="80" class="form">
		</TD>
	</TR>
	<TR>
		<TD COLSPAN=2 ALIGN="center">
		<INPUT TYPE="hidden" NAME="Action" VALUE="SEND">
		<INPUT TYPE="submit" NAME="_dummy" VALUE=" <% steTxt "Send Change Request" %> " class="form">
		</TD>
	</TR>
	</TABLE>
	</P>

	</FORM>

	<P>
	<FONT CLASS="small">(req) - <% steTxt "Indicates that a form field is required" %></FONT>
	</P>
<% Else %>
	<H3><% steTxt "Password Change Request Sent!" %></H3>

	<P>
	Thank you for your password change request for <%= Application("CompanyName") %>.
	Please allow some time for this password request to reach your mailbox. If the
	e-mail address you registered with no longer exists, simply re-register your
	account with us using a different e-mail address.
	</P>
<% End If %>

<!-- #include file="../../../footer.asp"-->
<%
' generate an authorization code for this user
Function locAuthorizationCode
	Dim sCode, sChar, I

	Randomize
	For I = 1 To 20
		If (I <> 5 And I <> 16) Then
			sChar = Int(36 * Rnd())
			If sChar < 10 Then sCode = sCode & Chr(sChar + 48) Else sCode = sCode & Chr(sChar + 55)
		Else
			sCode = sCode & "-"
		End If
	Next
	locAuthorizationCode = sCode
End Function

' check to see if the supplied e-mail address is valid
' RETURNS: True if it is valid, False otherwise
Function IsValidEmail(sEmail)
	Dim oRE

	' make sure a subject and body are present
	Set oRE = New RegExp
	oRE.Pattern = "(@.*@)|(\.\.)|(@\.)|(^\.)"
	If Not oRE.Test(sEmail) Then
		oRE.Pattern = "^.+\@(\[?)[a-zA-Z0-9\-\.]+\.([a-zA-Z]{2,3}|[0-9]{1,3})(\]?)$"
		If Not oRE.Test(sEmail) Then
			IsValidEmail = False
			Exit Function
		End If
	End If
	IsValidEmail = True
End Function
%>