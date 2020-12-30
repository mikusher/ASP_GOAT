<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/class/nukemail.asp"-->
<%
'--------------------------------------------------------------------
' member_add.asp
'	Adds a new member to the database.
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
Dim sUsername
Dim sPassword
Dim sConfirm
Dim sFirstName
Dim sLastName
Dim sAddress1
Dim sAddress2
Dim sCity
Dim sStateCode
Dim sZipCode
Dim nCountryID
Dim sEmail
Dim sCode			' authorization code
Dim rsState
Dim rsCountry

sAction = steForm("Action")
sUsername = steForm("Username")
sPassword = steForm("Password")
sConfirm = steForm("Confirm")
sFirstName = steForm("FirstName")
sLastName = steForm("LastName")
sAddress1 = steForm("Address1")
sAddress2 = steForm("Address2")
sCity = steForm("City")
sStateCode = steForm("StateCode")
sZipCode = steForm("ZipCode")
nCountryID = steNForm("CountryID")
sEmail = steForm("Email")

If sAction = "ADD" Then
	' check for required fields
	If (Trim(sFirstName) = "") Then
		sErrorMsg = sErrorMsg & steGetText("Please Enter your First Name") & "<BR>"
	End If
	If (Trim(sLastName) = "") Then
		sErrorMsg = sErrorMsg & steGetText("Please Enter your Last Name") & "<BR>"
	End If
	If (Trim(sUsername) = "") Then
		sErrorMsg = sErrorMsg & steGetText("Please Enter your Username") & "<BR>"
	End If
	If Len(Trim(sUsername)) < 6 Then
		sErrorMsg = sErrorMsg & steGetText("Your username must be at least 6 characters long") & "<BR>"
	End If
	If (Trim(sPassword) = "") Then
		sErrorMsg = sErrorMsg & steGetText("Please Enter your Password") & "<BR>"
	End If
	If Len(Trim(sPassword)) < 6 Then
		sErrorMsg = sErrorMsg & steGetText("Your password must be at least 6 characters long") & "<BR>"
	End If
	If (Trim(sConfirm) = "") Then
		sErrorMsg = sErrorMsg & steGetText("Please confirm the password that you entered") & "<BR>"
	ElseIf sPassword <> sConfirm Then
		sErrorMsg = sErrorMsg & steGetText("Confirmation password does not match password") & "<BR>"
	End If
	If (Trim(sEmail) = "") Then
		sErrorMsg = sErrorMsg & steGetText("Please Enter your E-Mail Address") & "<BR>"
	End If
	If sErrorMsg = "" Then
		' check to see if the username is already in use
		query = "SELECT MemberID " &_
				"FROM	tblMember " &_
				"WHERE	Username = '" & sUsername & "'"
		Set rsMember = adoOpenRecordset(query)
		If Not rsMember.EOF Then
			sErrorMsg = sErrorMsg & steGetText("The username you entered is already taken, please choose another") & "<BR>"
		Else
			' no errors occurred - add this member to the database
			sCode = locAuthorizationCode
			query = "INSERT INTO tblMember (" &_
					"		AuthCode, Username, Password, FirstName, LastName, " &_
					"		Address1, Address2, City, StateCode, ZipCode, CountryID, " &_
					"		EMailAddress, Active, Created" &_
					") VALUES (" &_
					"'" & Replace(sCode, "'", "''") & "'," &_
					"'" & Replace(sUsername, "'", "''") & "'," &_
					"'" & SHA256(sPassword) & "'," &_
					"'" & Replace(sFirstName, "'", "''") & "'," &_
					"'" & Replace(sLastName, "'", "''") & "'," &_
					"'" & Replace(sAddress1, "'", "''") & "'," &_
					"'" & Replace(sAddress2, "'", "''") & "'," &_
					"'" & Replace(sCity, "'", "''") & "'," &_
					"'" & sStateCode & "'," &_
					"'" & Replace(sZipCode, "'", "''") & "'," &_
					nCountryID & "," &_
					"'" & Replace(sEmail, "'", "''") & "', 0, " &_
					adoGetDate &_
					")"
			Call adoExecute(query)

			' send the e-mail to activate the account
			sBody = "<P>Dear " & sFirstName & " " & sLastName & ",</P>" & vbCrLf &_
				"<P>Thank you for registering with " & Application("CompanyName") & ".  All of your information " &_
				"is private and will not be shared with any outside sources.  To " &_
				"activate your account, simply enter the URL below in your browser.</P>" & vbCrLf &_
				"<BLOCKQUOTE>" & vbCrLf &_
				"<A HREF=""" & Application("ASPNukeURL") & "account/activate.asp?authcode=" & Server.URLEncode(sCode) & """>Activate Your Account</A>" & vbCrLf &_
				"</BLOCKQUOTE>" & vbCrLf &_
				"<P>Your account will be active immediately after you click the link above. " &_
				"Thank you for registering and we hope that you find our site useful!</P>"

			Set oMail = New NukeMail
			oMail.FromAddress = Application("SupportEmail")
			oMail.FromName = Application("CompanyName") & " Registration"
			oMail.ToAddress = sEmail
			oMail.ToName = sFirstName & " " & sLastName
			oMail.Subject = Application("CompanyName") & " Registration"
			oMail.HTMLBody = sBody
			If Not oMail.Send Then
				Response.Write "<p><b class=""error"">" & oMail.ErrorMsg & "</b></p>"
			End If
		End If
	End IF
End If

' build the selection list for the states
query = "SELECT StateCode, StateName " &_
		"FROM	tblState " &_
		"WHERE	Active <> 0 " &_
		"AND	Archive = 0 " &_
		"ORDER BY StateName"
Set rsState = adoOpenRecordset(query)

' build the selection ist for the countries
query = "SELECT CountryID, CountryName " &_
		"FROM	tblCountry " &_
		"WHERE	Active <> 0 " &_
		"AND	Archive = 0 " &_
		"ORDER BY CountryName"
Set rsCountry = adoOpenRecordset(query)
%>
<!-- #include file="../../../header.asp"-->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Members" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If sAction <> "ADD" Or sErrorMsg <> "" Then %>
	<H3><% steTxt "New Member Registration" %></H3>

	<P>
	<% steTxt "Registration with" %> <%= Application("CompanyName") %>
	<% steTxt "is completely free!" %>
	<% steTxt "Please fill out your personal information below." %>
	<% steTxt "Fields marked <I>(req)</I> are required and must be filled out." %>
	<% steTxt "Once your registration is complete, an e-mail will be sent to you with instructions on how to activate your membership." %>
	</P>

	<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
	<% End If %>

	<FORM METHOD="post" ACTION="member_add.asp">

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
	<TR>
		<TD class="forml"><% steTxt "First Name" %> (req)<BR>
		<INPUT TYPE="text" NAME="FirstName" VALUE="<%= sFirstName %>" SIZE="22" MAXLENGTH="32" class="form">
		</TD>
		<TD class="forml"><% steTxt "Last Name" %> (req)<BR>
		<INPUT TYPE="text" NAME="LastName" VALUE="<%= sLastName %>" SIZE="22" MAXLENGTH="32" class="form">
		</TD>
	</TR>
	<TR>
		<TD class="forml"><% steTxt "Username" %> (req)<BR>
		<INPUT TYPE="text" NAME="Username" VALUE="<%= sUserName %>" SIZE="22" MAXLENGTH="32" class="form">
		</TD>
	</TR><TR>
		<TD class="forml"><% steTxt "Password" %> (req)<BR>
		<INPUT TYPE="password" NAME="Password" VALUE="<%= sPassword %>" SIZE="22" MAXLENGTH="32" class="form">
		</TD>
		<TD class="forml"><% steTxt "Confirm Password" %> (req)<BR>
		<INPUT TYPE="password" NAME="Confirm" VALUE="<%= sConfirm %>" SIZE="22" MAXLENGTH="32" class="form">
		</TD>
	</TR>
	<TR>
		<TD class="forml" COLSPAN=2><% steTxt "Address (Line 1)" %><BR>
		<INPUT TYPE="text" NAME="Address1" VALUE="<%= sAddress1 %>" SIZE="52" MAXLENGTH="40" class="form">
		</TD>
	</TR>
	<TR>
		<TD class="forml" COLSPAN=2><% steTxt "Address (Line 2)" %><BR>
		<INPUT TYPE="text" NAME="Address2" VALUE="<%= sAddress2 %>" SIZE="52" MAXLENGTH="40" class="form">
		</TD>
	</TR>
		<TR>
		<TD COLSPAN=2>
		<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 WIDTH="100%">
		<TR>
			<TD class="forml"><% steTxt "City" %><BR>
			<INPUT TYPE="text" NAME="City" VALUE="<%= sCity %>" SIZE="14" MAXLENGTH="32" class="form">
			</TD>
			<TD class="forml"><% steTxt "State" %><BR>
			<SELECT NAME="StateCode" class="form">
			<OPTION VALUE=""> -- Choose --
			<% Do Until rsState.EOF %>
			<OPTION VALUE="<%= rsState.Fields("StateCode").Value %>"<% If sStateCode = rsState.Fields("StateCode").Value Then Response.Write(" SELECTED") %>> <%= rsState.Fields("StateName").Value %>
			<%	rsState.MoveNext
			   Loop %>
			</SELECT>
			</TD>
			<TD class="forml"><% steTxt "Zip Code" %><BR>
			<INPUT TYPE="text" NAME="ZipCode" VALUE="<%= sZipCode %>" SIZE="10" MAXLENGTH="10" class="form">
			</TD>
		</TR>
		</TABLE>
		</TD>
	</TR>
	<TR>
		<TD class="forml" COLSPAN=2><% steTxt "Country" %><BR>
		<SELECT NAME="CountryID" class="form">
		<OPTION VALUE="0"> -- Choose --
		<% Do Until rsCountry.EOF %>
		<OPTION VALUE="<%= rsCountry.Fields("CountryID").Value %>"<% If nCountryID = rsCountry.Fields("CountryID").Value Then Response.Write(" SELECTED") %>> <%= rsCountry.Fields("CountryName").Value %>
		<%	rsCountry.MoveNext
		   Loop %>
		</SELECT>
		</TD>
	</TR>
	<TR>
		<TD class="forml" COLSPAN=2><% steTxt "E-Mail Address" %> (req)<BR>
		<INPUT TYPE="text" NAME="Email" VALUE="<%= sEmail %>" SIZE="52" MAXLENGTH="80" class="form">
		</TD>
	</TR>
	<TR>
		<TD COLSPAN=2 ALIGN="right"><br>
		<INPUT TYPE="hidden" NAME="Action" VALUE="ADD">
		<INPUT TYPE="submit" NAME="_dummy" VALUE=" <% steTxt "Register" %> " class="form">
		</TD>
	</TR>
	</TABLE>
	</P>

	</FORM>

	<P>
	<FONT CLASS="small"><% steTxt "REQ - Indicates that a form field is required" %></FONT>
	</P>
<% Else %>
	<H3><% steTxt "Registration Complete" %></H3>

	<P>
	<% steTxt "The registration of the new member is now complete." %>
	<% steTxt "An e-mail was sent to the member's e-mail address asking them to verify the new account by activating it using a special code." %>
	<% steTxt "To continue administering the site, please use the admin links at the top of the screen." %>
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
%>