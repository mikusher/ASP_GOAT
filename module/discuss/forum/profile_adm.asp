<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' profile_adm.asp
'	create or modify a user profile
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

Dim sAction
Dim sStat
Dim nProfileID
Dim nMemberID
Dim rsEdit
Dim sStatusMsg
Dim sErrorMsg

sAction = Request.Form("action")
nMemberID = Request("MemberID")
sUsername = Request.Form("Username")
sPassword = Request.Form("Password")
sConfirm = Request.Form("Confirm")
nProfileID = Request.Form("ProfileID")
If IsNumeric(nMemberID) And nMemberID <> "" Then nMemberID = CInt(nMemberID) Else nMemberID = 0
If IsNumeric(nProfileID) And nProfileID <> "" Then nProfileID = CInt(nProfileID) Else nProfileID = 0

' validate the form first
If UCase(sAction) = "CREATE" Or UCase(sAction) = "UPDATE" Then
	If Trim(sUsername) = "" Then
		sErrorMsg = sErrorMsg & steGetText("Please enter your username") & "<BR>"
	End If
	If sPassword <> Trim(sPassword) Then
		sErrorMsg = sErrorMsg & steGetText("Invalid whitespace characters found in password") & "<BR>"
	End If
	If sPassword <> sConfirm Then
		sErrorMsg = sErrorMsg & steGetText("Your password onfirmation doesn't match") & "<BR>"
	End If
End If

If sErrorMsg = "" Then
If UCase(sAction) = "CREATE" Then
	' create a new forum profile
	sStat = "INSERT INTO tblMessageProfile (" &_
			"Username, Password, Location, Email, ForumIcon, " &_
			"ShowEmail, Biography, ThemeID, Created" &_
			") VALUES (" &_
			"'" & Replace(sUsername, "'", "''") & "'," &_
			"'" & Replace(sPassword, "'", "''") & "'," &_
			"'" & Replace(sLocation, "'", "''") & "'," &_
			"'" & Replace(sEmail, "'", "''") & "'," &_
			"'" & Replace(sForumIcon, "'", "''") & "'," &_
			"'" & Replace(sShowEmail, "'", "''") & "'," &_
			"'" & Replace(sBiography, "'", "''") & "'," &_
			nThemeID & "," & adoGetDate &_
			")"
	Call adoExecute(sStat)

	' get the newly generated profile ID
	sStat = "SELECT Max(ProfileID) AS ProfileID FROM tblMessageProfile"
	Set rsProfile = adoOpenRecordset(sStat)
	If Not rsProfile.EOF Then nProfileID = rsProfile.Fields("ProfileID").Value
	Set rsProfile = Nothing

	sStatusMsg = steGetText("Your profile was created successfully")

ElseIf UCase(sAction) = "UPDATE" Then
	' update an existing forum profile
	sStat = "UPDATE tblMessageProfile SET " &_
			"Username = '" & Replace(sUsername, "'", "''") & "'," &_
			"Password = '" & Replace(sPassword, "'", "''") & "'," &_
			"Location = '" & Replace(sLocation, "'", "''") & "'," &_
			"Email = '" & Replace(sEmail, "'", "''") & "'," &_
			"ForumIcon = '" & Replace(sForumIcon, "'", "''") & "'," &_
			"ShowEmail = '" & Replace(sShowEmail, "'", "''") & "'," &_
			"Biography = '" & Replace(sBiography, "'", "''") & "'," &_
			"ThemeID = " & nThemeID & " " &_
			"WHERE	ProfileID = " & nProfileID
	Call adoExecute(sStat)
End If
End If

' retrieve the profile to edit
If nProfileID > 0 Then
	sStat = "SELECT * FROM tblMessageProfile " &_
			"WHERE	ProfileID = " & nProfileID
	Set rsEdit = adoOpenRecordset(sStat)
ElseIf nMemberID > 0 Then
	sStat = "SELECT * FROM tblMessageProfile " &_
			"WHERE	MemberID = " & nMemberID
	Set rsEdit = adoOpenRecordset(sStat)
End If

%>
<!-- #include file="../../../header.asp" -->

<h3><% steTxt "Forum Registration" %></h3>

<p>
<% steTxt "Use this form to create an account necessary for posting in our forums." %>&nbsp;
<% steTxt "If you have already registered with" %>&nbsp;<%= Application("CompanyName") %>,
<% steTxt "you can simply use your" %>&nbsp;<%= Application("CompanyName") %>&nbsp;<% steTxt "login to access our forums." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B class="error"><%= sErrorMsg %></b></p>
<% ElseIf sStatusMsg <> "" Then %>
<P><B class="error"><%= sStatusMsg %></b></p>
<% End If %>

<form method="post" action="profile_adm.asp">
<input type="hidden" name="profileid" value="<%= nProfileID %>">
<input type="hidden" name="memberid" value="<%= nMemberID %>">

<table border=0 cellpadding=2 cellspacing=0>
<tr>
	<td class="forml"><% steTxt "Username" %><br>
	<input type="text" name="username" value="<%= steRecordEncValue(rsEdit, "Username") %>" size="20" maxlength="20" class="form"></td>
	<td class="forml"><% steTxt "Password" %><br>
	<input type="password" name="password" value="" size="20" maxlength="20" class="form"></td>
</tr><tr>
	<td></td>
	<td class="forml"><% steTxt "Confirm Password" %><br>
	<input type="password" name="confirm" value="" size="20" maxlength="20" class="form"></td>
</tr><tr>
	<td class="forml"><% steTxt "Location" %><br>
	<input type="text" name="location" value="<%= steRecordEncValue(rsEdit, "Location") %>" size="20" maxlength="50" class="form"></td>
	<td class="forml"><% steTxt "E-Mail Address" %><br>
	<input type="text" name="email" value="<%= steRecordEncValue(rsEdit, "Email") %>" size="20" maxlength="100" class="form"></td>
</tr><tr>
	<td colspan=2 class="forml"><% steTxt "Forum Icon" %><br>
	<input type="text" name="ForumIcon" value="<%= steRecordEncValue(rsEdit, "ForumIcon") %>" size="40" maxlength="100" class="form"></td>
</tr><tr>
	<td class="forml">Show E-Mail?<br>
	<input type="radio" name="ShowEmail" value="1"<% If steRecordValue(rsEdit, "ShowEmail") <> 0 Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "Yes" %>
	<input type="radio" name="ShowEmail" value="0"<% If steRecordValue(rsEdit, "ShowEmail") = 0 Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "No" %>
	</td>
</tr><tr>
	<td class="forml" colspan=2><% steTxt "Biography" %><br>
	<textarea name="biography" rows="14" cols="48" class="form"><%= steRecordEncValue(rsEdit, "Biography") %></textarea>
	</td>
</tr><tr>
	<td colspan=2 align="right">
	<input type="hidden" name="action" value="Create">
	<input type="submit" name="_submit" value=" <% steTxt "Create" %> " class="form">
	</td>
</tr>
</table>
</form>

<!-- #include file="../../../footer.asp" -->