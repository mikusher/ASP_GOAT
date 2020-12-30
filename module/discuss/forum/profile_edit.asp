<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/graphics/gfx_lib.asp" -->
<%
'--------------------------------------------------------------------
' profile_edit.asp
'	Creates a  profile for a member who is using our message forums.
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
Dim sAction
Dim sLocation
Dim sEmail
Dim sUsername		' username for profile
Dim nTopicID		' topic the user was browsing
Dim bForbidden
Dim sStatusMsg
Dim sForumIcon
Dim nMemberID
Dim sErrorMsg
Dim bPost

' process form post / http upload
steOverwrite = True
If steFormPost(Application("ASPNukeBasePath") & "img/forum/avatar", sErrorMsg) Then
	nMemberID = steForm("memberid")
	sUsername = steForm("username")
Else
	nMemberID = Request.Cookies("MemberID")
	sUsername = Request.Cookies("Username")
End If

' check to make sure the user can edit their profile
If Not steWasUpload Then
	bForbidden = True
	If nMemberID <> "" Then
		sStat = "SELECT MemberID " &_
				"FROM	tblMember " &_
				"WHERE	UserName = '" & Replace(sUsername, "'", "''") & "' " &_
				"AND	MemberID = " & nMemberID
		' Response.Write "<pre>" & sStat & "</pre>" : Response.End
		Set rsMember = adoOpenRecordset(sStat)
		If Not rsMember.EOF Then bForbidden = False
	Else
		' member not logged in yet
		Response.Redirect Application("ASPNukeBasePath") & "module/account/register/login.asp?goto=" &_
			Server.URLEncode(Application("ASPNukeBasePath") & "module/discuss/forum/profile_edit.asp") &_
			"&error=" & Server.URLEncode(steGetText("Please login in order to edit your profile"))
	End If
End If

' validate the image uploaded (if any)
If sErrorMsg = "" And steForm("action") = "update" Then
	Dim nWidth, nHeight, nColors, sImgType

	' define the forum icon
	If steForm("forumiconfile") <> "" Then
		If steUploadRename <> "" Then
			sForumIcon = steUploadRename
		Else
			sForumIcon = Application("ASPNukeBasePath") & "img/forum/avatar/" & steForm("forumiconfile")
		End If
		' check the size of the avatar image
		If gfxSpex(Server.MapPath(sForumIcon), nWidth, nHeight, nColors, sImgType) Then
			If nWidth <> modParam("Forum", "AvatarImageWidth") Or nHeight <> modParam("Forum", "AvatarImageHeight") Then
				sErrorMsg = steGetText("Invalid Avatar Image Size") & " (" & nWidth & "x" & nHeight &_
					") - " & steGetText("Should be") & " (" & modParam("Forum", "AvatarImageWidth") & "x" & modParam("Forum", "AvatarImageHeight") & ")<br>"
			End If
		Else
			sErrorMsg = steGetText("File is corrupt") & " (" & steForm("forumiconfile") & ") - " & steGetText("Expected GIF, JPG or PNG image") & "<br>"
		End If
	Else
		sForumIcon = steForm("forumicon")
	End If
End If

If sErrorMsg = "" And steForm("action") = "update" Then
		' check to see if the profile exists first
		Set rsCheck = adoOpenRecordset("SELECT * FROM tblMessageProfile WHERE MemberID = " & nMemberID)
		If rsCheck.EOF Then
			' insert the profile
			If sErrorMsg = "" Then
				sStat = "INSERT INTO tblMessageProfile (" &_
						"	MemberID, RankID, Location, Email, ForumIcon, Biography, HomePage, Created" &_
						") VALUES (" &_
						nMemberID & ", 0, " & steQForm("location") & ", " &_
						steQForm("email") & ", '" & Replace(sForumIcon, "'", "''") & "', '" &_
						Replace(steStripForm("biography"), "'", "''") & "', " & steQForm("homepage") &_
						"," & adoGetDate &_
						")"
				Call adoExecute(sStat)
				sStatusMsg = steGetText("Your message forum profile was created successfully")
			End If
		Else
			' update the profile here
			If sErrorMsg = "" Then
				sStat = "UPDATE tblMessageProfile SET " &_
						"	Location = " & steQForm("location") & ", " &_
						"	Email = " & steQForm("email") & ", " &_
						"	ForumIcon = '" & Replace(sForumIcon, "'", "''") & "', " &_
						"	Biography= '" & Replace(steStripForm("biography"), "'", "''") & "', " &_
						"	HomePage = " & steQForm("homepage") & " " &_
						"WHERE	MemberID = " & nMemberID
				Call adoExecute(sStat)
				sStatusMsg = steGetText("Your message forum profile was updated successfully")
			End If
		End If
End If

If Not bForbidden Then
	' retrieve the current profile for this user
	sStat = "SELECT * FROM tblMessageProfile WHERE MemberID = " & nMemberID
	Set rsEdit = adoOpenRecordset(sStat)
End If
%>
<!-- #include file="../../../header.asp" -->

<% If Not bForbidden Then %>

<% If LCase(steForm("action")) <> "update" Or sErrorMsg <> "" Then %>

<h3><% steTxt "Update Member Profile" %></h3>

<p>
<% steTxt "Please make the changes to your message forum profile using the form provided below." %>
</p>

<% If sErrorMsg <> "" Then %>
<P><B class="error"><%= sErrorMsg %></b></p>
<% ElseIf sStatusMsg <> "" Then %>
<P><B class="error"><%= sStatusMsg %></b></p>
<% End If %>

<form method="post" action="profile_edit.asp" enctype="multipart/form-data">
<input type="hidden" name="action" value="update">
<input type="hidden" name="topicid" value="<%= nTopicID %>">
<input type="hidden" name="threadid" value="<%= steEncForm("ThreadID") %>">
<input type="hidden" name="memberid" value="<%= nMemberID %>">
<input type="hidden" name="username" value="<%= Server.HTMLEncode(sUsername) %>">
<input type="hidden" name="uploadrename" value="<%= nMemberID %>">
<input type="hidden" name="steallowoverwrite" value="<%= Server.HTMLEncode(steRecordEncValue(rsEdit, "ForumIcon")) %>">

<table border=0 cellpadding=4 cellspacing=0>
<tr>
	<td class="forml"><% steTxt "Location" %></td><td>&nbsp;&nbsp;</td>
	<td><input type="text" name="location" value="<%= steRecordEncValue(rsEdit, "location") %>" size="42" maxlength="50" class="form"></td>
</tr><tr>
	<td class="forml"><% steTxt "E-Mail Address" %></td><td></td>
	<td><input type="text" name="email" value="<%= steRecordEncValue(rsEdit, "email") %>" size="42" maxlength="100" class="form"></td>
</tr><tr>
	<td class="forml"><% steTxt "Avatar URL" %><br><font class="tinytext">(<% steTxt "GIF or JPG Image" %>)</font></td><td></td>
	<td><input type="text" name="forumicon" value="<%= steRecordEncValue(rsEdit, "forumicon") %>" size="42" maxlength="255" class="form"></td>
</tr><tr>
	<td class="forml"><% steTxt "Upload Avatar" %><br><font class="tinytext">(<% steTxt "GIF or JPG Image" %>)</font></td><td></td>
	<td><input type="file" name="forumiconfile" size="42" maxlength="255" class="form"></td>
</tr><tr>
	<td class="forml" valign="top"><% steTxt "Biography" %></td><td></td>
	<td><textarea name="biography" cols="48" rows="14" class="form"><%= steRecordEncValue(rsEdit, "biography") %></textarea></td>
</tr><tr>
	<td class="forml"><% steTxt "Home Page" %></td><td></td>
	<td><input type="text" name="homepage" value="<%= steRecordEncValue(rsEdit, "homepage") %>" size="42" maxlength="32" class="form"></td>
</tr>
<tr>
	<td colspan=3 align="right"><br>
		<input type="submit" name="_edit" value=" <% steTxt "Update Profile" %> " class="form">
	</td>
</tr>
</table>
</form>


<% Else %>

<h3><% steTxt "Update Successful" %></h3>

<% If sErrorMsg <> "" Then %>
<P><B class="error"><%= sErrorMsg %></b></p>
<% ElseIf sStatusMsg <> "" Then %>
<P><B class="error"><%= sStatusMsg %></b></p>
<% End If %>

<p>
<% steTxt "The changes to your message forum profile were saved to the database." %>&nbsp;
<% steTxt "Visitors will now be able to view your forum profile by clicking on your username which is attached to your message forum posts." %>
</p>

<% End If %>

<% Else %>

<h3><% steTxt "Access Denied" %></h3>

<p>
<% steTxt "The profile you selected is not your own." %>&nbsp;
<% steTxt "You are not permitted to edit another member's profile." %>
</p>

<% End If %>

<p align="center">
<% If nTopicID > 0 Then %>
	<a href="topic.asp?topicid=<%= nTopicID %>" class="footerlink"><% steTxt "Topic Overview" %></a> &nbsp;
<% End If %>
<% If steNForm("ThreadID") > 0 Then %>
	<a href="thread.asp?topicid=<%= nTopicID %>&threadid=<%= steNForm("ThreadID") %>" class="footerlink"><% steTxt "Back to Thread" %></a> &nbsp;
<% End If %>
	<a href="profile.asp?topicid=<%= nTopicID %>&username=<%= Server.URLEncode(sUsername) %>" class="footerlink"><% steTxt "View Profile" %></a>
</p>
<!-- #include file="../../../footer.asp" -->