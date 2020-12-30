<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' user_rights.asp
'	Displays a list of assigned rights for a user
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
Dim rsUser
Dim nUserID
Dim sFullName
Dim sUsername
Dim rsRight
Dim I

nUserID = steNForm("UserID")
sAction = LCase(steForm("Action"))

If sAction = "update" Then
	' peform an update of the rights assigned to this user
	Call adoExecute("DELETE FROM tblUserToRight WHERE UserID = " & nUserID)

	sStat = ""
	aRight = Split(Replace(steForm("RightID"), " ", ""), ",")
	For I = 0 To UBound(aRight)
		sStat = sStat & "INSERT INTO tblUserToRight (UserID, RightID, Created) " &_
			"VALUES (" & nUserID & ", " & aRight(I) & ", " & adoGetDate & "); "
	Next
	' perform all access assignments at once
	Call adoExecute(sStat)

	' if admin modifies himself, update main menu
	If Request.Cookies("AdminUserID") = CStr(nUserID) Then
		Call navLogin(nUserID)
	End If
End If

' retrieve the user that we want to modify
sStat = "SELECT	FirstName, MiddleName, LastName, Username " &_
		"FROM	tblUser " &_
		"WHERE	UserID = " & nUserID
Set rsUser = adoOpenRecordset(sStat)
If Not rsUser.EOF Then
	sFullName = rsUser.Fields("FirstName") & " " & rsUser.Fields("MiddleName").Value &_
		" " & rsUser.Fields("LastName").Value
	sUsername = rsUser.Fields("Username").Value
End If
rsUser.Close
Set rsUser = Nothing

' retrieve the list of rights
sStat = "SELECT	ur.RightID, ur.RightName, utr.UserID " &_
		"FROM	tblUserRight ur " &_
		"LEFT JOIN tblUserToRight utr ON utr.RightID = ur.RightID " &_
		"		AND	utr.UserID = " & nUserID & " " &_
		"WHERE	ur.Active <> 0 " &_
		"AND	ur.Archive = 0 " &_
		"ORDER BY ur.RightName"
Set rsRight = adoOpenRecordset(sStat)

' make sure user has access to the access control module
navRight = "Access"
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Rights" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If sAction <> "update" Or sErrorMsg <> "" Then %>

<h3><% steTxt "Update User Rights" %></h3>

<p>
<% steTxt "Please select the rights that should be assigned to this user." %>&nbsp;
<% steTxt "The user will need to log out and log back in to see the changed access permissions." %>
</p>

<% If sErrorMsg <> "" Then %>
<p><b class="error"><%= sErrorMsg %></b></p>
<% End If %>

<table border=0 cellpadding=2 cellspacing=0>
<tr>
	<Td class="forml"><% steTxt "Username" %></td><Td>&nbsp;&nbsp;</td>
	<td class="formd"><%= Server.HTMLEncode(sUsername) %></td>
</tr>
<tr>
	<Td class="forml"><% steTxt "Full Name" %></td><Td></td>
	<td class="formd"><%= Server.HTMLEncode(sFullname) %></td>
</tr>
</table>

<% If Not rsRight.EOF Then %>

<form action="user_rights.asp" method="post">
<input type="hidden" name="UserID" value="<%= nUserID %>">
<input type="hidden" name="action" value="update">

<table border=0 cellpadding=2 cellspacing=0 class="list">
<tr>
	<td class="listhead"><% steTxt "Grant" %></td>
	<td class="listhead"><% steTxt "Access Right" %></td>
</tr>
<% I = 0
   Do Until rsRight.EOF %>
<tr class="list<%= I Mod 2 %>">
	<td><input type="checkbox" name="RightID" value="<%= rsRight.Fields("RightID").Value %>"<% If Not IsNull(rsRight.Fields("UserID").Value) Then Response.Write " CHECKED" %> class="form"></td>
	<td class="formd"><%= rsRight.Fields("RightName").Value %></td>
</tr>
<%	rsRight.MoveNext
	I = I + 1
   Loop
	rsRight.Close
	Set rsRight = Nothing %>
</table>

<p align="center">
	<input type="submit" name="_update" value=" <% steTxt "Update User Rights" %>" class="form">
</p>
</form>

<% Else %>

<p>
<b class="error"><% steTxt "No access rights have been defined yet" %></b>
</p>

<% End If %>

<% Else ' form submitted successfully %>

<h3><% steTxt "User Rights Updated" %></h3>

<p>
<% steTxt "The user rights were updated successfully in the database." %>&nbsp;
<% steTxt "Please use the admin links provided below to continue administering the site." %>
</p>

<% End If %>

<p align="center">
	<a href="user_list.asp" class="adminlink"><% steTxt "User List" %></a>
<% If sAction = "update" and sErrorMsg = "" Then %>
	&nbsp; <a href="user_rights.asp?userid=<%= nUserID %>" class="adminlink"><% steTxt "View Rights" %></a>
<% End If %>
</p>

<!-- #include file="../../../footer.asp" -->
		