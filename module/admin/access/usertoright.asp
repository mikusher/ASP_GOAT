<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' usertoright.asp
'	Displays a list of the admin rights for the site administration
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
Dim nRightID
Dim rsRight		' user right being modified
Dim sRightName	' name of right to modify
Dim rsUser	' users to choose from
Dim aUser
Dim sUserList
Dim nCanAdd
Dim nCanEdit
Dim nCanDelete
Dim nCanView
Dim nHasAdd
Dim nHasEdit
Dim nHasDelete
Dim nHasView
Dim rsAss	' assigned

nRightID = steNForm("RightID")
sAction = Trim(UCase(steForm("Action")))

Select Case sAction
	Case "ASSIGN"
		' remove the old associations first
		Call adoExecute("DELETE FROM tblUserToRight WHERE RightID = " & nRightID)

		If Trim(steForm("UserIDList")) <> "" Then
			Dim aUserID
			aUser = Split(steForm("UserIDList"), ",")

			For I = 0 To UBound(aUser)
				If Trim(steForm("UserID" & auser(I))) <> "" Then
					If InStr(1, steForm("UserID" & auser(I)), "A") Then nCanAdd = 1 Else nCanAdd = 0
					If InStr(1, steForm("UserID" & auser(I)), "E") Then nCanEdit = 1 Else nCanEdit = 0
					If InStr(1, steForm("UserID" & auser(I)), "D") Then nCanDelete = 1 Else nCanDelete = 0
					If InStr(1, steForm("UserID" & auser(I)), "V") Then nCanView = 1 Else nCanView = 0

					sStat = "INSERT INTO tblUserToRight (" &_
						"UserID, RightID, CanAdd, CanEdit, CanDelete, CanView" &_
						") VALUES (" &_
						aUser(I) & ", " & nRightID & ", " & nCanAdd & ", " & nCanEdit & ", " & nCanDelete & ", " & nCanView & ")"
					Call adoExecute(sStat)
				End If
			Next
		End If

		' re-login the administrator (to update nav)
		navLogin(Request.Cookies("AdminUserID"))
End Select

' retrieve the right name
sStat = "SELECT RightName, HasAdd, HasEdit, HasDelete, HasView FROM tblUserRight WHERE RightID = " & nRightID
Set rsRight = adoOpenRecordset(sStat)
If Not rsRight.EOF Then
	sRightName = rsRight.Fields("RightName").Value
	nHasAdd = rsRight.Fields("HasAdd").Value
	nHasEdit = rsRight.Fields("HasEdit").Value
	nHasDelete = rsRight.Fields("HasDelete").Value
	nHasView = rsRight.Fields("HasView").Value
End If
rsRight.Close
rsRight = Empty

' retrieve the list of users to choose from
sStat = "SELECT	UserID, FirstName, MiddleName, LastName, Username " &_
		"FROM	tblUser " &_
		"WHERE	Active <> 0 " &_
		"AND	Archive = 0 " &_
		"ORDER BY LastName, FirstName, MiddleName"
Set rsUser = adoOpenRecordset(sStat)
If Not rsUser.EOF Then aUser = rsUser.GetRows
rsUser.Close
rsUser = Empty

' retrieve the list of assigned users (granted the right)
sStat = "SELECT	u.UserID, u.FirstName, u.MiddleName, u.LastName, u.Username, " &_
		"		utr.CanAdd, utr.CanEdit, utr.CanDelete, utr.CanView " &_
		"FROM	tblUser u " &_
		"INNER JOIN	tblUserToRight utr ON utr.UserID = u.UserID " &_
		"WHERE	utr.RightID = " & nRightID & " " &_
		"AND	u.Active <> 0 " &_
		"AND	u.Archive = 0"
Set rsAss = adoOpenRecordset(sStat)
Set oAss = Server.CreateObject("Scripting.Dictionary")
Do Until rsAss.EOF
	If rsAss.Fields("CanAdd").Value = 1 Then oAss.Item("User" & CStr(rsAss.Fields("UserID").Value) & "A") = True
	If rsAss.Fields("CanEdit").Value = 1 Then oAss.Item("User" & CStr(rsAss.Fields("UserID").Value) & "E") = True
	If rsAss.Fields("CanDelete").Value = 1 Then oAss.Item("User" & CStr(rsAss.Fields("UserID").Value) & "D") = True
	If rsAss.Fields("CanView").Value = 1 Then oAss.Item("User" & CStr(rsAss.Fields("UserID").Value) & "V") = True
	rsAss.MoveNext
Loop
rsAss.Close
rsAss = Empty
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Rights" %>
<!-- #include file="pagetabs_inc.asp" -->

<h3><% steTxt "Users Assigned to Group" %></h3>

<p>
<% steTxt "The following rights are assigned to the group which is listed below." %>
</p>

<h4><%= sRightName %></h4>

<form name="formedit" method="post" action="usertoright.asp" ID="formedit">
<input type="hidden" name="RightID" value="<%= nRightID %>">
<input type="hidden" name="action" value=""

<table border=0 cellpadding=2 cellspacing=0 class="list" align="center">
<tr>
	<td class="listhead"><% steTxt "Full Name" %></td><td class="listhead">&nbsp;&nbsp;&nbsp;</td>
	<td class="listhead"><% steTxt "Username" %></td>
<% If nHasAdd = 1 Then %>
	<td class="listhead"><% steTxt "Add" %></td><td class="listhead">&nbsp;&nbsp;&nbsp;</td>
<% End If %>
<% If nHasEdit = 1 Then %>
	<td class="listhead"><% steTxt "Edit" %></td><td class="listhead">&nbsp;&nbsp;&nbsp;</td>
<% End If %>
<% If nHasDelete = 1 Then %>
	<td class="listhead"><% steTxt "Delete" %></td><td class="listhead">&nbsp;&nbsp;&nbsp;</td>
<% End If %>
<% If nHasView = 1 Then %>
	<td class="listhead"><% steTxt "View" %></td><td class="listhead">&nbsp;&nbsp;&nbsp;</td>
<% End If %>
</tr>
<% For I = 0 To UBound(aUser, 2)
	sUserList = sUserList & "," & aUser(0, I) %>
<tr class="list<%= I mod 2 %>">
	<td><%= aUser(1, I) & " " & aUser(2, I) & " " & aUser(3, I) %></td><td></td>
	<td><%= aUser(4, I) %></td>
<% If nHasAdd = 1 Then %>
	<td><input type="checkbox" name="UserID<%= aUser(0, I) %>" value="A"<% If oAss.Exists("User" & CStr(aUser(0, I))&"A") Then Response.Write " CHECKED" %> class="formcheck"></td><td></td>
<% End If %>
<% If nHasEdit = 1 Then %>
	<td><input type="checkbox" name="UserID<%= aUser(0, I) %>" value="E"<% If oAss.Exists("User" & CStr(aUser(0, I))&"E") Then Response.Write " CHECKED" %> class="formcheck"></td><td></td>
<% End If %>
<% If nHasDelete = 1 Then %>
	<td><input type="checkbox" name="UserID<%= aUser(0, I) %>" value="D"<% If oAss.Exists("User" & CStr(aUser(0, I))&"D") Then Response.Write " CHECKED" %> class="formcheck"></td><td></td>
<% End If %>
<% If nHasView = 1 Then %>
	<td><input type="checkbox" name="UserID<%= aUser(0, I) %>" value="V"<% If oAss.Exists("User" & CStr(aUser(0, I))&"V") Then Response.Write " CHECKED" %> class="formcheck"></td><td></td>
<% End If %>
</tr>
<% Next %>
</table>

<input type="hidden" name="UserIDList" value="<%= sUserList %>">
<% If nHasAdd = 1 Or nHasEdit = 1 Or nHasDelete = 1 Or nHasView = 1 Then %>
<p align="center">
	<input type="submit" name="_action" value=" <% steTxt "Assign" %> " class="form" onclick="document.formedit.action.value='assign'">
</p>
<% Else %>
<p align="center">
	<b class="error">No actions have been defined for this right. In order to
	edit the user rights, you must add an action such as "Add", "Edit", "Delete"
	or "View."</b>
</p>
<% End If %>
</form>

<p align="center">
	<a href="userright_list.asp" class="adminlink"><% steTxt "User Right List" %></a>
</p>

<!-- #include file="../../../footer.asp" -->