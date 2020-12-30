<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' userright_edit.asp
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

Dim sStat
Dim sPassword		' SQL to modify password
Dim rsUser
Dim rsRight
Dim aRight
Dim bFoundUser
Dim sErrorMsg
Dim nCanAdd
Dim nCanEdit
Dim nCanDelete
Dim nCanView
Dim I

If LCase(steForm("action")) = "assign" Then

	' delete all existing rights first
	sStat = "DELETE " &_
			"FROM	tblUserToRight " &_
			"WHERE	UserID = " & steNForm("UserID")
	Call adoExecute(sStat)
	' assign all selected rights to this user (if any)
	Dim sKey, nRightID
	For Each sKey In Request.Form
		nRightID = 0
		If Left(sKey, 5) = "Right" Then
			nRightID = Mid(sKey, 6)
			If IsNumeric(nRightID) And CStr(nRightID) <> "" Then nRightID = CInt(nRightID) Else nRightID = 0
		End If
		If nRightID > 0 And Trim(Request.Form(sKey)) <> "" Then
			If InStr(1, Request.Form(sKey), "A") Then nCanAdd = 1 Else nCanAdd = 0
			If InStr(1, Request.Form(sKey), "E") Then nCanEdit = 1 Else nCanEdit = 0
			If InStr(1, Request.Form(sKey), "D") Then nCanDelete = 1 Else nCanDelete = 0
			If InStr(1, Request.Form(sKey), "V") Then nCanView = 1 Else nCanView = 0
			sStat = "INSERT INTO tblUserToRight (" &_
					"	UserID, RightID, CanAdd, CanEdit, CanDelete, CanView" &_
					") VALUES (" &_
					steNForm("UserID") & ", " & aRight(I) & ", " & nCanAdd & ", " & nCanEdit & ", " & nCanDelete & ", " & nCanView &_
					")"
			Call adoExecute(sStat)
		End If
	Next
End If

' retrieve the user information to edit (if nec)
bFoundUser = False
If steForm("action") <> "assign" Or sErrorMsg <> "" Then
	sStat = "SELECT * FROM tblUser WHERE UserID = " & steNForm("UserID")
	Set rsUser = adoOpenRecordset(sStat)
	If Not rsUser.EOF Then bFoundUser = True
End If

' retrieve the user rights
sStat = "SELECT	ur.RightID, ur.ParentRightID, ur.RightName, ur.Hyperlink, utr.UserID, " &_
		"		ur.HasAdd, ur.HasEdit, ur.HasDelete, ur.HasView, " &_
		"		utr.CanAdd, utr.CanEdit, utr.CanDelete, utr.CanView " &_
		"FROM	tblUserRight ur " &_
		"LEFT JOIN tblUserToRight utr ON utr.RightID = ur.RightID " &_
		"	AND utr.UserID = " & steNForm("UserID") & " " &_
		"WHERE	ur.Active <> 0 " &_
		"AND	ur.Archive = 0 " &_
		"ORDER BY OrderNo"
Set rsRight =adoOpenRecordset(sStat)
If Not rsRight.EOF Then aRight = rsRight.GetRows
rsRight.Close
Set rsRight = Nothing
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Rights" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "assign" Or sErrorMsg <> "" Then %>

<% If bFoundUser Then  %>
<H3><% steTxt "Admin User Rights" %></H3>

<table border=0 cellpadding=2 cellspacing=0>
<tr>
	<td class="forml"><% steTxt "Full Name" %><br>
	<font class="formd"><%= steRecordEncValue(rsUser, "FirstName") %>&nbsp;<%= steRecordEncValue(rsUser, "LastName") %></font>
	</td>
	<td><img src="../../../img/pixel.gif" width=30 height=1></td>
	<td class="forml"><% steTxt "Username" %><br>
	<font class="formd"><%= steRecordEncValue(rsUser, "Username") %></font>
	</td>
</tr>
</table>

<p><b><% steTxt "Assigned Rights" %></b></p>

<form method="post" action="userright_edit.asp">
<input type="hidden" name="userid" value="<%= steForm("UserID") %>">

<% If IsArray(aRight) Then %>
<P>
<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 class="list">
<TR>
	<TD class="listhead"><% steTxt "Right Name" %></TD>
	<TD class="listhead"><% steTxt "Add" %></TD>
	<TD class="listhead"><% steTxt "Edit" %></TD>
	<TD class="listhead"><% steTxt "Delete" %></TD>
	<TD class="listhead"><% steTxt "View" %></TD>
</TR>
<% ' display the tree of rights here
	Call locRightTree(0, 0, 0) %>
</TABLE>
</P>
<% End If %>

<table border=0 cellpadding=2 cellspacing=0>
	<td align="right"><br>
		<input type="hidden" name="action" value="assign">
		<input type="submit" name="_submit" value=" <% steTxt "Update User Rights" %> " class="form">
	</td>
</tr>
</table>

</form>

<% Else %>

<H3><% steTxt "Admin User Rights" %></H3>

<p><b class="error"><% steTxt "No user has been selected to administer" %></b></p>

<p>
<% steTxt "In order to administer the user rights, you must first select a user." %>&nbsp;
<% steTxt "Do this by going through the user list and clicking <b>rights</b> next to the user to edit their access rights." %>
</p>

<p>
<% steTxt "You can also click on the <B>Rights</b> tab if you are in edit mode for a user." %>
</p>

<% End If %>

<% Else %>

<H3><% steTxt "Admin User Rights Modified" %></H3>

<P>
<% steTxt "The admin user rights have been modified according to the changes you requested." %>
</P>

<% End If %>

<!-- #include file="../../../footer.asp" -->
<%
Sub locRightTree(nParentID, nIndex, nLevel)
	Dim I

	If CStr(nIndex) = "" Then nIndex = 0
	' find all rights matching the parent and display
	For I = 0 To UBound(aRight, 2) 
		If aRight(1, I) = nParentID Then 
			'  onMouseOver="this.className='listsel'" onMouseOut="this.className='list<= nIndex mod 2 >'"
			%>
<TR CLASS="list<%= nIndex mod 2 %>">
	<TD><table border=0 cellpadding=2 cellspacing=0>
	<tr>
		<TD WIDTH="<%= nLevel * 15 %>"><IMG SRC="<%= Application("ASPNukeBasePath") %>img/pixel.gif" WIDTH="<%= nLevel * 15 %>" HEIGHT="1" ALT=""></TD>
		<TD WIDTH="100%"><%= aRight(2, I) %></TD>
	</tr>
	</table></TD>
	<TD align="center"><% If aRight(5, I) = 0 Then %><input type="checkbox" disabled class="formcheck" name="_disabled" value="1"><% Else %><input type="checkbox" name="Right<%= aRight(0, I) %>" value="A" class="formcheck"<% If aRight(9, I) = 1 Then Response.Write " checked" %>><% End If %></TD>
	<TD align="center"><% If aRight(6, I) = 0 Then %><input type="checkbox" disabled class="formcheck" name="_disabled" value="1"><% Else %><input type="checkbox" name="Right<%= aRight(0, I) %>" value="E" class="formcheck"<% If aRight(10, I) = 1 Then Response.Write " checked" %>><% End If %></TD>
	<TD align="center"><% If aRight(7, I) = 0 Then %><input type="checkbox" disabled class="formcheck" name="_disabled" value="1"><% Else %><input type="checkbox" name="Right<%= aRight(0, I) %>" value="D" class="formcheck"<% If aRight(11, I) = 1 Then Response.Write " checked" %>><% End If %></TD>
	<TD align="center"><% If aRight(8, I) = 0 Then %><input type="checkbox" disabled class="formcheck" name="_disabled" value="1"><% Else %><input type="checkbox" name="Right<%= aRight(0, I) %>" value="V" class="formcheck"<% If aRight(12, I) = 1 Then Response.Write " checked" %>><% End If %></TD>
</TR>
<%			' display the child rights here
			nIndex = nIndex + 1
			Call locRightTree(aRight(0, I), nIndex, nLevel+1)
		End If
	Next
End Sub
%>