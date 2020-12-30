<% Option Explicit %>
<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' help.asp
'	Display help information for the referring page (if it exists)
'	Provides links directly on the page to edit the help.
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
Dim rsDoc
Dim sReferrer
Dim sAction			' action to perform
Dim nCategoryID
Dim oList			' admin list object
Dim sWhere
Dim sStatusMsg		' status message to display
Dim I

sAction = steForm("action")
sReferrer = steForm("referrer")
nCategoryID = steNForm("CategoryID")

If sReferrer = "" Then
	sReferrer = Request.ServerVariables("HTTP_REFERER")
End If

' add or edit the documentation here (if nec)
If sAction = "doadd" Then
	If Trim(steForm("Title")) = "" Then
		sErrorMsg = "Please enter the title for your help document"
	ElseIf Trim(steForm("Body")) = "" Then
		sErrorMsg = "Please enter the body for your help document"
	Else
		' no error, create the help document
		Dim nAuthorID
	
		nAuthorID = locAuthor
		sStat = "INSERT INTO tblDoc (" &_
				"	AuthorID, FolderID, Title, Body, TypeID, ScriptName" &_
				") VALUES (" &_
					nAuthorID & ", 0, " & steQForm("Title") & "," & steQForm("Body") &_
					", 2, '" & Replace(sReferrer, "'", "''") & "'" &_
				")"
		Call adoExecute(sStat)
		sAction = ""
		sStatusMsg = "Help Documentation Created"
	End If
ElseIf sAction = "doedit" Then
	If Trim(steForm("Title")) = "" Then
		sErrorMsg = "Please enter the title for your help document"
	ElseIf Trim(steForm("Body")) = "" Then
		sErrorMsg = "Please enter the body for your help document"
	Else
		' no error, update the help document
		sStat = "UPDATE tblDoc SET " &_
				"	Title = " & steQForm("Title") & "," &_
				"	Body = " & steQForm("Body") & " " &_
				"WHERE	DocID = " & steNForm("DocID")
		Call adoExecute(sStat)
		sAction = ""
		sStatusMsg = "Help Documentation Updated"
	End If
End If

If sAction = "add" or sAction = "" Then
	' determine if documentation exists for the referring page
	sStat = "SELECT	DocID, Title, Body " &_
			"FROM	tblDoc " &_
			"WHERE	TypeID = 2 " &_
			"AND	ScriptName = '" & Replace(sReferrer, "'", "''") & "'"
	Set rsDoc = adoOpenRecordset(sStat)
End If
%>
<!-- #include file="../../../../header_popup.asp" -->

<% If sAction = "" Then %>

	<% If sStatusMsg <> "" Then %>
	<p><b class="error"><%= sStatusMsg %></b></p>
	<% End If %>

	<% If Not rsDoc.EOF Then %>
	
		<h3><%= rsDoc.Fields("Title").Value %></h3>
	
		<%= rsDoc.Fields("Body").Value %>
	
	<% Else %>
	
		<h3>No Documentation Written</h3>
	
		<p>
		No documentation has been written yet for this page
		</p>
	
	<% End If %>

<% ElseIf sAction = "add" Or sAction = "doadd" Then %>

	<h3>Add Help Documentation</h3>

	<% If sErrorMsg <> "" Then %>
	<p><b class="error"><%= sErrorMsg %></b></p>
	<% End If %>

	<% If sStatusMsg <> "" Then %>
	<p><b class="error"><%= sStatusMsg %></b></p>
	<% End If %>

	<form method="post" action="help.asp?action=doadd">
	<input type="hidden" name="referrer" value="<%= Server.HTMLEncode(sReferrer) %>">

	<table border=0 cellpadding=2 cellspacing=0>
	<tr>
		<Td class="forml"></td><td>&nbsp;&nbsp;</td>
		<td class="formd"><input type="text" name="title" value="<%= steEncForm("title") %>" size="32" maxlength="50" class="form"></td>
	</tr>
	<tr>
		<Td class="forml"></td><td>&nbsp;&nbsp;</td>
		<td class="formd"><textarea name="body" cols="48" rows="24" class="form" style="width:100%"><%= steEncForm("body") %></textarea></td>
	</tr>
	<tr>
		<td colspan="3" align="center"><br>
			<input type="submit" name="_action" value=" Add ">
		</td>
	</tr>
	</table>
	</form>

<% ElseIf sAction = "edit" Then %>

	<% If Not rsDoc.EOF Then %>

	<% If sErrorMsg <> "" Then %>
	<p><b class="error"><%= sErrorMsg %></b></p>
	<% End If %>

	<% If sStatusMsg <> "" Then %>
	<p><b class="error"><%= sStatusMsg %></b></p>
	<% End If %>

	<form method="post" action="help.asp?action=doedit">
	<input type="hidden" name="referrer" value="<%= Server.HTMLEncode(sReferrer) %>">
	<input type="hidden" name="DocID" value="<%= nDocID %>">

	<table border=0 cellpadding=2 cellspacing=0>
	<tr>
		<Td class="forml"></td><td>&nbsp;&nbsp;</td>
		<td class="formd"><input type="text" name="title" value="<%= Server.HTMLEncode(rsDoc.Fields("title").Value) %>" size="32" maxlength="50" class="form"></td>
	</tr>
	<tr>
		<Td class="forml"></td><td>&nbsp;&nbsp;</td>
		<td class="formd"><textarea name="body" cols="48" rows="24" class="form" style="width:100%"><%= Server.HTMLEncode(rsDoc.Fields("body").Value) %></textarea></td>
	</tr>
	<tr>
		<td colspan="3" align="center"><br>
			<input type="submit" name="_action" value=" Update ">
		</td>
	</tr>
	</table>
	</form>

	<% Else %>
	
		<h3>No Documentation Found</h3>
	
		<p>
		Could not find the help document to edit this page
		</p>
	
	<% End If %>

<% End If %>

<!-- #include file="../../../../footer_popup.asp" -->
<%
Function locAuthor
	Dim sStat, rsAuth

	' retrieve the author for this article
	sStat = "SELECT	AuthorID FROM tblDocAuthor WHERE UserID = " & Request.Cookies("UserID")
	Set rsAuth = adoOpenRecordset(sStat)
	If Not rsAuth.EOF Then
		locAuthor = rsAuth.Fields("AuthorID").Value
		Exit Function
	End If
	rsAuth.Close

	' retrieve the author for this article
	sStat = "SELECT	FirstName, MiddleName, LastName FROM tblUser WHERE UserID = " & Request.Cookies("UserID")
	Set rsAuth = adoOpenRecordset(sStat)
	If Not rsAuth.EOF Then
		' create a new document author for this user
		sStat = "INSERT INTO tblDocAuthor (" &_
				"	UserID, FirstName, MiddleName, LastName" &_
				") VALUES (" &_
				Request.Cookies("UserID") & "," &_
				"'" & Replace(rsAuth.Fields("FirstName").Value, "'", "''") & "," &_
				"'" & Replace(rsAuth.Fields("MiddleName").Value & "", "'", "''") & "," &_
				"'" & Replace(rsAuth.Fields("LastName").Value & "", "'", "''") &_
				")"
		rsAuth.Close
		Set rsAuth = Nothing
		Call adoExecute(sStat)
	
		' retrieve the new author ID			
		sStat = "SELECT	AuthorID FROM tblDocAuthor WHERE UserID = " & Request.Cookies("UserID")
		Set rsAuth = adoOpenRecordset(sStat)
		If Not rsAuth.EOF Then
			locAuthor = rsAuth.Fields("AuthorID").Value
		Else
			locAuthor = 0
		End If
	Else
		' unable to retrieve user information
		locAuthor = 0
	End If
End Function
%>