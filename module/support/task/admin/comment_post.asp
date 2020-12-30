<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' comment_post.asp
'	Enter a comment on an task.
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
Dim nTaskID
Dim rsMember
Dim rsArt
Dim nMemberID
Dim rsReply
Dim rsDup
Dim sSubject
Dim sBody
Dim nReplyID

nTaskID = steNForm("TaskID")
nReplyID = steNForm("ReplyID")
sSubject = steStripForm("subject")
sBody = steStripForm("body")

' make sure the user is logged in
If Request.Cookies("MemberID") <> "" Then
	nMemberID = Request.Cookies("MemberID")
Else
	nMemberID = 0
End If

If steForm("action") = "post" Then
	' check for the required fields here
	If sSubject = "" Then
		sErrorMsg = steGetText("Please enter the subject for your post")
	ElseIf sBody = "" Then
		sErrorMsg = steGetText("Please enter the body for your post")
	Else
		' all of the req. variables were passed
		' check the form variables
		If Trim(steForm("username")) <> "" And Trim(steForm("password")) <> "" Then
			' retrieve the user information here
			sStat = "SELECT	MemberID, FirstName, LastName, Username " & _
					"FROM	tblMember " &_
					"WHERE	Username = " & steQForm("username") & " " &_
					"AND	Password = '" & SHA256(steForm("password")) & "'"
			Set rsMember = adoOpenRecordset(sStat)
			If Not rsMember.EOF Then
				' login the user
				Response.Cookies("MemberID") = rsMember.Fields("MemberID").Value
				nMemberID = rsMember.Fields("MemberID").Value
				Response.Cookies("FullName") = rsMember.Fields("FirstName").Value & " " & rsMember.Fields("LastName").Value
				Response.Cookies("Username") = rsMember.Fields("Username").Value
			Else
				' display error message to user here
				sErrorMsg = steGetText("The Username and Password you entered is invalid")
			End If
		End If
		If sErrorMsg = "" Then
			' prevent dup posting here
			sStat = "SELECT	TaskID " &_
					"FROM	tblTaskComment " &_
					"WHERE	TaskID = " & nTaskID & " " &_
					"AND	Subject = '" & Replace(sSubject, "'", "''") & "' " &_
					"AND	Body LIKE '" & Replace(sBody, "'", "''") & "'"
			Set rsDup = adoOpenRecordset(sStat)
			If rsDup.EOF Then
				rsDup.Close
				rsDup = Empty
				' add the new post to the database
				sStat = "INSERT INTO tblTaskComment (" &_
						"	TaskID, ParentCommentID, MemberID, Subject, Body, Created" &_
						") VALUES (" &_
						nTaskID & ", " & nReplyID & ", " & nMemberID & ", '" &_
						Replace(sSubject, "'", "''") & "', '" & Replace(sBody, "'", "''") & "'," &_
						adoGetDate &_
						")"
				Call adoExecute(sStat)

				' update the comment count for the task
				sStat = "UPDATE	tblTask " &_
						"SET	CommentCount = CommentCount + 1, Modified = " & adoGetDate & " " &_
						"WHERE	TaskID = " & nTaskID
				Call adoExecute(sStat)
			Else
				rsDup.Close
				rsDup = Empty
			End If
		End If
	End If

End If

' retrieve the task synopsis here
sStat = "SELECT	tsk.TaskID, tsk.Title, tsk.Comments, tsk.PctComplete, usr.FirstName, usr.MiddleName, usr.LastName, " &_
		"		pri.PriorityName, sta.StatusName, 0 As CommentCount, " &_
		"		tsk.Created " &_
		"FROM	tblTask tsk " &_
		"INNER JOIN	tblUser usr ON tsk.UserID = usr.UserID " &_
		"INNER JOIN	tblTaskPriority pri ON pri.PriorityID = tsk.PriorityID " &_
		"INNER JOIN	tblTaskStatus sta ON sta.StatusID = tsk.StatusID " &_
		"WHERE	tsk.TaskID = " & steForm("taskid") & " " &_
		"AND	tsk.Active <> 0 " &_
		"AND	tsk.Archive = 0"
Set rsArt = adoOpenRecordset(sStat)

' retrieve the comment to reply to (if nec)
If nReplyID > 0 Then
	sStat = "SELECT	tblTaskComment.Subject, tblTaskComment.Body, tblMember.Username, " &_
			"		tblMember.FirstName, tblMember.MiddleName, tblMember.LastName, " &_
			"		tblTaskComment.Created " &_
			"FROM	tblTaskComment " &_
			"LEFT JOIN	tblMember ON tblMember.MemberID = tblTaskComment.MemberID " &_
			"WHERE	tblTaskComment.CommentID = " & nReplyID
	Set rsReply = adoOpenRecordset(sStat)
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Task" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "post" Or sErrorMsg <> "" Then %>

<P>
<FONT CLASS="articlehead"><%= rsArt.Fields("Title").Value %></FONT><BR>
<FONT CLASS="tinytext">by <%= rsArt.Fields("FirstName").Value & " " & Trim(rsArt.Fields("MiddleName").Value & " " & rsArt.Fields("LastName").Value) %> - <%= adoFormatDateTime(rsArt.Fields("Created").Value, vbLongDate) %></FONT><BR>
<FONT CLASS="tinytext"><%= rsArt.Fields("PriorityName").Value %> / <%= rsArt.Fields("StatusName").Value %></FONT>
</P>

<P>
<%= Replace(rsArt.Fields("Comments").Value, vbCrLf, "<BR>") %>
</P>

<h3><% steTxt "Post Comment" %></h3>

<% If IsObject(rsReply) Then
	If Not rsReply.EOF Then %>

<div><b>Original Message</b></div>
<br>

<table border=0 cellpadding=2 cellspacing=0 width="100%">
<tr>
	<td><img src="../../../../img/pixel.gif" width=20 height=1></td>
	<td class="commenthead">
	<div class="commentsubject"><%= rsReply.Fields("Subject").Value %></div>
	<font class="commentauthor"><%= rsReply.Fields("Created").Value %> - <% If Trim(rsReply.Fields("Username").Value & "") = "" Then %><% steTxt "Anonymous Coward" %><% Else %><%= rsReply.Fields("Username").Value %><% End If %></font>
	</td>
</tr>
<tr>
	<td></td>
	<td class="comment">
	<%= Replace(rsReply.Fields("Body").Value, vbCrLf, "<BR>") %><BR>
	<hr noshade style="color:#C0C0C0" size="1" width="100%">
	</td>
</tr>
</table>
<%
	rsReply.Close
	rsReply = Empty
	Else
		nReplyID = 0
	End If
   End If
%>

<% If sErrorMsg <> "" Then %>
<P><b class="error"><%= sErrorMsg %></b></p>
<% End If %>

<form method="post" action="comment_post.asp">
<input type="hidden" name="action" value="post">
<input type="hidden" name="TaskID" value="<%= nTaskID %>">
<input type="hidden" name="ReplyID" value="<%= nReplyID %>">

<table border=0 cellpadding=2 cellspacing=0>
<tr>
	<td><b class="forml"><% steTxt "Posted By User" %></b><br>
	<% If Request.Cookies("Username") <> "" Then %>
	<font class="formd"><%= Request.Cookies("Username") %></font>
	<% Else %>
	<font class="formd"><% steTxt "Anonymous Coward" %></font><br><br>
	<p>
	<% steTxt "We noticed that you are not logged in." %>
	<% steTxt "Please enter your member login to receive credit for your post." %>
	<% steTxt "Or you may" %>&nbsp;<a href="<%= Application("ASPNukeBasePath") %>module/account/register/register.asp"><% steTxt "register now for a free account" %></A>.
	</p>
	<table border=0 cellpadding=2 cellspacing=0>
	<TR>
		<TD class="forml"><% steTxt "Username" %><BR>
		<INPUT TYPE="text" NAME="username" VALUE="<%= steEncForm("username") %>" SIZE="16" MAXLENGTH="16" class="form" style="width:120px">
		</TD>
		<TD class="forml"><% steTxt "Password" %><BR>
		<INPUT TYPE="password" NAME="password" VALUE="" SIZE="16" MAXLENGTH="16" class="form" style="width:120px">
		</TD>
	</TR>
	</table>
	<% End If %>
	</td>
</tr><tr>
	<td><b class="forml"><% steTxt "Subject" %></b><br>
	<input type="text" name="subject" value="<%= Server.HTMLEncode(sSubject) %>" size="32" maxlength="100" class="form" style="width:440px">
	</td>
</tr><tr>
	<td><b class="forml"><% steTxt "Body" %></b><br>
	<textarea name="body" cols=80 rows=10 class="form" style="width:440px"><%= Server.HTMLEncode(sBody) %></textarea>
	</td>
</tr><tr>
	<td align="right"><br>
		<input type="submit" name="_submit" value=" <% steTxt "Post Comment" %> " class="form">
	</td>
</tr>
</table>

</form>

<% Else %>

<h3><% steTxt "Comments Posted" %></h3>

<p>
<% steTxt "Thank you for submitting your comments for this task." %>&nbsp;
<% steTxt "The new comments you posted should show up right away." %>&nbsp;
<% steTxt "Please remember to keep your posts clean and topical." %>&nbsp;
<% steTxt "Abuse of this privelege will not be tolerated on this web site." %>
</p>

<% End If %>

<p align="center">
	<a href="task_comments.asp?TaskID=<%= nTaskID %>" class="adminlink"><% steTxt "Comments" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->