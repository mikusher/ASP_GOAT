<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' comment_delete.asp
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

If steForm("action") = "delete" Then
	' check for the required fields here
	If steNForm("confirm") = 0 Then
		sErrorMsg = steGetText("Please confirm you want to delete this comment")
	Else
		' delete the comment from the database
		sStat = "DELETE FROM tblTaskComment " &_
				"WHERE CommentID = " & nReplyID
		Call adoExecute(sStat)

		' update the comment count for the task
		sStat = "UPDATE	tblTask " &_
				"SET	CommentCount = CommentCount - 1, Modified = " & adoGetDate & " " &_
				"WHERE	TaskID = " & nTaskID
		Call adoExecute(sStat)
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

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<P>
<FONT CLASS="articlehead"><%= rsArt.Fields("Title").Value %></FONT><BR>
<FONT CLASS="tinytext">by <%= rsArt.Fields("FirstName").Value & " " & Trim(rsArt.Fields("MiddleName").Value & " " & rsArt.Fields("LastName").Value) %> - <%= adoFormatDateTime(rsArt.Fields("Created").Value, vbLongDate) %></FONT><BR>
<FONT CLASS="tinytext"><%= rsArt.Fields("PriorityName").Value %> / <%= rsArt.Fields("StatusName").Value %></FONT>
</P>

<P>
<%= Replace(rsArt.Fields("Comments").Value, vbCrLf, "<BR>") %>
</P>

<h3><% steTxt "Delete Comment" %></h3>

<% If sErrorMsg <> "" Then %>
<P><b class="error"><%= sErrorMsg %></b></p>
<% End If %>

<% If IsObject(rsReply) Then
	If Not rsReply.EOF Then %>

<form method="post" action="comment_delete.asp">
<input type="hidden" name="action" value="delete">
<input type="hidden" name="TaskID" value="<%= nTaskID %>">
<input type="hidden" name="ReplyID" value="<%= nReplyID %>">

<table border=0 cellpadding=2 cellspacing=0 width="100%">
<tr>
	<td><img src="../../../../img/pixel.gif" width=20 height=1></td>
	<td class="commenthead" width="100%">
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
</tr><tr>
	<td></td>
	<td>
	<table border=0 cellpadding=2 cellspacing=0>
	<tr>
		<td class="forml">Confirm Delete</td>
		<td class="formd">
			<input type="radio" name="confirm" value="1" class="form"> Yes
			<input type="radio" name="confirm" value="0" class="form"> No
		</td>
	</tr>
	</table>
	</td>
</tr><tr>
	<td></td>
	<td align="center">
		<input type="submit" name="_submit" value="Delete Comment" class="form">
	</td>
</tr>
</table>
</form>

<%
	rsReply.Close
	rsReply = Empty
	Else
		nReplyID = 0
	End If
   End If
%>

<% Else %>

<h3><% steTxt "Delete Comment" %></h3>

<p>
<% steTxt "The comment was successfully deleted from the task." %>&nbsp;
<% steTxt "Please click the button below to return to the task overview." %>
</p>

<% End If %>

<p align="center">
	<a href="task_comments.asp?TaskID=<%= nTaskID %>" class="adminlink"><% steTxt "Comments" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->