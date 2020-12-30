<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/sha256.asp" -->
<%
'--------------------------------------------------------------------
' comment_post.asp
'	Enter a comment on an poll.
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
Dim nPollID
Dim rsMember
Dim rsAns
Dim rsPoll
Dim sQuestion
Dim nMemberID
Dim rsReply
Dim rsDup
Dim sSubject
Dim sBody
Dim nReplyID

nPollID = steNForm("PollID")
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
			sStat = "SELECT	PollID " &_
					"FROM	tblPollComment " &_
					"WHERE	PollID = " & nPollID & " " &_
					"AND	Subject = '" & Replace(sSubject, "'", "''") & "' " &_
					"AND	Body LIKE '" & Replace(sBody, "'", "''") & "'"
			Set rsDup = adoOpenRecordset(sStat)
			If rsDup.EOF Then
				rsDup.Close
				rsDup = Empty
				' add the new post to the database
				sStat = "INSERT INTO tblPollComment (" &_
						"	PollID, ParentCommentID, MemberID, Subject, Body, Created" &_
						") VALUES (" &_
						nPollID & ", " & nReplyID & ", " & nMemberID & ", '" &_
						Replace(sSubject, "'", "''") & "', '" & Replace(sBody, "'", "''") & "'," &_
						adoGetDate &_
						")"
				Call adoExecute(sStat)

				' update the comment count for the poll
				' DISABLED FOR NOW - no comment count tracked
				' sStat = "UPDATE	tblPoll " &_
				'		"SET	CommentCount = CommentCount + 1, Modified = " & adoGetDate & " " &_
				'		"WHERE	PollID = " & nPollID
				' Call adoExecute(sStat)
			Else
				rsDup.Close
				rsDup = Empty
			End If
		End If
	End If

End If

' retrieve the question here
sStat = "SELECT	tblPoll.PollID, tblPoll.Question " &_
			"FROM	tblPoll " &_
			"WHERE	PollID = " & nPollID & " " &_
			"AND	tblPoll.Active <> 0 " &_
			"AND	tblPoll.Archive = 0"
Set rsPoll = adoOpenRecordset(sStat)
If Not rsPoll.EOF Then
	sQuestion = rsPoll.Fields("Question").Value
Else
	sQuestion = steGetText("Unable to retrieve poll question") & " (PollID = " & nPollID & ")"
End If
Set rsPoll = Nothing

' retrieve the results of the poll here
sStat = "SELECT tblPollAnswer.AnswerID, tblPollAnswer.Answer, tblPollAnswer.Votes " &_
			"FROM	tblPollAnswer " &_
			"WHERE	PollID = " & nPollID & " " &_
			"AND	Active <> 0 " &_
			"AND	Archive = 0 " &_
			"ORDER BY OrderNo"
Set rsAns = adoOpenRecordset(sStat)
nTotal = 0
If not rsAns.EOF Then
	Do Until rsAns.EOF
		nTotal = nTotal + rsAns.Fields("Votes").Value
		rsAns.MoveNext
	Loop
	rsAns.MoveFirst
End If


' retrieve the comment to reply to (if nec)
If nReplyID > 0 Then
	sStat = "SELECT	tblPollComment.Subject, tblPollComment.Body, tblMember.Username, " &_
			"		tblMember.FirstName, tblMember.MiddleName, tblMember.LastName, " &_
			"		tblPollComment.Created " &_
			"FROM	tblPollComment " &_
			"LEFT JOIN	tblMember ON tblMember.MemberID = tblPollComment.MemberID " &_
			"WHERE	tblPollComment.CommentID = " & nReplyID & " " &_
			"AND	tblPollComment.Archive = 0 "
	Set rsReply = adoOpenRecordset(sStat)
End If
%>
<!-- #include file="../../../header.asp" -->

<% If steForm("action") <> "post" Or sErrorMsg <> "" Then %>

<H3><%= sQuestion %></H3>

<P>
<TABLE BORDER=0 CELLPADDING=5 CELLSPACING=0 WIDTH="100%">
<TR bgcolor="#E0B070">
	<TH ALIGN="left"><% steTxt "Answer" %></TH>
	<TH ALIGN="left"><% steTxt "Result" %></TH>
</TR>
<% Do Until rsAns.EOF
	If nTotal = 0 Then
		nPct = 0
	Else
		nPct = CInt(100 * rsAns.Fields("Votes").Value / nTotal)
	End If %>
<TR>
	<TD><NOBR><%= rsAns.Fields("Answer").Value %></NOBR></TD>
	<TD ALIGN="left" width="100%">
	<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
	<TR>
		<TD BGCOLOR="#C0E080" WIDTH="<%= nPct %>%"><IMG SRC="../../../img/pixel.gif" width=1 height=1></TD>
		<TD WIDTH="<%= 100 - nPct %>%">&nbsp;<%= rsAns.Fields("Votes").Value %>&nbsp;(<%= nPct %>%)</TD>
	</TR>
	</TABLE>
	</TD>
</TR>
<%	rsAns.MoveNext
   Loop %>
</TABLE>
</P>

<h3><% steTxt "Post Comment" %></h3>

<p>
<% steTxt "Please enter your new poll comment using the form below." %>&nbsp;
<% steTxt "The synopsis of the poll is shown below for your reference." %>
</p>

<% If sErrorMsg <> "" Then %>
<P><b class="error"><%= sErrorMsg %></b></p>
<% End If %>

<form method="post" action="comment_post.asp">
<input type="hidden" name="action" value="post">
<input type="hidden" name="PollID" value="<%= nPollID %>">
<input type="hidden" name="ReplyID" value="<%= nReplyID %>">

<table border=0 cellpadding=2 cellspacing=0>
<tr>
	<td><b class="forml"><% steTxt "Posted By User" %></b><br>
	<% If Request.Cookies("Username") <> "" Then %>
	<font class="formd"><%= Request.Cookies("Username") %></font>
	<% Else %>
	<font class="formd"><% steTxt "Anonymous Coward" %></font><br><br>
	<p>
	<% steTxt "We noticed that you are not logged in, Please enter your member login to	receive credit for your post." %>
	<% steTxt "Or you may" %> <a href="../../../account/register.asp"><% steTxt "register now for a free account" %></A>.
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

<h3><% steTxt "Poll Comment Posted" %></h3>

<p>
<% steTxt "Thank you for submitting your comments for this poll." %>&nbsp;
<% steTxt "The new comments you posted should show up right away." %>&nbsp;
<% steTxt "Please remember to keep your posts clean and topical." %>&nbsp;
<% steTxt "Abuse of this privelege will not be tolerated on this web site." %>
</p>

<% End If %>

<p align="center">
	<a href="detail.asp?pollid=<%= nPollID %>" class="commentlink"><% steTxt "Comments" %></a> |
	<a href="javascript:history.go(-2)" class="commentlink"><% steTxt "Back" %></a>
</p>

<!-- #include file="../../../footer.asp" -->