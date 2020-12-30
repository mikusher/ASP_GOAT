<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' detail.asp
'	Place a vote / or view the details for the poll.
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
Dim rsArt
Dim sStat
Dim sAction		' action to be performed
Dim nAnswerID	' answer the user voted
Dim nPollID		' poll voted on
Dim nTotal		' total votes on all items
Dim rsPoll		' poll from the database
Dim rsAns		' answer from the database
Dim sQuestion	' question from the database
Dim rsComment	' comments on the poll
Dim aComment	' array of comment results
Dim nPct		' percentage for bar to use
Dim sReferrer	' referring page
Dim sErrorMsg	' error to display to user

Const DEF_INDENT_SIZE = 15

sAction = steForm("Action")
nPollID = steNForm("PollID")
nAnswerID = steNForm("AnswerID")
sReferrer = steForm("Referrer")
If sReferrer = "" And Not (InStr(1, Request.ServerVariables("HTTP_REFERER"), "/poll/comment_post.asp") > 0) Then
	sReferrer = Request.ServerVariables("HTTP_REFERER")
End If

If sAction = "vote" Then
	If Not (nPollID > 0) Then
		sErrorMsg = steGetText("Invalid Poll ID Specified") & " (PollID = " & nPollID & ")"
	ElseIf Not (nAnswerID > 0) Then
		sErrorMsg = steGetText("Invalid Answer ID Specified") & " (AnswerID = " & nAnswerID & ")"
	Else
		' clean up all "old poll logs" older than 24 hours
		Dim dtYesterday
		dtYesterday = DateAdd("d", -1, Now())
		sStat = "DELETE FROM tblPollIPAddress WHERE Created < '" & dtYesterday & "'"
		Call adoExecute(sStat)

		' check to see if user has voted already
		sStat = "SELECT Created " &_
				"FROM	tblPollIPAddress " &_
				"WHERE	PollID = " & nPollID & " " &_
				"AND	IPAddress = '" & Replace(Request.ServerVariables("REMOTE_HOST"), "'", "''") & "'"
		Set rsPoll = adoOpenRecordset(sStat)
		If rsPoll.EOF Then
			' user is allowed to vote, no votge in past 24 hours
			sStat = "UPDATE	tblPollAnswer " &_
					"SET	Votes = Votes + 1, " &_
					"		Modified = " & adoGetDate & " " &_
					"WHERE	PollID = " & nPollID & " " &_
					"AND	AnswerID = " & nAnswerID
			Call adoExecute(sStat)

			sStat = "INSERT INTO tblPollIPAddress (" &_
					"	PollID, IPAddress, Modified, Created" &_
					") VALUES (" &_
					nPollID & ", '" & Replace(Request.ServerVariables("REMOTE_HOST"), "'", "''") & "'," &_
					adoGetDate & "," & adoGetDate &_
					")"
			Call adoExecute(sStat)
		Else
			sErrorMsg = steGetText("Your vote has been recorded")
		End if
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

' retrieve all of the comments posted
sStat = "SELECT	c.CommentID, c.ParentCommentID, c.Subject, c.Body, c.Created, " &_
		"		m.Username " &_
		"FROM	tblPollComment c " &_
		"LEFT JOIN tblMember m ON m.MemberID = c.MemberID " &_
		"WHERE	c.PollID = " & nPollID & " " &_
		"ORDER BY c.Created"
Set rsComment = adoOpenRecordset(sStat)
If Not rsComment.EOF Then aComment = rsComment.GetRows
rsComment.Close
rsComment = Empty
%>
<!-- #include file="../../../header.asp" -->

<H3><%= sQuestion %></H3>

<% If sErrorMsg <> "" Then %>
<P><B class="error"><%= sErrorMsg %></B></P>
<% ElseIf nAnswerID > 0 Then %>
<P><% steTxt "Your vote has been recorded, thanks for participating" %></P>
<% End If %>

<TABLE BORDER=0 CELLPADDING=5 CELLSPACING=0 WIDTH="100%" CLASS="list">
<TR>
	<TD CLASS="listhead"><% steTxt "Answer" %></TD>
	<TD CLASS="listhead"><% steTxt "Result" %></TD>
</TR>
<% Do Until rsAns.EOF
	If nTotal = 0 Then
		nPct = 0
	Else
		nPct = CInt(100 * rsAns.Fields("Votes").Value / nTotal)
	End If %>
<TR>
	<TD nowrap><%= rsAns.Fields("Answer").Value %></TD>
	<TD ALIGN="left" width="100%">
	<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
	<TR>
		<TD BGCOLOR="#C0E080" WIDTH="<%= nPct %>%"><IMG SRC="../../../img/pixel.gif" width=1 height=1 ALT=""></TD>
		<TD WIDTH="<%= 100 - nPct %>%">&nbsp;<%= rsAns.Fields("Votes").Value %>&nbsp;(<%= nPct %>%)</TD>
	</TR>
	</TABLE>
	</TD>
</TR>
<%	rsAns.MoveNext
   Loop %>
</TABLE>

<div><a href="comment_post.asp?pollid=<%= nPollID %>" class="commentlink"><% steTxt "Post Comment" %></A></div>

<hr noshade style="color:#C0C0C0" size="1" width="100%"><BR>

<% If IsArray(aComment) Then
	locComment aComment
Else %>
<P><b class="error"><% steTxt "No comments have been posted for this poll yet" %></b></p>
<% End If %>

<% If sReferrer <> "" Then %>
<P ALIGN="center">
	<A HREF="<%= sReferrer %>" CLASS="footerlink"><% steTxt "Previous Page" %></A>
</P>
<% End If %>

<!-- #include file="../../../footer.asp" -->
<%
'----------------------------------------------------------------------------
' locCommentLevel
'	Output a list of comments
'	Calls itself recursively to do nested comment layout

Sub locCommentLevel(oMesg, nParentID, ByVal nLevelNo)
	Dim aMesg

	aMesg = Split(Mid(oMesg(CStr(nParentID)), 2), ",")
	With Response
		' build the proper identing for this level
		If nLevelNo > 0 Then
			.Write "<table border=0 cellpadding=0 cellspacing=0 width=""100%"">" & vbCrLf
			.Write "<tr>" & vbCrLf
			.Write "<td width=""" & DEF_INDENT_SIZE & """><img src=""../../../img/pixel.gif"" width=""" &_
					DEF_INDENT_SIZE & """>"
			.Write "</td>" & vbCrLf
			.Write "	<td width=""100%"">"
		End If

		' iterate over all comments at this level
		For I = 0 To UBound(aMesg)
			' show the current message
			.Write oMesg("M" & aMesg(I)) & vbCrLf

			' check for any children
			If oMesg.Exists(aMesg(I)) Then
				If oMesg.Item(aMesg(I)) <> "" Then
					Call locCommentLevel(oMesg, aMesg(I), nLevelNo + 1)
				End If
			End If	
		Next

		If nLevelNo > 0 Then
			.Write " </td>" & vbCrLf
			.Write "</tr>" & vbCrLf
			.Write "</table>"
		End If
	End With
End Sub

'----------------------------------------------------------------------------
' locComment
'	Display all of the comments using a nested syntax
' TODO: paging of comments

Sub locComment(aComment)
	Dim I, sUsername, oMesg

	Set oMesg = Server.CreateObject("Scripting.Dictionary")
	For I = 0 To UBound(aComment, 2)
		' build the list of comment IDs
		oMesg.Item(CStr(aComment(1, I))) = oMesg.Item(CStr(aComment(1, I))) & "," & CStr(aComment(0, I))
		If Trim(aComment(5, I) & "") = "" Then
			sUsername = "Anonymous Coward"
		Else
			sUsername = aComment(5, I)
		End If
		oMesg.Item("M" & aComment(0, I)) = "<table border=0 cellpadding=2 cellspacing=0 width=""100%"">" & vbCrLf &_
			"<tr><td class=""commenthead"">" & vbCrLf &_
			"<div class=""commentsubject"">" & aComment(2, I) & "</div>" & vbCrLf &_
			"<font class=""commentauthor"">" & aComment(4, I) & " - " & sUsername & "</font>" & vbCrLf &_
			"</td></tr>" & vbCrLf &_
			"<tr><td class=""comment"">" & vbCrLf &_
			Replace(aComment(3, I), vbCrLf, "<BR>") & "<BR>" & vbCrLf &_
			"<div align=""right""><a href=""comment_post.asp?pollid=" & nPollID & "&replyid=" & aComment(0, I) & """ class=""commentlink"">" & steGetText("Reply") & "</A></div>" & vbCrLf &_
			"<hr noshade style=""color:#C0C0C0"" size=""1"" width=""100%"">" & vbCrLf &_
			"</td></tr>" & vbCrLf &_
			"</table>"
	Next

	' output the comments here (indenting where necessary)
	Call locCommentLevel(oMesg, 0, 0)
End Sub
%>