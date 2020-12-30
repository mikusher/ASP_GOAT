<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' comments.asp
'	Create the task comments section
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

Const DEF_INDENT_SIZE = 20		' pixels to indent each level

Dim rsArt
Dim sStat
Dim nTaskID

nTaskID = steNForm("TaskID")

' retrieve all of the comments posted
sStat = "SELECT	c.CommentID, c.ParentCommentID, c.Subject, c.Body, c.Created, " &_
		"		m.Username " &_
		"FROM	tblTaskComment c " &_
		"LEFT JOIN tblMember m ON m.MemberID = c.MemberID " &_
		"WHERE	c.TaskID = " & nTaskID & " " &_
		"ORDER BY c.Created"
Set rsComment = adoOpenRecordset(sStat)
If Not rsComment.EOF Then aComment = rsComment.GetRows
rsComment.Close
rsComment = Empty

' retrieve the task synopsis here
sStat = "SELECT	tsk.TaskID, tsk.Title, tsk.Comments, tsk.PctComplete, usr.FirstName, usr.LastName, " &_
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
%>

<!-- #include file="../../../header.asp" -->

<P>
<FONT CLASS="articlehead"><%= rsArt.Fields("Title").Value %></FONT><BR>
<FONT CLASS="tinytext">by <%= rsArt.Fields("FirstName").Value & " " & rsArt.Fields("LastName").Value %> - <%= adoFormatDateTime(rsArt.Fields("Created").Value, vbLongDate) %></FONT><BR>
<FONT CLASS="tinytext"><%= rsArt.Fields("PriorityName").Value %> / <%= rsArt.Fields("StatusName").Value %></FONT>
</P>

<P>
<%= Replace(rsArt.Fields("Comments").Value, vbCrLf, "<BR>") %>
</P>

<table border=0 cellpadding=0 cellspacing=0 width="100%">
<tr>
	<td>
	<a href="comment_post.asp?TaskID=<%= nTaskID %>" class="commentlink"><% steTxt "Post Comment" %></A>
	</td><td align="right">
	</td>
</tr>
</table>

<hr noshade style="color:#C0C0C0" size="1" width="100%"><BR>

<% If IsArray(aComment) Then
	locComment aComment
Else %>
<P><b class="error"><% steTxt "No comments have been posted for this task yet" %></b></p>
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
			.Write "<td width=""" & DEF_INDENT_SIZE & """><img src=""" & Application("ASPNukeBasePath") & "img/pixel.gif"" width=""" &_
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
			sUsername = steGetText("Anonymous Coward")
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
			"<div align=""right""><a href=""comment_post.asp?TaskID=" & nTaskID & "&replyid=" & aComment(0, I) & "&subject=" & Server.URLEncode("re: " & aComment(2, I)) & """ class=""commentlink"">" & steGetText("Reply") & "</A></div>" & vbCrLf &_
			"<hr noshade style=""color:#C0C0C0"" size=""1"" width=""100%"">" & vbCrLf &_
			"</td></tr>" & vbCrLf &_
			"</table>"
	Next

	' output the comments here (indenting where necessary)
	Call locCommentLevel(oMesg, 0, 0)
End Sub
%>