<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/sha256.asp" -->
<%
'--------------------------------------------------------------------
' comment_post.asp
'	Enter a comment on an article.
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
Dim nArticleID
Dim rsMember
Dim rsArt
Dim nMemberID
Dim rsReply
Dim rsDup
Dim sSubject
Dim sBody
Dim nReplyID

nArticleID = steNForm("ArticleID")
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
			sStat = "SELECT	ArticleID " &_
					"FROM	tblArticleComment " &_
					"WHERE	ArticleID = " & nArticleID & " " &_
					"AND	Subject = '" & Replace(sSubject, "'", "''") & "' " &_
					"AND	Body LIKE '" & Replace(sBody, "'", "''") & "'"
			Set rsDup = adoOpenRecordset(sStat)
			If rsDup.EOF Then
				rsDup.Close
				rsDup = Empty
				' add the new post to the database
				sStat = "INSERT INTO tblArticleComment (" &_
						"	ArticleID, ParentCommentID, MemberID, Subject, Body, Created" &_
						") VALUES (" &_
						nArticleID & ", " & nReplyID & ", " & nMemberID & ", '" &_
						Replace(sSubject, "'", "''") & "', '" & Replace(sBody, "'", "''") & "'," &_
						adoGetDate &_
						")"
				Call adoExecute(sStat)

				' update the comment count for the article
				sStat = "SELECT COUNT(*) AS CommentCount FROM tblArticleComment WHERE ArticleID = " & nArticleID & " AND Archive = 0"
				Set rsCount = adoOpenRecordset(sStat)
				If Not rsCount.EOF Then
					Call adoExecute("UPDATE tblArticle SET CommentCount = " & rsCount.Fields("CommentCount").Value & " WHERE ArticleID = " & nArticleID)
				End If
				rsCount.Close : Set rsCount = Nothing
			Else
				rsDup.Close
				rsDup = Empty
			End If
		End If
	End If

End If

' retrieve the article to work with
sStat = "SELECT	art.ArticleID, art.Title, art.LeadIn, " &_
		"		auth.FirstName, auth.LastName, " &_
		"		cat.CategoryName, art.Created " &_
		"FROM	tblArticle art " &_
		"INNER JOIN	tblArticleAuthor auth ON art.AuthorID = auth.AuthorID " &_
		"INNER JOIN	tblArticleToCategory atc ON atc.ArticleID = art.ArticleID " &_
		"INNER JOIN	tblArticleCategory cat ON atc.CategoryID = cat.CategoryID " &_
		"WHERE	art.ArticleID = " & nArticleID & " " &_
		"AND	art.Active <> 0 " &_
		"AND	art.Archive = 0"
Set rsArt = adoOpenRecordset(sStat)

' retrieve the comment to reply to (if nec)
If nReplyID > 0 Then
	sStat = "SELECT	tblArticleComment.Subject, tblArticleComment.Body, tblMember.Username, " &_
			"		tblMember.FirstName, tblMember.MiddleName, tblMember.LastName, " &_
			"		tblArticleComment.Created " &_
			"FROM	tblArticleComment " &_
			"LEFT JOIN	tblMember ON tblMember.MemberID = tblArticleComment.MemberID " &_
			"WHERE	tblArticleComment.CommentID = " & nReplyID & " " &_
			"AND	tblArticleComment.Archive = 1 "
	Set rsReply = adoOpenRecordset(sStat)
End If
%>
<!-- #include file="../../../header.asp" -->

<% If steForm("action") <> "post" Or sErrorMsg <> "" Then %>

<p>
<font class="articlehead"><%= rsArt.Fields("Title").Value %></font><br>
<font class="tinytext"><%= adoFormatDateTime(rsArt.Fields("Created").Value, vbGeneralDate) %> - <i><%= rsArt.Fields("FirstName").Value & " " & rsArt.Fields("LastName").Value %></i><br>
<%= rsArt.Fields("LeadIn").Value %>
</p>
<%
	rsArt.Close
	rsArt = Empty
%>

<h3><% steTxt "Post Comment" %></h3>

<% If IsObject(rsReply) Then
	If Not rsReply.EOF Then %>

<div><b><% steTxt "Original Message" %></b></div>
<br>

<table border=0 cellpadding=2 cellspacing=0 width="100%">
<tr>
	<td><img src="../../../img/pixel.gif" width=20 height=1></td>
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
<input type="hidden" name="ArticleID" value="<%= nArticleID %>">
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
	<% steTxt "Or you may" %> <a href="<%= Application("ASPNukeBasePath") %>module/account/register/register.asp"><% steTxt "register now for a free account" %></A>.
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

<h3><% steTxt "Comment Posted" %></h3>

<p>
<% steTxt "Thank you for submitting your comments for this article." %>&nbsp;
<% steTxt "The new comments you posted should show up right away." %>&nbsp;
<% steTxt "Please remember to keep your posts clean and topical." %>&nbsp;
<% steTxt "Abuse of this privelege will not be tolerated on this web site." %>
</p>

<% End If %>

<p align="center">
	<a href="comments.asp?articleid=<%= nArticleID %>" class="commentlink"><% steTxt "Comments" %></a> |
	<a href="javascript:history.go(-1)" class="commentlink"><% steTxt "Back" %></a>
</p>

<!-- #include file="../../../footer.asp" -->