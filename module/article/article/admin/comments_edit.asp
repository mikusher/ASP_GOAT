<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' comments_edit.asp
'	Edit an existing comment on an article
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
Dim rsComment
Dim rsDup
Dim sSubject
Dim sBody
Dim nCommentID

nArticleID = steNForm("ArticleID")
nCommentID = steNForm("CommentID")
sSubject = steStripForm("subject")
sBody = steStripForm("body")

If steForm("action") = "edit" Then
	' check for the required fields here
	If sSubject = "" Then
		sErrorMsg = steGetText("Please enter the subject for this post")
	ElseIf sBody = "" Then
		sErrorMsg = steGetText("Please enter the body for this post")
	Else
		' all of the req. variables were passed
		' add the new post to the database
		sStat = "UPDATE tblArticleComment SET " &_
				"Subject = '" & Replace(sSubject, "'", "''") & "', " &_
				"Body = '" & Replace(sBody, "'", "''") & "' " &_
				"WHERE	ArticleID = " & nArticleID & " " &_
				"AND	CommentID = " & nCommentID
		Call adoExecute(sStat)
	End If
End If

' retrieve the article to work with
If steForm("action") <> "edit" Or sErrorMsg <> "" Then
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
	
	' retrieve the comment to edit
	If nCommentID > 0 Then
		sStat = "SELECT	tblArticleComment.Subject, tblArticleComment.Body, tblMember.Username, " &_
				"		tblMember.FirstName, tblMember.MiddleName, tblMember.LastName, " &_
				"		tblArticleComment.Created " &_
				"FROM	tblArticleComment " &_
				"LEFT JOIN	tblMember ON tblMember.MemberID = tblArticleComment.MemberID " &_
				"WHERE	tblArticleComment.CommentID = " & nCommentID & " " &_
				"AND	tblArticleComment.ArticleID = " & nArticleID
		Set rsComment = adoOpenRecordset(sStat)
	End If
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Comments" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<p>
<font class="articlehead"><%= rsArt.Fields("Title").Value %></font><br>
<font class="tinytext"><%= adoFormatDateTime(rsArt.Fields("Created").Value, vbGeneralDate) %> - <i><%= rsArt.Fields("FirstName").Value & " " & rsArt.Fields("LastName").Value %></i><br>
<%= rsArt.Fields("LeadIn").Value %>
</p>
<%
	rsArt.Close
	rsArt = Empty
%>

<h3><% steTxt "Edit Article Comment" %></h3>

<% If IsObject(rsComment) Then
	If Not rsComment.EOF Then %>

<div><b><% steTxt "Original Message" %></b></div>
<br>

<table border=0 cellpadding=2 cellspacing=0 width="100%">
<tr>
	<td><img src="../../../../img/pixel.gif" width=20 height=1></td>
	<td class="commenthead">
	<div class="commentsubject"><%= rsComment.Fields("Subject").Value %></div>
	<font class="commentauthor"><%= rsComment.Fields("Created").Value %> - <% If Trim(rsComment.Fields("Username").Value & "") = "" Then %><% steTxt "Anonymous Coward" %><% Else %><%= rsComment.Fields("Username").Value %><% End If %></font>
	</td>
</tr>
<tr>
	<td></td>
	<td class="comment">
	<%= Replace(rsComment.Fields("Body").Value, vbCrLf, "<BR>") %><BR>
	<hr noshade style="color:#C0C0C0" size="1" width="100%">
	</td>
</tr>
</table>
<%
	Else
		nCommentID = 0
	End If
   End If
%>

<% If sErrorMsg <> "" Then %>
<P><b class="error"><%= sErrorMsg %></b></p>
<% End If %>

<form method="post" action="comments_edit.asp">
<input type="hidden" name="action" value="edit">
<input type="hidden" name="ArticleID" value="<%= nArticleID %>">
<input type="hidden" name="CommentID" value="<%= nCommentID %>">

<table border=0 cellpadding=2 cellspacing=0>
<tr>
	<td><b class="forml"><% steTxt "Posted By User" %></b><br>
	<font class="formd"><%= steRecordEncValue(rsComment, "Username") %></font>
	</td>
</tr><tr>
	<td><b class="forml"><% steTxt "Subject" %></b><br>
	<input type="text" name="subject" value="<%= steRecordEncValue(rsComment, "subject") %>" size="32" maxlength="100" class="form" style="width:440px">
	</td>
</tr><tr>
	<td><b class="forml"><% steTxt "Body" %></b><br>
	<textarea name="body" cols=80 rows=10 class="form" style="width:440px"><%= steRecordEncValue(rsComment, "body") %></textarea>
	</td>
</tr><tr>
	<td align="right"><br>
		<input type="submit" name="_submit" value=" <% steTxt "Edit Comment" %> " class="form">
	</td>
</tr>
</table>

</form>

<% Else %>

<H3><% steTxt "Article Comment Updated" %></H3>

<P>
<% steTxt "The changes to the comments were made successfully." %>&nbsp;
<% steTxt "You may use the admin links above to continue administering the site." %>
</P>

<% End If %>

<p align="center">
	<a href="comments_list.asp?articleid=<%= nArticleID %>" class="adminlink"><% steTxt "Comments List" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->