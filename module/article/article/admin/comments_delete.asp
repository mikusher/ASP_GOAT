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

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If steNForm("confirm") <> 1 Then
		sErrorMsg = steGetText("Please confirm the deletion of this article comment")
	Else
		' check for any child replies
		sStat = "SELECT CommentID " &_
				"FROM	tblArticleComment " &_
				"WHERE	ArticleID = " & nArticleID & " " &_
				"AND	ParentCommentID = " & nCommentID
		Set rsChild = adoOpenRecordset(sStat)
		If Not rsChild.EOF Then
			sErrorMsg = steGetText("Cannot delete a comment that has replies to it, you must delete the replies first")
			rsChild.Close : Set rsChild = Nothing
		Else
			' create the new author in the database
			rsChild.Close : Set rsChild = Nothing

			sStat = "DELETE FROM tblArticleComment " &_
					"WHERE	ArticleID = " & nArticleID & " " &_
					"AND	CommentID = " & nCommentID
			Call adoExecute(sStat)

			' update the comment count for the article
			sStat = "SELECT COUNT(*) AS CommentCount FROM tblArticleComment WHERE ArticleID = " & nArticleID & " AND Archive = 0"
			Set rsCount = adoOpenRecordset(sStat)
			If Not rsCount.EOF Then
				Call adoExecute("UPDATE tblArticle SET CommentCount = " & rsCount.Fields("CommentCount").Value & " WHERE ArticleID = " & nArticleID)
			End If
			rsCount.Close : Set rsCount = Nothing

			' update the article comment count
		End If
	End If
End If

' retrieve the article to work with
If steForm("action") <> "delete" Or sErrorMsg <> "" Then
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

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<p>
<font class="articlehead"><%= rsArt.Fields("Title").Value %></font><br>
<font class="tinytext"><%= adoFormatDateTime(rsArt.Fields("Created").Value, vbGeneralDate) %> - <i><%= rsArt.Fields("FirstName").Value & " " & rsArt.Fields("LastName").Value %></i><br>
<%= rsArt.Fields("LeadIn").Value %>
</p>
<%
	rsArt.Close
	rsArt = Empty
%>

<h3><% steTxt "Delete Article Comment" %></h3>

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

<form method="post" action="comments_delete.asp">
<input type="hidden" name="action" value="delete">
<input type="hidden" name="ArticleID" value="<%= nArticleID %>">
<input type="hidden" name="CommentID" value="<%= nCommentID %>">

<table border=0 cellpadding=2 cellspacing=0>
<tr>
	<td><b class="forml"><% steTxt "Confirm Delete" %></b><br>
	<td class="formd">
		<input type="radio" name="confirm" value="1" class="formradio"> <% steTxt "Yes" %>
		<input type="radio" name="confirm" value="0" checked class="formradio"> <% steTxt "No" %>
	</td>
</tr><tr>
	<td colspan="2" align="right"><br>
		<input type="submit" name="_submit" value=" <% steTxt "Delete Comment" %> " class="form">
	</td>
</tr>
</table>

</form>

<% Else %>

<H3><% steTxt "Article Comment Deleted" %></H3>

<P>
<% steTxt "The article comment was successfully deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align="center">
	<a href="comments_list.asp?articleid=<%= nArticleID %>" class="adminlink"><% steTxt "Comments List" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->