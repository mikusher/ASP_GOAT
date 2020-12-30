<%
'--------------------------------------------------------------------
' content.asp
'	Create the news article section which will appear in the main
'	content area of the site.
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

Dim rsArtArt
Dim sArtStat

sArtStat = "SELECT	" & adoTop(10) & " tblArticle.ArticleID, tblArticle.Title, tblArticle.LeadIn, " &_
			"		tblArticleAuthor.FirstName, tblArticleAuthor.LastName, " &_
			"		tblArticleCategory.CategoryName, tblArticle.CommentCount, tblArticle.Created " &_
			"FROM	tblArticle " &_
			"INNER JOIN	tblArticleAuthor ON tblArticle.AuthorID = tblArticleAuthor.AuthorID " &_
			"INNER JOIN	tblArticleCategory ON tblArticle.CategoryID = tblArticleCategory.CategoryID " &_
			"WHERE	tblArticle.Active <> 0 " &_
			"AND	tblArticle.Archive = 0 " &_
			"ORDER BY tblArticle.Created DESC" & adoTop2(10)
Set rsArtArt = adoOpenRecordset(sArtStat)
%>

<% Do Until rsArtArt.EOF %>
<A HREF="../../news/article.asp?articleid=<%= rsArtArt.Fields("ArticleID").Value %>" CLASS="articlehead2"><%= rsArtArt.Fields("Title").Value %></A><BR>
<FONT CLASS="articleauthor"><%= rsArtArt.Fields("Created").Value %> - <%= rsArtArt.Fields("FirstName").Value & " " & rsArtArt.Fields("LastName").Value %></FONT><BR>
<font class="articleleadin"><%= rsArtArt.Fields("LeadIn").Value %></font>
<BR>
<div align="right" class="articlelink">
	<A HREF="../../news/article.asp?articleid=<%= rsArtArt.Fields("ArticleID").Value %>" class="articlelink">...(More)</A> |
	<A HREF="../../news/comments.asp?articleid=<%= rsArtArt.Fields("ArticleID").Value %>" class="articlelink">Comments (<%= rsArtArt.Fields("CommentCount").Value %>)</A>
</div>
<hr noshade width="100%" SIZE="1" style="color:#F8E8D8">
<%	rsArtArt.MoveNext
   Loop %>
<P ALIGN="center">
	<A HREF="../../news/archive.asp" class="commentlink">More Articles</A>
</P>