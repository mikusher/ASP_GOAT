<%
'--------------------------------------------------------------------
' capsule.asp
'	Create the news article capsule which will appear on all pages of
'	the site.
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

sArtStat = "SELECT	" & adoTop(3) & " tblArticle.ArticleID, tblArticle.Title, tblArticle.LeadIn, " &_
			"		tblArticleAuthor.FirstName, tblArticleAuthor.LastName, " &_
			"		tblArticleCategory.CategoryName, tblArticle.Created " &_
			"FROM	tblArticle " &_
			"INNER JOIN	tblArticleAuthor ON tblArticle.AuthorID = tblArticleAuthor.AuthorID " &_
			"INNER JOIN	tblArticleCategory ON tblArticle.CategoryID = tblArticleCategory.CategoryID " &_
			"WHERE	tblArticle.Active <> 0 " &_
			"AND	tblArticle.Archive = 0 " &_
			"ORDER BY tblArticle.Created DESC" & adoTop2(3)
Set rsArtArt = adoOpenRecordset(sArtStat)
%>

<%= Replace(Application("ModCapHeader"), "<CAPSULETITLE/>", "News Articles") %>
<%= Application("ModCapLeft") %>
<% Do Until rsArtArt.EOF %>
<P>
<A HREF="../../news/article.asp?articleid=<%= rsArtArt.Fields("ArticleID").Value %>"><%= rsArtArt.Fields("Title").Value %></A><BR>
<FONT CLASS="tinytext"><%= rsArtArt.Fields("Created").Value %> - <%= rsArtArt.Fields("FirstName").Value & " " & rsArtArt.Fields("LastName").Value %></FONT><BR>
<%= rsArtArt.Fields("LeadIn").Value %>
<A HREF="../../news/article.asp?articleid=<%= rsArtArt.Fields("ArticleID").Value %>">( more )</A>
</P>
<%	rsArtArt.MoveNext
   Loop %>
<P ALIGN="center">
	<A HREF="../../news/archive.asp">More Articles</A>
</P>
<%= Application("ModCapRight") %>
<%= Application("ModCapFooter") %>