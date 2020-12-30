<!-- #include file="../../../lib/ado_lib.asp" -->
<!-- #include file="../../../lib/lang_lib.asp" -->
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

Dim rsArt
Dim sStat

sStat = "SELECT	" & adoTop(3) & " art.ArticleID, art.Title, art.LeadIn, " &_
			"		auth.FirstName, auth.LastName, " &_
			"		cat.CategoryName, art.Created " &_
			"FROM	tblArticle art " &_
			"INNER JOIN	tblArticleAuthor auth ON art.AuthorID = auth.AuthorID " &_
			"INNER JOIN	tblArticleToCategory atc ON atc.ArticleID = art.ArticleID " &_
			"INNER JOIN	tblArticleCategory cat ON atc.CategoryID = cat.CategoryID " &_
			"WHERE	art.Active <> 0 " &_
			"AND	art.Archive = 0 " &_
			"ORDER BY art.Created DESC" & adoTop2(3)
Set rsArt = adoOpenRecordset(sStat)
%>

<%= Replace(Application("ModCapHeader"), "<CAPSULETITLE/>", steGetText("News Articles")) %>
<%= Application("ModCapLeft") %>
<% Do Until rsArt.EOF %>
<P>
<A HREF="../../../news/article.asp?articleid=<%= rsArt.Fields("ArticleID").Value %>"><%= rsArt.Fields("Title").Value %></A><BR>
<FONT CLASS="tinytext"><%= rsArt.Fields("Created").Value %> - <%= rsArt.Fields("FirstName").Value & " " & rsArt.Fields("LastName").Value %></FONT><BR>
<%= rsArt.Fields("LeadIn").Value %>
<A HREF="../../../news/article.asp?articleid=<%= rsArt.Fields("ArticleID").Value %>">( <% steTxt "more" %> )</A>
</P>
<%	rsArt.MoveNext
   Loop %>
<P ALIGN="center">
	<A HREF="../../../news/archive.asp" class="footerlink"><% steTxt "More Articles" %></A>
</P>
<%= Application("ModCapRight") %>
<%= Application("ModCapFooter") %>