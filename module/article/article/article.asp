<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' article.asp
'	Display an individual article from the database.
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
Dim sParams

sParams = "action="&steForm("Action")&"&pageno="&steNForm("PageNo")&"&results="&steNForm("Results")&"&keywords="&Server.URLEncode(steForm("Keywords"))

sStat = "SELECT	art.ArticleID, art.Title, art.ArticleBody, " &_
		"		auth.FirstName, auth.LastName, " &_
		"		cat.CategoryName, art.CommentCount, " &_
		"		art.Created " &_
		"FROM	tblArticle art " &_
		"INNER JOIN	tblArticleAuthor auth ON art.AuthorID = auth.AuthorID " &_
		"INNER JOIN	tblArticleToCategory atc ON atc.ArticleID = art.ArticleID " &_
		"INNER JOIN	tblArticleCategory cat ON atc.CategoryID = cat.CategoryID " &_
		"WHERE	art.ArticleID = " & steForm("articleid") & " " &_
		"AND	art.Active <> 0 " &_
		"AND	art.Archive = 0"
Set rsArt = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->

<% If Not rsArt.EOF Then %>

<P>
<FONT CLASS="articlehead"><%= rsArt.Fields("Title").Value %></FONT><BR>
<FONT CLASS="tinytext"><% steTxt "by" %>&nbsp;<%= rsArt.Fields("FirstName").Value & " " & rsArt.Fields("LastName").Value %> - <%= adoFormatDateTime(rsArt.Fields("Created").Value, vbLongDate) %></FONT><BR>
<FONT CLASS="tinytext"><% steTxt "from the" %>&nbsp;<%= rsArt.Fields("CategoryName").Value %>&nbsp;<% steTxt "dept." %></FONT>
</P>

<P>
<%= Replace(rsArt.Fields("ArticleBody").Value, vbCrLf, "<BR>") %>
</P>
<div><A HREF="comments.asp?articleid=<%= steForm("articleid") %>" class="articlelink"><% steTxt "Comments" %> (<%= rsArt.Fields("CommentCount").Value %>)</A></div>
<% Else %>

<H3><% steTxt "Article No Longer Available" %></H3>

<P>
<% steTxt "Sorry, but the article that you requested is no longer available." %>&nbsp;
<% steTxt "Although we try to maintain an archive of all of our old articles," %>&nbsp;
<% steTxt "sometimes it becomes necessary to remove an article from our site." %>&nbsp;
<% steTxt "Please update your bookmarks accordingly." %>
</P>

<% End If %>

<% If steForm("src") = "search" Then %>
<p align="center">
	<a href="search.asp?<%= sParams %>" class="adminlink">&lt;&lt; <% steTxt "Back to Search" %></A>
</p>
<% End If %>

<!-- #include file="../../../footer.asp" -->
