<!-- #include file="../../../lib/ado_lib.asp" -->
<!-- #include file="../../../lib/module_lib.asp" -->
<!-- #include file="../../../lib/lang_lib.asp" -->
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

Dim rsArt
Dim sArtStat
Dim nKeepDays
Dim sCritCutoff

If CStr(modParam("Articles", "HomeKeepArticles")) <> "" And CStr(modParam("Articles", "HomeKeepArticles")) <> "0" Then
	nTopNum = modParam("Articles", "HomeKeepArticles")
Else
	nTopNum = 10
	nKeepDays = modParam("Articles", "HomePageDays")
	If CStr(nKeepDays) = "" Then nKeepDays = 7
	sCritCutoff = "AND	art.Created > '" & DateAdd("d", -nKeepDays, Now()) & "' "
End If

sArtStat = "SELECT	" & adoTop(nTopNum) & " art.ArticleID, art.Title, art.LeadIn, " &_
			"		auth.FirstName, auth.LastName, " &_
			"		cat.CategoryName, cat.IconImage, art.CommentCount, art.Created " &_
			"FROM	tblArticle art " &_
			"INNER JOIN	tblArticleAuthor auth ON art.AuthorID = auth.AuthorID " &_
			"INNER JOIN	tblArticleToCategory atc ON atc.ArticleID = art.ArticleID " &_
			"INNER JOIN	tblArticleCategory cat ON atc.CategoryID = cat.CategoryID " &_
			"WHERE	art.Active <> 0 " &_
			"AND	art.Archive = 0 " &_
			sCritCutoff &_
			"ORDER BY art.Created DESC" & adoTop2(nTopNum)
Set rsArt = adoOpenRecordset(sArtStat)
%>

<% Do Until rsArt.EOF %>
<table border=0 cellpadding=0 cellspacing=10 class="articlesummary">
<tr>
	<td valign="top">
	<A HREF="<%= Application("ASPNukeBasePath") %>module/article/article/article.asp?articleid=<%= rsArt.Fields("ArticleID").Value %>" CLASS="articlehead2"><%= rsArt.Fields("Title").Value %></A><BR>
	<FONT CLASS="articleauthor"><%= rsArt.Fields("Created").Value %> - <%= rsArt.Fields("FirstName").Value & " " & rsArt.Fields("LastName").Value %></FONT><BR>
	<font class="articleleadin"><%= rsArt.Fields("LeadIn").Value %></font>
	<BR>
	<div align="right" class="articlelink">
		<A HREF="<%= Application("ASPNukeBasePath") %>module/article/article/comments.asp?articleid=<%= rsArt.Fields("ArticleID").Value %>" class="articlelink"><% steTxt "Comments" %> (<%= rsArt.Fields("CommentCount").Value %>)</A> |
		<A HREF="<%= Application("ASPNukeBasePath") %>module/article/article/article.asp?articleid=<%= rsArt.Fields("ArticleID").Value %>" class="articlelink">...(<% steTxt "More" %>)</A>
	</div>
	</td>
	<% If Trim(rsArt.Fields("IconImage").Value & "") <> "" Then %>
	<td>
	<img src="<%= Replace(Application("ASPNukeBasePath") & rsArt.Fields("IconImage").Value, "//", "/") %>" width="<%= modParam("Articles", "IconImageWidth") %>" height="<%= modParam("Articles", "IconImageHeight") %>" border="0" alt="">
	</td>
	<% End If %>
</tr>
</table>

<hr noshade width="100%" SIZE="1" style="color:#F8E8D8">
<%	rsArt.MoveNext
   Loop %>
<P ALIGN="center">
	<A HREF="<%= modParam("Articles", "RSSFeedFile") %>"><img src="<%= Application("SiteRoot") %>/img/articles/xml.gif" width=36 height=14 border=0 alt="<% steTxt "RSS Feed XML" %>"></A> .
	<A HREF="<%= Application("ASPNukeBasePath") %>module/article/article/archive.asp" class="commentlink"><% steTxt "Older Articles" %></A>
</P>