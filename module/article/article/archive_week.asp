<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' article_week.asp
'	Display a list of archived articles for the specified week.
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
Dim nWeek
Dim nYear
Dim nMonth
Dim sMonthName
Dim sErrorMsg

nWeek = Request.QueryString("Week")
nYear = Request.QueryString("Year")
nMonth = Request.QueryString("Month")
If IsNumeric(nWeek) Then nWeek = CInt(nWeek) Else nWeek = 0
If IsNumeric(nMonth) Then nMonth = CInt(nMonth) Else nMonth = 0
If IsNumeric(nYear) Then nYear = CInt(nYear) Else nYear = 0

If nWeek <> 0 And nYear <> 0 And nMonth <> 0 Then
	sMonthName = MonthName(nMonth)
	sStat = "SELECT	art.ArticleID, art.Title, art.LeadIn, " &_
			"		auth.FirstName, auth.LastName, " &_
			"		cat.CategoryName, art.CommentCount, art.Created " &_
			"FROM	tblArticle art " &_
			"INNER JOIN	tblArticleAuthor auth ON art.AuthorID = auth.AuthorID " &_
			"INNER JOIN	tblArticleToCategory atc ON atc.ArticleID = art.ArticleID " &_
			"INNER JOIN	tblArticleCategory cat ON atc.CategoryID = cat.CategoryID " &_
			"WHERE	art.Active <> 0 " &_
			"AND	art.Archive = 0 " &_
			"AND	DatePart(week, art.Created) = " & nWeek & " " &_
			"AND	Month(art.Created) = " & nMonth & " " &_
			"AND	Year(art.Created) = " & nYear & " " &_
			"ORDER BY art.Created DESC"
	Set rsArt = adoOpenRecordset(sStat)
Else
	sMonthName = ""
	sErrorMsg = steGetText("Both the year and month must be specified")
End If
%>
<!-- #include file="../../../header.asp" -->

<% If Not rsArt.EOF Then %>

<H3><% steTxt "Week" %>&nbsp;<%= nWeek %> - <%= " " & MonthName(nMonth) & " " & nYear & " " %><% steTxt "Articles" %></H3>

<% Do Until rsArt.EOF %>
<A HREF="article.asp?articleid=<%= rsArt.Fields("ArticleID").Value %>" CLASS="articlehead2"><%= rsArt.Fields("Title").Value %></A><BR>
<FONT CLASS="articleauthor"><%= rsArt.Fields("Created").Value %> - <%= rsArt.Fields("FirstName").Value & " " & rsArt.Fields("LastName").Value %></FONT><BR>
<font class="articleleadin"><%= rsArt.Fields("LeadIn").Value %></font>
<BR>
<table border=0 cellpadding=0 cellspacing=0 width="100%">
<tr>
	<td align="right" class="articlelink">
	<A HREF="article.asp?articleid=<%= rsArt.Fields("ArticleID").Value %>" class="articlelink">...(<% steTxt "More" %>)</A> |
	<A HREF="comments.asp?articleid=<%= rsArt.Fields("ArticleID").Value %>" class="articlelink"><% steTxt "Comments" %> (<%= rsArt.Fields("CommentCount").Value %>)</A>
	</td>
</tr>
</table>
<hr noshade width="100%" SIZE="1" style="color:#F8E8D8">
<%	rsArt.MoveNext
   Loop %>

<% Else %>

<H3><% steTxt "No Articles to Display" %></H3>

<P>
<% steTxt "Sorry, but no articles could be found to display here." %>&nbsp;
<% steTxt "Please use the link below to return to the article archive." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="archive.asp" class="footerlink"><% steTxt "Article Archive" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
