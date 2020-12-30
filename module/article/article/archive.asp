<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' archive.asp
'	Display the article archives from the database.
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
Dim nRow
Dim sPeriod

sPeriod = modParam("Articles", "ArchivePeriod")
If sPeriod = "" Then sPeriod = "MONTHLY"

Select Case sPeriod
	Case "WEEKLY"
		' retrieve all months for the archive
		sStat = "SELECT	DatePart(week, art.Created) As PublishWeek, Month(art.Created) AS PublishMonth, Year(art.Created) As PublishYear, " &_
				"		Count(*) AS ArticleCount " &_
				"FROM	tblArticle art " &_
				"INNER JOIN	tblArticleAuthor auth ON art.AuthorID = auth.AuthorID " &_
				"INNER JOIN	tblArticleToCategory atc ON atc.ArticleID = art.ArticleID " &_
				"INNER JOIN	tblArticleCategory cat ON art.CategoryID = cat.CategoryID " &_
				"WHERE	art.Active <> 0 " &_
				"AND	art.Archive = 0 " &_
				"GROUP BY DatePart(week, art.Created), Month(art.Created), Year(art.Created)"
				' "AND	art.Created < DateAdd(m, -1, GetDate()) " &_
	Case "MONTHLY"
		' retrieve all months for the archive
		sStat = "SELECT	Month(art.Created) AS PublishMonth, Year(art.Created) As PublishYear, " &_
				"		Count(*) AS ArticleCount " &_
				"FROM	tblArticle art " &_
				"INNER JOIN	tblArticleAuthor auth ON art.AuthorID = auth.AuthorID " &_
				"INNER JOIN	tblArticleToCategory atc ON atc.ArticleID = art.ArticleID " &_
				"INNER JOIN	tblArticleCategory cat ON atc.CategoryID = cat.CategoryID " &_
				"WHERE	art.Active <> 0 " &_
				"AND	art.Archive = 0 " &_
				"GROUP BY Month(art.Created), Year(art.Created)"
				' "AND	art.Created < DateAdd(m, -1, GetDate()) " &_
	Case "YEARLY"
		' retrieve all years for the archive
		sStat = "SELECT	Year(art.Created) As PublishYear, " &_
				"		Count(*) AS ArticleCount " &_
				"FROM	tblArticle art " &_
				"INNER JOIN	tblArticleAuthor auth ON art.AuthorID = auth.AuthorID " &_
				"INNER JOIN	tblArticleToCategory atc ON atc.ArticleID = art.ArticleID " &_
				"INNER JOIN	tblArticleCategory cat ON atc.CategoryID = cat.CategoryID " &_
				"WHERE	art.Active <> 0 " &_
				"AND	art.Archive = 0 " &_
				"GROUP BY Year(art.Created)"
				' "AND	art.Created < DateAdd(m, -1, GetDate()) " &_
End Select

Set rsArt = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->

<% If Not rsArt.EOF Then %>

<h3><% steTxt "Article Archive" %></h3>

<p>
<% steTxt "Shown below is a list of all previous months from the article archive." %>&nbsp;
<% steTxt "Please click on a month to obtain a listing of articles for that month." %>
</p>

<table border=0 cellpadding=2 cellspacing=0 class="list" align="center" width="300">
<tr>
	<td class="listhead"><% steTxt "Month" %></td>
	<td class="listhead" align="right"><% steTxt "Articles" %></td>
</tr>
<% nRow = 0
Do Until rsArt.EOF
	Select Case sPeriod
		Case "WEEKLY"
			' TODO - Show the week name %>
<tr class="list<%= nRow Mod 2 %>">
	<td><A HREF="archive_week.asp?week=<%= rsArt.Fields("PublishWeek").Value %>&month=<%= rsArt.Fields("PublishMonth").Value %>&year=<%= rsArt.Fields("PublishYear").Value %>" class="articlelink"><%= locMonth(rsArt.Fields("PublishMonth").Value) & " " & rsArt.Fields("PublishYear").Value %></A></td>
	<td align="right"><%= rsArt.fields("ArticleCount").Value %></td>
</tr>
<%		Case "MONTHLY" %>
<tr class="list<%= nRow Mod 2 %>">
	<td><A HREF="archive_mon.asp?month=<%= rsArt.Fields("PublishMonth").Value %>&year=<%= rsArt.Fields("PublishYear").Value %>" class="articlelink"><%= locMonth(rsArt.Fields("PublishMonth").Value) & " " & rsArt.Fields("PublishYear").Value %></A></td>
	<td align="right"><%= rsArt.fields("ArticleCount").Value %></td>
</tr>
<%		Case "YEARLY" %>
<tr class="list<%= nRow Mod 2 %>">
	<td><A HREF="archive_year.asp?year=<%= rsArt.Fields("PublishYear").Value %>" class="articlelink"><%= rsArt.Fields("PublishYear").Value %></A></td>
	<td align="right"><%= rsArt.fields("ArticleCount").Value %></td>
</tr>
<%	End Select
	rsArt.MoveNext
	nRow = nRow + 1
   Loop %>
</table>

<% Else %>

<H3><% steTxt "Article Archive Empty" %></H3>

<P>
<% steTxt "Sorry, the article archive is currently empty and there are no articles to view here." %>
<%
Select Case sPeriod
	Case "WEEKLY" : steTxt "Articles will be archived on a week-by-week basis."
	Case "MONTHLY" : steTxt "Articles will be archived on a month-by-month basis."
	Case "YEARLY" : steTxt "Articles will be archived on a year-by-year basis."
End Select %>
</P>

<% End If %>

<!-- #include file="../../../footer.asp" -->
<%
Function locMonth(nMonth)
	If Not IsNull(nMonth) Then
		If IsNumeric(CInt(nMonth)) Then
			If (CInt(nMonth) > 0 And CInt(nMonth) < 13) Then
				locMonth = MonthName(CInt(nMonth))
			End If
		End If
	End If
End Function
%>