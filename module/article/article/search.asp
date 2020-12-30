<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' search.asp
'	Perform a search on the articles.
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

Const DEF_PAGE_SIZE = 20

Dim rsArt
Dim sStat
Dim sAction		' action to be performed
Dim rsSearch	' search results for the articles
Dim rsAns		' answer from the database
Dim aWord		' keywords to search for
Dim sWhere		' where clause
Dim nPageNo		' page number
Dim nResults	' total results for the search
Dim sParams
Dim aResult		' array of results from search

sAction = steForm("Action")
nPageNo = steNForm("PageNo")
nResults = steNForm("Results")
sKeywords = steForm("Keywords")

If sAction = "search" Then
	If Trim(sKeywords) = "" Then
		sErrorMsg = steGetText("Please enter valid search terms")
	Else
		' split up the keywords and build the where clause
		aWord = Split(sKeywords, " ")
		For I = 0 To UBound(aWord)
			sWhere = sWhere & " OR art.Title LIKE '%" & Replace(aWord(I), "'", "''") & "%'" &_
				"OR art.ArticleBody LIKE '%" & Replace(aWord(I), "'", "''") & "%'"
		Next

		' perform the search of the articles
		sStat = "SELECT	art.ArticleID, art.Title, art.LeadIn, " &_
				"		auth.FirstName, auth.LastName, " &_
				"		cat.CategoryName, art.Created " &_
				"FROM	tblArticle art " &_
				"INNER JOIN	tblArticleAuthor auth ON art.AuthorID = auth.AuthorID " &_
				"INNER JOIN	tblArticleToCategory atc ON atc.ArticleID = art.ArticleID " &_
				"INNER JOIN	tblArticleCategory cat ON atc.CategoryID = cat.CategoryID " &_
				"WHERE	art.Active <> 0 " &_
				"AND	art.Archive = 0 AND (" & Mid(sWhere, 4) & ")"

		Set rsSearch = adoOpenRecordset(sStat)
		If Not rsSearch.EOF Then
			If nPageNo > 0 Then rsSearch.Move(nPageNo * DEF_PAGE_SIZE)
			aResult = rsSearch.GetRows(DEF_PAGE_SIZE)
		End If
		rsSearch.Close
		rsSearch = Empty
	End If
End If
%>
<!-- #include file="../../../header.asp" -->

<H3><% steTxt "Article Search Results" %></H3>

<form method="post" action="search.asp">
<input type="hidden" name="action" value="search">

<table border=0 cellpadding=5 cellspacing=0>
<tr>
	<td><b class="forml"><% steTxt "Search For" %></b></td>
	<td><input type="text" name="Keywords" value="<%= steEncForm("keywords") %>" size="32" maxlength="100" class="form"></td>
	<td><input type="submit" name="_submit" value=" <% steTxt "GO" %> " class="form"></td>
</tr>
</table>
</form>

<% If sErrorMsg <> "" Then %>
<P><B class="error"><%= sErrorMsg %></B></P>
<% End If %>

<% If sAction = "search" Then %>

<% If IsArray(aResult) Then %>

<p align="center"><b><% steTxt "Displaying results" %><%= " " & (nPageNo * DEF_PAGE_SIZE + 1) & " " %><% steTxt " to " %><%= (nPageNo * DEF_PAGE_SIZE + UBound(aResult, 2) + 1) & " " %><% steTxt "of" %><%= " " & (UBound(aResult, 2) + 1) %></b></p>

<%= locPageNav(sKeywords, sWhere, DEF_PAGE_SIZE, nResults, nPageNo) %>
<% sParams = "action="&sAction&"&pageno="&nPageNo&"&results="&nResults&"&keywords="&Server.URLEncode(sKeywords) %>

<table border=0 cellpadding=2 cellspacing=0 width="100%">
<% For I = 0 To UBound(aResult, 2) %>
<tr>
	<td valign=top><%= I + nPageNo * DEF_PAGE_SIZE + 1 %>.&nbsp;&nbsp;</td>
	<td>
	<a href="article.asp?articleid=<%= aResult(0, I) %>&src=search&<%= sParams %>" CLASS="articlehead2"><%= aResult(1, I) %></a><br>
	<font class="tinytext">
		<%= adoFormatDateTime(aResult(6, I), vbGeneralDate) %> - <i><%= aResult(3, I) & " " & aResult(4, I) %></i>
	</font><br>
	<%= aResult(2, I) %>
	</td>
</tr><tr>
	<td>&nbsp;</td>
</tr>
<% Next %>
</table>

<% Else %>

<p><b class="error"><% steTxt "No articles matched your search criteria" %></b></p>

<% End If %>

<% End If %>

<P ALIGN="center">
	<A HREF="javascript:history.back()" class="footerlink"><% steTxt "Back to Previous Page" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
<%
Function locPageNav(sKeywords, sWhere, nPageSize, nResults, nPageNo)
	Dim sStat, rsCount, nNumPages, sHTML, I

	' count the results returned (if needed)
	If nResults = 0 Then
		sStat = "SELECT	COUNT(*) AS ResultCount " &_
				"FROM	tblArticle art " &_
				"INNER JOIN	tblArticleAuthor auth ON art.AuthorID = auth.AuthorID " &_
				"INNER JOIN	tblArticleToCategory atc ON atc.ArticleID = art.ArticleID " &_
				"INNER JOIN	tblArticleCategory cat ON atc.CategoryID = cat.CategoryID " &_
				"WHERE	art.Active <> 0 " &_
				"AND	art.Archive = 0 AND (" & Mid(sWhere, 4) & ")"
		Set rsCount = adoOpenRecordset(sStat)
		If rsCount.EOF Then
			locPageNav = steGetText("Failed to count search results")
			Exit Function
		End If
		nResults = rsCount.Fields("ResultCount").Value
		rsCount.Close
		rsCount = Empty
	End If

	nNumPages = Int((nResults - 1) / nPageSize) + 1
	If nNumPages = 0 Then
		locPageNav = ""
		Exit Function
	End If
	For I = 1 To nNumPages
		If I > 1 Then sHTML = sHTML & " | "
		If nPageNo = I - 1 Then
			sHTML = sHTML & "<B>" & I & "</B>"
		Else
			sHTML = sHTML & "<a href=""search.asp?keywords=" & Server.URLEncode(sKeywords) &_
				"&action=search&results=" & nResults & "&pageno=" & (I - 1) & """>" & I & "</A>"
		End If
	Next
	locPageNav = "<p align=""center""><I>" & steGetText("Page No:") & "</I> " & sHTML & "</p>"
End Function
%>