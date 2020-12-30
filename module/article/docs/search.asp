<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' search.asp
'	Perform a search on the document repository.
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
Dim sParams		' search params
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
			sWhere = sWhere & " OR doc.Title LIKE '%" & Replace(aWord(I), "'", "''") & "%'" &_
				"OR doc.Body LIKE '%" & Replace(aWord(I), "'", "''") & "%'"
		Next

		' perform the search of the documentation
		sStat = "SELECT	doc.DocID, doc.Title, doc.SubTitle, b.Version, " &_
				"		a.FirstName, a.MiddleName, a.LastName, " &_
				"		doc.Created, b.PublishDate, doc.ShortDescription, doc.OrderNo " &_
				"FROM	tblDoc doc " &_
				"INNER JOIN	tblDocAuthor a ON doc.AuthorID = a.AuthorID " &_
				"LEFT JOIN tblDocBook b ON b.BookID = doc.BookID " &_
				"WHERE	doc.Archive = 0 " &_
				"AND	doc.Active = 1 " &_
				"AND	Coalesce(b.Archive, 0) = 0 AND (" & Mid(sWhere, 4) & ") " &_
				"ORDER BY doc.OrderNo"
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

<H3><% steTxt "Document Search Results" %></H3>

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
	<td valign="top"><%= I + nPageNo * DEF_PAGE_SIZE + 1 %>.&nbsp;&nbsp;</td>
	<td width="100%">
	<a href="document.asp?docid=<%= aResult(0, I) %>&src=search&<%= sParams %>" CLASS="articlehead2"><%= aResult(1, I) %></a><br>
	<font class="tinytext">
	<% If IsDate(aResult(8, I)) Then %><%= adoFormatDateTime(aResult(8, I), vbGeneralDate) %> -<% ElseIf IsDate(aResult(7, I)) Then %><%= adoFormatDateTime(aResult(7, I), vbGeneralDate) %> -<% End If %> <i><%= aResult(4, I) & " " & aResult(5, I) & " " & aResult(6, I) %></i>
	</font><br>
	<% If Trim(aResult(2, I)&"") <> "" Then %><%= aResult(2, I) %><br><% End If %>
	<% If Trim(aResult(9, I)&"") <> "" Then %>
	<%= Replace(Server.HTMLEncode(aResult(9, I)&""), vbCrLf, "<br>") %>
	<% Else %>
	<b class="error">No description is avaliable</b>
	<% End If %>
	</td>
</tr><tr>
	<td>&nbsp;</td>
</tr>
<% Next %>
</table>

<% Else %>

<p><b class="error"><% steTxt "No documents matched your search criteria" %></b></p>

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
				"FROM	tblDoc doc " &_
				"INNER JOIN	tblDocAuthor a ON doc.AuthorID = a.AuthorID " &_
				"LEFT JOIN tblDocBook b ON b.BookID = doc.BookID " &_
				"WHERE	doc.Archive = 0 " &_
				"AND	doc.Active = 1 " &_
				"AND	Coalesce(b.Archive, 0) = 0 AND (" & Mid(sWhere, 4) & ")"
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
				"&results="&nResults&"&action=search&pageno=" & (I - 1) & """>" & I & "</A>"
		End If
	Next
	locPageNav = "<p align=""center""><I>" & steGetText("Page No:") & "</I> " & sHTML & "</p>"
End Function
%>