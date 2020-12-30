<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' search.asp
'	Perform a search of the ASP Nuke FAQ documents.  Only search the
'	question and answers.
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

Dim rsQuest
Dim sStat
Dim nLastDocID
Dim sSearchText

sSearchText = steForm("searchtext")

' perform a search for the search text in the FAQ questions
sStat = "SELECT	fd.DocumentID, fd.Title FAQTitle, fq.QuestionID, fq.Question, fq.Answer, " &_
		"		fa.Title, fa.FirstName, fa.MiddleName, fa.LastName " &_
		"FROM	tblFaqQuestion fq " &_
		"INNER JOIN	tblFaqDocument fd ON fq.DocumentID = fd.DocumentID " &_
		"INNER JOIN	tblFaqAuthor fa ON fd.AuthorID = fa.AuthorID " &_
		"WHERE	(fq.Question like '%" & Replace(sSearchText, "'", "''") & "%' " &_
		"		OR fq.Answer like '%" & Replace(sSearchText, "'", "''") & "%') " &_
		"AND	fq.Active <> 0 " &_
		"AND	fq.Archive = 0 " &_
		"ORDER BY fd.OrderNo, fq.OrderNo"
Set rsQuest = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->

<h3><% steTxt "FAQ Search Results" %></h3>

<form method="post" action="search.asp">
<input type="hidden" name="action" value="GO">
<p align="center">
<table border=0 cellpadding=2 cellspacing=0>
<tr>
	<td class="forml"><% steTxt "Search For" %>:</td><td>&nbsp;&nbsp;</td>
	<td><input type="text" name="searchtext" value="<%= steEncForm("searchtext") %>" size="32" maxlength="100" class="form"></td><td>&nbsp;&nbsp;</td>
	<td><input type="submit" name="_action" value=" <% steTxt "GO" %> " class="form"></td>
</tr>
</table>
</p>
</form>

<p align="center"><i><% steTxt "Searching for" %> <B><%= Server.HTMLEncode(sSearchText) %></B></i></p>

<%
If Not rsQuest.EOF Then
	nLastDocID = 0
	nQuestNo = 1
	With Response
	Do Until rsQuest.EOF
		If nLastDocID <> rsQuest.Fields("DocumentID").Value Then
			.Write "<p>" & vbCrLf
			.Write "<font class=""maintitle"">" & Server.HTMLEncode(rsQuest.Fields("FAQTitle").Value) & "</font><br>" & vbCrLf
			.Write "<font class=""articleauthor"">" & Server.HTMLEncode(Trim(rsQuest.Fields("Title").Value & " " & rsQuest.Fields("FirstName").Value) & " " & Trim(rsQuest.Fields("MiddleName").Value & " " & rsQuest.Fields("LastName").Value)) & "</font>" & vbCrLf
			.Write "</p>" & vbCrLf
			nLastDocID = rsQuest.Fields("DocumentID").Value
		End If
		.Write	"<p><b>"
		.Write nQuestNo
		.Write ". &nbsp;  "
		.Write Replace(rsQuest.Fields("Question").Value, vbCrLf, "<br>")
		.Write "</b>" & vbCrLf
		.Write "<blockquote>" & vbCrLf
		.Write Replace(rsQuest.Fields("Answer").Value, vbCrLf, "<br>")
		.Write "</blockquote></p>" & vbCrLf
		rsQuest.MoveNext
		nQuestNo = nQuestNo + 1
	Loop
	End With
	rsQuest.Close
	Set rsQuest = Nothing
Else
%>

<p><b class="error"><% steTxt "No FAQ results found" %></b></p>

<p>
Sorry, but the text you entered found no matching questions
from the various Frequently Asked Question documents.
</p>

<% End If %>

<p align="center">
	<A href="index.asp" class="footerlink"><% steTxt "FAQ Index" %></a>
</p>
<br>

<!-- #include file="../../../footer.asp" -->
