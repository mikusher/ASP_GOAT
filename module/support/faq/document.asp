<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' document.asp
'	Display a FAQ document on the page
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

Dim rsDoc
Dim rsQuest
Dim nDocumentID
Dim nQuestNo
Dim sStat
Dim sErrorMsg

nDocumentID = steNForm("DocumentID")

sStat = "SELECT	fd.DocumentID, fd.Title As DocTitle, fd.Introduction, fd.Epilogue, fd.Modified, " &_
			"		fa.Title, fa.FirstName, fa.MiddleName, fa.LastName " &_
			"FROM	tblFaqDocument fd " &_
			"INNER JOIN	tblFaqAuthor fa on fd.AuthorID = fa.AuthorID " &_
			"WHERE	fd.DocumentID = " & nDocumentID & " " &_
			"AND	fd.Active <> 0 " &_
			"AND	fd.Archive = 0"
Set rsDoc = adoOpenRecordset(sStat)
If Not rsDoc.EOF Then
	' retrieve the list of questions for this document
	sStat = "SELECT	fq.QuestionID, fq.Question, fq.Answer " &_
			"FROM	tblFaqQuestion fq " &_
			"WHERE	fq.DocumentID = " & nDocumentID & " " &_
			"AND	fq.Active <> 0 " &_
			"AND	fq.Archive = 0 " &_
			"ORDER BY fq.OrderNo"
	Set rsQuest = adoOpenRecordset(sStat)
Else
	sErrorMsg = steGetText("Sorry, document could not be found in the database") & " (ID = " & nDocumentID & ")"
End If
%>
<!-- #include file="../../../header.asp" -->

<% If sErrorMsg = "" Then %>

<p>
<font class="maintitle"><%= Server.HTMLEncode(rsDoc.Fields("DocTitle").Value) %></font><br>
<font class="articleauthor"><%= Server.HTMLEncode(Trim(rsDoc.Fields("Title").Value & " " & rsDoc.Fields("FirstName").Value) & " " & Trim(rsDoc.Fields("MiddleName").Value & " " & rsDoc.Fields("LastName").Value)) %></font>
</p>

<% If rsDoc.Fields("Introduction").Value & "" <> "" Then %>
<p>
<%= Replace(rsDoc.Fields("Introduction").Value, vbCrLf, "<br>") %>
<p>
<% End If %>

<%
nQuestNo = 1
With Response
Do Until rsQuest.EOF
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
%>

<% If rsDoc.Fields("Epilogue").Value & "" <> "" Then %>
<p>
<%= Replace(rsDoc.Fields("Epilogue").Value, vbCrLf, "<br>") %>
<p>
<% End If %>

<font class="tinytext"><% steTxt "Last Updated" %>: <%= adoFormatDateTime(rsDoc.Fields("Modified").Value, vbLongDate) %></font>
</p>

<% Else %>

<h3><% steTxt "FAQ Document Unavailable" %></h3>

<p class="error"><%= sErrorMsg %></p>

<p>
<% steTxt "Sorry, but the FAQ document you requested could not be retrieved from the database." %>&nbsp;
<% steTxt "Please refer to the error message above for the cause of this error." %>
</p>

<% End If %>


<p align="center">
	<A href="index.asp" class="footerlink"><% steTxt "FAQ Index" %></a>
</p>
<br>

<!-- #include file="../../../footer.asp" -->
