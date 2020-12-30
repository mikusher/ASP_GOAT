<!-- #include file="../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' detail.asp
'	Displays the details for a FAQ document
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

Dim sStat
Dim nDocumentID
Dim rsDoc
Dim rsQuest
Dim K

nDocumentID = steNForm("DocumentID")

If nDocumentID > 0 Then
	sStat = "SELECT	Title, AuthorName, Introduction, Epilogue " &_
			"FROM	tblFaqDocument " &_
			"WHERE	DocumentID = " & nDocumentID & " " &_
			"AND	Active <> 0 " &_
			"AND	Archive = 0"
	Set rsDoc = adoOpenRecordset(sStat)
	If rsDoc.EOF Then
		sErrorMsg = "Unable to retrieve FAQ Document (ID = " & nDocumentID & ")"
	Else
		' retrieve questions
		sStat = "SELECT	Question, Answer " &_
				"FROM	tblFaqQuestion " &_
				"WHERE	DocumentID = " & nDocumentID & " " &_
				"AND	Active <> 0 " &_
				"AND	Archive = 0 " &_
				"ORDER BY OrderNo"
		Set rsQuest = adoOpenRecordset(sStat)
	End If
End If
%>
<!-- #include file="../header.asp" -->

<% If Not rsDoc.EOF Then %>

	<H3><%= rsDoc.Fields("Title").Value %></H3>
	
	<P>
	<%= Replace(rsDoc.Fields("Introduction").Value, vbCrLf, "<BR>") %>
	</P>
	
	<% If Not rsQuest.EOF Then
		K = 1
		Do Until rsQuest.EOF %>
	<P><B><%= K %>. <%= rsQuest.Fields("Question").Value %></B></P>
	<blockquote>
		<%= Replace(rsQuest.Fields("Answer").Value, vbCrLf, "<BR>") %>
	</blockquote>
	<%		rsQuest.MoveNext
			K = K + 1
		Loop
		rsQuest.Close
		rsQuest = Empty
	Else %>
	<P>
		<B class="error">No questions were found for this FAQ Document</B>
	</P>
	<% End If %>
	
	<P>
	<%= Replace(rsDoc.Fields("Epilogue").Value, vbCrLf, "<BR>") %>
	</P>

<% Else %>

	<h3>Error Occurred</h3>

	<% If sErrorMsg <> "" Then %>
	<P><B class="error"><%= sErrorMsg %></B></P>
	<% Else %>
	<P><B class="error">Unable to retrieve the FAQ Document</B></P>
	<% End If %>

<% End If %>

<!-- #include file="../footer.asp" -->