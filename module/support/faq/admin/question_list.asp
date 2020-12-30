<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' question_list.asp
'	Displays a list of the faq questions for the document
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
Dim rsQuest
Dim I

nDocumentID = steNForm("DocumentID")

Select Case LCase(steForm("action"))
	Case "moveup"
			Dim rsPrev, nPrevOrder

			sStat = "SELECT " & adoTop(1) & " OrderNo FROM tblFAQQuestion " &_
					"WHERE DocumentID = " & nDocumentID & " " &_
					"AND	OrderNo < " & steNForm("OrderNo") & " " &_
					"ORDER BY OrderNo DESC " & adoTop2(1)
			Set rsPrev = adoOpenRecordset(sStat)
			If Not rsPrev.EOF Then
				nPrevOrder = rsPrev.Fields("OrderNo").Value
				' increment orders above the new order no (to make room)
				sStat = "UPDATE	tblFAQQuestion " &_
						"SET	OrderNo = OrderNo + 1, Modified = " & adoGetDate & " " &_
						"WHERE	DocumentID = " & nDocumentID & " " &_
						"AND	OrderNo = " & nPrevOrder
				Call adoExecute(sStat)
	
				sStat = "UPDATE	tblFAQQuestion " &_
						"SET	OrderNo = " & nPrevOrder & ", Modified = " & adoGetDate & " " &_
						"WHERE	QuestionID = " & steNForm("QuestionID")
				Call adoExecute(sStat)
			End If
			rsPrev.Close
			rsPrev = ""
	Case "movedown"
			Dim rsNext, nNextOrder

			sStat = "SELECT " & adoTop(1) & " OrderNo FROM tblFAQQuestion " &_
					"WHERE DocumentID = " & nDocumentID & " " &_
					"AND	OrderNo > " & steNForm("OrderNo") & " " &_
					"ORDER BY OrderNo " & adoTop2(1)
			Set rsNext = adoOpenRecordset(sStat)
			If Not rsNext.EOF Then
				sNextOrder = rsNext.Fields("OrderNo").Value
				' increment orders above the new order no (to make room)
				sStat = "UPDATE	tblFAQQuestion " &_
						"SET	OrderNo = OrderNo - 1, Modified = " & adoGetDate & " " &_
						"WHERE	DocumentID = " & nDocumentID & " " &_
						"AND	OrderNo = " & sNextOrder
				Call adoExecute(sStat)
	
				sStat = "UPDATE	tblFAQQuestion " &_
						"SET	OrderNo = " & sNextOrder & ", Modified = " & adoGetDate & " " &_
						"WHERE	QuestionID = " & steNForm("QuestionID")
				Call adoExecute(sStat)
			End If
			rsNext.Close
			rsNext = ""
End Select

sStat = "SELECT	fq.OrderNo, fq.QuestionID, fq.Question, fq.Answer, " &_
		"		fq.Created, fq.Modified " &_
		"FROM	tblFaqQuestion fq " &_
		"WHERE	fq.DocumentID = " & nDocumentID & " " &_
		"ORDER BY fq.OrderNo"
Set rsQuest = adoOpenRecordset(sStat)
If nDocumentID = 0 Then
	sStat = "SELECT	fd.DocumentID, fd.Title " &_
			"FROM	tblFaqDocument fd " &_
			"WHERE	fd.Active <> 0 " &_
			"AND	fd.Archive = 0 " &_
			"ORDER BY fd.Title"
	Set rsDoc = adoOpenRecordset(sStat)
End If

%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Questions" %>
<!-- #include file="pagetabs_inc.asp" -->

<script language="javascript" type="text/javascript">
function pickDoc(nDocID) {
	if (nDocID != '0')
		location.href="question_list.asp?documentid=" + nDocID;
}
</script>

<H3><% steTxt "FAQ Question List" %></H3>

<% If nDocumentID = 0 Then %>

<p>
<% steTxt "Please select a FAQ document to work with from the list below." %>
</p>

<p>
<form method="post" action="question_list.asp">
<table border=0 cellpadding=2 cellspacing=0>
<tr>
	<td class="forml"><% steTxt "Document" %></td><Td>&nbsp;&nbsp;</td>
	<td><select name="DocumentID" class="form" onChange="pickDoc(this.options[this.selectedIndex].value)">
	<option value="0"> -- <% steTxt "Choose" %> --
	<% Do Until rsDoc.EOF %>
	<option value="<%= rsDoc.Fields("DocumentID").Value %>"><%= Server.HTMLEncode(rsDoc.Fields("Title").Value) %>
	<%	rsDoc.MoveNext
	   Loop %>
	</select>
</tr>
</table>
</form>

<% Else %>

<P>
<% steTxt "Shown below are all of the current FAQ questions defined in the database." %>
</P>

<% If Not rsQuest.EOF Then %>

<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 CLASS="list">
<TR>
	<TD class="listhead" ALIGN="left"><% steTxt "Order" %></TD>
	<TD class="listhead" ALIGN="left"><% steTxt "Question" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Created" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Modified" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<% I = 0
Do Until rsQuest.EOF %>
<TR CLASS="list<%= I mod 2 %>">
	<TD><%= rsQuest.Fields("OrderNo").Value %></TD>
	<TD><%= steRecordEncValue(rsQuest, "Question") %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsQuest.Fields("Created").Value, vbShortDate) %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsQuest.Fields("Modified").Value, vbShortDate) %></TD>
	<TD>
		<A HREF="question_list.asp?DocumentID=<%= nDocumentID %>&questionid=<%= rsQuest.Fields("QuestionID").Value %>&orderno=<%= rsQuest.Fields("OrderNo").Value %>&action=moveup" class="actionlink"><% steTxt "up" %></A> .
		<A HREF="question_list.asp?DocumentID=<%= nDocumentID %>&questionid=<%= rsQuest.Fields("QuestionID").Value %>&orderno=<%= rsQuest.Fields("OrderNo").Value %>&action=movedown" class="actionlink"><% steTxt "down" %></A> .
		<A HREF="question_edit.asp?DocumentID=<%= nDocumentID %>&QuestionID=<%= rsQuest.Fields("QuestionID").Value %>" class="actionlink"><% steTxt "edit" %></A> .
		<A HREF="question_delete.asp?DocumentID=<%= nDocumentID %>&QuestionID=<%= rsQuest.Fields("QuestionID").Value %>" class="actionlink"><% steTxt "delete" %></A>
	</TD>
</TR>
<%	rsQuest.MoveNext
	I = I + 1
   Loop %>
</TABLE>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No FAQ questions exist in the database" %></B></P>

<% End If %>

<% End If %>

<P ALIGN="center">
	<A HREF="document_list.asp" class="adminlink"><% steTxt "Document List" %></A> &nbsp;
	<A HREF="question_add.asp?documentid=<%= nDocumentID %>" class="adminlink"><% steTxt "Add New Question" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->