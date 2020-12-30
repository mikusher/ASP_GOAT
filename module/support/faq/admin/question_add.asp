﻿<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' question_add.asp
'	Adds a new FAQ question to the database.
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
Dim nDocumentID	' document ID to work with
Dim rsDoc		' faq document newly added to database
Dim rsCat		' list of categories to choose from
Dim sCatList	' list of currently selected categories
Dim rsOrder
Dim nOrderNo
Dim sErrorMsg	' error message to display to user

nDocumentID = steNForm("DocumentID")

If UCase(steForm("action")) = "ADD" Then
	' check for required fields here
	If Trim(steForm("Question")) = "" Then
		sErrorMsg = steGetText("Please enter the question for this FAQ entry")
	ElseIf Trim(steForm("Answer")) = "" Then
		sErrorMsg = steGetText("Please enter the answer for this FAQ entry")
	Else
		' get the new order number
		sStat = "SELECT Coalesce(Max(OrderNo) + 1, 1) As OrderNo FROM tblFaqQuestion " &_
				"WHERE	DocumentID = " & nDocumentID
		Set rsOrder = adoOpenRecordset(sStat)
		If Not rsOrder.EOF Then nOrderNo = rsOrder.Fields("OrderNo").Value Else nOrderNo = 1
		rsOrder.Close
		rsOrder = Empty

		' insert the new faq document into the database
		sStat = "INSERT INTO tblFaqQuestion (" &_
				"	DocumentID, Question, Answer, OrderNo, Created" &_
				") VALUES (" &_
				steNForm("DocumentID") & "," & steQForm("Question") & "," &_
				steQForm("Answer") & "," & nOrderNo & "," & adoGetDate &_
				")"
		Call adoExecute(sStat)
	End If
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Questions" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If UCase(steForm("action")) <> "ADD" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New FAQ Question" %></H3>

<P>
<% steTxt "Please enter the information for the new FAQ question in the form below." %>
</P>

<FORM METHOD="post" ACTION="question_add.asp">
<input type="hidden" name="DocumentID" value="<%= nDocumentID %>">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Question" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="Question" VALUE="<%= steEncForm("question") %>" SIZE="32" MAXLENGTH="255" class="form" style="width:440px"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Answer" %></TD><TD></TD>
	<TD><TEXTAREA NAME="Answer" COLS="42" ROWS="14" WRAP="Virtual" class="form" style="width:440px"><%= steEncForm("Answer") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add FAQ Question" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "FAQ Question Added" %></H3>

<P>
<% steTxt "The new faq entry was successfully added." %>&nbsp;
<% steTxt "You may use the admin links above to continue administering the site." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="document_list.asp" class="adminlink"><% steTxt "Document List" %></A>
<% If nDocumentID > 0 Then %>
	&nbsp; <A HREF="question_list.asp?documentid=<%= nDocumentID %>" class="adminlink"><% steTxt "Questions" %></A>
	&nbsp; <A HREF="question_add.asp?documentid=<%= nDocumentID %>" class="adminlink"><% steTxt "Add Another" %></A>
<% End If %>
</P>

<!-- #include file="../../../../footer.asp" -->