<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' question_edit.asp
'	Edit an existing FAQ question to the database
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

Dim sErrorMsg
Dim sStat
Dim rsQuestion
Dim nDocumentID
Dim nQuestionID

nDocumentID = steNForm("documentid")
nQuestionID = steNForm("questionid")

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("question")) = ""	Then
		sErrorMsg = steGetText("Please enter the Question for this FAQ entry")
	ElseIf Trim(steForm("answer")) = "" Then
		sErrorMsg = steGetText("Please enter the Answer for this FAQ entry")
	Else
		' update the FAQ question in the database
		sStat = "UPDATE tblFaqQuestion " &_
				"SET	Question = " & steQForm("Question") & "," &_
				"		Answer = " & steQForm("Answer") & " " &_
				"WHERE	QuestionID = " & nQuestionID
		Call adoExecute(sStat)
	End If
End If

' retrieve the question to edit
sStat = "SELECT	* FROM tblFaqQuestion WHERE QuestionID = " & nQuestionID
Set rsQuestion = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Questions" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit FAQ Question" %></H3>

<P>
<% steTxt "Please make your changes to the FAQ question using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="question_edit.asp">
<input type="hidden" name="DocumentID" value="<%= nDocumentID %>">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="questionid" VALUE="<%= nQuestionID %>">
<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD class="forml" nowrap><% steTxt "Question" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="Question" VALUE="<%= steRecordEncValue(rsQuestion, "Question") %>" SIZE="32" MAXLENGTH="255" class="form" style="width:440px"></TD>
</TR><TR>
	<TD CLASS="forml" valign="top" nowrap><% steTxt "Answer" %></TD><TD></TD>
	<TD><textarea NAME="Answer" cols=80 rows=14 class="form" style="width:440px"><%= steRecordEncValue(rsQuestion, "Answer") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3 align="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update FAQ Question" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "FAQ Question Updated" %></H3>

<P>
<% steTxt "The FAQ question was successfully updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="question_list.asp?documentid=<%= nDocumentID %>" class="adminlink"><% steTxt "Questions" %></A> &nbsp;
	<A HREF="document_list.asp" class="adminlink"><% steTxt "Document List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->
