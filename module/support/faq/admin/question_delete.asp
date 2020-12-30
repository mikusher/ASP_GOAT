<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' question_delete.asp
'	Delete an existing FAQ question from the database
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
Dim rsQuest
Dim nDocumentID
Dim nQuestionID

nDocumentID = steNForm("DocumentID")
nQuestionID = steNForm("QuestionID")

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If steNForm("confirm") <> 1 Then
		sErrorMsg = steGetText("Please confirm the deletion of this FAQ Question")
	Else
		' delete the faq question from the database
		sStat = "DELETE FROM tblFaqQuestion " &_
				"WHERE	QuestionID = " & nQuestionID & " " &_
				"AND	DocumentID = " & nDocumentID
		Call adoExecute(sStat)
	End If
End If

If nQuestionID > 0 Then
	sStat = "SELECT * FROM tblFaqQuestion " &_
			"WHERE QuestionID = " & nQuestionID
	Set rsQuest = adoOpenRecordset(sStat)
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Questions" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete FAQ Question" %></H3>

<P>
<% steTxt "Please confirm the deletion of the FAQ question by clicking the yes button next to <I>Confirm</I>." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="question_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<input type="hidden" name="DocumentID" value="<%= nDocumentID %>">
<INPUT TYPE="hidden" NAME="QuestionID" VALUE="<%= nQuestionID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD class="forml" nowrap><% steTxt "Question" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD class="formd"><%= steRecordEncValue(rsQuest, "Question") %></TD>
</TR><TR>
	<TD valign=top class=forml nowrap><% steTxt "Answer" %></TD><TD></TD>
	<TD class="formd"><%= Replace(steRecordEncValue(rsQuest, "Answer"), vbCrLf, "<br>") %></TD>
</TR><TR>
	<TD valign=top class=forml nowrap><% steTxt "Confirm Delete?" %></TD><TD></TD>
	<TD class="formd">
		<input type="radio" name="confirm" value="1" class="formradio"> <% steTxt "Yes" %>
		<input type="radio" name="confirm" value="0" class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 align=right><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete FAQ Question" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "FAQ Question Deleted" %></H3>

<P>
<% steTxt "The FAQ entry was successfully deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
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
