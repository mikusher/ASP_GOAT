<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' document_edit.asp
'	Edit an existing FAQ Document from the database
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
Dim rsDoc
Dim rsAuth
Dim nDocumentID

nDocumentID = steNForm("DocumentID")

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("Title")) = "" Then
		sErrorMsg = steGetText("Please enter the title for this FAQ document")
	ElseIf steNForm("AuthorID") = 0 Then
		sErrorMsg = steGetText("Please select an author for this FAQ document")
	ElseIf Trim(steForm("Synopsis")) = "" Then
		sErrorMsg = steGetText("Please enter the synopsis for this FAQ document")
	ElseIf Trim(steForm("Introduction")) = "" Then
		sErrorMsg = steGetText("Plese enter the Introduction for this FAQ document")
	ElseIf Trim(steForm("Epilogue")) = "" Then
		sErrorMsg = steGetText("Please enter the Epilogue for this FAQ document")
	Else
		' create the author in the database
		sStat = "UPDATE tblFAQDocument " &_
				"SET	Title = " & steQForm("Title") & "," &_
				"		AuthorID = " & steNForm("AuthorID") & "," &_
				"		Synopsis = " & steQForm("Synopsis") & "," &_
				"		Introduction = " & steQForm("Introduction") & "," &_
				"		Epilogue = " & steQForm("Epilogue") & " " &_
				"WHERE	DocumentID = " & nDocumentID
		Call adoExecute(sStat)
	End If
End If

' retrieve the document to edit
sStat = "SELECT	* FROM tblFaqDocument WHERE DocumentID = " & nDocumentID
Set rsDoc = adoOpenRecordset(sStat)

' retrieve the list of authors to choose from
sStat = "SELECT	AuthorID, Title, FirstName, MiddleName, LastName " &_
		"FROM	tblFaqAuthor " &_
		"WHERE	Archive = 0 " &_
		"ORDER BY LastName, FirstName"
Set rsAuth = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Document" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit FAQ Document" %></H3>

<P>
<% steTxt "Please make your changes to the document using the form below." %>&nbsp;
<% steTxt "You may use the ""Questions"" tab above to add questions to this FAQ document." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="document_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="DocumentID" VALUE="<%= nDocumentID %>">
<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 WIDTH="100%">
<TR>
	<TD class="forml" nowrap><% steTxt "Title" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Title" VALUE="<%= steRecordEncValue(rsDoc, "Title") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Author Name" %></TD><TD></TD>
	<TD><select name="AuthorID" class="form">
		<option value="0"> -- <% steTxt "Choose" %> --
		<% With Response
		Do Until rsAuth.EOF
			.Write "<option value="""
			.Write rsAuth.Fields("AuthorID").Value
			.Write """"
			If steRecordEncValue(rsDoc, "AuthorID") = CStr(rsAuth.Fields("AuthorID").Value) Then .Write " SELECTED"
			.Write ">"
			.Write Server.HTMLEncode(Trim(rsAuth("Title") & " " & rsAuth("FirstName") & " " & rsAuth("MiddleName") & " " & rsAuth("LastName")))
			.Write vbCrLf
			rsAuth.MoveNext
		Loop
		End With
		rsAuth.Close
		Set rsAuth = Nothing %>
		</select>
	</TD>
</TR><TR>
	<TD class="forml" nowrap><% steTxt "Synopsis" %></TD><TD></TD>
	<TD><textarea NAME="Synopsis" cols="80" rows="8" class="form" style="width:100%"><%= steRecordEncValue(rsDoc, "Synopsis") %></textarea></TD>
</TR><TR>
	<TD class="forml" VALIGN="top" nowrap><% steTxt "Introduction" %></TD><TD></TD>
	<TD><textarea NAME="Introduction" cols="80" rows="8" class="form" style="width:100%"><%= steRecordEncValue(rsDoc, "Introduction") %></textarea></TD>
</TR><TR>
	<TD class="forml" VALIGN="top" nowrap><% steTxt "Epilogue" %></TD><TD></TD>
	<TD><textarea NAME="Epilogue" cols="80" rows="8" class="form" style="width:100%"><%= steRecordEncValue(rsDoc, "Epilogue") %></textarea></TD>
</TR><TR>
	<TD COLSPAN=3 align="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update FAQ Document" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "FAQ Document Updated" %></H3>

<P>
<% steTxt "The FAQ document was successfully updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
<% If nDocumentID > 0 Then %>
	<A HREF="question_list.asp?documentid=<%= nDocumentID %>" class="adminlink"><% steTxt "Questions" %></A> &nbsp;
<% End If %>
	<A HREF="document_list.asp" class="adminlink"><% steTxt "Document List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->
