<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' document_add.asp
'	Adds a new FAQ document to the database.
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
Dim nDocumentID	' document ID newly added to database
Dim rsDoc		' faq document newly added to database
Dim rsCat		' list of categories to choose from
Dim sCatList	' list of currently selected categories
Dim rsAuth
Dim sErrorMsg	' error message to display to user

nDocumentID = 0

If UCase(steForm("action")) = "ADD" Then
	' check for required fields here
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
		' insert the new faq document into the database
		sStat = "INSERT INTO tblFaqDocument (" &_
				"	Title, AuthorID, Synopsis, Introduction, Epilogue, Created" &_
				") VALUES (" &_
				steQForm("Title") & "," & steNForm("AuthorID") & "," &_
				steQForm("Synopsis") & "," & steQForm("Introduction") & "," &_
				steQForm("Epilogue") & "," & adoGetDate &_
				")"
		Call adoExecute(sStat)

		' retrieve the new DocumentID
		sStat = "SELECT Max(DocumentID) As DocumentID FROM tblFaqDocument"
		Set rsDoc = adoOpenRecordset(sStat)
		If Not rsDoc.EOF Then nDocumentID = rsDoc.Fields("DocumentID").Value
	End If
End If

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

<% If UCase(steForm("action")) <> "ADD" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New FAQ Document" %></H3>

<P>
<% steTxt "Please enter the information for the new FAQ document in the form below." %>
</P>

<FORM METHOD="post" ACTION="document_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 width="100%">
<TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Title" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="Title" VALUE="<%= steEncForm("title") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Author Name" %></TD><TD></TD>
	<TD><select name="AuthorID" class="form">
		<option value="0"> -- Choose --
		<% With Response
		Do Until rsAuth.EOF
			.Write "<option value="""
			.Write rsAuth.Fields("AuthorID").Value
			.Write """"
			If steEncForm("AuthorID") = CStr(rsAuth.Fields("AuthorID").Value) Then .Write " SELECTED"
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
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Synopsis" %></TD><TD></TD>
	<TD><TEXTAREA NAME="Synopsis" COLS="42" ROWS="10" WRAP="Virtual" class="form" style="width:100%"><%= steEncForm("Synopsis") %></TEXTAREA></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Introduction" %></TD><TD></TD>
	<TD><TEXTAREA NAME="Introduction" COLS="42" ROWS="10" WRAP="Virtual" class="form" style="width:100%"><%= steEncForm("Introduction") %></TEXTAREA></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Epilogue" %></TD><TD></TD>
	<TD><TEXTAREA NAME="Epilogue" COLS="42" ROWS="10" WRAP="Virtual" class="form" style="width:100%"><%= steEncForm("Epilogue") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add FAQ Document" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "FAQ Document Added" %></H3>

<P>
<% steTxt "The new faq document was successfully added." %>&nbsp;
<% steTxt "You may use the admin links above to continue administering the site." %>
</P>

<% End If %>

<P ALIGN="center">
<% If nDocumentID > 0 Then %>
	<A HREF="question_list.asp?documentid=<%= nDocumentID %>" class="adminlink"><% steTxt "Questions" %></A> &nbsp;
<% End If %>
	<A HREF="document_list.asp" class="adminlink"><% steTxt "Document List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->