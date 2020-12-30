<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' document_delete.asp
'	Delete an existing FAQ document from the database
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
Dim sAuthName
Dim rsDoc
Dim nDocumentID

nDocumentID = steNForm("DocumentID")

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If steNForm("confirm") <> 1 Then
		sErrorMsg = steGetText("Please confirm the deletion of this FAQ Document")
	Else
		' delete the faq document from the database
		sStat = "DELETE FROM tblFaqDocument " &_
				"WHERE	DocumentID = " & nDocumentID
		Call adoExecute(sStat)
	End If
End If

sAuthName = "<i>n/a</i>"
If nDocumentID > 0 Then
	sStat = "SELECT * FROM tblFaqDocument " &_
			"WHERE DocumentID = " & nDocumentID
	Set rsDoc = adoOpenRecordset(sStat)
	If Not rsDoc.EOF Then
		Dim rsAuth
		sStat = "SELECT	Title, FirstName, MiddleName, LastName FROM tblFaqAuthor WHERE AuthorID = " & rsDoc.Fields("AuthorID").Value & " AND Archive = 0"
		sET rsAuth = adoOpenRecordset(sStat)
		If Not rsAuth.EOF Then
			sAuthName  = Trim(rsAuth("Title") & " " & rsAuth("FirstName") & " " & rsAuth("MiddleName") & " " & rsAuth("LastName"))
		End If
		rsAuth.Close
		Set rsAuth = Nothing
	End If
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Document" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete FAQ Document" %></H3>

<P>
<% steTxt "Please confirm the deletion of the FAQ document by clicking the yes button next to <I>Confirm</I>." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="document_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="DocumentID" VALUE="<%= nDocumentID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD class="forml" nowrap><% steTxt "Title" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsDoc, "Title") %></TD>
</TR><TR>
	<TD class="forml" nowrap><% steTxt "Author Name" %></TD><TD></TD>
	<TD class="formd"><%= Server.HTMLEncode(sAuthName) %></TD>
</TR><TR>
	<TD class="forml" valign=top nowrap><% steTxt "Synopsis" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsDoc, "Synopsis") %></TD>
</TR><TR>
	<TD class="forml" VALIGN="top" nowrap><% steTxt "Introduction" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsDoc, "Introduction") %><%= steRecordEncValue(rsDoc, "Introduction") %></textarea></TD>
</TR><TR>
	<TD class="forml" VALIGN="top" nowrap><% steTxt "Epilogue" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsDoc, "Epilogue") %></TD>
</TR><TR>
	<TD class="forml" nowrap><% steTxt "Confirm Delete" %></TD><TD></TD>
	<TD class="formd">
			<INPUT TYPE="radio" NAME="confirm" VALUE="1" class="formradio" > <% steTxt "Yes" %>
			<INPUT TYPE="radio" NAME="confirm" VALUE="0" checked class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 align="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete FAQ Document" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "FAQ Document Deleted" %></H3>

<P>
<% steTxt "The FAQ document was successfully deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align=center>
	<a href="document_list.asp" class="adminlink"><% steTxt "FAQ Document List" %></a>
</p>
<!-- #include file="../../../../footer.asp" -->
