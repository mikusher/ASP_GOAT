<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' doc_delete.asp
'	Displays a list of the current documents for the site
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
Dim rsArt
Dim rsAuth		' list of authors to choose from
Dim rsCat		' list of categories to choose from
Dim rsBook		' list of books to choose from
Dim nBookID		' book to work with
Dim nDocID		' document to edit
Dim rsDoc		' parent doc
Dim nOrderNo	' current order no
Dim sSectionNo	' section no (eg: "1.3.13")
Dim rsDoc2		' parent document
Dim sParentDoc	' parent document name
Dim sCatList	' list of currently selected categories
Dim sBookTitle	' title for the book (if any)
Dim sErrorMsg	' error message to display to user

nDocID = steNForm("DocID")
nBookID = steNForm("BookID")

If steForm("Action") = "delete" Then
	' check for required fields here
	If steNForm("confirm") = 0 Then
		sErrorMsg = steGetText("Please confirm the deletion of this document")
	Else
		' delete the doc book from the database
		sStat = "DELETE FROM tblDoc WHERE DocID = " & steNForm("DocID")
		Call adoExecute(sStat)
	End If
End If

' retrieve the document to edit
Set rsDoc = adoOpenRecordset("SELECT * FROM tblDoc WHERE DocID = " & nDocID)
If Not rsDoc.EOF Then
	' build the author name
	sStat = "SELECT	AuthorID, Title, FirstName, MiddleName, LastName, Surname " &_
			"FROM	tblDocAuthor " &_
			"WHERE	Active <> 0 " &_
			"AND	Archive = 0 " &_
			"AND	AuthorID = " & rsDoc.Fields("AuthorID").Value
	Set rsAuth = adoOpenRecordset(sStat)
	If Not rsAuth.EOF Then
		sAuthName = Trim((rsAuth.Fields("Title").Value & "") &_
			rsAuth.Fields("FirstName").Value & " " &_
			(rsAuth.Fields("MiddleName").Value & "") &_
			" " & rsAuth.Fields("LastName").Value &_
			" " & (rsAuth.Fields("Surname").Value & ""))
	End If
	rsAuth.Close
	Set rsAuth = Nothing

	' build the book name (if nec)
	If Not IsNull(rsDoc.Fields("BookID").Value) Then
		If rsDoc.Fields("BookID").Value > 0 Then
			sStat = "SELECT	BookID, Title " &_
					"FROM	tblDocBook " &_
					"WHERE	Archive = 0 " &_
					"AND	BookID = " & rsDoc.Fields("BookID").Value
			Set rsBook = adoOpenRecordset(sStat)
			If Not rsBook.EOF Then sBookTitle = rsBook.Fields("Title").Value
			rsBook.Close
			Set rsBook = Nothing
		Else
			sBookTitle = steGetText("n/a")
		End If
	Else
		sBookTitle = steGetText("n/a")
	End If

	' build the parent document name (if nec)
	If Not IsNull(rsDoc.Fields("ParentDocID").Value) Then
		If rsDoc.Fields("ParentDocID").Value > 0 Then
			sStat = "SELECT SectionNo, Title " &_
					"FROM	tblDoc " &_
					"WHERE  DocID = " & rsDoc.Fields("ParentDocID").Value
			Set rsDoc2 = adoOpenRecordset(sStat)
			If Not rsDoc2.EOF Then sParentDoc = Trim(rsDoc2.Fields("SectionNo").Value & " " & rsDoc2.Fields("Title").Value)
			rsDoc2.Close
			Set rsDoc2 = Nothing
		Else
			sBookTitle = steGetText("n/a")
		End If
	Else
		sBookTitle = steGetText("n/a")
	End If

End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->
<script language="javascript" type="text/javascript">
<!-- // hide
// referesh the parent doc list
function pickBook(nBookID) {
	if (nBookID != '' && nBookID != '0') {
		document.form1.action.value = '';
		document.form1.submit();
	}
}
// unhide -->
</script>

<% sCurrentTab = "Document" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If UCase(steForm("action")) <> "DELETE" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Document" %></H3>

<P>
<% steTxt "Please confirm that you would like to delete the document shown below." %>&nbsp;
<% steTxt "You must check <I>Yes</I> next to <B>Confirm Delete</B> in order to delete this record permanently." %>
</P>

<FORM NAME="form1" METHOD="post" ACTION="doc_delete.asp">
<input type="hidden" name="DocID" value="<%= nDocID %>">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="OrderNo" value="<%= steRecordEncValue(rsDoc, "OrderNo") %>">
<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Book" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD CLASS="formd"><%= Server.HTMLEncode(sBookTitle) %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Parent Section" %></TD><TD></TD>
	<TD CLASS="formd"><%= Server.HTMLEncode(sParentDoc) %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Author" %></TD><TD></TD>
	<TD CLASS="formd"><%= Server.HTMLEncode(sAuthName) %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Title" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsDoc, "title") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Short Description" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsDoc, "ShortDescription") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Body" %></TD><TD></TD>
	<TD><%= Replace(steRecordEncValue(rsDoc, "body"), vbCrLf, "<br>") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Section Name" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsDoc, "SectionName") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Is Inline Content?" %></TD><TD></TD>
	<TD class="formd">
		<% If steRecordBoolValue(rsDoc, "IsInlineContent") Then Response.Write steGetText("Yes") Else Response.Write steGetText("No") %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Author Notes" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsDoc, "AuthorNotes") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Confirm Delete" %></B></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="confirm" VALUE="1" CLASS="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="confirm" VALUE="0" CLASS="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Document" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Document Deleted" %></H3>

<P>
<% steTxt "The document was deleted successfully from the database." %>&nbsp;
<% steTxt "You may use the admin links above to continue administering the site." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="doc_list.asp" class="adminlink"><% steTxt "Document List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->