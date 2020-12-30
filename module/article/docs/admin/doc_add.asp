<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' doc_add.asp
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
Dim rsDoc		' parent doc
Dim sSectionNo	' section no (eg: "1.3.13")
Dim sCatList	' list of currently selected categories
Dim sErrorMsg	' error message to display to user

nBookID = steNForm("BookID")

If UCase(steForm("action")) = "ADD" Then
	' check for required fields here
	If steNForm("authorid") = 0 Then
		sErrorMsg = steGetText("Please select the Author for this document")
	ElseIf Trim(steForm("Title")) = "" Then
		sErrorMsg = steGetText("Please enter the Title for this document")
	Else
		' determine the new order no
		Set rsOrder= adoOpenRecordset("SELECT Coalesce(Max(OrderNo) + 1, 1) As OrderNo FROM tblDoc")
		If Not rsOrder.EOF Then nOrderNo = rsOrder.Fields("OrderNo").Value Else nOrderNo = 1
		rsOrder.Close
		Set rsOrder = Nothing

		' build the section no from the parent (if nec)
		If steNForm("ParentDocID") <> 0 Then
			Set rsDoc = adoOpenRecordset("SELECT SectionNo FROM tblDoc WHERE DocID = " & steNForm("ParentDocID"))
			If Not rsDoc.EOF Then sSectionNo = rsDoc.Fields("SectionNo").Value & "." & nOrderNo Else sSectionNo = nOrderNo
		Else
			sSectionNo = nOrderNo
		End If

		' insert the new book into the database
		sStat = "INSERT INTO tblDoc (" &_
				"	BookID, ParentDocID, Title, Body, AuthorID, AuthorNotes, " &_
				"	SectionName, SectionNo, IsInlineContent, OrderNo, Created" &_
				") VALUES (" &_
				steNForm("BookID") & "," & steNForm("ParentDocID") & "," &_
				steQForm("Title") & "," & steQForm("Body") & "," &_
				steNForm("AuthorID") & "," & steQForm("AuthorNotes") & "," &_
				steQForm("SectionName") & ", '" & sSectionNo & "', " &_
				steNForm("IsInlineContent") & "," & nOrderNo & "," & adoGetDate &_
				")"
		Call adoExecute(sStat)
	End If
End If

' build the list of authors to choose from
sStat = "SELECT	AuthorID, Title, FirstName, MiddleName, LastName, Surname " &_
		"FROM	tblDocAuthor " &_
		"WHERE	Active <> 0 " &_
		"AND	Archive = 0 " &_
		"ORDER BY LastName, FirstName, MiddleName"
Set rsAuth = adoOpenRecordset(sStat)

' build the list of books to choose from
sStat = "SELECT	BookID, Title " &_
		"FROM	tblDocBook " &_
		"WHERE	Archive = 0 " &_
		"ORDER BY Title"
Set rsBook = adoOpenRecordset(sStat)
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

<% If UCase(steForm("action")) <> "ADD" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New Document" %></H3>

<P>
<% steTxt "Please enter the information for the new document in the form below." %>
</P>

<FORM NAME="form1" METHOD="post" ACTION="doc_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 WIDTH="100%">
<TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Book" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><SELECT NAME="BookID" class="form" onChange="pickBook(this.options[this.selectedIndex].value);">
	<OPTION VALUE="0"> -- Choose --
	<% Do Until rsBook.EOF %>
	<OPTION VALUE="<%= rsBook.Fields("BookID").Value %>"<% If CStr(steForm("BookID")) = CStr(rsBook.Fields("BookID").Value) Then Response.Write " SELECTED" %>> <%= rsBook.Fields("Title").Value %>
	<%	rsBook.MoveNext
	   Loop
	rsBook.Close
	Set rsBook = Nothing %>
	</SELECT>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Parent Section" %></TD><TD></TD>
	<TD><SELECT NAME="ParentDocID" class="form">
	<OPTION VALUE="0"> -- Choose --
	<% locParentDoc nBookID, steNForm("ParentDocID") %>
	</SELECT>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Author" %></TD><TD></TD>
	<TD><SELECT NAME="AuthorID" class="form">
	<OPTION VALUE="0"> -- Choose --
	<% Do Until rsAuth.EOF %>
	<OPTION VALUE="<%= rsAuth.Fields("AuthorID").Value %>"<% If CStr(steForm("AuthorID")) = CStr(rsAuth.Fields("AuthorID").Value) Then Response.Write " SELECTED" %>> <%= rsAuth.Fields("Title").Value & " " & rsAuth.Fields("FirstName").Value & " " & rsAuth.Fields("MiddleName").Value & " " & rsAuth.Fields("LastName").Value & " " & rsAuth.Fields("Surname").Value %>
	<%	rsAuth.MoveNext
	   Loop
	rsAuth.Close
	Set rsAuth = Nothing %>
	</SELECT>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Title" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Title" VALUE="<%= steEncForm("title") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Short Description" %></TD><TD></TD>
	<TD><TEXTAREA NAME="ShortDescription" COLS="48" ROWS="5" class="form"><%= steEncForm("ShortDescription") %></TEXTAREA></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Body" %></TD><TD></TD>
	<TD><TEXTAREA NAME="Body" COLS="48" ROWS="16" class="form"><%= steEncForm("body") %></TEXTAREA></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap>Section Name</TD><TD></TD>
	<TD><INPUT TYPE="TEXT" NAME="SectionName" VALUE="<%= steEncForm("SectionName") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Is Inline Content?" %></TD><TD></TD>
	<TD class="formd">
		<INPUT TYPE="radio" NAME="IsInlineContent" VALUE="1"<% If steForm("IsInlineContent") = "1" Then Response.Write " SELECTED" %> class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="IsInlineContent" VALUE="0"<% If steForm("IsInlineContent") = "0" Then Response.Write " SELECTED" %> class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Author Notes" %></TD><TD></TD>
	<TD WIDTH="100%"><TEXTAREA NAME="AuthorNotes" COLS="42" ROWS="10" WRAP="Virtual" class="form" style="width:100%"><%= steEncForm("AuthorNotes") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Document" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Document Added" %></H3>

<P>
<% steTxt "The new document was successfully added." %>&nbsp;
<% steTxt "You may use the admin links above to continue administering the site." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="doc_list.asp" class="adminlink"><% steTxt "Document List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->
<%
'--------------------------------------------------------------------
' display the hierarchical drop-list options

Sub locOptions(oDoc, nParentID, nLevel)
	Dim sIndent, aDoc, I

	If oDoc.Item("Chd" & nParentID) <> "" Then
		' build the ident code for this level
		For I = 1 To nLevel
			sIndent = sIndent & " &nbsp; &nbsp;"
		Next

		aDoc = Split(Mid(oDoc.Item("Chd" & nParentID), 2), ",")
		For I = 0 To UBound(aDoc)
			Response.Write sIndent & oDoc.Item("Doc" & aDoc(I)) & vbCrLf
			Call locOptions(oDoc, aDoc(I), nLevel + 1)
		Next
	End If
End Sub

'--------------------------------------------------------------------
' display the hierarchical drop-list options

Sub locParentDoc(nBookID, nSelectedID)
	Dim sStat, rsDoc, oDoc

	If nBookID = 0 Then Exit Sub
	Set oDoc = Server.CreateObject("Scripting.Dictionary")
	sStat = "SELECT	DocID, ParentDocID, SectionName, Title " &_
			"FROM	tblDoc " &_
			"WHERE	BookID = " & nBookID & " " &_
			"AND	Active <> 0 " &_
			"AND	Archive = 0"
	Set rsDoc = adoOpenRecordset(sStat)
	Do Until rsDoc.EOF
		' build the option
		sOption = "<option value=""" & rsDoc.Fields("DocID").Value & """"
		If nSelectedID = rsDoc.Fields("DocID").Value Then sOption = sOption & " SELECTED"
		sOption = sOption & ">" & rsDoc.Fields("Title").Value & "</option>"

		' add the option to the dict
		oDoc.Item("Doc" & rsDoc.Fields("DocID").Value) = sOption
		oDoc.Item("Chd" & rsDoc.Fields("ParentDocID").Value) = oDoc.Item("Chd" & rsDoc.Fields("ParentDocID").Value) &_
			"," & rsDoc.Fields("DocID").Value
		rsDoc.MoveNext
	Loop
	rsDoc.Close
	Set rsDoc = Nothing

	Call locOptions(oDoc, 0, 0)
End Sub
%>