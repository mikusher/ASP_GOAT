<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/class/date.asp" -->
<!-- #include file="../../../../lib/class/form/listinput.asp" -->
<%
'--------------------------------------------------------------------
' book_edit.asp
'	Edit a document book for the site
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
Dim sCatList	' list of currently selected categories
Dim rsBook
Dim nBookID
Dim sErrorMsg	' error message to display to user

nBookID = steNForm("BookID")

If UCase(steForm("action")) = "EDIT" Then
	' check for required fields here
	If steNForm("authorid") = 0 Then
		sErrorMsg = steGetText("Please select the Author for this book")
	ElseIf steNForm("folderid") = 0 Then
		sErrorMsg = steGetText("Please select a Document Folder for this book")
	ElseIf Trim(steForm("Title")) = "" Then
		sErrorMsg = steGetText("Please enter the Title for this book")
	ElseIf steForm("Version") = "" Then
		sErrorMsg = steGetText("Plese enter the Version for this book")
	ElseIf steForm("ShowSectionNo") = "" Then
		sErrorMsg = steGetText("Please enter Show Section No for this book")
	Else
		' determine the new order no
		Set rsOrder= adoOpenRecordset("SELECT Coalesce(Max(OrderNo) + 1, 1) As OrderNo FROM tblDocBook")
		If Not rsOrder.EOF Then nOrderNo = rsOrder.Fields("OrderNo").Value Else nOrderNo = 1
		rsOrder.Close
		Set rsOrder = Nothing

		' insert the new book into the database
		sStat = "UPDATE tblDocBook SET " &_
				"	Title = " & steQForm("Title") & "," &_
				"	SubTitle = " & steQForm("SubTitle") & "," &_
				"	AuthorID = " & steNForm("AuthorID") & "," &_
				"	PublishDate = " & datQForm("PublishDate") & "," &_
				"	ShowSectionNo = " & steNForm("ShowSectionNo") & "," &_
				"	AuthorNotes = " & steQForm("AuthorNotes") & "," &_
				"	Version = " & steQForm("version") & " " &_
				"WHERE	BookID = " & nBookID
		' Response.Write sStat : Response.End
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

' build the list of categories to choose from
'sStat = "SELECT	CategoryID, CategoryName " &_
'		"FROM	tblDocCategory " &_
'		"WHERE	Active <> 0 " &_
'		"AND	Archive = 0 " &_
'		"ORDER BY CategoryName"
'Set rsCat = adoOpenRecordset(sStat)

' retrieve the book to edit
sStat = "SELECT * FROM tblDocBook WHERE BookID = " & nBookID
Set rsBook = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Book" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If UCase(steForm("action")) <> "EDIT" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Book" %></H3>

<P>
<% steTxt "Please enter the changes to the book properties in the form below." %>
</P>

<FORM METHOD="post" ACTION="book_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="BookID" VALUE="<%= nBookID %>">

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<TABLE BORDER=0 CELLPEDITING=2 CELLSPACING=0 WIDTH="100%">
<TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Document Folder" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><%
	Dim oList
	Set oList = New clsListInput
	oList.TreeListInput "FolderID", "tblDocFolder", "FolderID", "ParentFolderID", "", _
		"OrderNo", "FolderID", "FolderName", CInt(steRecordEncValue(rsBook, "FolderID")), "", False
	%>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Author" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><SELECT NAME="AuthorID" class="form">
	<OPTION VALUE="0"> -- Choose --
	<% Do Until rsAuth.EOF %>
	<OPTION VALUE="<%= rsAuth.Fields("AuthorID").Value %>"<% If CStr(steRecordEncValue(rsBook, "AuthorID")) = CStr(rsAuth.Fields("AuthorID").Value) Then Response.Write " SELECTED" %>> <%= rsAuth.Fields("Title").Value & " " & rsAuth.Fields("FirstName").Value & " " & rsAuth.Fields("MiddleName").Value & " " & rsAuth.Fields("LastName").Value & " " & rsAuth.Fields("Surname").Value %>
	<%	rsAuth.MoveNext
	   Loop %>
	</SELECT>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Title" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Title" VALUE="<%= steRecordEncValue(rsBook, "title") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Sub-Title" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="SubTitle" VALUE="<%= steRecordEncValue(rsBook, "SubTitle") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Version" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Version" VALUE="<%= steRecordEncValue(rsBook, "version") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Publish Date" %></TD><TD></TD>
	<TD><%
		Set oDat = New clsDate
		oDat.Selected = steRecordEncValue(rsBook, "publishdate")
		oDat.DateInput("publishdate") %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Show Section No's?" %></TD><TD></TD>
	<TD class="formd">
		<INPUT TYPE="radio" NAME="ShowSectionNo" VALUE="1"<% If steRecordBoolValue(rsBook, "ShowSectionNo") Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="ShowSectionNo" VALUE="0"<% If Not steRecordBoolValue(rsBook, "ShowSectionNo") Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Author Notes" %></TD><TD></TD>
	<TD WIDTH="100%"><TEXTAREA NAME="AuthorNotes" COLS="42" ROWS="10" WRAP="Virtual" class="form" style="width:100%"><%= steRecordEncValue(rsBook, "AuthorNotes") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update Book" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Book Updated" %></H3>

<P>
<% steTxt "The new book was modified successfully in the database." %>&nbsp;
<% steTxt "You may use the admin links above to continue administering the site." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="book_list.asp" class="adminlink"><% steTxt "Book List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->