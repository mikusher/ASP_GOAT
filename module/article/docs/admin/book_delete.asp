<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' book_delete.asp
'	Delete a document book for the site
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
Dim sAuthName
Dim rsCat		' list of categories to choose from
Dim sCatList	' list of currently selected categories
Dim rsBook
Dim nBookID
Dim sErrorMsg	' error message to display to user

nBookID = steNForm("BookID")
sAction = steForm("action")

If sAction = "delete" Then
	' check for required fields here
	If steForm("confirm") <> "1" Then
		sErrorMsg = steGetText("Please confirm the deletion of this book")
	Else
		' delete the doc book from the database
		sStat = "DELETE FROM tblDocBook WHERE BookID = " & steNForm("BookID")
		Call adoExecute(sStat)
	End If
End If

' build the list of categories to choose from
'sStat = "SELECT	CategoryID, CategoryName " &_
'		"FROM	tblDocCategory " &_
'		"WHERE	Active <> 0 " &_
'		"AND	Archive = 0 " &_
'		"ORDER BY CategoryName"
'Set rsCat = adoOpenRecordset(sStat)

' retrieve the book to delete
sStat = "SELECT * FROM tblDocBook WHERE BookID = " & nBookID
Set rsBook = adoOpenRecordset(sStat)
If Not rsBook.EOF Then
	' retrieve the author name
	sStat = "SELECT	AuthorID, Title, FirstName, MiddleName, LastName, Surname " &_
			"FROM	tblDocAuthor " &_
			"WHERE	Active <> 0 " &_
			"AND	Archive = 0 " &_
			"AND	AuthorID = " & rsBook.Fields("AuthorID").Value
	Set rsAuth = adoOpenRecordset(sStat)
	If Not rsAuth.EOF Then
		sAuthName = rsAuth.Fields("FirstName").Value & " " & rsAuth.Fields("MiddleName").Value &_
			" " & rsAuth.Fields("FirstName").Value
	End If
	rsAuth.Close
	Set rsAuth = Nothing
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Book" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If UCase(steForm("action")) <> "DELETE" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Book" %></H3>

<P>
<% steTxt "Please confirm that you would like to delete the book shown below." %>&nbsp;
<% steTxt "You must check <I>Yes</I> next to <B>Confirm Delete</B> in order to delete this book permanently." %>
</P>

<FORM METHOD="post" ACTION="book_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="BookID" VALUE="<%= nBookID %>">

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<TABLE BORDER=0 CELLPEDITING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Author" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD class="formd"><%= Server.HTMLEncode(sAuthName) %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Title" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsBook, "title") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Sub-Title" %></TD><TD></TD>
	<TD><%= steRecordEncValue(rsBook, "SubTitle") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Version" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsBook, "version") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Publish Date" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsBook, "publishdate") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Show Section No's?" %></TD><TD></TD>
	<TD class="formd"><% If steRecordBoolValue(rsBook, "ShowSectionNo") Then Response.Write steGetText("Yes") Else Response.Write steGetText("No") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Author Notes" %></TD><TD></TD>
	<TD class="formd"><TEXTAREA NAME="LeadIn" COLS="42" ROWS="10" WRAP="Virtual" class="form"><%= steRecordEncValue(rsBook, "AuthorNotes") %></TEXTAREA></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Confirm Delete" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="confirm" VALUE="1" CLASS="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="confirm" VALUE="0" CLASS="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Book" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Book Deleted" %></H3>

<P>
<% steTxt "The book was deleted successfully from the database." %>&nbsp;
<% steTxt "You may use the admin links above to continue administering the site." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="book_list.asp" class="adminlink"><% steTxt "Book List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->