<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/tree_lib.asp" -->
<%
'--------------------------------------------------------------------
' doc_browse.asp
'	Displays a list of folders and files for the document manager
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
Dim sAction
Dim sWhere		' where clause for search
Dim sCriteria	' criteria we are searching on
Dim rsCount		' count of search results
Dim rsList
Dim rsFolder
sAction = LCase(steForm("action"))

If sAction = "search" Then
	If Trim(steForm("FirstName")) <> "" Then
		sWhere = sWhere & " AND a.FirstName LIKE '%" & Replace(steForm("FirstName"), "'", "''") & "%'"
		sCriteria = sCriteria & steGetText("<I>FirstName</I> like") & " <B>" & steForm("FirstName") & "</B><BR>"
	End If
	If Trim(steForm("LastName")) <> "" Then
		sWhere = sWhere & " AND a.LastName LIKE '%" & Replace(steForm("LastName"), "'", "''") & "%'"
		sCriteria = sCriteria & steGetText("<I>LastName</I> like") & " <B>" & steForm("LastName") & "</B><BR>"
	End If
	If Trim(steForm("Title")) <> "" Then
		sWhere = sWhere & " AND d.Title LIKE '%" & Replace(steForm("Title"), "'", "''") & "%'"
		sCriteria = sCriteria & steGetText("<I>Title</I> like") & " <B>" & steForm("Title") & "</B><BR>"
	End If
	If Trim(steNForm("FolderID")) <> 0 Then
		sWhere = sWhere & " AND d.FolderID = " & steForm("FolderID")
		' retrieve the name of this Folder
		sStat = "SELECT	FolderName " &_
				"FROM	tblFolder " &_
				"WHERE	FolderID = " & steForm("FolderID")
		Set rsFolder = adoOpenRecordset(sStat)
		If Not rsFolder.EOF Then sCriteria = sCriteria & steGetText("<I>Folder</I> is") & " <B>" & rsFolder.Fields("FolderName").Value & "</B><BR>"
	End If
End If

' retrieve the list of documents to display here (for search)
sStat = "SELECT	" & adoTop(30) & " d.DocID, d.Title, d.ShortDescription, a.FirstName, a.LastName, " &_
		"		d.Modified " &_
		"FROM	tblDoc d " &_
		"INNER JOIN	tblDocAuthor a on a.AuthorID = d.AuthorID " &_
		"WHERE	d.Archive = 0 " &_
		"AND	d.Active = 1 " & sWhere & adoTop2(30)
Set rsList = adoOpenRecordset(sStat)

' count the total number of records matched here
sStat = "SELECT COUNT(*) AS DocCount " &_
		"FROM	tblDoc d " &_
		"INNER JOIN	tblDocAuthor a on a.AuthorID = d.AuthorID " &_
		"WHERE	d.Archive = 0 " &_
		"AND	d.Active = 1 " & sWhere
Set rsCount = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Document" %>
<!-- #include file="pagetabs_inc.asp" -->

<!-- form for document search -->

<!-- display all of the documents here (max of 30) -->
<H3><% steTxt "Document List" %></H3>

<FORM METHOD="post" ACTION="doc_browse.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="search">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD class="forml"><% steTxt "Author First Name" %><BR>
	<INPUT TYPE="text" NAME="FirstName" VALUE="<%= steEncForm("FirstName") %>" SIZE=16 MAXLENGTH=32 class="form">
	</TD>
	<TD class="forml"><% steTxt "Author Last Name" %><BR>
	<INPUT TYPE="text" NAME="LastName" VALUE="<%= steEncForm("LastName") %>" SIZE=16 MAXLENGTH=32 class="form">
	</TD>
	<TD class="forml"><% steTxt "Title" %><BR>
	<INPUT TYPE="text" NAME="Title" VALUE="<%= steEncForm("Title") %>" SIZE=16 MAXLENGTH=32 class="form">
	</TD>
</TR><TR>
	<TD class="forml"><% steTxt "Folder" %><BR>
	<% treTreeSelect "FolderID", "tblDocFolder", "FolderID", "ParentFolderID", "FolderName", "", steNForm("FolderID") %>
	</TD>
	<TD>
	</TD><TD>
	<INPUT TYPE="submit" NAME="_submit" ACTION=" <% steTxt "Search" %> " class="form">
	</TD>
</TR>
</TABLE>
</FORM>

<BLOCKQUOTE>
<%= sCriteria %>
</BLOCKQUOTE>

<P ALIGN="center">
<I><% steTxt "Total of" %> <%= " " & rsCount.Fields("DocCount").Value & " " %> <% steTxt "Documents Found" %></I>
</P>
<% If Not rsList.EOF Then %>

<P>
<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 class="list">
<TR>
	<TD class="listhead">Title</TD>
	<TD class="listhead">Author Name</TD>
	<TD class="listhead" ALIGN="right">Modified</TD>
	<TD class="listhead" ALIGN="right">Action</TD>
</TR>
<% I = 0
	Do Until rsList.EOF %>
<TR class="list<%= I mod 2 %>">
	<TD><%= rsList.Fields("Title").Value %></TD>
	<TD><%= rsList.Fields("FirstName").Value & " " & rsList.Fields("LastName").Value %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsList.Fields("Modified").Value, vbShortDate) %></TD>
	<TD ALIGN="right"><A HREF="doc_edit.asp?docid=<%= rsList.Fields("DocID").Value %>" class="actionlink">edit</A> . <A HREF="doc_delete.asp?docid=<%= rsList.Fields("docid").Value %>" class="actionlink">delete</A></TD>
</TR>
<%	rsList.MoveNext
	I = I + 1
   Loop %>
</TABLE>
</P>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No documents found to display here" %></B></P>

<% End If %>

<P ALIGN="center">
	<A HREF="doc_add.asp" class="adminlink"><% steTxt "Add New Document" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->