<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' book_list.asp
'	Displays a list of the document books for the site
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
Dim rsBook

sStat = "SELECT	tblDocBook.BookID, tblDocBook.Title, tblDocBook.SubTitle, tblDocBook.Version, " &_
		"		tblDocAuthor.FirstName, tblDocAuthor.MiddleName, tblDocAuthor.LastName, " &_
		"		tblDocBook.Created, tblDocBook.PublishDate, tblDocBook.OrderNo " &_
		"FROM	tblDocBook " &_
		"INNER JOIN	tblDocAuthor ON tblDocBook.AuthorID = tblDocAuthor.AuthorID " &_
		"ORDER BY tblDocBook.OrderNo"
Set rsBook = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->

<H3><% steTxt "Book Listing" %></H3>

<P>
<% steTxt "The following books are available for you to browse." %>&nbsp;
<% steTxt "These books have been created and are maintained by the ASP Nuke book authoring feature which is part of the ""documentation"" module." %>&nbsp;
<% steTxt "Click on a book title to view its table-of-contents." %>
</P>

<% If Not rsBook.EOF Then %>

<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 class="list">
<TR>
	<TD class="listhead"><% steTxt "Title" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Author" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Version" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Published" %></TD>
</TR>
<%	I = 0
	Do Until rsBook.EOF %>
<TR class="list<%= I mod 2 %>">
	<TD nowrap><a href="book_toc.asp?bookid=<%= rsBook.Fields("BookID").Value %>"><%= rsBook.Fields("Title").Value %></a><% If rsBook.Fields("SubTitle").Value & "" <> "" Then %><br><font class="tinytext"><%= rsBook.Fields("SubTitle").Value %></font><% End If %></TD>
	<TD nowrap><%= rsBook.Fields("FirstName").Value & " " & Trim(rsBook.Fields("MiddleName").Value & " " & rsBook.Fields("LastName").Value) %></TD>
	<TD nowrap align="right"><%= rsBook.Fields("Version").Value %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsBook.Fields("PublishDate").Value, vbShortDate) %></TD>
</TR>
<%	rsBook.MoveNext
   Loop %>
</TABLE>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No books exist in the database" %></B></P>

<% End If %>

<!-- #include file="../../../footer.asp" -->