<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' document_list.asp
'	Displays a list of the faq documents for the site
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
Dim rsDoc
Dim I

sStat = "SELECT	fd.DocumentID, fd.Title, " &_
		"		fa.Title AS AuthTitle, fa.FirstName, fa.MiddleName, fa.LastName, " &_
		"		fd.Created, fd.Modified " &_
		"FROM	tblFaqDocument fd " &_
		"INNER JOIN	tblFaqAuthor fa ON fd.AuthorID = fa.AuthorID " &_
		"ORDER BY fd.Created DESC"
Set rsDoc = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Document" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "FAQ Document List" %></H3>

<P>
<% steTxt "Shown below are all of the current FAQ documents defined in the database." %>
</P>

<% If Not rsDoc.EOF Then %>

<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 CLASS="list">
<TR>
	<TD CLASS="listhead" ALIGN="left"><% steTxt "Title" %></TD>
	<TD CLASS="listhead" ALIGN="left"><% steTxt "Author Name" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Created" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Modified" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<% I = 0
	Do Until rsDoc.EOF %>
<TR CLASS="list<%= I mod 2 %>">
	<TD><%= Server.HTMLEncode(rsDoc.Fields("Title").Value) %></TD>
	<TD><%= Server.HTMLEncode(Trim(rsDoc.Fields("AuthTitle").Value & " " & rsDoc.Fields("FirstName").Value) & " " & Trim(rsDoc.Fields("MiddleName").Value & " " & rsDoc.Fields("LastName").Value)) %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsDoc.Fields("Created").Value, vbShortDate) %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsDoc.Fields("Modified").Value, vbShortDate) %></TD>
	<TD>
		<A HREF="document_edit.asp?DocumentID=<%= rsDoc.Fields("DocumentID").Value %>" CLASS="actionlink"><% steTxt "edit" %></A> .
		<A HREF="document_delete.asp?DocumentID=<%= rsDoc.Fields("DocumentID").Value %>" CLASS="actionlink"><% steTxt "delete" %></A>
	</TD>
</TR>
<%	rsDoc.MoveNext
	I = I + 1
   Loop %>
</TABLE>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No FAQ documents exist in the database" %></B></P>

<% End If %>

<P ALIGN="center">
	<A HREF="document_add.asp" class="adminlink"><% steTxt "Add New Document" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->