<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' author_list.asp
'	Displays a list of the current admins for the site
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
Dim rsAuthor
Dim I

sStat = "SELECT	AuthorID, FirstName, MiddleName, LastName, Modified " &_
		"FROM	tblArticleAuthor " &_
		"ORDER BY Lastname, FirstName"
Set rsAuthor = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Author" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "Article Author List" %></H3>

<P>
<% steTxt "Shown below are all of the article authors defined in the database." %>
</P>

<% If Not rsAuthor.EOF Then %>

<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 class="list">
<TR>
	<TD CLASS="listhead"><% steTxt "Full Name" %></TD>
	<TD CLASS="listhead" align="right"><% steTxt "Modified" %></TD>
	<TD CLASS="listhead" align="right"><% steTxt "Action" %></TD>
</TR>
<% I = 0
Do Until rsAuthor.EOF %>
<TR CLASS="list<%= I mod 2 %>">
	<TD><%= rsAuthor.Fields("FirstName").Value & " " & rsAuthor.Fields("MiddleName").Value & " " & rsAuthor.Fields("LastName").Value %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsAuthor.Fields("Modified").Value, vbShortDate) %></TD>
	<TD>
		<A HREF="author_edit.asp?authorid=<%= rsAuthor.Fields("AuthorID").Value %>" class="actionlink"><% steTxt "edit" %></A> . 
		<A HREF="author_delete.asp?authorid=<%= rsAuthor.Fields("AuthorID").Value %>" class="actionlink"><% steTxt "delete" %></A></TD>
</TR>
<%	rsAuthor.MoveNext
	I = I + 1
   Loop %>
</TABLE>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No authors exist in the database" %></B></P>

<% End If %>

<P ALIGN="center">
	<a href="article_list.asp" class="adminlink"><% steTxt "Article List" %></a> &nbsp;
	<A HREF="author_add.asp" class="adminlink"><% steTxt "Add New Article Author" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->