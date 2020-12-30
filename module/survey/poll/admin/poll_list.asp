<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' poll_list.asp
'	Displays a list of the poll questions for the site
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
Dim rsPoll
Dim I

sStat = "SELECT	PollID, Question, Created, Modified " &_
		"FROM	tblPoll " &_
		"ORDER BY Created DESC"
Set rsPoll = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Poll" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "Poll List" %></H3>

<P>
<% steTxt "Shown below are all of the current polls defined in the database." %>
</P>

<% If Not rsPoll.EOF Then %>

<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 class="list">
<TR BGCOLOR="#E0C0A0">
	<TD class="listhead"><% steTxt "Poll Question" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Created" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Modified" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<% I = 0
Do Until rsPoll.EOF %>
<TR class="list<%= I mod 2 %>">
	<TD><%= rsPoll.Fields("Question").Value %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsPoll.Fields("Created").Value, vbShortDate) %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsPoll.Fields("Modified").Value, vbShortDate) %></TD>
	<TD>
		<A HREF="poll_edit.asp?pollid=<%= rsPoll.Fields("PollID").Value %>" class="actionlink"><% steTxt "edit" %></A> .
		<A HREF="poll_delete.asp?pollid=<%= rsPoll.Fields("PollID").Value %>" class="actionlink"><% steTxt "delete" %></A> .
		<A HREF="answer_list.asp?pollid=<%= rsPoll.Fields("PollID").Value %>" class="actionlink"><% steTxt "answers" %></A>
	</TD>
</TR>
<%	rsPoll.MoveNext
	I = I + 1
   Loop %>
</TABLE>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No poll questions exist in the database" %></B></P>

<% End If %>

<P ALIGN="center">
<A HREF="poll_add.asp" class="adminlink"><% steTxt "Add New Poll Question" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->