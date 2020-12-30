<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' state_list.asp
'	Displays a list of the states for the member registration
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
Dim rsList
Dim I

sStat = "SELECT	StateCode, StateName, Modified " &_
		"FROM	tblState " &_
		"WHERE	Active <> 0 " &_
		"AND	Archive = 0 " &_
		"ORDER BY StateName"
Set rsList = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "State" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "States" %></H3>

<% If Not rsList.EOF Then %>

<P>
<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 CLASS="list">
<TR>
	<TD class="listhead"><% steTxt "Code" %></TD>
	<TD class="listhead"><% steTxt "State Name" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Modified" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<% I = 0
Do Until rsList.EOF %>
<TR class="list<%= I mod 2 %>">
	<TD><%= rsList.Fields("StateCode").Value %></TD>
	<TD><%= rsList.Fields("StateName").Value %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsList.Fields("Modified").Value, vbShortDate) %></TD>
	<TD ALIGN="right">
		<A HREF="state_edit.asp?statecode=<%= rsList.Fields("StateCode").Value %>" class="actionlink"><% steTxt "edit" %></A> . 
		<A HREF="state_delete.asp?statecode=<%= rsList.Fields("StateCode").Value %>" class="actionlink"><% steTxt "delete" %></A>
	</TD>
</TR>
<%	rsList.MoveNext
	I = I + 1
   Loop %>
</TABLE>
</P>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No states found to display here" %></B></P>

<% End If %>

<P ALIGN="Center">
	<A HREF="state_add.asp" class="adminlink"><% steTxt "Add New State" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
