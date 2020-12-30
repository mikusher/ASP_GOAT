<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' table_list.asp
'	List all of the database tables (in the MySQL database)
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

Dim sAction
Dim sStat
Dim rsTab
Dim aTab
Dim nStorage
Dim nRows

nStorage = 0
nRows = 0

Set rsTab = adoOpenRecordset("show table status;")
If Not rsTab.EOF Then aTab = rsTab.GetRows
rsTab.Close
Set rsTab = Nothing
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Tables" %>
<!-- #include file="pagetabs_inc.asp" -->

<h4>MySQL Tables</h4>

<p>
Shown below are all of the tables defined in the MySQL database for the
ASP Nuke project.
</p>

<% If IsArray(aTab) Then %>

<table border=0 cellpadding=2 cellspacing=0 class="list">
<tr>
	<td class="listhead">Table Name</td>
	<td class="listhead">Type</td>
	<td class="listhead">Rows</td>
	<td class="listhead">Storage</td>
	<td class="listhead">Modified</td>
	<td class="listhead">Action</td>
</tr>
<% For I = 0 To UBound(aTab, 2) %>
<tr class="list<%=I mod 2 %>">
	<td><a href="table_browse.asp?table=<%= Server.URLEncode(aTab(0, I)) %>"><%= Server.HTMLEncode(aTab(0, I)) %></a></td>
	<td><%= aTab(1, I) %></td>
	<td><%= aTab(3, I) %></td>
	<td><%= aTab(5, I) %></td>
	<td><% If IsDate(aTab(11, I)) Then Response.Write FormatDateTime(aTab(11, I), vbGeneralDate) Else Response.Write "<i>n/a</i>" %></td>
	<td><a href="table_drop.asp?table=<%= Server.URLEncode(aTab(0, I)) %>" class="actionlink">drop</a> . <a href="table_alter.asp?table=<%= Server.URLEncode(aTab(0, I)) %>" class="actionlink">alter</a></td>
</tr>
<%	nRows = nRows + CInt(aTab(3, I))
	nStorage = nStorage + CLng(aTab(5, I))
   Next %>
<tr class="list<%= I mod 2 %>">
	<td class="listhead">TOTALS</td>
	<td>&nbsp;</td>
	<td><b><%= nRows %></b></td>
	<td><b><%= nStorage %></b></td>
	<td colspan=2>&nbsp;</td>
</tr>
</table>

<% Else %>

<p>
<b class="error">No tables could be found in the local database</b>
</p>

<% End If %>

<p align="center">
	<input type="button" name="_create" value="Create Table" onclick="location.href='table_add.asp';" class="form">
</p>

<!-- #include file="../../../footer.asp" -->
