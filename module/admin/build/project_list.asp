<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' project_list.asp
'	List all of the build projects in the database
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

sStat = "SELECT	bp.ProjectID, bp.ProjectName, u.FirstName, u.LastName, " &_
		"		bp.VersionNo, bp.BuildDate, bp.Modified " &_
		"FROM	tblBuildProject bp " &_
		"LEFT JOIN tblUser u on u.UserID = bp.UserID " &_
		"WHERE	bp.Active <> 0 " &_
		"AND	bp.Archive = 0 " &_
		"ORDER BY bp.ProjectName"
Set rsTab = adoOpenRecordset("show table status;")
If Not rsTab.EOF Then aTab = rsTab.GetRows
rsTab.Close
Set rsTab = Nothing
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<h4>Build Projects</h4>

<p>
Shown below are all of the build projects defined in the database for the
ASP Nuke project.
</p>

<% If IsArray(aTab) Then %>

<table border=0 cellpadding=2 cellspacing=0 class="list">
<tr>
	<td class="listhead">Project Name</td>
	<td class="listhead">User</td>
	<td class="listhead">Version</td>
	<td class="listhead">Built On</td>
	<td class="listhead">Modified</td>
	<td class="listhead">Action</td>
</tr>
<% For I = 0 To UBound(aTab, 2) %>
<tr class="list<%=I mod 2 %>">
	<td><a href="project_edit.asp?projectid=<%= aTab(0, I) %>"><%= Server.HTMLEncode(aTab(1, I)) %></a></td>
	<td><%= aTab(2, I) & " " & aTab(3, I) %></td>
	<td><%= aTab(4, I) %></td>
	<td><% If IsDate(aTab(5, I)) Then Response.Write FormatDateTime(aTab(5, I), vbGeneralDate) Else Resonse.Write "<i>n/a</i>" %></td>
	<td><% If IsDate(aTab(6, I)) Then Response.Write FormatDateTime(aTab(6, I), vbGeneralDate) Else Response.Write "<i>n/a</i>" %></td>
	<td><a href="project_delete.asp?projectid=<%= aTab(0, I) %>" class="actionlink"><% steText "delete" %></a> . <a href="project_edit.asp?projectid=<%= aTab(0, I) %>" class="actionlink"><% steText "edit" %></a></td>
</tr>
<%	nRows = nRows + CInt(aTab(3, I))
	nStorage = nStorage + CInt(aTab(5, I))
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
<b class="error">No build projects could be found in the local database</b>
</p>

<% End If %>

<p align="center">
	<input type="button" name="_create" value="Create Build Project" onclick="location.href='project_add.asp';" class="form">
</p>

<!-- #include file="../../../footer.asp" -->
