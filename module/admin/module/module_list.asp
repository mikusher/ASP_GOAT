<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' module_list.asp
'	Displays a list of the current modules defined in the database
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
Dim rsModule

' build the param counts
Set oCount = Server.CreateObject("Scripting.Dictionary")
sStat = "SELECT ModuleID, Count(*) AS ParamCount " &_
		"FROM	tblModuleParam " &_
		"WHERE	Archive = 0 " &_
		"GROUP BY ModuleID"
Set rsCount = adoOpenRecordset(sStat)
Do Until rsCount.EOF
	oCount.Item(CStr(rsCount.Fields("ModuleID").Value)) = rsCount.Fields("ParamCount").Value
	rsCount.MoveNext
Loop
rsCount.Close : Set rsCount = Nothing

' retrieve the module list
sStat = "SELECT	ModuleID, Title, VersionNo, Modified " &_
		"FROM	tblModule " &_
		"ORDER BY Title, VersionNo"
Set rsModule = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Module" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "Module List" %></H3>

<p>
<% steTxt "Shown below are all of the nuke modules defined in the database." %>
</P>

<% If Not rsModule.EOF Then %>

<table border="0" cellpadding="4" cellspacing="0" class="list">
<tr>
	<td class="listhead"><% steTxt "Title" %></td>
	<td class="listhead"><% steTxt "Version" %></td>
	<td class="listhead" align="center"><% steTxt "Params" %></td>
	<td class="listhead" align="right"><% steTxt "Modified" %></td>
	<td class="listhead" align="right"><% steTxt "Action" %></td>
</tr>
<% I = 0
Do Until rsModule.EOF %>
<tr class="list<%= I mod 2 %>">
	<td><%= rsModule.Fields("Title").Value %></td>
	<td><%= rsModule.Fields("VersionNo").Value %></td>
	<td align="center"><% If oCount.Exists(CStr(rsModule.Fields("ModuleID").Value)) Then Response.Write oCount.Item(CStr(rsModule.Fields("ModuleID").Value)) Else Response.Write "0" %></td>
	<td align="right"><%= adoFormatDateTime(rsModule.Fields("Modified").Value, vbShortDate) %></td>
	<td>
		<a href="param/param_list.asp?moduleid=<%= rsModule.Fields("ModuleID").Value %>" class="actionlink"><% steTxt "params" %></A> .
		<a href="module_edit.asp?moduleid=<%= rsModule.Fields("ModuleID").Value %>" class="actionlink"><% steTxt "edit" %></A> .
		<a href="module_delete.asp?moduleid=<%= rsModule.Fields("ModuleID").Value %>" class="actionlink"><% steTxt "delete" %></A>
	</td>
</tr>
<%	rsModule.MoveNext
	I = I + 1
   Loop %>
</table>

<% Else %>

<p><b class="error"><% steTxt "Sorry, No modules exist in the database" %></b></P>

<% End If %>

<p align="center">
	<a href="module_add.asp" class="adminlink"><% steTxt "Add New Module" %></A>
</P>

<!-- #include file="../../../footer.asp" -->