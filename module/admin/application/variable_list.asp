<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' variable_list.asp
'	Displays a list of the application variables for the site
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
Dim rsVar
Dim nTabID
Dim rsTab
Dim sWhere

' retrieve the list of application variables to display
If steNForm("TabID") > 0 Then sWhere = "WHERE av.TabID = " & steNForm("TabID") & " "
sStat = "SELECT	av.VarID, av.VarName, av.VarValue, avt.HasOptions, av.Created, av.Modified " &_
		"FROM	tblApplicationVar av " &_
		"INNER JOIN	tblApplicationVarType avt ON av.TypeID = avt.TypeID " & sWhere &_
		"ORDER BY av.VarName"
Set rsVar = adoOpenRecordset(sStat)

%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->
<script language="Javascript" type="text/javascript">
function pickTab(nTabID) {
	if (nTabID != '')
		location.href='variable_list.asp?tabid=' + nTabID;
}
</script>
<% sCurrentTab = "Configure" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "Application Variable List" %></H3>

<P>
<% steTxt "Shown below are all of the current application variables defined in the database." %>&nbsp;
<% steTxt "These are used for general configuration of the entire ASP Nuke application." %>
</P>

<P>
<form method="post" action="#">
<table border=0 cellpadding=2 cellspacing="0">
<tr>
	<td class="forml"><% steTxt "Configuration Tab" %></td><td>&nbsp;&nbsp;</td>
	<td class="formd">
	<select name="TabID" class="form" onchange="pickTab(this.options[this.selectedIndex].value)">
	<option value="">-- All Tabs Combined --
<%
' retrieve the tab options to choose from
sStat = "SELECT	TabID, TabName, Title " &_
		"FROM	tblApplicationVarTab " &_
		"WHERE	Archive = 0 " &_
		"ORDER BY OrderNo"
Set rsTab = adoOpenRecordset(sStat)
Do Until rsTab.EOF %>
<option value="<%= rsTab.Fields("TabID").Value %>"<% If rsTab.Fields("TabID").Value = steNForm("TabID") Then Response.Write " SELECTED" %>> <%= rsTab.Fields("TabName").Value %>
<%	rsTab.MoveNext
Loop
rsTab.Close
Set rsTab = Nothing %></td>
</tr>
</table>
</P>
<% If Not rsVar.EOF Then %>

<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 CLASS="list">
<TR>
	<TD CLASS="listhead"><% steTxt "Variable Name" %></TD>
	<TD CLASS="listhead"><% steTxt "Variable Value" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Modified" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<%
I = 0
Do Until rsVar.EOF %>
<TR CLASS="list<%= I mod 2 %>">
	<TD><%= rsVar.Fields("VarName").Value %></TD>
	<TD><% If Len(rsVar.Fields("VarValue").Value) > 30 Then
			Response.Write Server.HTMLEncode(Left(rsVar.Fields("VarValue").Value, 30)) & "..."
		Else
			Response.Write Server.HTMLEncode(rsVar.Fields("VarValue").Value)
		End If %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsVar.Fields("Modified").Value, vbShortDate) %></TD>
	<TD ALIGN="right" nowrap>
		<% If steRecordBoolValue(rsVar, "HasOptions") Then %>
		<A HREF="option_list.asp?varid=<%= rsVar.Fields("VarID").Value %>" CLASS="actionlink"><% steTxt "options" %></A> .
		<% End If %>
		<A HREF="variable_edit.asp?varid=<%= rsVar.Fields("VarID").Value %>" CLASS="actionlink"><% steTxt "edit" %></A> .
		<A HREF="variable_delete.asp?varid=<%= rsVar.Fields("VarID").Value %>" CLASS="actionlink"><% steTxt "delete" %></A>
	</TD>
</TR>
<%	rsVar.MoveNext
	I = I + 1
   Loop %>
</TABLE>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No application variables exist in the database" %></B></P>

<% End If %>

<P ALIGN="center">
	<A HREF="variable_add.asp" CLASS="adminlink"><% steTxt "Add Variable" %></A>
</P>

<!-- #include file="../../../footer.asp" -->