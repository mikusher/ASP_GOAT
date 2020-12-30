<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' param_list.asp
'	Displays a list of the current param params defined in the database
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
Dim rsParam
Dim rsMod
Dim sAction
Dim sModuleName
Dim sBackURL

sAction = LCase(steForm("Action"))
sBackURL = steForm("BackURL")
sModuleName = "Module"


Select Case sAction
	Case "moveup"
		Dim rsPrev, sPrevOrder

		' retrieve the previous order no
		sStat = "SELECT	OrderNo " &_
				"FROM	tblModuleParam " &_
				"WHERE	OrderNo < " & steNForm("OrderNo") & " " &_
				"ORDER BY OrderNo DESC"
		Set rsPrev = adoOpenRecordset(sStat)
		If Not rsPrev.EOF Then
			sPrevOrder = rsPrev.Fields("OrderNo").Value
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblModuleParam " &_
					"SET	OrderNo = OrderNo + 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & sPrevOrder
			Call adoExecute(sStat)

			sStat = "UPDATE	tblModuleParam " &_
					"SET	OrderNo = " & sPrevOrder & ", Modified = " & adoGetDate & " " &_
					"WHERE	ParamID = " & steNForm("ParamID")
			Call adoExecute(sStat)
			modRefresh True
		End If
		rsPrev.Close
		Set rsPrev = Nothing
	Case "movedown"
		Dim rsNext, sNextOrder

		' retrieve the next order no
		sStat = "SELECT	OrderNo " &_
				"FROM	tblModuleParam " &_
				"WHERE	OrderNo > " & steNForm("OrderNo") & " " &_
				"ORDER BY OrderNo"
		Set rsNext = adoOpenRecordset(sStat)
		If Not rsNext.EOF Then
			sNextOrder = rsNext.Fields("OrderNo").Value
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblModuleParam " &_
					"SET	OrderNo = OrderNo - 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & sNextOrder
			Call adoExecute(sStat)

			sStat = "UPDATE	tblModuleParam " &_
					"SET	OrderNo = " & sNextOrder & ", Modified = " & adoGetDate & " " &_
					"WHERE	ParamID = " & steNForm("ParamID")
			Call adoExecute(sStat)
			modRefresh True
		End If
		rsNext.Close
		Set rsNext = Nothing
End Select

' retrieve the list of parameters for the module (if nec)
If steNForm("ModuleID") > 0 Then
	' get the module name here
	Set rsMod = adoOpenRecordset("SELECT Title FROM tblModule WHERE ModuleID = " & steNForm("ModuleID"))
	If Not rsMod.EOF Then sModuleName = rsMod.Fields("Title").Value

	sStat = "SELECT	mp.OrderNo, mp.ParamID, mp.ParamName, mpt.TypeName, mp.IsRequired, " &_
			"		mp.Archive, mp.Modified " &_
			"FROM	tblModuleParam mp " &_
			"INNER JOIN	tblModuleParamType mpt ON mpt.TypeID = mp.TypeID " &_
			"WHERE	mp.ModuleID = " & steNForm("ModuleID") & " " &_
			"ORDER BY mp.OrderNo"
	Set rsParam = adoOpenRecordset(sStat)
End If

' create the list of modules to choose from
sStat = "SELECT	ModuleID, Title " &_
		"FROM	tblModule " &_
		"ORDER BY Title"
Set rsMod = adoOpenRecordset(sStat)

If Not (InStr(1, Request.ServerVariables("HTTP_REFERER"), "/module/param/") > 0) And sBackURL = "" Then
	sBackURL = Request.ServerVariables("HTTP_REFERER")
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Param" %>
<!-- #include file="pagetabs_inc.asp" -->

<script language="javascript" type="text/javascript">
<!-- // hide
function pickModule(nModuleID) {
	if (nModuleID != '')
		location.href='param_list.asp?backurl=<%= Server.URLEncode(sBackURL) %>&moduleid=' + nModuleID;
}
// unhide -->
</script>

<H3><%= sModuleName %>&nbsp;<% steTxt "Param List" %></H3>

<p>
<form method="post" action="#">
<table border=0 cellpadding=4 cellspacing=0>
<tr>
	<td class="forml"><% steTxt "Module to Display" %></td><td>&nbsp;&nbsp;</td>
	<td>
		<select name="ModuleID" class="form" onChange="pickModule(this.options[this.selectedIndex].value)">
		<option value=""> -- Choose --
		<% Do Until rsMod.EOF %>
		<option value="<%= rsMod.Fields("ModuleID").Value %>"<% If steNForm("ModuleID") = rsMod.Fields("ModuleID").Value Then Response.Write " SELECTED" %>> <%= rsMod.Fields("Title").Value %>
		<%	rsMod.MoveNext
		   Loop
		rsMod.Close
		Set rsMod = Nothing %>
		</select>
	</td>
</tr>
</table>
</form>
</p>

<% If steNForm("ModuleID") > 0 Then %>

<P>
<% steTxt "The following module parameters types are defined in the database." %>&nbsp;
<% steTxt "Configuration parameters are defined by the module authors to allow you to control the behavior of a module." %>
</P>

<% If Not rsParam.EOF Then %>

<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 class="list">
<TR>
	<TD CLASS="listhead"><% steTxt "Order" %></TD>
	<TD CLASS="listhead"><% steTxt "Param Name" %></TD>
	<TD CLASS="listhead"><% steTxt "Type" %></TD>
	<TD CLASS="listhead"><% steTxt "Req" %></TD>
	<TD CLASS="listhead"><% steTxt "Active" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Modified" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<% I = 0
Do Until rsParam.EOF %>
<TR CLASS="list<%= I mod 2 %>">
	<TD><%= rsParam.Fields("OrderNo").Value %></TD>
	<TD><%= rsParam.Fields("ParamName").Value %></TD>
	<TD><%= rsParam.Fields("TypeName").Value %></TD>
	<TD><% If rsParam.Fields("IsRequired").Value Then Response.Write "Y" Else Response.Write "N" %></TD>
	<TD><% If rsParam.Fields("Archive").Value Then Response.Write "N" Else Response.Write "Y" %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsParam.Fields("Modified").Value, vbShortDate) %></TD>
	<TD>
		<A HREF="param_list.asp?backurl=<%= Server.URLEncode(sBackURL) %>&ModuleID=<%= steNForm("ModuleID") %>&paramid=<%= rsParam.Fields("ParamID").Value %>&orderno=<%= rsParam.Fields("OrderNo").Value %>&action=moveup" class="actionlink"><% steTxt "up" %></A> .
		<A HREF="param_list.asp?backurl=<%= Server.URLEncode(sBackURL) %>&ModuleID=<%= steNForm("ModuleID") %>&paramid=<%= rsParam.Fields("ParamID").Value %>&orderno=<%= rsParam.Fields("OrderNo").Value %>&action=movedown" class="actionlink"><% steTxt "down" %></A> .
		<A HREF="option_list.asp?ModuleID=<%= steNForm("ModuleID") %>&paramid=<%= rsParam.Fields("ParamID").Value %>" class="actionlink"><% steTxt "options" %></A> .
		<A HREF="param_edit.asp?backurl=<%= Server.URLEncode(sBackURL) %>&ModuleID=<%= steNForm("ModuleID") %>&paramid=<%= rsParam.Fields("ParamID").Value %>" class="actionlink"><% steTxt "edit" %></A> .
		<A HREF="param_delete.asp?backurl=<%= Server.URLEncode(sBackURL) %>&ModuleID=<%= steNForm("ModuleID") %>&paramid=<%= rsParam.Fields("ParamID").Value %>" class="actionlink"><% steTxt "delete" %></A>
	</TD>
</TR>
<%	rsParam.MoveNext
	I = I + 1
   Loop
	rsParam.Close
	Set rsParam = Nothing %>
</TABLE>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No params exist in the database" %></B></P>

<% End If %>

<P ALIGN="center">
	<% If sBackURL <> "" Then %>
	<A HREF="<%= sBackURL %>" class="adminlink">&lt;&lt; Back</A>
	<% End If %>
	<A HREF="param_add.asp?backurl=<%= Server.URLEncode(sBackURL) %>&moduleid=<%= steNForm("ModuleID") %>" class="adminlink"><% steTxt "Add New Param" %></A>
</P>

<% End If %>


<!-- #include file="../../../../footer.asp" -->