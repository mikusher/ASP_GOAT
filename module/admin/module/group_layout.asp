<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' group_layout.asp
'	Performs a layout (positioning) of modules within a module
'	group.
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
Dim sAction
Dim rsGroup
Dim rsPos
Dim nGroupID
Dim nModuleID
Dim nOrderNo
Dim sErrorMsg
Dim I

sAction = steForm("action")
nGroupID = steNForm("GroupID")
nModuleID = steNForm("ModuleID")


Select Case LCase(sAction)
	Case "add"
		If Trim(steForm("InsertBeforeID")) = "" Then
			sErrorMsg = steGetText("Please select where you would like to insert the module")
		Else
			' make sure module does not already exist
			sStat = "SELECT	ModuleID FROM tblModuleGroupPos WHERE GroupID = " & nGroupID & " AND ModuleID = " & nModuleID
			Set rsPos = adoOpenRecordset(sStat)
			If rsPos.EOF Then
				' add a new module to the layout group

				' set the order no and insert the new module position
				If steNForm("InsertBeforeID") = 0 Then
					' determine the new order no
					sStat = "SELECT Coalesce(Max(OrderNo) + 1, 1) As OrderNo FROM tblModuleGroupPos WHERE GroupID = " & nGroupID
					Set rsOrder = adoOpenRecordset(sStat)
					If Not rsOrder.EOF Then nOrderNo = rsOrder.Fields("OrderNo").Value Else nOrderNo = 1
					rsOrder.Close
					Set rsOrder = Nothing
				Else
					nOrderNo = steNForm("InsertBeforeID")
				End If

				' increment orders above the new order no (to make room)
				sStat = "UPDATE	tblModuleGroupPos " &_
						"SET	OrderNo = OrderNo + 1, Modified = " & adoGetDate & " " &_
						"WHERE	OrderNo >= " & steNForm("InsertBeforeID")
				Call adoExecute(sStat)

				' insert the new module here
				sStat = "INSERT INTO tblModuleGroupPos (" &_
						"	GroupID, ModuleID, OrderNo" &_
						") VALUES (" &_
						nGroupID & ", " & steNForm("ModuleID") & ", " & nOrderNo &_
						")"
				Call adoExecute(sStat)
			Else
				sErrorMsg = steGetText("The module has been added to the layout group")
			End If
		End If
	Case "delete"
		sStat = "SELECT	OrderNo FROM tblModuleGroupPos WHERE GroupID = " & nGroupID & " AND ModuleID = " & nModuleID
		Set rsPos = adoOpenRecordset(sStat)
		If Not rsPos.EOF Then
			' re-order any positions after the one we are deleting
			sStat = "UPDATE	tblModuleGroupPos " &_
					"SET	OrderNo = OrderNo - 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo > " & rsPos.Fields("OrderNo").Value
			Call adoExecute(sStat)

			' delete the position item from this group
			sStat = "DELETE FROM tblModuleGroupPos WHERE GroupID = " & nGroupID & " AND ModuleID = " & steNForm("ModuleID")
			Call adoExecute(sStat)
			modRefresh True
		Else
			sErrorMsg = steGetText("The module has been deleted from the layout group")
		End If
	Case "moveup"
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblModuleGroupPos " &_
					"SET	OrderNo = OrderNo + 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & steNForm("OrderNo") - 1
			Call adoExecute(sStat)

			sStat = "UPDATE	tblModuleGroupPos " &_
					"SET	OrderNo = " & (steNForm("OrderNo") - 1) & ", Modified = " & adoGetDate & " " &_
					"WHERE	ModuleID = " & steNForm("ModuleID")
			Call adoExecute(sStat)
			modRefresh True
	Case "movedown"
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblModuleGroupPos " &_
					"SET	OrderNo = OrderNo - 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & steNForm("OrderNo") + 1
			Call adoExecute(sStat)

			sStat = "UPDATE	tblModuleGroupPos " &_
					"SET	OrderNo = " & (steNForm("OrderNo") + 1) & ", Modified = " & adoGetDate & " " &_
					"WHERE	ModuleID = " & steNForm("ModuleID")
			Call adoExecute(sStat)
			modRefresh True
	Case "activ"
			' archive a module
			sStat = "UPDATE	tblModuleGroupPos " &_
					"SET	Archive = 1, Modified = " & adoGetDate & " " &_
					"WHERE	ModuleID = " & steNForm("ModuleID")
			Call adoExecute(sStat)
			modRefresh True
	Case "deactiv"
			sStat = "UPDATE	tblModuleGroupPos " &_
					"SET	Archive = 0, Modified = " & adoGetDate & " " &_
					"WHERE	ModuleID = " & steNForm("ModuleID")
			Call adoExecute(sStat)
			modRefresh True
End Select

' retrieve the group user is working with
sStat = "SELECT	GroupID, GroupCode, GroupName, Modified " &_
		"FROM	tblModuleGroup " &_
		"WHERE	GroupID = " & nGroupID
Set rsGroup = adoOpenRecordset(sStat)
If Not rsGroup.EOF Then
	sGroupName = rsGroup.Fields("GroupName").Value
Else
	sErrorMsg = steGetText("Invalid Group ID Specified") & " (ID = " & nGroupID & ")"
End If
rsGroup.Close
Set rsGroup = Nothing

' retrieve the list of positions that we may insert before
sStat = "SELECT m.ModuleID, m.Title, mgp.OrderNo, mgp.Modified, mgp.Archive " &_
		"FROM	tblModuleGroupPos mgp " &_
		"INNER JOIN	tblModule m ON m.ModuleID = mgp.ModuleID " &_
		"WHERE	mgp.GroupID = " & nGroupID & " " &_
		"AND	mgp.Active <> 0 " &_
		"ORDER BY mgp.OrderNo"
Set rsIns = adoOpenRecordset(sStat)

' retrieve the list of modules which the user may add
sStat = "SELECT	DISTINCT m.ModuleID, m.Title " &_
		"FROM	tblModule m " &_
		"LEFT JOIN tblModuleGroupPos mgp ON mgp.ModuleID = m.ModuleID " &_
		"		AND mgp.GroupID = " & nGroupID & " " &_
		"		AND mgp.Active <> 0 " &_
		"		AND mgp.Archive = 0 " &_
		"WHERE	mgp.ModuleID IS NULL " &_
		"AND	m.Archive = 0 " &_
		"ORDER BY m.Title"
Set rsAdd = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Group" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3>"<%= sGroupName %>" <% steTxt "Layout" %></H3>

<P>
<% steTxt "Define the modules which will appear in the module group" %> <b><%= sGroupName %></b>.
<% steTxt "The changes you make to the module group layout will take effect immediately." %>
</P>

<!-- ADD FORM -->
<form method="post" action="group_layout.asp">
<input type="hidden" name="Action" value="add">
<input type="hidden" name="GroupID" value="<%= nGroupID %>">

<h4><% steTxt "Add New Module" %></h4>

<p>
<% steTxt "Select the module and position where you would like to add a module to this layout group." %>
</p>

<% If sErrorMsg <> "" Then %>
<p><b class="error"><%= sErrorMsg %></b></p>
<% End If %>

<table border=0 cellpadding=4 cellspacing=0>
<tr>
	<td class="forml"><% steTxt "Module Name" %><br>
	<select name="ModuleID" class="form">
	<option value=""> -- Choose --
	<% Do Until rsAdd.EOF %>
	<option value="<%= rsAdd.Fields("ModuleID").Value %>"<% If rsAdd.Fields("ModuleID").Value = steNForm("ModuleID") Then Response.Write " SELECTED" %>> <%= rsAdd.Fields("Title").Value %>
	<%	rsAdd.MoveNext
	   Loop %>
	</select>
	</td>
	<td class="forml"><% steTxt "Insert Before" %><br>
	<select name="InsertBeforeID" class="form">
	<option value=""> -- Choose --
	<% If Not rsIns.EOF Then %>
	<% Do Until rsIns.EOF %>
	<option value="<%= rsIns.Fields("OrderNo").Value %>"<% If rsIns.Fields("OrderNo").Value = steNForm("InsertBeforeID") Then Response.Write " SELECTED" %>> <%= rsIns.Fields("Title").Value %>
	<%	rsIns.MoveNext
	   Loop %>
	<% rsIns.MoveFirst %>
	<% End If %>
	<option value="0">* End Of List *
	</select>
	</td><td valign="bottom">
		<input type="submit" name="_add" value=" <% steTxt "Add Module" %> " class="form">
	</td>
</tr>
</table><br>

</form>

<h4><% steTxt "Module Group Layout" %></h4>

<% If Not rsIns.EOF Then %>

<form action="#" method="post">
<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 class="list">
<TR>
	<TD CLASS="listhead"><% steTxt "Order" %></TD>
	<TD CLASS="listhead"><% steTxt "Module Name" %></TD>
	<TD CLASS="listhead"><% steTxt "Active" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Modified" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<% I = 0
	Do Until rsIns.EOF %>
<TR CLASS="list<%= I mod 2 %>">
	<TD><%= rsIns.Fields("OrderNo").Value %></TD>
	<TD><%= rsIns.Fields("Title").Value %></TD>
	<TD><% If Not steRecordBoolValue(rsIns, "Archive") Then %>
		<INPUT NAME="archive" TYPE="checkbox" CHECKED onClick="location.href='group_layout.asp?GroupID=<%= nGroupID %>&moduleid=<%= rsIns.Fields("ModuleID").Value %>&action=activ'">
		<% Else %>
		<INPUT NAME="archive" TYPE="checkbox" onClick="location.href='group_layout.asp?GroupID=<%= nGroupID %>&moduleid=<%= rsIns.Fields("ModuleID").Value %>&action=deactiv'">
		<% End If %>
	</TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsIns.Fields("Modified").Value, vbShortDate) %></TD>
	<TD>
		<A HREF="group_layout.asp?GroupID=<%= nGroupID %>&moduleid=<%= rsIns.Fields("ModuleID").Value %>&orderno=<%= rsIns.Fields("OrderNo").Value %>&action=moveup" class="actionlink"><% steTxt "up" %></A> .
		<A HREF="group_layout.asp?GroupID=<%= nGroupID %>&moduleid=<%= rsIns.Fields("ModuleID").Value %>&orderno=<%= rsIns.Fields("OrderNo").Value %>&action=movedown" class="actionlink"><% steTxt "down" %></A> .
		<A HREF="group_layout.asp?GroupID=<%= nGroupID %>&moduleid=<%= rsIns.Fields("ModuleID").Value %>&action=delete" class="actionlink"><% steTxt "delete" %></A>
	</TD>
</TR>
<%	rsIns.MoveNext
	I = I + 1
   Loop %>
</TABLE>
</form>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No groups exist in the database" %></B></P>

<% End If %>

<p align="center">
	<a href="group_list.asp" class="adminlink">Group List</a>
</p>

<!-- #include file="../../../footer.asp" -->