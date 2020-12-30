<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' group_list.asp
'	Displays a list of the current groups defined in the database
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
Dim rsGroup
Dim sAction

sAction = LCase(steForm("Action"))

Select Case sAction
	Case "moveup"
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblModuleGroup " &_
					"SET	OrderNo = OrderNo + 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & steNForm("OrderNo") - 1
			Call adoExecute(sStat)

			sStat = "UPDATE	tblModuleGroup " &_
					"SET	OrderNo = " & (steNForm("OrderNo") - 1) & ", Modified = " & adoGetDate & " " &_
					"WHERE	GroupID = " & steNForm("GroupID")
			Call adoExecute(sStat)
			modRefresh True
	Case "movedown"
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblModuleGroup " &_
					"SET	OrderNo = OrderNo - 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & steNForm("OrderNo") + 1
			Call adoExecute(sStat)

			sStat = "UPDATE	tblModuleGroup " &_
					"SET	OrderNo = " & (steNForm("OrderNo") + 1) & ", Modified = " & adoGetDate & " " &_
					"WHERE	GroupID = " & steNForm("GroupID")
			Call adoExecute(sStat)
			modRefresh True
End Select

sStat = "SELECT	OrderNo, GroupID, GroupCode, GroupName, OrderNo, Modified " &_
		"FROM	tblModuleGroup " &_
		"ORDER BY OrderNo"
Set rsGroup = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Group" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "Group List" %></H3>

<P>
<% steTxt "The following module layout groups are defined in the database." %>&nbsp;
<% steTxt "These control the main areas of your content layout." %>&nbsp;
<% steTxt "You cannot just create a new group and have it show up on the site," %>&nbsp;
<% steTxt "you need to add special code in the header and footer templates to pull in your layout group." %>
</P>

<% If Not rsGroup.EOF Then %>

<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 class="list">
<TR>
	<TD CLASS="listhead"><% steTxt "Order" %></TD>
	<TD CLASS="listhead"><% steTxt "Code" %></TD>
	<TD CLASS="listhead"><% steTxt "Group Name" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Modified" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<% I = 0
Do Until rsGroup.EOF %>
<TR CLASS="list<%= I mod 2 %>">
	<TD><%= rsGroup.Fields("OrderNo").Value %></TD>
	<TD><%= rsGroup.Fields("GroupCode").Value %></TD>
	<TD><%= rsGroup.Fields("GroupName").Value %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsGroup.Fields("Modified").Value, vbShortDate) %></TD>
	<TD>
		<A HREF="group_list.asp?GroupID=<%= rsGroup.Fields("GroupID").Value %>&orderno=<%= rsGroup.Fields("OrderNo").Value %>&action=moveup" class="actionlink"><% steTxt "up" %></A> .
		<A HREF="group_list.asp?GroupID=<%= rsGroup.Fields("GroupID").Value %>&orderno=<%= rsGroup.Fields("OrderNo").Value %>&action=movedown" class="actionlink"><% steTxt "down" %></A> .
		<A HREF="group_layout.asp?GroupID=<%= rsGroup.Fields("GroupID").Value %>" class="actionlink"><% steTxt "layout" %></A> .
		<A HREF="group_edit.asp?GroupID=<%= rsGroup.Fields("GroupID").Value %>" class="actionlink"><% steTxt "edit" %></A> .
		<A HREF="group_delete.asp?GroupID=<%= rsGroup.Fields("GroupID").Value %>" class="actionlink"><% steTxt "delete" %></A>
	</TD>
</TR>
<%	rsGroup.MoveNext
	I = I + 1
   Loop %>
</TABLE>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No groups exist in the database" %></B></P>

<% End If %>

<P ALIGN="center">
	<A HREF="group_add.asp" class="adminlink"><% steTxt "Add New Group" %></A>
</P>

<!-- #include file="../../../footer.asp" -->