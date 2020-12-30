<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' tab_list.asp
'	Displays a list of the application variable tabs which group
'	the application variables into a tabbed interface or wizard.
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
Dim rsTab
Dim sAction

sAction = LCase(steForm("action"))

Select Case sAction
	Case "moveup"
		Dim rsPrev, sPrevOrder

		' retrieve the previous order no
		sStat = "SELECT	OrderNo " &_
				"FROM	tblApplicationVarTab " &_
				"WHERE	OrderNo < " & steNForm("OrderNo") & " " &_
				"ORDER BY OrderNo DESC"
		Set rsPrev = adoOpenRecordset(sStat)
		If Not rsPrev.EOF Then
			sPrevOrder = rsPrev.Fields("OrderNo").Value
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblApplicationVarTab " &_
					"SET	OrderNo = OrderNo + 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & sPrevOrder
			Call adoExecute(sStat)

			sStat = "UPDATE	tblApplicationVarTab " &_
					"SET	OrderNo = " & sPrevOrder & ", Modified = " & adoGetDate & " " &_
					"WHERE	TabID = " & steNForm("TabID")
			Call adoExecute(sStat)
			modRefresh True
		End If
		rsPrev.Close
		Set rsPrev = Nothing
	Case "movedown"
		Dim rsNext, sNextOrder

		' retrieve the next order no
		sStat = "SELECT	OrderNo " &_
				"FROM	tblApplicationVarTab " &_
				"WHERE	OrderNo > " & steNForm("OrderNo") & " " &_
				"ORDER BY OrderNo"
		Set rsNext = adoOpenRecordset(sStat)
		If Not rsNext.EOF Then
			sNextOrder = rsNext.Fields("OrderNo").Value
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblApplicationVarTab " &_
					"SET	OrderNo = OrderNo - 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & sNextOrder
			Call adoExecute(sStat)

			sStat = "UPDATE	tblApplicationVarTab " &_
					"SET	OrderNo = " & sNextOrder & ", Modified = " & adoGetDate & " " &_
					"WHERE	TabID = " & steNForm("TabID")
			Call adoExecute(sStat)
			modRefresh True
		End If
		rsNext.Close
		Set rsNext = Nothing
End Select

' retrieve the tab to edit
sStat = "SELECT	OrderNo, TabID, TabName, Title, Archive, Modified " &_
		"FROM	tblApplicationVarTab " &_
		"ORDER BY OrderNo"
Set rsTab = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Tabs" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "Application Variable Tabs" %></H3>

<P>
<% steTxt "Configure the list of the application variable tabs which group the application variables into a tabbed interface or wizard." %>
</P>

<% If Not rsTab.EOF Then %>

<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 CLASS="list">
<TR>
	<TD CLASS="listhead"><% steTxt "Order" %></TD>
	<TD CLASS="listhead"><% steTxt "Tab Name" %></TD>
	<TD CLASS="listhead"><% steTxt "Title" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Modified" %></TD>
	<TD CLASS="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<%
I = 0
Do Until rsTab.EOF %>
<TR CLASS="list<%= I mod 2 %>">
	<TD><%= rsTab.Fields("OrderNo").Value %></TD>
	<TD><%= Server.HTMLEncode(rsTab.Fields("TabName").Value) %></TD>
	<TD><% If Len(rsTab.Fields("Title").Value) > 30 Then
			Response.Write Server.HTMLEncode(Left(rsTab.Fields("Title").Value, 30)) & "..."
		Else
			Response.Write Server.HTMLEncode(rsTab.Fields("Title").Value)
		End If %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsTab.Fields("Modified").Value, vbShortDate) %></TD>
	<TD>
		<A HREF="tab_list.asp?tabid=<%= rsTab.Fields("TabID").Value %>&orderno=<%= rsTab.Fields("OrderNo").Value %>&action=moveup" CLASS="actionlink"><% steTxt "up" %></A> .
		<A HREF="tab_list.asp?tabid=<%= rsTab.Fields("TabID").Value %>&orderno=<%= rsTab.Fields("OrderNo").Value %>&action=movedown" CLASS="actionlink"><% steTxt "down" %></A> .
		<A HREF="tab_edit.asp?tabid=<%= rsTab.Fields("TabID").Value %>&orderno=<%= rsTab.Fields("OrderNo").Value %>" CLASS="actionlink"><% steTxt "edit" %></A> .
		<A HREF="tab_delete.asp?tabid=<%= rsTab.Fields("TabID").Value %>&orderno=<%= rsTab.Fields("OrderNo").Value %>" CLASS="actionlink"><% steTxt "delete" %></A>
	</TD>
</TR>
<%	rsTab.MoveNext
	I = I + 1
   Loop %>
</TABLE>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No application variables tabs exist in the database" %></B></P>

<% End If %>

<P ALIGN="center">
	<A HREF="tab_add.asp" CLASS="adminlink"><% steTxt "Add Variable Tab" %></A>
</P>

<!-- #include file="../../../footer.asp" -->