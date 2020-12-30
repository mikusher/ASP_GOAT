<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/order_lib.asp" -->
<%
'--------------------------------------------------------------------
' right_list.asp
'	Displays a list of the admin rights for the site administration
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
Dim rsCount
Dim rsState
Dim rsCountry
Dim sAction
Dim sWhere		' where clause for search
Dim sCriteria	' criteria we are searching on
Dim sErrorMsg
Dim I

sAction = LCase(steForm("action"))

Select Case sAction
	Case "moveup"
		Call ordMoveUp("access right", "tblUserRight", "RightID", "ParentRightID", steNForm("ParentID"), steNForm("OrderNo"), steNForm("RightID"), sErrorMsg)
	Case "movedown"
		Call ordMoveDown("access right", "tblUserRight", "RightID", "ParentRightID", steNForm("ParentID"), steNForm("OrderNo"), steNForm("RightID"), sErrorMsg)
End Select

' retrieve the list of user rights to display here
sStat = "SELECT	RightID, ParentRightID, RightName, " &_
		"		Hyperlink, Modified, OrderNo, HasAdd, HasEdit, HasDelete, HasView " &_
		"FROM	tblUserRight " &_
		"WHERE	Archive = 0 " &_
		"ORDER BY OrderNo"
Set rsList = adoOpenRecordset(sStat)
If Not rsList.EOF Then aRight = rsList.GetRows
rsList.Close

%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Rights" %>
<!-- #include file="pagetabs_inc.asp" -->

<!-- display all of the rights here (max of 30) -->

<H3>User Right List</H3>

<P ALIGN="center">
<I>Total of <%= UBound(aRight, 2)+1 %> User Rights Found</I>
</P>

<% If sErrorMsg <> "" Then %>
<p><b class="error"><%= sErrorMsg %></b></p>
<% End If %>

<% If IsArray(aRight) Then %>

<P>
<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 class="list">
<TR>
	<TD class="listhead"><% steTxt "Right Name" %></TD>
	<TD class="listhead"><% steTxt "Add" %></TD>
	<TD class="listhead"><% steTxt "Edit" %></TD>
	<TD class="listhead"><% steTxt "Delete" %></TD>
	<TD class="listhead"><% steTxt "View" %></TD>

	<TD class="listhead" ALIGN="right"><% steTxt "Modified" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<% ' display the tree of rights here
	Call locRightTree(0, 0, 0) %>
</TABLE>
</P>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No user rights found to display here" %></B></P>

<% End If %>

<P ALIGN="center">
	<A HREF="userright_add.asp" class="adminlink"><% steTxt "Add New User Right" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
<%
Sub locRightTree(nParentID, nIndex, nLevel)
	Dim I

	If CStr(nIndex) = "" Then nIndex = 0
	' find all rights matching the parent and display
	For I = 0 To UBound(aRight, 2) 
		If aRight(1, I) = nParentID Then 
			'  onMouseOver="this.className='listsel'" onMouseOut="this.className='list<= nIndex mod 2 >'"
			%>
<TR CLASS="list<%= nIndex mod 2 %>">
	<TD><table border=0 cellpadding=2 cellspacing=0>
	<tr>
		<TD WIDTH="<%= nLevel * 15 %>"><IMG SRC="<%= Application("ASPNukeBasePath") %>img/pixel.gif" WIDTH="<%= nLevel * 15 %>" HEIGHT="1" ALT=""></TD>
		<TD WIDTH="100%"><%= aRight(2, I) %></TD>
	</tr>
	</table></TD>
	<TD align="center"><img src="<%= Application("ASPNukeBasePath") %>img/<% If aRight(6, I) = 1 Then Response.Write "check.gif" Else Response.Write "redex.gif" %>" alt=""></TD>
	<TD align="center"><img src="<%= Application("ASPNukeBasePath") %>img/<% If aRight(7, I) = 1 Then Response.Write "check.gif" Else Response.Write "redex.gif" %>" alt=""></TD>
	<TD align="center"><img src="<%= Application("ASPNukeBasePath") %>img/<% If aRight(8, I) = 1 Then Response.Write "check.gif" Else Response.Write "redex.gif" %>" alt=""></TD>
	<TD align="center"><img src="<%= Application("ASPNukeBasePath") %>img/<% If aRight(9, I) = 1 Then Response.Write "check.gif" Else Response.Write "redex.gif" %>" alt=""></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(aRight(4, I), vbShortDate) %></TD>
	<TD ALIGN="right">
		<A HREF="usertoright.asp?RightID=<%= aRight(0, I) %>" class="actionlink"><% steTxt "users" %></A> .
		<A HREF="userright_list.asp?RightID=<%= aRight(0, I) %>&ParentID=<%= aRight(1, I) %>&orderno=<%= aRight(5, I) %>&action=moveup" class="actionlink"><img src="<%= Application("ASPNukeBasePath") %>img/moveup.gif" alt="<% steTxt "up" %>"></A>
		<A HREF="userright_list.asp?RightID=<%= aRight(0, I) %>&ParentID=<%= aRight(1, I) %>&orderno=<%= aRight(5, I) %>&action=movedown" class="actionlink"><img src="<%= Application("ASPNukeBasePath") %>img/movedown.gif" alt="<% steTxt "down" %>"></A>
		<A HREF="userright_edit.asp?RightID=<%= aRight(0, I) %>" class="actionlink"><img src="<%= Application("ASPNukeBasePath") %>img/edit.gif" alt="<% steTxt "edit" %>"></A>
		<A HREF="userright_delete.asp?RightID=<%= aRight(0, I) %>" class="actionlink"><img src="<%= Application("ASPNukeBasePath") %>img/delete.gif" alt="<% steTxt "delete" %>"></A>
	</TD>
</TR>
<%			' display the child rights here
			nIndex = nIndex + 1
			Call locRightTree(aRight(0, I), nIndex, nLevel+1)
		End If
	Next
End Sub
%>