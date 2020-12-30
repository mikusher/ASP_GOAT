<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' group_add.asp
'	Add a new module group to the database
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

Dim sErrorMsg
Dim sStat
Dim rsCat

If steForm("action") = "add" Then
	' make sure the required fields are present
	If Trim(steForm("GroupCode")) = "" Then
		sErrorMsg = steGetText("Please enter the Group Code for this group")
	ElseIf Trim(steForm("GroupName")) = ""	Then
		sErrorMsg = steGetText("Please enter the Group Name for this group")
	ElseIf steNForm("HasSize140Module") = 0 And steNForm("HasSizeFullModule") = 0 Then
		sErrorMsg = steGetText("You must choose Has Size 140 or Has Full Size Modules")
	Else
		' determine the new orderno for the entry
		Dim rsOrder, nOrder
		Set rsOrder = adoOpenRecordset("SELECT Coalesce(Max(OrderNo) + 1, 1) AS OrderNo FROM tblModuleGroup")
		If Not rsOrder.EOF Then nOrderNo = rsOrder.Fields("OrderNo").Value Else nOrderNo = 1
		rsOrder.Close
		Set rsOrder = Nothing

		' create the new module group in the database
		sStat = "INSERT INTO tblModuleGroup (" &_
				"	GroupCode, GroupName, HasSize140Module, HasSizeFullModule, OrderNo, Created " &_
				") VALUES (" &_
				steQForm("GroupCode") & "," &_
				steQForm("GroupName") & "," &_
				steNForm("HasSize140Module") & "," &_
				steNForm("HasSizeFullModule") & "," &_
				nOrderNo & "," & adoGetDate &_
				")"
		Call adoExecute(sStat)
	End If
End If
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Group" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New Module Group" %></H3>

<P>
<% steTxt "Please enter the properties for the new module group using the form below." %>&nbsp;
<% steTxt "You should only create new groups if you understand the concept of adding layout groups to the site templates." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="group_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Group Code" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="GroupCode" VALUE="<%= steEncForm("GroupCode") %>" SIZE="6" MAXLENGTH="4" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Group Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="GroupName" VALUE="<%= steEncForm("GroupName") %>" SIZE="32" MAXLENGTH="40" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has 140 (pixel) Size Modules?" %></TD><TD></TD>
	<TD class="formd">
		<INPUT TYPE="radio" NAME="HasSize140Module" VALUE="1"<% If steNForm("HasSize140Module") <> 0 Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="HasSize140Module" VALUE="0"<% If steNForm("HasSize140Module") = 0 Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has Full Size Modules?" %></TD><TD></TD>
	<TD class="formd">
		<INPUT TYPE="radio" NAME="HasSizeFullModule" VALUE="1"<% If steNForm("HasSizeFullModule") <> 0 Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="HasSizeFullModule" VALUE="0"<% If steNForm("HasSizeFullModule") = 0 Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Group" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "New Module Group Added" %></H3>

<P>
<% steTxt "The new module group has been added to the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="group_list.asp" class="adminlink"><% steTxt "Group List" %></A> &nbsp;
	<A HREF="module_list.asp" class="adminlink"><% steTxt "Module List" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
