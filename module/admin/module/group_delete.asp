<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' group_delete.asp
'	Delete an existing module group from the database
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
Dim rsGroup
Dim nGroupID

nGroupID = steNForm("groupid")

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If Trim(steNForm("Confirm")) <> 1	Then
		sErrorMsg = steGetText("Please confirm the deletion of this group")
	Else
		' delete the module group from the database
		sStat = "DELETE FROM tblModuleGroup WHERE GroupID = " & nGroupID
		Call adoExecute(sStat)
	End If
End If

sStat = "SELECT * FROM tblModuleGroup WHERE GroupID = " & nGroupID
Set rsGroup = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Group" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Module Group" %></H3>

<P>
<% steTxt "Please confirm that you would like to delete this module group by clicking <I>Yes</I> next to <B>Confirm</B> below." %>&nbsp;
<% steTxt "Once the group has been deleted, it may not be recovered." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="group_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="groupid" VALUE="<%= nGroupID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD class="forml"><% steTxt "Group Code" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD class="formd"><%= steRecordEncValue(rsGroup, "GroupCode") %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Group Name" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsGroup, "GroupName") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has 140 (pixel) Size Modules?" %></TD><TD></TD>
	<TD class="formd"><% If steRecordBoolValue(rsGroup, "HasSize140Module") Then Response.Write "Yes" Else Response.Write "No" %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has Full Size Modules?" %></TD><TD></TD>
	<TD class="formd"><% If steRecordBoolValue(rsGroup, "HasSizeFullModule") Then Response.Write "Yes" Else Response.Write "No" %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Confirm Delete" %></TD><TD></TD>
	<TD><INPUT TYPE="radio" NAME="confirm" VALUE="1" class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="confirm" VALUE="0" checked class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Group" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Module Group Deleted" %></H3>

<P>
<% steTxt "The module group was permanently deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align="center">
	<a href="group_list.asp" class="adminlink">Group List</a>
</p>

<!-- #include file="../../../footer.asp" -->
