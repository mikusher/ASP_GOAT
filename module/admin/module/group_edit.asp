<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' group_edit.asp
'	Update existing module group in the database
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

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("GroupName")) = ""	Then
		sErrorMsg = steGetText("Please enter the name for this group")
	ElseIf Trim(steForm("GroupName")) = ""	Then
		sErrorMsg = steGetText("Please enter the Group Name for this group")
	ElseIf steNForm("HasSize140Module") = 0 And steNForm("HasSizeFullModule") = 0 Then
		sErrorMsg = steGetText("You must choose Has Size 140 or Has Full Size Modules")
	Else
		' create the new module group in the database
		sStat = "UPDATE tblModuleGroup SET " &_
				"	GroupCode = " & steQForm("GroupCode") & "," &_
				"	GroupName = " & steQForm("GroupName") & "," &_
				"	HasSize140Module = " & steNForm("HasSize140Module") & "," &_
				"	HasSizeFullModule = " & steNForm("HasSizeFullModule") & " " &_
				"WHERE	GroupID = " & nGroupID
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

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Module Group" %></H3>

<P>
<% steTxt "Please make your changes to the module group using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="group_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="groupid" VALUE="<%= nGroupID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Group Code" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="GroupCode" VALUE="<%= steRecordEncValue(rsGroup, "GroupCode") %>" SIZE="6" MAXLENGTH="4" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Group Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="GroupName" VALUE="<%= steRecordEncValue(rsGroup, "GroupName") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has 140 (pixel) Size Modules?" %></TD><TD></TD>
	<TD class="formd">
		<INPUT TYPE="radio" NAME="HasSize140Module" VALUE="1"<% If steRecordBoolValue(rsGroup, "HasSize140Module") Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="HasSize140Module" VALUE="0"<% If Not steRecordBoolValue(rsGroup, "HasSize140Module") Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has Full Size Modules?" %></TD><TD></TD>
	<TD class="formd">
		<INPUT TYPE="radio" NAME="HasSizeFullModule" VALUE="1"<% If steRecordBoolValue(rsGroup, "HasSizeFullModule") Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="HasSizeFullModule" VALUE="0"<% If Not steRecordBoolValue(rsGroup, "HasSizeFullModule") Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update Group" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Module Group Updated" %></H3>

<P>
<% steTxt "The module group was successfully updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align="center">
	<a href="group_list.asp" class="adminlink">Group List</a>
</p>

<!-- #include file="../../../footer.asp" -->
