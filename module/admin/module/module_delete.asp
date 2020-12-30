<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/module_lib.asp" -->
<%
'--------------------------------------------------------------------
' module_delete.asp
'	Deletes an existing module from the database
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
Dim rsEdit
Dim nModuleID

nModuleID = steNForm("ModuleID")

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If steNForm("Confirm") = 0	Then
		sErrorMsg = steGetText("Please confirm the deletion of this module")
	Else
		' delete module from the database
		sStat = "DELETE FROM tblModule WHERE ModuleID = " & nModuleID
		Call adoExecute(sStat)
		Call modUpdateCategoryCounts
	End If
End If

' retrieve the record to delete here
sStat = "SELECT * FROM tblModule WHERE ModuleID = " & nModuleID
Set rsEdit = adoOpenRecordset(sStat)

If Not rsEdit.EOF Then
	' retrieve the category selected for this module
	sStat = "SELECT	CategoryName " &_
			"FROM	tblModuleCategory " &_
			"WHERE	CategoryID = " & rsEdit.Fields("CategoryID").Value
	Set rsCat = adoOpenRecordset(sStat)
	If Not rsCat.EOF Then sCatName = rsCat.Fields("CategoryName").Value
	rsCat.Close
	Set rsCat = Nothing
End If
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Module" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Module" %></H3>

<P>
<% steTxt "Please confirm that you would like to delete the module shown below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="module_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="ModuleID" Value="<%= nModuleID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Category" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD>
		<% If Trim(sCatName) = "" Then %><%= sCatName %><% Else %><I>n/a</I><% End If %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Title" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsEdit, "Title") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Synopsis" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsEdit, "Synopsis") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Description" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsEdit, "Description") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Folder Name" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsEdit, "FolderName") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Version No" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsEdit, "VersionNo") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Small Capsule Module" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsEdit, "Size140Module") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Full Module" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsEdit, "SizeFullModule") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Update URL" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsEdit, "UpdateURL") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Do Update Check" %></TD><TD></TD>
	<TD CLASS="formd">
		<% If steRecordBoolValue(rsEdit, "DoUpdateCheck") Then Response.Write steGetText("Yes") Else Response.Write steGetText("No") %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Check Days" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsEdit, "CheckDays") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Confirm Delete?" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="Confirm" VALUE="1" CLASS="formradio"<% If steEncForm("Confirm") = "True" Then Response.Write " CHECKED" %>"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="Confirm" VALUE="0" CLASS="formradio"<% If steEncForm("Confirm") <> "True" Then Response.Write " CHECKED" %>"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Module" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Module Deleted" %></H3>

<P>
<% steTxt "The module has been deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="module_list.asp" class="adminlink"><% steTxt "Module List" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
