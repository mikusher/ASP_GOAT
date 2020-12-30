<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' module_edit.asp
'	Updates an existing module from the database
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

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("Title")) = ""	Then
		sErrorMsg = steGetText("Please enter the title for this module")
	ElseIf Trim(steForm("Synopsis")) = ""	Then
		sErrorMsg = steGetText("Please enter the synopsis for this module")
	ElseIf Trim(steForm("Description")) = ""	Then
		sErrorMsg = steGetText("Please enter the description for this module")
	ElseIf Trim(steForm("FolderName")) = ""	Then
		sErrorMsg = steGetText("Please enter the folder name for this module")
	ElseIf Trim(steForm("VersionNo")) = ""	Then
		sErrorMsg = steGetText("Please enter the version no for this module")
	Else
		' create the new module in the database
		sStat = "UPDATE tblModule SET " &_
				"CategoryID = " & steNForm("CategoryID") & ", " &_
				"Title = " & steQForm("Title") & ", " &_
				"Synopsis = " & steQForm("Synopsis") & ", " &_
				"Description = " & steQForm("Description") & ", " &_
				"FolderName = " & steQForm("FolderName") & ", " &_
				"VersionNo = " & steQForm("VersionNo") & ", " &_
				"Size140Module = " & steQForm("Size140Module") & "," &_
				"SizeFullModule = " & steQForm("SizeFullModule") & "," &_
				"UpdateURL = " & steQForm("UpdateURL") & "," &_
				"DoUpdateCheck = " & steNForm("DoUpdateCheck") & "," &_
				"CheckDays = " & steNForm("CheckDays") & ", " &_
				"Modified = " & adoGetDate & " " &_
				"WHERE	ModuleID = " & nModuleID
		Call adoExecute(sStat)
	End If
End If

' retrieve the list of categories to choose from
sStat = "SELECT	CategoryID, CategoryName " &_
		"FROM	tblModuleCategory " &_
		"WHERE	Archive = 0 " &_
		"ORDER BY CategoryName"
Set rsCat = adoOpenRecordset(sStat)

' retrieve the record to edit here
sStat = "SELECT * FROM tblModule WHERE ModuleID = " & nModuleID
Set rsEdit = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Module" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Module" %></H3>

<P>
<% steTxt "Please enter the new properties for the new module using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="module_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="ModuleID" Value="<%= nModuleID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Category" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD>
		<SELECT NAME="CategoryID" class="form">
		<OPTION VALUE="0"> -- Choose --
		<% Do Until rsCat.EOF
			Response.Write "<OPTION VALUE=""" & rsCat.Fields("CategoryID").Value & """"
			If CStr(steRecordValue(rsEdit, "CategoryID")) = CStr(rsCat.Fields("CategoryID").Value) Then Response.Write " SELECTED"
			Response.Write "> " & rsCat.Fields("CategoryName").Value
			rsCat.MoveNext
		   Loop %>
		</SELECT>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Title" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Title" VALUE="<%= steRecordEncValue(rsEdit, "Title") %>" SIZE="32" MAXLENGTH="40" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Synopsis" %></TD><TD></TD>
	<TD><TEXTAREA NAME="Synopsis" COLS="38" ROWS="8" CLASS="form"><%= steRecordEncValue(rsEdit, "Synopsis") %></TEXTAREA></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Description" %></TD><TD></TD>
	<TD><TEXTAREA NAME="Description" COLS="52" ROWS="10" WRAP="virtual" class="form"><%= steRecordEncValue(rsEdit, "Description") %></TEXTAREA></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Folder Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="FolderName" VALUE="<%= steRecordEncValue(rsEdit, "FolderName") %>" SIZE="20" MAXLENGTH="20" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Version No" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="VersionNo" VALUE="<%= steRecordEncValue(rsEdit, "VersionNo") %>" SIZE="20" MAXLENGTH="16" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Small Capsule Module" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Size140Module" VALUE="<%= steRecordEncValue(rsEdit, "Size140Module") %>" SIZE="32" MAXLENGTH="200" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Full Module" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="SizeFullModule" VALUE="<%= steRecordEncValue(rsEdit, "SizeFullModule") %>" SIZE="32" MAXLENGTH="200" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Update URL" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="UpdateURL" VALUE="<%= steRecordEncValue(rsEdit, "UpdateURL") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Do Update Check" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="DoUpdateCheck" VALUE="1" CLASS="formradio"<% If steRecordBoolValue(rsEdit, "DoUpdateCheck") Then Response.Write " CHECKED" %>> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="DoUpdateCheck" VALUE="0" CLASS="formradio"<% If Not steRecordBoolValue(rsEdit, "DoUpdateCheck") Then Response.Write " CHECKED" %>> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Check Days" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="CheckDays" VALUE="<%= steRecordEncValue(rsEdit, "CheckDays") %>" SIZE="10" MAXLENGTH="8" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update Module" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "New Module Updated" %></H3>

<P>
<% steTxt "The new module has been updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="module_list.asp" class="adminlink"><% steTxt "Module List" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
