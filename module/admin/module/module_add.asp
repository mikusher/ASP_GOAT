<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/module_lib.asp" -->
<%
'--------------------------------------------------------------------
' module_add.asp
'	Add a new module to the database
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
	If steNForm("CategoryID") = 0 Then
		sErrorMsg = steGetText("Please select the category for the module")
	ElseIf Trim(steForm("Title")) = ""	Then
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
		sStat = "INSERT INTO tblModule (" &_
				"	CategoryID, Title, Synopsis, Description, FolderName, VersionNo, " &_
				"	Size140Module, SizeFullModule, UpdateURL, DoUpdateCheck, " &_
				"	CheckDays, Created" &_
				") VALUES (" &_
				steNForm("CategoryID") & "," &_
				steQForm("Title") & "," &_
				steQForm("Synopsis") & "," &_
				steQForm("Description") & "," &_
				steQForm("FolderName") & "," &_
				steQForm("VersionNo") & "," &_
				steQForm("Size140Module") & "," &_
				steQForm("SizeFullModule") & "," &_
				steQForm("UpdateURL") & "," &_
				steNForm("DoUpdateCheck") & "," &_
				steNForm("CheckDays") & "," & adoGetDate &_
				")"
		Call adoExecute(sStat)
		Call modUpdateCategoryCounts
	End If
End If

' build the list of categories to choose from
sStat = "SELECT	CategoryID, CategoryName " &_
		"FROM	tblModuleCategory " &_
		"WHERE	Archive = 0 " &_
		"ORDER BY CategoryName"
Set rsCat = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Module" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New Module" %></H3>

<P>
<% steTxt"Please enter the new properties for the new module using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="module_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Category" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><SELECT NAME="CategoryID" CLASS="form">
		<option value="0">-- Choose --
		<% Do Until rsCat.EOF %>
		<option value="<%= rsCat.Fields("CategoryID").Value %>"<% If steEncForm("CategoryID") = CStr(rsCat.Fields("CategoryID").Value) Then Response.Write " SELECTED" %>> <%= rsCat.Fields("CategoryName").Value %>
		<% rsCat.MoveNext
		   Loop %>
		</SELECT>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Title" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Title" VALUE="<%= steEncForm("Title") %>" SIZE="32" MAXLENGTH="40" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Synopsis" %></TD><TD></TD>
	<TD><TEXTAREA NAME="Synopsis" COLS="38" ROWS="8" CLASS="form"><%= steEncForm("Synopsis") %></TEXTAREA></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Description" %></TD><TD></TD>
	<TD><TEXTAREA NAME="Description" COLS="52" ROWS="10" WRAP="virtual" class="form"><%= steEncForm("Description") %></TEXTAREA></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Folder Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="FolderName" VALUE="<%= steEncForm("FolderName") %>" SIZE="20" MAXLENGTH="20" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Version No" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="VersionNo" VALUE="<%= steEncForm("VersionNo") %>" SIZE="20" MAXLENGTH="16" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Small Capsule Module" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Size140Module" VALUE="<%= steEncForm("Size140Module") %>" SIZE="32" MAXLENGTH="200" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Full Module" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="SizeFullModule" VALUE="<%= steEncForm("SizeFullModule") %>" SIZE="32" MAXLENGTH="200" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Update URL" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="UpdateURL" VALUE="<%= steEncForm("UpdateURL") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Do Update Check" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="DoUpdateCheck" VALUE="1" CLASS="formradio"<% If steNForm("DoUpdateCheck") <> 0 Then Response.Write " CHECKED" %>> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="DoUpdateCheck" VALUE="0" CLASS="formradio"<% If steNForm("DoUpdateCheck") = 0 Then Response.Write " CHECKED" %>> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Check Days" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="CheckDays" VALUE="<%= steEncForm("CheckDays") %>" SIZE="10" MAXLENGTH="8" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Module" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "New Module Added" %></H3>

<P>
<% steTxt "The new module has been added to the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="module_list.asp" class="adminlink"><% steTxt "Module List" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
