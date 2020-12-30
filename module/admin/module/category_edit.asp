<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' category_edit.asp
'	Update existing module category in the database
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
Dim nCategoryID

nCategoryID = steNForm("categoryid")

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("CategoryName")) = ""	Then
		sErrorMsg = steGetText("Please enter the name for this category")
	Else
		' create the new module category in the database
		sStat = "UPDATE tblModuleCategory SET " &_
				"	CategoryName = " & steQForm("CategoryName") & "," &_
				"	FolderName = " & steQForm("FolderName") & "," &_
				"	Description = " & steQForm("Description") & " " &_
				"WHERE	CategoryID = " & nCategoryID
		Call adoExecute(sStat)
	End If
End If

sStat = "SELECT * FROM tblModuleCategory WHERE CategoryID = " & nCategoryID
Set rsCat = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Category" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Module Category" %></H3>

<P>
<% steTxt "Please make your changes to the module category using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="category_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="categoryid" VALUE="<%= nCategoryID %>">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Name" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="CategoryName" VALUE="<%= steRecordEncValue(rsCat, "CategoryName") %>" SIZE="32" MAXLENGTH="40" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Folder Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="FolderName" VALUE="<%= steRecordEncValue(rsCat, "FolderName") %>" SIZE="32" MAXLENGTH="20" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Description" %></TD><TD></TD>
	<TD><TEXTAREA NAME="Description" COLS="52" ROWS="10" WRAP="virtual" class="form"><%= steRecordEncValue(rsCat, "Description") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update Category" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Module Category Updated" %></H3>

<P>
<% steTxt "The module category was successfully updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align="center">
	<a href="category_list.asp" class="adminlink">Category List</a>
</p>

<!-- #include file="../../../footer.asp" -->
