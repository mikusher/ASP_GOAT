<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' category_add.asp
'	Add a new module category to the database
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
	If Trim(steForm("CategoryName")) = ""	Then
		sErrorMsg = steGetText("Please enter the name for this category")
	Else
		' create the new module category in the database
		sStat = "INSERT INTO tblModuleCategory (" &_
				"	CategoryName, FolderName, Description, Created" &_
				") VALUES (" &_
				steQForm("CategoryName") & "," &_
				steQForm("FolderName") & "," &_
				steQForm("Description") & "," & adoGetDate &_
				")"
		Call adoExecute(sStat)
	End If
End If
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Category" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New Module Category" %></H3>

<P>
<% steTxt "Please enter the new properties for the new module category using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="category_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Name" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="CategoryName" VALUE="<%= steEncForm("CategoryName") %>" SIZE="32" MAXLENGTH="40" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Folder Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="FolderName" VALUE="<%= steEncForm("FolderName") %>" SIZE="32" MAXLENGTH="20" CLASS="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Comments" %></TD><TD></TD>
	<TD><TEXTAREA NAME="Description" COLS="52" ROWS="10" WRAP="virtual" class="form"><%= steEncForm("Description") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Category" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "New Module Category Added" %></H3>

<P>
<% steTxt "The new module category has been added to the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="category_list.asp" class="adminlink"><% steTxt "Category List" %></A> &nbsp;
	<A HREF="module_list.asp" class="adminlink"><% steTxt "Module List" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
