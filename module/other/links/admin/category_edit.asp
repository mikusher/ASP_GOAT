<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' category_edit.asp
'	Edit an existing link caegory from the database
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

nCategoryID = steNForm("CategoryID")

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("categoryname")) = ""	Then
		sErrorMsg = steGetText("Please enter the category name for this link category")
	Else
		' create the link category in the database
		sStat = "UPDATE tblLinkCategory " &_
				"SET	CategoryName= " & steQForm("CategoryName") & " " &_
				"WHERE	CategoryID = " & nCategoryID
		Call adoExecute(sStat)
	End If
End If

' retrieve the category to edit
sStat = "SELECT	* FROM tblLinkCategory WHERE CategoryID = " & nCategoryID
Set rsCat = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Category" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Link Category" %></H3>

<P>
<% steTxt "Please enter the properties for the link category using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="category_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="CategoryID" VALUE="<%= nCategoryID %>">
<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD class="forml"><% steTxt "Category Name" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="CategoryName" VALUE="<%= steRecordEncValue(rsCat, "CategoryName") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 align="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update Category" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Link Category Updated" %></H3>

<P>
<% steTxt "The link category was successfully updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align="center">
	<a href="category_list.asp" class="adminlink"><% steTxt "Category List" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->
