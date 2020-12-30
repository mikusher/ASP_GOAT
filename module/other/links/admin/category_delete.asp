<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' category_delete.asp
'	Delete an existing link category from the database
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

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If steNForm("confirm") <> 1 Then
		sErrorMsg = steGetText("Please confirm the deletion of this link category")
	Else
		' delete the existing link category in the database
		sStat = "DELETE FROM tblLinkCategory " &_
				"WHERE	CategoryID = " & nCategoryID
		Call adoExecute(sStat)
	End If
End If

' retrieve the category information to delete
Set rsCat = adoOpenRecordset("SELECT * FROM tblLinkCategory WHERE CategoryID = " & nCategoryID)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Category" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Link Category" %></H3>

<P>
<% steTxt "Please confirm the deletion of the link category by clicking the yes button next to <I>Confirm</I>." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="category_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="CategoryID" VALUE="<%= nCategoryID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD class="forml"><% steTxt "Category Name" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD class="formd"><%= rsCat.Fields("CategoryName").Value %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Confirm Delete" %></B></TD><TD></TD>
	<TD class="formd">
		<INPUT TYPE="radio" NAME="confirm" VALUE="1" class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="confirm" VALUE="0" checked class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 align="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Link Category" %> " CLASS="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Link Category Deleted" %></H3>

<P>
<% steTxt "The link category was successfully deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align="center">
	<a href="category_list.asp" class="adminlink"><% steTxt "Category List" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->
