<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' category_delete.asp
'	Delete an existing article category to the database
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

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If Trim(steNForm("Confirm")) <> 1	Then
		sErrorMsg = steGetText("Please confirm the deletion of this category")
	Else
		' create the new article category in the database
		sStat = "DELETE FROM tblArticleCategory WHERE CategoryID = " & nCategoryID
		Call adoExecute(sStat)
	End If
End If

sStat = "SELECT * FROM tblArticleCategory WHERE CategoryID = " & nCategoryID
Set rsCat = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Category" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Article Category" %></H3>

<P>
<% steTxt "Please confirm that you would like to delete this article category by clicking <I>Yes</I> next to <B>Confirm</B> below." %>&nbsp;
<% steTxt "Once the category has been deleted, it may not be recovered." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="category_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="categoryid" VALUE="<%= nCategoryID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD class="forml"><% steTxt "Name" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD class="formd"><%= steRecordEncValue(rsCat, "CategoryName") %></TD>
</TR><TR>
	<TD class="forml" VALIGN="top"><% steTxt "Comments" %></TD><TD></TD>
	<TD><%= steRecordEncValue(rsCat, "Comments") %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Confirm Delete" %></TD><TD></TD>
	<TD><INPUT TYPE="radio" NAME="confirm" VALUE="1" class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="confirm" VALUE="0" CHECKED class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Category" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Article Category Deleted" %></H3>

<P>
<% steTxt "The article category was permanently deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<!-- #include file="../../../../footer.asp" -->
