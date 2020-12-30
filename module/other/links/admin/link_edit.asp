<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' link_edit.asp
'	Update an existing link in the database
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

Dim sStat
Dim sAction
Dim rsLink
Dim rsCat		' list of categories to choose from
Dim sErrorMsg	' error message to display to user

sAction = LCase(steForm("action"))

If sAction = "edit" Then
	' check for required fields here
	If steNForm("categoryid") = 0 Then
		sErrorMsg = steGetText("Please select a category for this link")
	ElseIf Trim(steForm("URL")) = "" Then
		sErrorMsg = steGetText("Please enter the URL for this link")
	ElseIf Trim(steForm("Label")) = "" Then
		sErrorMsg = steGetText("Plese enter the Label for this link")
	Else
		' update the existing article in the database
		sStat = "UPDATE tblLink SET " &_
				"CategoryID = " & steNForm("CategoryID") & "," &_
				"URL = " & steQForm("URL") & "," &_
				"Label = " & steQForm("Label") & " " &_
				"WHERE	LinkID = " & steNForm("LinkID")
		Call adoExecute(sStat)
	End If
End If

' build the list of categories to choose from
sStat = "SELECT	CategoryID, CategoryName " &_
		"FROM	tblLinkCategory " &_
		"WHERE	Active <> 0 " &_
		"AND	Archive = 0 " &_
		"ORDER BY CategoryName"
Set rsCat = adoOpenRecordset(sStat)

' retrieve the article we are editing
sStat = "SELECT * FROM tblLink WHERE LinkID = " & steNForm("LinkID")
Set rsLink = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Link" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If sAction <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Link" %></H3>

<P>
<% steTxt "Please enter the changes for the link in the form below." %>
</P>

<FORM METHOD="post" ACTION="link_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="LinkID" VALUE="<%= steNForm("LinkID") %>">

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Category" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><SELECT NAME="CategoryID" class="form">
	<OPTION VALUE="0"> -- <% steTxt "Choose" %> --
	<% Do Until rsCat.EOF %>
	<OPTION VALUE="<%= rsCat.Fields("CategoryID").Value %>"<% If CStr(rsCat.Fields("CategoryID").Value) = steRecordEncValue(rsLink, "CategoryID") Then Response.Write " SELECTED" %>> <%= rsCat.Fields("CategoryName").Value %>
	<%	rsCat.MoveNext
	   Loop %>
	</SELECT>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "URL" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="URL" VALUE="<%= steRecordEncValue(rsLink, "URL") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Label" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Label" VALUE="<%= steRecordEncValue(rsLink, "Label") %>" SIZE="32" MAXLENGTH="50" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update Link" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Link Updated" %></H3>

<P>
<% steTxt "The changes to the link were made successfully." %>&nbsp;
<% steTxt "You may use the admin links above to continue administering the site." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="link_list.asp" class="adminlink"><% steTxt "Link List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->