<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' author_edit.asp
'	Edit an existing article author to the database
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
Dim rsAuthor
Dim nAuthorID

nAuthorID = steNForm("authorid")

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("firstname")) = ""	Then
		sErrorMsg = steGetText("Please enter the first name for this article author")
	ElseIf Trim(steForm("lastname")) = "" Then
		sErrorMsg = steGetText("Please enter the last name for this article author")
	Else
		' create the author in the database
		sStat = "UPDATE tblArticleAuthor " &_
				"SET	Title = " & steQForm("Title") & "," &_
				"		FirstName = " & steQForm("FirstName") & "," &_
				"		MiddleName = " & steQForm("MiddleName") & "," &_
				"		LastName = " & steQForm("lastName") & "," &_
				"		Surname = " & steQForm("Surname") & " " &_
				"WHERE	AuthorID = " & nAuthorID
		Call adoExecute(sStat)
	End If
End If

' retrieve the author to edit
sStat = "SELECT	* FROM tblArticleAuthor WHERE AuthorID = " & nAuthorID
Set rsAuthor = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Author" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Article Author" %></H3>

<P>
<% steTxt "Please enter the properties for the article author using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="author_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="authorid" VALUE="<%= nAuthorID %>">
<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD class="forml" nowrap><% steTxt "Title" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Title" VALUE="<%= steRecordEncValue(rsAuthor, "Title") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "First Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="FirstName" VALUE="<%= steRecordEncValue(rsAuthor, "FirstName") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD class="forml" nowrap><% steTxt "Middle Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="MiddleName" VALUE="<%= steRecordEncValue(rsAuthor, "MiddleName") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD class="forml" nowrap><% steTxt "Last Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="LastName" VALUE="<%= steRecordEncValue(rsAuthor, "LastName") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD class="forml" nowrap><% steTxt "Surname" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Surname" VALUE="<%= steRecordEncValue(rsAuthor, "Surname") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update Author" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Article Author Updated" %></H3>

<P>
<% steTxt "The article author was successfully updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align="center">
	<a href="article_list.asp" class="adminlink"><% steTxt "Article List" %></a> &nbsp;
	<a href="author_list.asp" class="adminlink"><% steTxt "Author List" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->
