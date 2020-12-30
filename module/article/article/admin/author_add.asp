<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' author_add.asp
'	Add a new article author to the database
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

If steForm("action") = "add" Then
	' make sure the required fields are present
	If Trim(steForm("firstname")) = ""	Then
		sErrorMsg = steGetText("Please enter the first name for this new article author")
	ElseIf Trim(steForm("lastname")) = "" Then
		sErrorMsg = steGetText("Please enter the last name for this new article author")
	Else
		' create the new author in the database
		sStat = "INSERT INTO tblArticleAuthor (" &_
				"	Title, FirstName, MiddleName, LastName, Surname, Created " &_
				") VALUES (" &_
				steQForm("Title") & "," &_
				steQForm("FirstName") & "," & steQForm("MiddleName") & "," &_
				steQForm("lastName") & "," & steQForm("Surname") & "," & adoGetDate &_
				")"
		Call adoExecute(sStat)
	End If
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Author" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New Article Author" %></H3>

<P>
<% steTxt "Please enter the new properties for the new article author using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="author_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD CLASS="forml"><% steTxt "Title" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="Title" VALUE="<%= steEncForm("Title") %>" SIZE="32" MAXLENGTH="32" CLASS="form"></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "First Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="FirstName" VALUE="<%= steEncForm("FirstName") %>" SIZE="32" MAXLENGTH="32" CLASS="form"></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Middle Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="MiddleName" VALUE="<%= steEncForm("MiddleName") %>" SIZE="32" MAXLENGTH="32" CLASS="form"></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Last Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="LastName" VALUE="<%= steEncForm("LastName") %>" SIZE="32" MAXLENGTH="32" CLASS="form"></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Surname" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Surname" VALUE="<%= steEncForm("Surname") %>" SIZE="32" MAXLENGTH="32" CLASS="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Author" %> " CLASS="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "New Article Author Added" %></H3>

<P>
<% steTxt "The new article author has been added to the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align="center">
	<a href="article_list.asp" class="adminlink"><% steTxt "Article List" %></a> &nbsp;
	<a href="author_list.asp" class="adminlink"><% steTxt "Author List" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->
