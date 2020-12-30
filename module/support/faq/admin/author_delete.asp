<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' author_delete.asp
'	Delete an existing FAQ author from the database
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

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If steNForm("confirm") <> 1 Then
		sErrorMsg = steGetText("Please confirm the deletion of this FAQ author")
	Else
		' create the new author in the database
		sStat = "DELETE FROM tblFAQAuthor " &_
				"WHERE	AuthorID = " & nAuthorID
		Call adoExecute(sStat)
	End If
End If

If nAuthorID > 0 Then
	sStat = "SELECT * FROM tblFAQAuthor " &_
			"WHERE AuthorID = " & nAuthorID
	Set rsAuthor = adoOpenRecordset(sStat)
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Author" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete FAQ Author" %></H3>

<P>
<% steTxt "Please confirm the deletion of the FAQ author by clicking the yes button next to <I>Confirm</I>." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="author_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="authorid" VALUE="<%= nAuthorID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml"><% steTxt "Title" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsAuthor, "Title") %></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "First Name" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsAuthor, "FirstName") %></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Middle Name" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsAuthor, "MiddleName") %></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Last Name" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsAuthor, "LastName") %></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Email" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsAuthor, "Email") %></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Confirm Delete?" %></TD><TD></TD>
	<TD CLASS="formd"><INPUT TYPE="radio" NAME="confirm" VALUE="1" CLASS="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="confirm" VALUE="0" CHECKED CLASS="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Author" %> " CLASS="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "FAQ Author Deleted" %></H3>

<P>
<% steTxt "The FAQ author was successfully deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align="center">
	<a href="document_list.asp" class="adminlink"><% steTxt "FAQ List" %></a> &nbsp;
	<a href="author_list.asp" class="adminlink"><% steTxt "Author List" %></a>
</p>
<!-- #include file="../../../../footer.asp" -->
