<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' topic_add.asp
'	Add a new forum topic to the database.
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
Dim rsTopic

If steForm("action") = "add" Then
	' make sure the required fields are present
	If Trim(steForm("title")) = ""	Then
		sErrorMsg = steGetText("Please enter the title for this new forum topic")
	ElseIf Trim(steForm("shortcomments")) = "" Then
		sErrorMsg = steGetText("Please enter the short description for this new forum topic")
	Else
		' create the new forum topic in the database
		sStat = "INSERT INTO tblMessageTopic (" &_
				"	Title, ShortComments, Created " &_
				") VALUES (" &_
				steQForm("Title") & "," &_
				steQForm("ShortComments") & "," & adoGetDate &_
				")"
		Call adoExecute(sStat)
	End If
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Topic" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New Forum Topic" %></H3>

<P>
<% steTxt "Please enter the new properties for the new forum topic using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="topic_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD CLASS="forml"><% steTxt "Title" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="Title" VALUE="<%= steEncForm("Title") %>" SIZE="32" MAXLENGTH="100" CLASS="form"></TD>
</TR><TR>
	<TD CLASS="forml" valign="top"><% steTxt "Short Description" %></TD><TD></TD>
	<TD><TEXTAREA NAME="ShortComments" COLS="48" ROWS="6" CLASS="form"><%= steEncForm("ShortComments") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Forum Topic" %> " CLASS="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "New Forum Topic Added" %></H3>

<P>
<% steTxt "The new forum topic has been added to the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align="center">
	<a href="topic_list.asp" class="adminlink"><% steTxt "Topic List" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->
