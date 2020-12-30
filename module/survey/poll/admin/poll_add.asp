<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../cache.asp" -->
<%
'--------------------------------------------------------------------
' poll_add.asp
'	Add a new poll question to the database
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
Dim rsPoll

If steForm("action") = "add" Then
	' make sure the required fields are present
	If Trim(steForm("Question")) = ""	Then
		sErrorMsg = steGetText("Please enter the question for this poll")
	Else
		' create the new poll question in the database
		sStat = "INSERT INTO tblPoll (" &_
				"	Question, Created " &_
				") VALUES (" &_
				steQForm("Question") & "," & adoGetDate &_
				")"
		Call adoExecute(sStat)

		Call modCapsuleCache(True)
	End If
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Poll" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New Poll Question" %></H3>

<P>
<% steTxt "Please enter the new properties for the new poll question using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="poll_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Question" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="Question" VALUE="<%= steEncForm("Question") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR CLASS="forml">
	<TD COLSPAN=3 align="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Question" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "New Poll Question Added" %></H3>

<P>
<% steTxt "The new poll question has been added to the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="poll_list.asp" class="adminlink"><% steTxt "Poll List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->
