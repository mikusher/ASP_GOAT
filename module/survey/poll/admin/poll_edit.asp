<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../cache.asp" -->
<%
'--------------------------------------------------------------------
' poll_edit.asp
'	Update existing poll question in the database
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
Dim nPollID

nPollID = steNForm("pollid")

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("Question")) = ""	Then
		sErrorMsg = steGetText("Please enter the question for this poll")
	Else
		' update the existing poll question in the database
		sStat = "UPDATE tblPoll SET " &_
				"	Question = " & steQForm("Question") & " " &_
				"WHERE	CategoryID = " & nPollID
		Call adoExecute(sStat)

		Call modCapsuleCache(True)
	End If
End If

sStat = "SELECT * FROM tblPoll WHERE PollID = " & nPollID
Set rsPoll = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Poll" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Poll Answer" %></H3>

<P>
<% steTxt "Please make your changes to the poll answer using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="poll_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="pollid" VALUE="<%= nPollID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Question" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="Question" VALUE="<%= steRecordEncValue(rsPoll, "Question") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update Question" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Poll Question Updated" %></H3>

<P>
<% steTxt "The poll question was successfully updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align="center">
	<a href="poll_list.asp" class="adminlink"><% steTxt "Poll List" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->
