<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../cache.asp" -->
<%
'--------------------------------------------------------------------
' answer_add.asp
'	Add a new poll answer to the database
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
Dim nPollID
Dim rsPoll

nPollID = steNForm("pollid")

' get the question being asked by this poll
sStat = "SELECT	Question " &_
		"FROM	tblPoll " &_
		"WHERE	PollID = " & nPollID
Set rsPoll =adoOpenRecordset(sStat)

If steForm("action") = "add" Then
	' make sure the required fields are present
	If Trim(steForm("Answer")) = ""	Then
		sErrorMsg = steGetText("Please enter the answer for this poll")
	Else
		' create the new poll answer in the database
		sStat = "INSERT INTO tblPollAnswer (" &_
				"	PollID, Answer, Created " &_
				") VALUES (" &_
				nPollID & "," &_
				steQForm("Answer") & "," & adoGetDate &_
				")"
		Call adoExecute(sStat)

		Call modCapsuleCache(True)
	End If
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Answers" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New Poll Answer" %></H3>

<P>
<% steTxt "Please enter the new properties for the new poll question using the form below." %>
</P>

<P>
Q: <B><%= rsPoll.Fields("Question").Value %></B>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="answer_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">
<INPUT TYPE="hidden" NAME="pollid" VALUE="<%= nPollID %>">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Answer" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="Answer" VALUE="<%= steEncForm("Answer") %>" SIZE="32" MAXLENGTH="50" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 align="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Answer" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<P ALIGN="center">
	<A HREF="answer_list.asp?pollid=<%= nPollID %>" class="adminlink"><% steTxt "Answer List" %></A> &nbsp;
	<A HREF="poll_list.asp" class="adminlink"><% steTxt "Poll List" %></A>
</P>

<% Else %>

<H3><% steTxt "New Poll Answer Added" %></H3>

<P>
<% steTxt "The new poll question has been added to the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<P ALIGN="center">
	<A HREF="answer_add.asp?pollid=<%= nPollID %>" class="adminlink"><% steTxt "Add Another" %></A> &nbsp;
	<A HREF="answer_list.asp?pollid=<%= nPollID %>" class="adminlink"><% steTxt "Answer List" %></A> &nbsp;
	<A HREF="poll_list.asp" class="adminlink"><% steTxt "Poll List" %></A>
</P>

<% End If %>

<!-- #include file="../../../../footer.asp" -->
