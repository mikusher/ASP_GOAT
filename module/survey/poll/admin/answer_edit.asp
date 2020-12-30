<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../cache.asp" -->
<%
'--------------------------------------------------------------------
' answer_edit.asp
'	Update existing poll answer in the database
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
Dim rsAns
Dim nPollID
Dim nAnswerID

nPollID = steNForm("pollid")
nAnswerID = steNForm("answerid")

' get the question being asked by this poll
sStat = "SELECT	Question " &_
		"FROM	tblPoll " &_
		"WHERE	PollID = " & nPollID
Set rsAns =adoOpenRecordset(sStat)

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("Answer")) = ""	Then
		sErrorMsg = steGetText("Please enter the poll answer below")
	Else
		' update the existing poll answer in the database
		sStat = "UPDATE tblPollAnswer SET " &_
				"	Answer = " & steQForm("Answer") & " " &_
				"WHERE	AnswerID = " & nAnswerID & " " &_
				"AND	PollID = " & nPollID
		Call adoExecute(sStat)

		Call modCapsuleCache(True)
	End If
End If

' get the answer to edit
sStat = "SELECT * FROM tblPollAnswer " &_
		"WHERE	AnswerID = " & steForm("AnswerID") & " " &_
		"AND	PollID = " & nPollID
Set rsAns = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Answers" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Poll Answer" %></H3>

<P>
<% steTxt "Please make your changes to the poll answer using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="answer_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="pollid" VALUE="<%= nPollID %>">
<INPUT TYPE="hidden" NAME="answerid" VALUE="<%= nAnswerID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD VALIGN="top" class="forml"><% steTxt "Answer" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="Answer" VALUE="<%= steRecordEncValue(rsAns, "Answer") %>" SIZE="32" MAXLENGTH="50" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update Poll Answer" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Poll Answer Updated" %></H3>

<P>
<% steTxt "The poll answer was successfully updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="answer_list.asp?pollid=<%= nPollID %>" class="adminlink"><% steTxt "Answer List" %></A> &nbsp;
	<A HREF="poll_list.asp" class="adminlink"><% steTxt "Poll List" %></A>
</P>


<!-- #include file="../../../../footer.asp" -->
