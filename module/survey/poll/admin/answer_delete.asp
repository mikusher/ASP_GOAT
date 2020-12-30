<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../cache.asp" -->
<%
'--------------------------------------------------------------------
' answer_delete.asp
'	Delete an existing poll answer from the database
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
Dim rsAns
Dim nPollID
Dim nAnswerID

nPollID = steNForm("pollid")
nAnswerID = steNForm("answerid")

' get the question being asked by this poll
sStat = "SELECT	Question " &_
		"FROM	tblPoll " &_
		"WHERE	PollID = " & nPollID
Set rsPoll =adoOpenRecordset(sStat)

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If Trim(steNForm("Confirm")) <> 1	Then
		sErrorMsg = steGetText("Please confirm the deletion of this poll answer")
	Else
		' delete the existing poll answer in the database
		sStat = "DELETE FROM tblPollAnswer WHERE AnswerID = " & nAnswerID
		Call adoExecute(sStat)

		Call modCapsuleCache(True)
	End If
End If

sStat = "SELECT * FROM tblPollAnswer WHERE PollID = " & nPollID & " AND AnswerID = " & nAnswerID
Set rsAns = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Answers" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Poll Answer" %></H3>

<P>
<% steTxt "Please confirm that you would like to delete this poll answer by clicking <I>Yes</I> next to <B>Confirm</B> below." %>&nbsp;
<% steTxt "Once the answer has been deleted, it may not be recovered." %>
</P>

<P>
Q: <B><%= rsPoll.Fields("Question").Value %></B>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="answer_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="pollid" VALUE="<%= nPollID %>">
<INPUT TYPE="hidden" NAME="answerid" VALUE="<%= nAnswerID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD class="forml"><% steTxt "Answer" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD class="formd"><%= steRecordEncValue(rsAns, "Answer") %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Votes" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD class="formd"><%= steRecordEncValue(rsAns, "Votes") %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Confirm Delete" %></TD><TD></TD>
	<TD><INPUT TYPE="radio" NAME="confirm" VALUE="1" class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="confirm" VALUE="0" class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 align="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Answer" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Poll Answer Deleted" %></H3>

<P>
<% steTxt "The poll answer was permanently deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="answer_list.asp?pollid=<%= nPollID %>" class="adminlink"><% steTxt "Answer List" %></A> &nbsp;
	<A HREF="poll_list.asp" class="adminlink"><% steTxt "Poll List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->
