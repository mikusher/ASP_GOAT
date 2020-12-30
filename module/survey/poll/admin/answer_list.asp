<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' answer_list.asp
'	Displays a list of the answers for a specific poll question.
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
Dim rsPoll		' poll to show answers for
Dim rsAns		' answers for the poll question
Dim nPollID
Dim I

nPollID = steNForm("pollid")

If nPollID > 0 Then
	' get the question being asked by this poll
	sStat = "SELECT	Question " &_
			"FROM	tblPoll " &_
			"WHERE	PollID = " & nPollID
	Set rsPoll =adoOpenRecordset(sStat)
	
	sStat = "SELECT	AnswerID, Answer, Votes, Created, Modified " &_
			"FROM	tblPollAnswer " &_
			"WHERE	PollID = " & nPollID & " " &_
			"ORDER BY OrderNo"
	Set rsAns = adoOpenRecordset(sStat)
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<script language="Javascript" type="text/javascript">
function pickQuestion(nPollID) {
	if (nPollID != '')
		window.location.href='answer_list.asp?pollid=' + nPollID;
}
</script>

<% sCurrentTab = "Answers" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3>Answers List</H3>

<p>
<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0>
<TR>
	<TD class="forml"><% steTxt "Poll Displayed" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD class="formd">
	<select name="PollID" onChange="pickQuestion(this.options[this.selectedIndex].value)" style="{width:300px}">
	<option value=""> -- <% steTxt "Choose" %> --
	<% ' retrieve the list of poll questions to work with
	Dim rsList
	sStat = "SELECT	PollID, Question " &_
			"FROM	tblPoll " &_
			"ORDER BY Created DESC"
	Set rsList =adoOpenRecordset(sStat)
	Do Until rsList.EOF %>
	<option value="<%= rsList.Fields("PollID").Value %>"<% If nPollID = rsList.Fields("PollID").Value Then Response.Write " SELECTED" %>> <%= rsList.Fields("Question").Value %>
	<%	rsList.MoveNext
	Loop
	rsList.Close
	Set rsList = Nothing %>
	</select>
	</TD>
</TR>
</TABLE>
</p>

<% If nPollID > 0 Then %>

<P>Q:
<B><%= rsPoll.Fields("Question").Value %></B></P>

<% If Not rsAns.EOF Then %>

<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 class="list">
<TR>
	<TD class="listhead"><% steTxt "Answer" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Votes" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Created" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Modified" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<%	I = 0
	Do Until rsAns.EOF %>
<TR CLASS="list<%= I mod 2 %>">
	<TD><%= rsAns.Fields("Answer").Value %></TD>
	<TD><%= rsAns.Fields("Votes").Value %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsAns.Fields("Created").Value, vbShortDate) %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsAns.Fields("Modified").Value, vbShortDate) %></TD>
	<TD>
		<A HREF="answer_edit.asp?pollid=<%= nPollID %>&answerid=<%= rsAns.Fields("AnswerID").Value %>" class="actionlink"><% steTxt "edit" %></A> . 
		<A HREF="answer_delete.asp?pollid=<%= nPollID %>&answerid=<%= rsAns.Fields("AnswerID").Value %>" class="actionlink"><% steTxt "delete" %></A>
	</TD>
</TR>
<%	rsAns.MoveNext
	I = I + 1
   Loop %>
</TABLE>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No poll answers exist in the database" %></B></P>

<% End If %>

<P ALIGN="center">
	<A HREF="answer_add.asp?pollid=<%= nPollID %>" class="adminlink"><% steTxt "Add New Poll Answer" %></A>
</P>

<% Else %>

<p><b class="error">
<% steTxt "To get started, please select a poll question above." %>
</b></p>

<% End If %>

<!-- #include file="../../../../footer.asp" -->