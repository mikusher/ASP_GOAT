﻿<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/class/form/listinput.asp" -->
<%
'--------------------------------------------------------------------
' message_add.asp
'	Add a new forum message in the database
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
Dim rsMessage
Dim nTopicID
Dim nThreadID
Dim nMessageID
Dim nMemberID
Dim sBody
Dim sErrorMsg
Dim sStatusMsg

nTopicID = steNForm("topicid")
nThreadID = steNForm("threadid")
nMessageID = steNForm("messageid")
sBody = steStripForm("htmledit")

If steForm("action") = "add" Then
	' make sure the required fields are present
	If nTopicID = 0 Then
		sErrorMsg = steGetText("Please select a topic to work with")
	ElseIf Trim(steForm("Subject")) = ""	Then
		sErrorMsg = steGetText("Please enter the subject for this forum message")
	ElseIf Trim(sBody) = "" Then
		sErrorMsg = steGetText("Please enter the body for this forum message")
	Else
		' check to make sure the username is valid
		sStat = "SELECT	MemberID, Active, Archive FROM tblMember WHERE Username = " & steQForm("Username")
		Set rsMember = adoOpenRecordset(sStat)
		If Not rsMember.EOF Then
			If Not steRecordBoolValue(rsMember, "Active") Then
				sStatusMsg = sStatusMsg & steGetText("*** WARNING *** - User has been set inactive") & "<Br>"
			End If
			If steRecordBoolValue(rsMember, "Archive") Then
				sStatusMsg = sStatusMsg & steGetText("*** WARNING *** - User has been archived") & "<Br>"
			End If
			nMemberID = rsMember.Fields("MemberID").Value
		Else
			sErrorMsg = steGetText("Unrecognized member username:") & " """ & steEncForm("Username") & """<br>"
		End If
		rsMember.Close
		Set rsMember = Nothing

		If sErrorMsg = "" Then
			Dim rsCount
			Dim nMessageCount
			Dim nThreadCount

			' update the forum message in the database
			sStat = "INSERT INTO tblMessage (" &_
					"	ParentMessageID, TopicID, ThreadID, MemberID, Subject, MessageBody, Created" &_
					") VALUES (" &_
					steNForm("ParentMessageID") & ", " & nTopicID & ", " & nThreadID & ", " & nMemberID & "," &_
					steQForm("Subject") & ", '" & Replace(sBody, "'", "''") & "'," & adoGetDate &_
					")"
			Call adoExecute(sStat)

			' count the total number of message in the topic
			sStat = "SELECT Count(*) AS MessageCount FROM tblMessage WHERE TopicID = " & nTopicID &_
					"		AND	Active <> 0 AND Archive = 0"
			Set rsCount = adoOpenRecordset(sStat)
			If Not rsCount.EOF Then
				nMessageCount = rsCount.Fields("MessageCount").Value
			End If
			rsCount.Close

			' count the total number of threads in the topic
			sStat = "SELECT Count(*) AS ThreadCount FROM tblMessage WHERE TopicID = " & nTopicID &_
					"		AND	ParentMessageID = 0 AND Active <> 0 AND Archive = 0"
			Set rsCount = adoOpenRecordset(sStat)
			If Not rsCount.EOF Then
				nThreadCount = rsCount.Fields("ThreadCount").Value
			End If
			rsCount.Close
			Set rsCount = Nothing
			
			' update the counts of messages for all parent messages
			If steNForm("ParentMessageID") <> 0 Then
				sStat = "SELECT Count(*) AS MessageCount FROM tblMessage WHERE ParentMessageID = " & steNForm("ParentMessageID")
				Set rsCount = adoOpenRecordset(sStat)
				If Not rsCount.EOF Then
					sStat = "UPDATE tblMessage SET Messages = " & rsCount.Fields("MessageCount").Value & " WHERE MessageID = " & steNForm("ParentMessageID")
					Call adoExecute(sStat)
				End If
				rsCount.Close
				Set rsCount = Nothing
			End If

			' update the counts of threads / messages for this topic
			sStat = "UPDATE	tblMessageTopic " &_
					"SET	LastPost = " & adoGetDate & ", " &_
					"		Messages = " & nMessageCount & ", " &_
					"		Threads = " & nThreadCount & ", " &_
					"		Modified = " & adoGetDate & " " &_
					"WHERE	TopicID = " & nTopicID
			Call adoExecute(sStat)
		End If
	End If
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->
<!-- #include file="../../../../lib/htmledit_lib.asp" -->

<script language="Javascript" type="text/javascript">
  function pickTopic(nTopicID)
  {
	if (nTopicID != '0')
		location.href='message_add.asp?topicid=' + nTopicID;
  }

  function pickThread(nThreadID)
  {
	if (nThreadID != '0')
		location.href='message_add.asp?topicid=<%= nTopicID %>&threadid=' + nThreadID;
  }
</script>

<% sCurrentTab = "Message" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Post Forum Message" %></H3>

<P>
<% steTxt "Please enter the properties for the forum message using the form below." %>
</P>

<FORM METHOD="post" ACTION="message_add.asp" NAME="formedit">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<p>
<% ' build the topic droplist to choose a topic
Dim rsTopic
sStat = "SELECT	TopicID, Title, Threads " &_
		"FROM	tblMessageTopic " &_
		"WHERE	Active <> 0 " &_
		"AND	Archive = 0 " &_
		"ORDER BY OrderNo"
Set rsTopic = adoOpenRecordset(sStat)
%>
<table border=0 cellpadding=2 cellspacing=0>
<tr>
	<td class="forml"><% steTxt "Forum Topic:" %></td><Td>&nbsp;&nbsp;</td>
	<td>
	<select name="TopicID" class="form" onChange="pickTopic(this.options[this.selectedIndex].value)">
	<option value="0"> -- Choose --
	<% Do Until rsTopic.EOF %>
	<option value="<%= rsTopic.Fields("TopicID").Value %>"<% If nTopicID = rsTopic.Fields("TopicID").Value Then Response.Write " SELECTED" %>> <%= rsTopic.Fields("Title").Value & " (" & rsTopic.Fields("Threads").Value & ")" %>
	<%	rsTopic.MoveNext
	   Loop
	   rsTopic.Close
	   Set rsTopic = Nothing %>
	</select>
	</td>
</tr>
</table>
</p>

<p>
<% ' build the thread droplist to choose a topic
If nTopicID > 0 Then
	Dim rsThread
	sStat = "SELECT	tm.MessageID, tm.Subject, tm.Messages, tm.LastPost, tm.Modified, " &_
			"		m.Username " &_
			"FROM	tblMessage tm " &_
			"INNER JOIN	tblMember m on m.MemberID = tm.MemberID " &_
			"WHERE	tm.TopicID = " & nTopicID & " " &_
			"AND	tm.ParentMessageID = 0 " &_
			"ORDER BY tm.Created DESC"
	Set rsThread = adoOpenRecordset(sStat)
%>
<table border=0 cellpadding=2 cellspacing=0>
<tr>
	<td class="forml"><% steTxt "Forum Thread:" %></td><Td>&nbsp;&nbsp;</td>
	<td>
	<select name="ThreadID" class="form" onChange="pickThread(this.options[this.selectedIndex].value)">
	<option value="0"> -- <% steTxt "Choose" %> --
	<% Do Until rsThread.EOF %>
	<option value="<%= rsThread.Fields("MessageID").Value %>"<% If nThreadID = rsThread.Fields("MessageID").Value Then Response.Write " SELECTED" %>> <%= rsThread.Fields("Subject").Value & " (" & rsThread.Fields("Messages").Value & ")" %>
	<%	rsThread.MoveNext
	   Loop
	   rsThread.Close
	   Set rsThread = Nothing %>
	</select>
	</td>
</tr>
</table>
</p>
<% End If %>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>
<% If sStatusMsg <> "" Then %>
<P><B CLASS="error"><%= sStatusMsg %></B></P>
<% End If %>

<% If nTopicID > 0 Then %>

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD class="forml" nowrap><% steTxt "Parent Message" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD>
	<% ' build the drop-list for the parent message
	If nThreadID <> 0 Then
		Dim oList
		Set oList = New clsListInput
		oList.TreeListInput "ParentMessageID", "tblMessage", "MessageID", _
			"ParentMessageID", "TopicID = " & nTopicID & " AND (ThreadID = " & nThreadID & " OR MessageID = " & nThreadID & ")", _
			"Created", "MessageID", "Subject", steNForm("ParentMessageID"), _
			"topicid=" & nTopicID & "&threadid=" & nThreadID, False
	Else
		Response.Write "<I>* " & steGetText("New Thread") & " *</I>"
	End If %>
	</TD>
</TR><TR>
	<TD class="forml" nowrap><% steTxt "Username" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Username" VALUE="<%= steEncForm("Username") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD class="forml" nowrap><% steTxt "Subject" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Subject" VALUE="<%= steEncForm("Subject") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD COLSPAN="3">
	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 WIDTH="100%">
	<TR>
		<TD class="forml"><% steTxt "Message Body" %></TD>
		<TD ALIGN="right" VALIGN="top"><% HTMLCommandButtons %></TD>
	</TR>
	</TABLE><BR>
	<TEXTAREA NAME="htmledit" COLS=58 ROWS=10 WRAP="virtual" class="form" style="width:500px"><%= Server.HTMLEncode(sBody) %></TEXTAREA>
	</TD>
</TR><TR>
	<TD COLSPAN=3><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Post Forum Message" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% End If %>

<% Else %>

<H3><% steTxt "Forum Message Added" %></H3>

<P>
<% steTxt "The forum message was successfully created in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align="center">
	<a href="message_list.asp?topicid=<%= nTopicID %>&threadid=<%= nThreadID %>" class="adminlink"><% steTxt "Forum Thread List" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->
