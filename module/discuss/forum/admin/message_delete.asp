<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/class/form/listinput.asp" -->
<!-- #include file="../../../../lib/bbforum_lib.asp" -->
<%
'--------------------------------------------------------------------
' message_delete.asp
'	Delete an existing forum message in the database
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
Dim rsMess
Dim rsCount
Dim nMessageCount
Dim nThreadCount
Dim rsParent
Dim nParentID
Dim sErrorMsg
Dim sStatusMsg

nTopicID = steNForm("topicid")
nThreadID = steNForm("threadid")
nMessageID = steNForm("messageid")

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If steNForm("Confirm") = 0 Then
		sErrorMsg = steGetText("You must confirm the deletion of this forum message")
	Else
		Set rsMess = adoOpenRecordset("SELECT * FROM tblMessage WHERE TopicID = " & nTopicID & " AND MessageID = " & nMessageID)
		If Not rsMess.EOF Then
			rsMess.Close

			' retrieve the parent message ID for the one being deleted
			sStat = "SELECT ParentMessageID FROM tblMessage WHERE MessageID = " & nMessageID
			Set rsParent = adoOpenRecordset(sStat)
			If Not rsParent.EOF Then
				nParentID = rsParent.Fields("ParentMessageID").Value
			End If
			rsParent.Close
			Set rsParent = Nothing

			' delete the forum message in the database
			sStat = "DELETE FROM tblMessage WHERE TopicID = " & nTopicID & " AND MessageID = " & nMessageID
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

			' update the message count and last post for the thread
			If nParentID <> 0 Then
				sStat = "SELECT Count(*) AS MessageCount FROM tblMessage WHERE ParentMessageID = " & nParentID
				Set rsCount = adoOpenRecordset(sStat)
				If Not rsCount.EOF Then
					sStat = "UPDATE tblMessage SET Messages = " & rsCount.Fields("MessageCount").Value & " WHERE MessageID = " & nParentID
					Call adoExecute(sStat)
				End If
				rsCount.Close
				Set rsCount = Nothing
			End If

			' update the counts of threads / messages for this topic
			sStat = "UPDATE	tblMessageTopic " &_
					"SET	Messages = " & nMessageCount & ", " &_
					"		Threads = " & nThreadCount & ", " &_
					"		Modified = " & adoGetDate & " " &_
					"WHERE	TopicID = " & nTopicID
			Call adoExecute(sStat)
		Else
			sErrorMsg = steGetText("Unable to find message") & " (TopicID = " & nTopicID & " and MessageID = " & nMessageID & ")"
		End If
		Set rsMess = Nothing
	End If
End If

If nTopicID = 0 Then
	sStat = "SELECT TopicID FROM tblMessage WHERE MessageID = " & nMessageID
	Set rsMess = adoOpenRecordset(sStat)
	If Not rsMess.EOF Then nTopicID = rsMess.Fields("TopicID").Value
	rsMess.Close
	Set rsMess = Nothing
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Message" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Forum Message" %></H3>

<P>
<% steTxt "Please confirm that you want to delete the forum message shown below." %>
</P>

<FORM METHOD="post" ACTION="message_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="memberid" VALUE="<%= nMemberID %>">
<INPUT TYPE="hidden" NAME="topicid" VALUE="<%= nTopicID %>">
<INPUT TYPE="hidden" NAME="threadid" VALUE="<%= nThreadID %>">
<INPUT TYPE="hidden" NAME="messageid" VALUE="<%= nMessageID %>">

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<% forSingleMessage nTopicID, nMessageID %>

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml"><% steTxt "Confirm Delete" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="confirm" VALUE="1" CLASS="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="confirm" VALUE="0" CLASS="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" Delete Forum Message " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Forum Message Deleted" %></H3>

<P>
<% steTxt "The forum message was successfully deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align="center">
	<a href="message_list.asp?topicid=<%= nTopicID %>&threadid=<%= nThreadID %>" class="adminlink"><% steTxt "Forum Thread List" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->
<%
' update the message counts for the parent messages
Sub locUpdateMessageCount(nMessageID)
	Dim sStat, rsCount

	sStat = "SELECT Count(*) AS MessageCount FROM tblMessage WHERE ParentMessageID = " & nMessageID
	Set rsCount = adoOpenRecordset(sStat)
	If Not rsCount.EOF Then
		sStat = "UPDATE tblMessage SET Messages = " & rsCount.Fields("MessageCount").Value & " WHERE MessageID = " & nMessageID
		Call adoExecute(sStat)
	End If
	sStat = "SELECT	ParentMessageID FROM tblMessage WHERE MessageID = " & nMessageID
	Set rsMessage = adoOpenRecordset(sStat)
	If Not rsMessage.EOF Then
		If rsMessage.Fields("ParentMessageID").Value <> 0 Then
			Call locUpdateMessageCount(rsMessage.Fields("ParentMessageID").Value)
		End If
	End If
End Sub
%>