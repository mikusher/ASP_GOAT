<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/class/treelist.asp" -->
<%
'--------------------------------------------------------------------
' message_list.asp
'	Displays a list of the messages for a forum thread.
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
Dim rsThread
Dim rsTopic
Dim nTopicID
Dim nThreadID
Dim sErrorMsg
Dim oList

nTopicID = steNForm("topicid")
nThreadID = steNForm("threadid")

' retrieve the thread or topic list
If nTopicID > 0 Then
	sStat = "SELECT	MessageID, Subject, Messages " &_
			"FROM	tblMessage " &_
			"WHERE	TopicID = " & nTopicID & " " &_
			"AND	Active <> 0 " &_
			"AND	Archive = 0 " &_
			"AND	ParentMessageID = 0 " &_
			"ORDER BY Created DESC"
	Set rsThread = adoOpenRecordset(sStat)
Else
	sStat = "SELECT	TopicID, Title, Threads " &_
			"FROM	tblMessageTopic " &_
			"WHERE	Active <> 0 " &_
			"AND	Archive = 0 " &_
			"ORDER BY OrderNo"
	Set rsTopic = adoOpenRecordset(sStat)
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Message" %>
<!-- #include file="pagetabs_inc.asp" -->

<script language="Javascript" type="text/javascript">
  function pickThread(nThreadID)
  {
	if (nThreadID != '0')
		location.href='message_list.asp?topicid=<%= nTopicID %>&threadid=' + nThreadID;
  }

  function pickTopic(nTopicID)
  {
	if (nTopicID != '0')
		location.href='message_list.asp?topicid=' + nTopicID;
  }

</script>

<h3>Forum Thread List</h3>

<% If nTopicID > 0 Then %>
<p>
<table border=0 cellpadding=2 cellspacing=0>
<tr>
	<td class="forml"><% steTxt "Thread Displayed" %>:</td><Td>&nbsp;&nbsp;</td>
	<td>
	<select name="TopicID" class="form" onChange="pickThread(this.options[this.selectedIndex].value)">
	<option value="0"> -- Choose --
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

<% Else %>

<p>
<table border=0 cellpadding=2 cellspacing=0>
<tr>
	<td class="forml"><% steTxt "Forum Topic" %>:</td><Td>&nbsp;&nbsp;</td>
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
<% End If %>

<%
Set oList = New clsTreeList

sStat = "SELECT	m.MessageID, m.TopicID, m.ParentMessageID, m.Subject, mbr.Username, m.Created " &_
		"FROM	tblMessage m " &_
		"INNER JOIN	tblMember mbr on m.MemberID = mbr.MemberID " &_
		"WHERE	((m.TopicID = " & nTopicID & " " &_
		"		AND	m.ThreadID = " & nThreadID & ") " &_
		"		OR m.MessageID = " & nThreadID & ") " &_
		"AND	m.Active <> 0 " &_
		"AND	m.Archive = 0 " &_
		"ORDER BY m.created"
' Response.Write sStat : Response.End
oList.Query = sStat
oList.AddColumn "Subject", steGetText("Message Subject"), ""
oList.AddColumn "Username", steGetText("Username"), ""
oList.AddColumn "Created", steGetText("Posted"), ""

oList.PrimaryKey = "MessageID"
oList.ParentField = "ParentMessageID"
oList.ActionLink = "<a href=""message_edit.asp?messageid=##messageid##"" class=""actionlink"">" & steGetText("edit") & "</a> . <a href=""message_delete.asp?topicid=##topicid##&messageid=##messageid##"" class=""actionlink"">" & steGetText("delete") & "</a>" 
oList.QueryString = "topicid=" & nTopicID & "&threadid=" & nThreadID
oList.Display
If oList.ErrorMsg <> "" Then
	.Write "<p><b class=""error"">" & oList.ErrorMsg & "</b></p>" & vbCrLf
End If
%>

<P ALIGN="center">
	<A HREF="topic_list.asp" class="adminlink"><% steTxt "Topic List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->