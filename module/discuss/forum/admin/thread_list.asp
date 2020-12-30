<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/class/adminlist.asp" -->
<%
'--------------------------------------------------------------------
' thread_list.asp
'	Displays a list of the messages for a forum topics.
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
Dim rsTopic
Dim nTopicID

nTopicID = steNForm("topicid")

sStat = "SELECT	TopicID, Title, Threads " &_
		"FROM	tblMessageTopic " &_
		"WHERE	Active <> 0 " &_
		"AND	Archive = 0 " &_
		"ORDER BY OrderNo"
Set rsTopic = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Threads" %>
<!-- #include file="pagetabs_inc.asp" -->

<script language="Javascript" type="text/javascript">
  function pickTopic(nTopicID)
  {
	if (nTopicID != '0')
		location.href='thread_list.asp?topicid=' + nTopicID;
  }
</script>

<h3><% steTxt "Forum Thread List" %></h3>

<p>
<table border=0 cellpadding=2 cellspacing=0>
<tr>
	<td class="forml"><% steTxt "Forum Topic" %>:</td><Td>&nbsp;&nbsp;</td>
	<td>
	<select name="TopicID" class="form" onChange="pickTopic(this.options[this.selectedIndex].value)">
	<option value="0"> -- <% steTxt "Choose" %> --
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

<%
sStat = "SELECT	me.MessageID, me.MemberID, me.Subject, me.Messages, me.LastPost, me.Modified, " &_
		"		m.Username " &_
		"FROM	tblMessage me " &_
		"INNER JOIN	tblMember m on m.MemberID = me.MemberID " &_
		"WHERE	me.TopicID = " & nTopicID & " " &_
		"AND	me.ParentMessageID = 0 " &_
		"ORDER BY me.Created DESC"
Set oList = New clsAdminList
oList.Query = sStat
oList.AddColumn "<a href=""message_list.asp?topicid=" & nTopicID & "&threadid=##messageid##"">##Subject##</A>", steGetText("Subject"), ""
oList.AddColumn "Username", steGetText("Username"), ""
oList.AddColumn "Messages", steGetText("Msgs"), ""
oList.AddColumn "LastPost", steGetText("Modified"), ""
oList.ActionLink = "<a href=""message_edit.asp?topicid=" & nTopicID &_
	"&messageid=##messageid##"" class=""actionlink"">" & steGetText("edit") & "</a> . <a href=""message_delete.asp?topicid=" &_
	nTopicID & "&messageid=##messageid##&memberid=##memberid##"" class=""actionlink"">" & steGetText("delete") & "</a>"
oList.QueryString = "topicid=" & nTopicID
Call oList.Display
%>

<P ALIGN="center">
	<a HREF="topic_list.asp" class="adminlink"><% steTxt "Topic List" %></a> &nbsp;
	<a href="message_add.asp?topicid=<%= nTopicID %>" class="adminlink"><% steTxt "Start New Thread" %></a>
</P>

<!-- #include file="../../../../footer.asp" -->