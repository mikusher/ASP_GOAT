<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/class/adminlist.asp" -->

<%
'--------------------------------------------------------------------
' topic.asp
'	Displays the message list for a particular topic
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
Dim oList
Dim rsTopic		' topic information to display
Dim nTopicID
Dim sTitle		' title for the topic
Dim sComments	' comments for the topic
Dim nThreads	' count of threads in the topic

nTopicID = steNForm("TopicID")

' retrieve the thread information
sStat = "SELECT	Title, ShortComments, Threads " &_
		"FROM	tblMessageTopic " &_
		"WHERE	TopicID = " & nTopicID & " " &_
		"AND	Active <> 0 " &_
		"AND	Archive = 0"
Set rsTopic = adoOpenRecordset(sStat)
If Not rsTopic.EOF Then
	sTitle = rsTopic.Fields("Title").Value
	sComments = rsTopic.Fields("ShortComments").Value
	nThreads = rsTopic.Fields("Threads").Value
End If
%>
<!-- #include file="../../../header.asp" -->

<H3><%= sTitle %></H3>

<P><%= sComments %></P>


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
oList.AddColumn "<a href=""thread.asp?topicid=" & nTopicID & "&threadid=##messageid##"">##Subject##</A>", steGetText("Subject"), ""
oList.AddColumn "<a href=""profile.asp?topicid=" & nTopicID & "&thradid=##messageid##&username=##Username##"">##Username##</a>", steGetText("Author"), ""
oList.AddColumn "Messages", steGetText("Messages"), ""
oList.AddColumn "LastPost", steGetText("Last Post"), ""
oList.QueryString = "topicid=" & nTopicID
Call oList.Display
%>

<P ALIGN="center">
	<A HREF="index.asp" class="footerlink"><% steTxt "Forum Overview" %></A> &nbsp;
	<A HREF="post.asp?topicid=<%= nTopicID %>" class="footerlink"><% steTxt "Post Message" %></A>
</P>

<!-- #include file="../../../footer.asp" -->