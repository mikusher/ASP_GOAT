<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/bbforum_lib.asp" -->
<%
'--------------------------------------------------------------------
' thread.asp
'	Displays the message list for a particular thread
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

Dim query
Dim rsTopic		' topic information to display
Dim nTopicID	' topic that we want to display
Dim nThreadID	' thread that we want to display
Dim sTitle		' title for the topic
Dim sComments	' comments for the topic
Dim nThreads	' count of threads in the topic

nTopicID = steNForm("TopicID")
nThreadID = steNForm("ThreadID")

' retrieve the thread information
query = "SELECT	mt.TopicID, mt.Title, mt.ShortComments, mt.Threads " &_
		"FROM	tblMessageTopic mt " &_
		"INNER JOIN tblMessage m ON m.MessageID = " & nThreadID & " " &_
		"WHERE	mt.TopicID = " & nTopicID & " " &_
		"AND	mt.Active <> 0 " &_
		"AND	mt.Archive = 0"
Set rsTopic = adoOpenRecordset(query)
If Not rsTopic.EOF Then
	nTopicID = CInt(rsTopic.Fields("TopicID").Value)
	sTitle = rsTopic.Fields("Title").Value
	sComments = rsTopic.Fields("ShortComments").Value
	nThreads = rsTopic.Fields("Threads").Value
End If
%>
<!-- #include file="../../../header.asp" -->

<% If sTitle <> "" Then %>

<H3><%= sTitle %></H3>

<P><%= sComments %></P>

<% forThread nTopicID, nThreadID %>

<P ALIGN="center">
	<A HREF="post.asp?topicid=<%= nTopicID %>" class="footerlink"><% steTxt "Post Message" %></A> &nbsp;
	<A HREF="topic.asp?topicid=<%= nTopicID %>" class="footerlink"><% steTxt "Topic Overview" %></A> &nbsp;
	<A HREF="index.asp" class="footerlink"><% steTxt "Forum Overview" %></A>
</P>

<% Else %>

<H3><% steTxt "Forum Topic Invalid" %></H3>

<P><B class="error"><% steTxt "Sorry, but the forum topic you entered is invalid" %> (ID = <%= nTopicID %>)</b></P>

<P ALIGN="center">
	<A HREF="index.asp" class="footerlink"><% steTxt "Forum Overview" %></A>
</P>

<% End If %>

<!-- #include file="../../../footer.asp" -->