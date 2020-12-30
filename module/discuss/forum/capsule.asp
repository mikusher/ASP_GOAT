<!-- #include file="../../../lib/ado_lib.asp" -->
<!-- #include file="../../../lib/lang_lib.asp" -->
<%
'--------------------------------------------------------------------
' capsule.asp
'	Create the links capsule which will appear on all pages of
'	the site.
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
%>

<%= Replace(Application("ModCapHeader"), "<CAPSULETITLE/>", steGetText("Message Forums")) %>
<%= Application("ModCapLeft") %>
<DIV class="forumcapsule">

<% Call locCacheForum
	If Application("FORUMCAPSULE") <> "" Then %>

<%= Application("FORUMCAPSULE") %>

<% Else %>

<P><B CLASS="Error"><% steTxt "No forum topics are defined yet" %></B></P>

<% End If %>
</DIV>
<%= Application("ModCapRight") %>
<%= Application("ModCapFooter") %>
<%
'----------------------------------------------------------------------------
' locCacheForum
'	Cache the forum capsule content to eliminate excessive database hits

Sub locCacheForum
	Dim query, sHTML, rsTopic, bFirst

	' check to see if we need to refresh (every 15 mins)
	If IsDate(Application("FORUMCAPSULEREFRESH")) Then
		If DateDiff("n", Application("FORUMCAPSULEREFRESH"), Now()) < 15 Then Exit Sub
	End If

	' retrieve the list of topics from the database
	query = "SELECT	TopicID, Title, Threads, Messages " &_
			"FROM	tblMessageTopic " &_
			"WHERE	Active <> 0 " &_
			"AND	Archive = 0 " &_
			"ORDER BY OrderNo"
	Set rsTopic = adoOpenRecordset(query)
	bFirst = False
	Do Until rsTopic.EOF
		If bFirst Then sHTML = sHTML & "<hr noshade class=""forumcapsulesep"">" & vbCrLf
		sHTML = sHTML & "<a href=""" & Application("ASPNukeBasePath") & "module/discuss/forum/topic.asp?topicid=" & rsTopic.Fields("TopicID").Value & """ class=""forumtopic"">" &_
			rsTopic.Fields("Title").Value & "</a>" &_
			"&nbsp;&nbsp;<font class=""tinytext"">(" & rsTopic.Fields("Threads").Value & "/" & rsTopic.Fields("Messages").Value & ")</font>" & vbCrLf
		rsTopic.MoveNext
		bFirst = True
	Loop
	rsTopic.Close
	rsTopic = Empty
	Application("FORUMCAPSULEREFRESH") = Now()
	Application("FORUMCAPSULE") = sHTML
End Sub
%>