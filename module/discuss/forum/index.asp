<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' index.asp
'	Main index page for the discussion forums.
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
Dim rsTopic

' retrieve the list of topics from the database
query = "SELECT	TopicID, Title, ShortComments, Threads, Messages, LastPost " &_
		"FROM	tblMessageTopic " &_
		"WHERE	Active <> 0 " &_
		"AND	Archive = 0 " &_
		"ORDER BY OrderNo"
Set rsTopic = adoOpenRecordset(query)
%>
<!-- #include file="../../../header.asp" -->

<H3><% steTxt "Message Forums" %></H3>

<P>
<% steTxt "Please click on a topic below to view or post a message in our message forums." %>&nbsp;
<% steTxt "You will be required to sign up for our" %> <A HREF="<%= Application("ASPNukeBasePath") %>module/account/register/register.asp"><% steTxt "free registration" %></A>
<% steTxt "if you wish to post new messages or replies." %>
</P>

<% If rsTopic.EOF Then %>
<P><B class="error"><% steTxt "Sorry, The discussion forum is temporarily unavailable" %></B></P>
<% End If %>

<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 CLASS="list">
<TR>
	<TD class="listhead"><% steTxt "Topic" %></TD>
	<TD class="listhead"><% steTxt "Comments" %></TD>
	<TD class="listhead"><% steTxt "Threads" %></TD>
	<TD class="listhead" ALIGN="right" nowrap><% steTxt "Last Post" %></TD>
</TR>
<%
If Not rsTopic.EOF Then
I = 0
   Do Until rsTopic.EOF %>
<TR class="list<%= I mod 2 %>">
	<TD VALIGN="top"><NOBR><B><A HREF="topic.asp?topicid=<%= rsTopic.Fields("TopicID").Value %>" class="forumtopic"><%= rsTopic.Fields("Title").Value %></A></B></NOBR></TD>
	<TD VALIGN="top"><%= rsTopic.Fields("ShortComments").Value %></TD>
	<TD ALIGN="center" VALIGN="top"><B><%= rsTopic.Fields("Threads").Value %>/<%= rsTopic.Fields("Messages").Value %></B></TD>
	<TD ALIGN="right" VALIGN="top"><% If Not IsNull(rsTopic.Fields("LastPost").Value) Then %><%= adoFormatDateTime(rsTopic.Fields("LastPost").Value, vbShortDate) %><% Else %><i><% steTxt "n/a" %></i><% End If %></TD>
</TR>
<%	rsTopic.MoveNext
	I = I + 1
  Loop
Else %>
<TR>
	<TD COLSPAN="4" ALIGN="center"><b style="color:#c0c0c0">&nbsp;<% steTxt "No topics are available to display here" %>...&nbsp;</b></TD>
</TR>
<% End If %>
</TABLE>

<DIV>
<FONT CLASS="tinytext">* 13 / 56 <% steTxt "means 13 threads containing 56 individual messages" %></FONT>
</DIV>
<!-- #include file="../../../footer.asp" -->