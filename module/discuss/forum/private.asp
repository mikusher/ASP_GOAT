<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/bbforum_lib.asp" -->
<%
'--------------------------------------------------------------------
' private.asp
'	Sends a private message (instant message) to the member who posted
'	the original message.
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

Dim query			' SQL query
Dim sAction			' action to be performed
Dim nMessageID		' message user is responding to
Dim nThreadID		' thread for the message
Dim nTopicID		' top for this thread
Dim sMessage		' message to send to user
Dim sUsername		' username for the member to contact
Dim nToMemberID		' member id for recipient

nMessageID = steNForm("MessageID")
nThreadID = steNForm("ThreadID")
nTopicID = steNForm("TopicID")
sUsername = steForm("Username")
nToMemberID = steNForm("ToMemberID")
sAction = steForm("action")
sMessage = steStripForm("htmledit")

' make sure the user is logged in
If Request.Cookies("MemberID") <> "" Then
	nMemberID = Request.Cookies("MemberID")
Else
	nMemberID = 0
End If

' retrieve the recipient information from the database
query = "SELECT	MemberID, Firstname, Lastname, EmailAddress " &_
		"FROM	tblMember " &_
		"WHERE	Username = '" & Replace(sUsername, "'", "''") & "' " &_
		"AND	Active <> 0 " &_
		"AND	Archive = 0"
Set rsMember = adoOpenRecordset(query)
If Not rsMember.EOF Then
	nToMemberID = rsMember.Fields("MemberID")
End If

If UCase(sAction) = "SEND" Then
	If Trim(sMessage) = "" Then
		sErrorMsg = "Please enter the message for this member"
	Else
		' store the message in the database
		query = "INSERT INTO tblMessagePrivate (" &_
			"ThreadID, MessageID, FromMemberID, ToMemberID, Body, Created" &_
			") VALUES (" &_
			nThreadID & ", " & nMessageID & ", " & nMemberID & ", " &_
			nToMemberID & ", '" & Replace(sMessage, "'", "''") & "'," & adoGetDate &_
			")"
		Call adoExecute(query)
	End If
End If
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/htmledit_lib.asp" -->

<% If Request.Cookies("MemberID") = "" Then %>

<H3><% steTxt "User Login Required" %></H3>

<P>
<% steTxt "Before you can send and receive private messages between yourself and other forum members," %>&nbsp;
<% steTxt "you must first login using the login box shown on the left." %>
</p>

<% ElseIf UCase(sAction) <> "SEND" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Send Instant Message to" %> <%= sUsername %></H3>

<h4><% steTxt "Original Message" %></h4>

<% forSingleMessage nTopicID, nMessageID %>

<h4><% steTxt "Instant Message to Send" %></h4>

<P>
<% steTxt "Enter your message for the member below and an instant message will be sent to the member." %>&nbsp;
<% steTxt "If the member is currently not online, the message will be saved until they log into the site." %>&nbsp;
<% steTxt "If you prefer, you may" %>
<A HREF="email.asp?threadid=<%= nThreadID %>&messageid=<%= nMessageID %>&username=<%= Server.URLEncode(sUsername) %>"><% steTxt "send the member an e-mail" %></A>.
</P>

<% If sErrorMsg <> "" Then %>
<p><b class="error"><%= sErrorMsg %></b></p>
<% End If %>

<FORM NAME="formedit" METHOD="post" ACTION="private.asp">
<INPUT TYPE="hidden" NAME="username" value="<%= Server.HTMLEncode(sUsername) %>">
<input type="hidden" name="ToMemberID" value="<%= nToMemberID %>">

<TABLE BORDER="0" CELLPADDING="2" CELLSPACING="0" width="500px">
<TR>
	<TD class="forml"><% steTxt "From:" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD class="formd"><%= Request.Cookies("Username") %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "To:" %></TD><TD></TD>
	<TD class="formd"><%= sUsername %></TD>
</TR><TR>
	<TD colspan="3">
	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 WIDTH="100%">
	<TR>
		<TD class="forml"><% steTxt "Message" %></TD>
		<TD ALIGN="right" VALIGN="top"><% HTMLCommandButtons %></TD>
	</TR>
	</TABLE><BR>
	<TEXTAREA NAME="htmledit" COLS=58 ROWS=10 WRAP="virtual" class="form" style="width:500px"><%= Server.HTMLEncode(sBody) %></TEXTAREA>
	</TD>
</TR><TR>
	<TD COLSPAN="3" ALIGN="right"><br>
	<INPUT TYPE="hidden" NAME="action" VALUE="send">
	<INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Send Message" %> " CLASS="form">
	</TD>
</TR>
</TABLE>

</FORM>

<% Else %>

<H3><% steTxt "Message Sent to" %> <%= sUsername %></H3>

<P>
<% steTxt "Your instant message has been sent to the specified member." %>&nbsp;
<% steTxt "The recipient will see a flashing message in their <I>instant messanger</I> window that appears at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="thread.asp?threadid=<%= nThreadID %>" CLASS="footerlink"><% steTxt "Back to Thread" %></A> &nbsp;
	<A HREF="topic.asp?topicid=<%= nTopicID %>" CLASS="footerlink"><% steTxt "Topic Overview" %></A> &nbsp;
	<A HREF="index.asp" CLASS="footerlink"><% steTxt "Forum Overview" %></A>
</P>

<!-- #include file="../../../footer.asp" -->