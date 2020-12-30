<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/SHA256.asp" -->
<!-- #include file="../../../lib/bbforum_lib.asp" -->
<%
'--------------------------------------------------------------------
' reply.asp
'	Adds a new reply to the database
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
Dim sAction
Dim sSubject
Dim sBody
Dim nMessageID
Dim nTopicID
Dim nThreadID		' thread we are posting a reply to
Dim rsTopic
Dim rsMessage
Dim sTitle			' title for the message topic
Dim sComments		' comments for the message topic
Dim nMessageCount	' total messages for the topic
Dim nThreadCount	' total threads for the topic
Dim rsCount

sAction = Trim(LCase(Request("Action")))
sSubject = steStripForm("subject")
sBody = steStripForm("htmledit")
nMessageID = steNForm("MessageID")
nTopicID = steNForm("TopicID")
nThreadID = steNForm("ThreadID")

' make sure the user is logged in
If Request.Cookies("MemberID") <> "" Then
	nMemberID = Request.Cookies("MemberID")
Else
	nMemberID = 0
End If

' check for a user login here
If sAction = "reply" Or sAction = "preview" Then
	' check for user login here
	If Trim(steForm("username")) <> "" And Trim(steForm("password")) <> "" Then
		' retrieve the user information here
		sStat = "SELECT	MemberID, FirstName, LastName, Username " & _
				"FROM	tblMember " &_
				"WHERE	Username = " & steQForm("username") & " " &_
				"AND	Password = '" & SHA256(steForm("password")) & "'"
		Set rsMember = adoOpenRecordset(sStat)
		If Not rsMember.EOF Then
			' login the user
			Response.Cookies("MemberID") = rsMember.Fields("MemberID").Value
			nMemberID = rsMember.Fields("MemberID").Value
			Response.Cookies("FullName") = rsMember.Fields("FirstName").Value & " " & rsMember.Fields("LastName").Value
			Response.Cookies("Username") = rsMember.Fields("Username").Value
		Else
			' display error message to user here
			sErrorMsg = steGetText("The Username and Password you entered is invalid")
		End If
	End If
End If

If sAction = "reply" Then
	' check the form variables
	If Trim(sSubject) = "" Then
		sErrorMsg = steGetText("Please enter the message subject")
	ElseIf Trim(sBody) = "" Then
		sErrorMsg = steGetText("Please enter the message body")
	Else
		If sErrorMsg = "" Then
			' insert the new reply into the database
			query = "INSERT INTO tblMessage (" &_
					"	ParentMessageID, TopicID, ThreadID, MemberID, Subject, MessageBody, Created" &_
					") VALUES (" &_
					nMessageID & "," &_
					nTopicID & "," &_
					nThreadID & "," &_
					nMemberID & "," &_
					"'" & Replace(sSubject, "'", "''") & "'," &_
					"'" & Replace(sBody, "'", "''") & "'," & adoGetDate &_
					")"
			Call adoExecute(query)

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

			' update the count of messages for the parent message
			If nMessageID <> 0 Then
				sStat = "SELECT Count(*) AS MessageCount FROM tblMessage WHERE ParentMessageID = " & nMessageID
				Set rsCount = adoOpenRecordset(sStat)
				If Not rsCount.EOF Then
					sStat = "UPDATE tblMessage SET Messages = " & rsCount.Fields("MessageCount").Value & " WHERE MessageID = " & nMessageID
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

' retrieve the original message here
If nMessageID > 0 Then
	query = "SELECT	tblMessage.TopicID, tblMessage.Subject, tblMessage.MessageBody, " &_
			"		tblMessage.ThreadID, tblMember.Username " &_
			"FROM	tblMessage " &_
			"INNER JOIN	tblMember ON tblMessage.MemberID = tblMember.MemberID " &_
			"WHERE	tblMessage.MessageID = " & nMessageID & " " &_
			"AND	tblMessage.Active <> 0 " &_
			"AND	tblMessage.Archive = 0"
	Set rsMessage = adoOpenRecordset(query)
	If Not rsMessage.EOF Then
		nThreadID = rsMessage.Fields("ThreadID").Value
		If nTopicID = 0 Then nTopicID = rsMessage.Fields("TopicID").Value
		If nThreadID = 0 Then nThreadID = nMessageID
		If sSubject = "" Then sSubject = "RE: " & rsMessage.Fields("Subject").Value
		If Len(sSubject) > 80 Then sSubject = Left(sSubject, 77) & "..."
	End If
End If

If nTopicID > 0 Then
	' retrieve the topic information here
	query = "SELECT	Title, ShortComments " &_
			"FROM	tblMessageTopic " &_
			"WHERE	TopicID = " & nTopicID
	Set rsTopic = adoOpenRecordset(query)
	If Not rsTopic.EOF Then
		sTitle = rsTopic.Fields("Title").Value
		sComments = rsTopic.Fields("ShortComments").Value
	End If
End If
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/htmledit_lib.asp" -->

<% If sAction <> "reply" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Post Reply" %></H3>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM NAME="formedit" METHOD="post" ACTION="reply.asp" ID="formedit">
<INPUT TYPE="hidden" NAME="TopicID" VALUE="<%= nTopicID %>">
<INPUT TYPE="hidden" NAME="MessageID" VALUE="<%= nMessageID %>">
<INPUT TYPE="hidden" NAME="ThreadID" VALUE="<%= nThreadID %>">
<input type="hidden" name="action" value="">

<% If IsObject(rsMessage) Then %>
	<% If Not rsMessage.EOF Then %>
	<H4><% steTxt "Original Message" %></H4>

	<% forSingleMessage nTopicID, nMessageID %>

	<% End If %>
<% End If %>

<% If sAction = "preview" Then %>

<input type="hidden" name="subject" value="<%= Server.HTMLEncode(sSubject) %>">
<input type="hidden" name="htmledit" value="<%= Server.HTMLEncode(sBody) %>">

<H4><% steTxt "Preview Message" %></H4>

<% forPreviewMessage nMemberID, nTopicID, sSubject, ConvertUBB(sBody) %>

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 width="500">
<TR>
	<TD align="right">
	<input type="button" name="_submit" value=" <% steTxt "Make Changes" %> " onclick="javascript:history.go(-1)" class="form">
	<input type="submit" name="_submit2" value=" <% steTxt "Reply" %> " class="form" onclick="document.formedit.action.value='reply'">
	</TD>
</TR>
</TABLE>

<% Else ' form to enter the reply %>

<H4><% steTxt "Your Reply" %></H4>

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 width="500">
<TR>
	<TD><b class="forml"><% steTxt "Posted By User" %></b><br>
	<% If Request.Cookies("Username") <> "" Then %>
	<font class="formd"><%= Request.Cookies("Username") %></font>
	<% Else %>
	<font class="formd"><% steTxt "Anonymous Coward" %></font><br><br>
	<p>
	<% steTxt "We noticed that you are not logged in." %>
	<% steTxt "Please enter your member login to receive credit for your post." %>
	<% steTxt "Or you may" %>
	<a href="<%= Application("ASPNukeBasePath") %>module/account/register/register.asp"><% steTxt "register	now for a free account" %></A>.
	</p>
	<table border=0 cellpadding=2 cellspacing=0>
	<TR>
		<TD class="forml"><% steTxt "Username" %><BR>
		<INPUT TYPE="text" NAME="username" VALUE="<%= steEncForm("username") %>" SIZE="16" MAXLENGTH="16" class="form" style="width:120px">
		</TD>
		<TD class="forml"><% steTxt "Password" %><BR>
		<INPUT TYPE="password" NAME="password" VALUE="" SIZE="16" MAXLENGTH="16" class="form" style="width:120px">
		</TD>
	</TR>
	</table>
	<% End If %>
	</TD>
</TR><TR>
	<TD class="forml">
	<% steTxt "Message Subject" %><BR>
	<INPUT TYPE="text" NAME="Subject" VALUE="<%= Server.HTMLEncode(sSubject) %>" SIZE="44" MAXLENGTH="80" class="form" style="width:500px">
	</TD>
</TR>
<TR>
	<TD>
	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 WIDTH="100%">
	<TR>
		<TD class="forml"><% steTxt "Message Body" %></TD>
		<TD ALIGN="right" VALIGN="top"><% HTMLCommandButtons %></TD>
	</TR>
	</TABLE><BR>
	<TEXTAREA NAME="htmledit" COLS=58 ROWS=10 WRAP="virtual" class="form" style="width:500px"><%= Server.HTMLEncode(sBody) %></TEXTAREA>
	</TD>
</TR><TR>
	<TD ALIGN="right"><br>
	<INPUT TYPE="submit" NAME="_action" VALUE=" <% steTxt "Preview" %> " class="form" onclick="document.formedit.action.value='preview'">
	<INPUT TYPE="submit" NAME="_action" VALUE=" <% steTxt "Reply" %> " class="form" onclick="document.formedit.action.value='reply'">
	</TD>
</TR>
</TABLE>

<% End If %>

</FORM>

<% Else %>

<H3><% steTxt "Reply Posted" %></H3>

<P>
<% steTxt "Your new message has been posted to the forum." %>&nbsp;
<% steTxt "You can view your message by clicking on the <I>Topic Overview</I> button below." %>&nbsp;
<% steTxt "Thank you for posting your reply!" %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="topic.asp?topicid=<%= nTopicID %>" class="footerlink"><% steTxt "Topic Overview" %></A> &nbsp;
	<A HREF="index.asp" class="footerlink"><% steTxt "Forum Overview" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
