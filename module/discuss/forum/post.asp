<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/sha256.asp" -->
<!-- #include file="../../../lib/bbforum_lib.asp" -->
<%
'--------------------------------------------------------------------
' post.asp
'	Posts a new message to the message forums
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
Dim nTopicID
Dim sBody		' body of the message post
Dim sSubject	' subject for the message post
Dim sAction
Dim rsTopic		' retrieve the topic information
Dim sTitle		' title for the specified topic
Dim sComments	' short comments for the specified topic
Dim nMemberID	' member posting the message
Dim sErrorMsg	' error message to display to the user

sAction = LCase(Trim(steForm("Action")))
nTopicID = steNForm("TopicID")
sSubject = steStripForm("subject")
sBody = steStripForm("htmledit")

' make sure the user is logged in
If Request.Cookies("MemberID") <> "" Then
	nMemberID = Request.Cookies("MemberID")
Else
	nMemberID = 0
End If

' check for a user login here
If sAction = "post" Or sAction = "preview" Then
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

If sAction = "post" Then
	' check for the required fields here
	If Trim(sSubject) = "" Then
		sErrorMsg = steGetText("Please enter the subject of your message")
	End If
	If Trim(sBody) = "" Then
		If sErrorMsg <> "" Then sErrorMsg = sErrorMsg & "<BR>"
		sErrorMsg = sErrorMsg & steGetText("Please enter the message body for your post")
	End If
	If nTopicID = 0 Then
		If sErrorMsg <> "" Then sErrorMsg = sErrorMsg & "<BR>"
		sErrorMsg = sErrorMsg & steGetText("Invalid topic specified") & " (ID = " & nTopicID & ")"
	End If
	If sErrorMsg = "" Then
		' check the form variables
		If Trim(steForm("username")) <> "" And Trim(steForm("password")) <> "" Then
			' retrieve the user information here
			sStat = "SELECT	MemberID, Username " & _
					"FROM	tblMember " &_
					"WHERE	Username = " & steQForm("username") & " " &_
					"AND	Password = '" & SHA256(steForm("password")) & "'"
			Set rsMember = adoOpenRecordset(sStat)
			If Not rsMember.EOF Then
				' login the user
				Response.Cookies("MemberID") = rsMember.Fields("MemberID").Value
				nMemberID = rsMember.Fields("MemberID").Value
				Response.Cookies("Username") = rsMember.Fields("Username").Value
			Else
				' display error message to user here
				sErrorMsg = steGetText("The Username and Password you entered is invalid")
			End If
		End If
		If sErrorMsg = "" Then
			' build the query to insert the message post
			query = "INSERT INTO tblMessage (" &_
					"	TopicID, ThreadID, MemberID, Subject, MessageBody, Messages, Created" &_
					") VALUES (" &_
					nTopicID & "," &_
					"0," &_
					nMemberID & "," &_
					"'" & Replace(sSubject, "'", "''") & "'," &_
					"'" & Replace(sBody, "'", "''") & "'," &_
					"1," & adoGetDate &_
					")"
			Call adoExecute(query)

			' update the counts of threads / messages for this topic
			Dim rsCount, nMessageCount, nThreadCount
			sStat = "SELECT Count(*) MessageCount FROM tblMessage " &_
					"WHERE TopicID = " & nTopicID & " " &_
					"AND	Active <> 0 AND Archive = 0"
			Set rsCount = adoOpenRecordset(sStat)
			If Not rsCount.EOF Then nMessageCount = rsCount.Fields("MessageCount").Value Else nMessageCount = 0

			sStat = "SELECT Count(*) ThreadCount FROM tblMessage " &_
					"WHERE TopicID = " & nTopicID & " " &_
					"AND ParentMessageID = 0 AND Active <> 0 AND Archive = 0"
			Set rsCount = adoOpenRecordset(sStat)
			If Not rsCount.EOF Then nThreadCount = rsCount.Fields("ThreadCount").Value Else nThreadCount = 0
			rsCount.Close

			' update the topic statistics
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

' retrieve information about the topic
query = "SELECT	Title, ShortComments " &_
		"FROM	tblMessageTopic " &_
		"WHERE	TopicID = " & nTopicID & " " &_
		"AND	Active <> 0 " &_
		"AND	Archive = 0"
Set rsTopic = adoOpenRecordset(query)
If Not rsTopic.EOF Then
	sTitle = rsTopic.Fields("Title").Value
	sComments = rsTopic.Fields("ShortComments").Value
End If
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/htmledit_lib.asp" -->

<% If sAction <> "post" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Post to" %>&nbsp;<%= sTitle %></H3>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<P>
<% steTxt "Enter your new message post in the form below." %>&nbsp;
<% steTxt "When you are done click on the <I>Post Message</I> button." %>
</P>

<FORM NAME="formedit" METHOD="Post" ACTION="post.asp" ID="formedit">
<INPUT TYPE="hidden" NAME="TopicID" VALUE="<%= nTopicID %>">
<input type="hidden" name="action" value="">

<% If sAction = "preview" Then %>

<input type="hidden" name="subject" value="<%= Server.HTMLEncode(sSubject) %>">
<input type="hidden" name="htmledit" value="<%= Server.HTMLEncode(sBody) %>">

<H4><% steTxt "Preview Message" %></H4>

<% forPreviewMessage nMemberID, nTopicID, sSubject, ConvertUBB(sBody) %>

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 width="500">
<TR>
	<TD align="right">
	<input type="button" name="_change" value=" <% steTxt "Make Changes" %> " onclick="javascript:history.go(-1)" class="form">
	<input type="submit" name="_action" value=" <% steTxt "Post" %> " class="form" onclick="document.formedit.action.value='post'">
	</TD>
</TR>
</TABLE>

<% Else ' form to enter the reply %>

<TABLE BORDER=0 CELLPADDING=6 CELLSPACING=0>
<tr>
	<td><b class="forml"><% steTxt "Posted By User" %></b><br>
	<% If Request.Cookies("Username") <> "" Then %>
	<font class="formd"><%= Request.Cookies("Username") %></font>
	<% Else %>
	<font class="formd"><% steTxt "Anonymous Coward" %></font><br><br>
	<p>
	<% steTxt "We noticed that you are not logged in." %>
	<% steTxt "Please enter your member login to receive credit for your post." %>
	<% steTxt "Or you may" %>&nbsp;<a href="<%= Application("ASPNukeBasePath") %>module/account/register/register.asp"><% steTxt "register now for a free account" %></A>.
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
	</td>
</tr><TR>
	<TD class="forml"><% steTxt "Message Subject" %><BR>
	<INPUT TYPE="text" NAME="Subject" VALUE="<%= sSubject %>" SIZE=45 MAXLENGTH="50" class="form" style="width:500px">
	</TD>
</TR><TR>
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
	<TD class="forml" ALIGN="right"><BR>
	<INPUT TYPE="submit" NAME="_action" VALUE=" <% steTxt "Post" %> " class="form" onclick="document.formedit.action.value='post'">
	</TD>
</TR>
</TABLE>

<% End If %>

</FORM>

<% Else %>

<H3><% steTxt "New Message Posted" %></H3>

<P>
<% steTxt "Your new message was successfully posted to the forum" %>&nbsp;<%= sTitle %>.
<% steTxt "Thank you for contributing your forum posting." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="topic.asp?topicid=<%= nTopicID %>" CLASS="footerlink"><% steTxt "Topic Overview" %></A> &nbsp;
	<A HREF="index.asp" CLASS="footerlink"><% steTxt "Forum Overview" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
