<%
'--------------------------------------------------------------------
' bbforum_lib.asp
'	This library of functions is useful for maintaining a bulletin board
'	system on your site.
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
' REQUIRES:
'	ado_lib.asp			for database calls

Const FOR_GRAPHIC_ACTIONLINKS = False		' action links for indiv. msgs

Dim for_bInit			' has the forum been initialized?

for_bInit = False

'------------------------------------------------------------------
' forThreadStart
'	Displays the HTML code for the start of a thread

Sub forThreadStart
%>
<P>
<TABLE border=0 CELLPADDING=0 CELLSPACING=0 CLASS="threadlist" width="100%">
<tr>
	<td class="threadlisthead">Author</td>
	<td class="threadlisthead" colspan="2">Message</td>
</tr>
<%
End Sub

'------------------------------------------------------------------
' forThreadEnd
'	Displays the HTML code for the end of a thread

Sub forThreadEnd
%>
</TABLE>
</P>
<%
End Sub

'------------------------------------------------------------------
' forActionLinks
'	Display the action links for an individual message

Sub forActionLinks(nThreadID, nMessageID, sUsername)
	Dim sEmailIcon, sPrivIcon, sEditIcon, sReplyIcon

	If FOR_GRAPHIC_ACTIONLINKS Then
		sEmailIcon = "<img src=""" & Application("ASPNukeBasePath") & "img/emailicon.gif"" width=16 height=16 border=""0"" alt=""E-mail user who posted this message"">"
		sPrivIcon = "<img src=""" & Application("ASPNukeBasePath") & "img/pvticon.gif"" width=16 height=16 border=""0"" alt=""Send a private message to this user"">"
		sEditIcon = "<img src=""" & Application("ASPNukeBasePath") & "img/editicon.gif"" width=16 height=16 border=""0"" alt=""Edit this message"">"
		sReplyIcon = "<img src=""" & Application("ASPNukeBasePath") & "img/replyicon.gif"" width=16 height=16 border=""0"" alt=""Reply to this message"">"
	Else
		sEmailIcon = "e-mail ."
		sPrivIcon = "msg ."
		sEditIcon = "edit ."
		sReplyIcon = "reply"
	End If
	With Response
		If Trim(sUsername & "") <> "" Then
			.Write "<a href=""email.asp?threadid="
			.Write nThreadID
			.Write "&messageid="
			.Write nMessageID
			.Write "&username="
			.Write Server.URLEncode(sUsername)
			.Write """ class=""forumaction"">"
			.Write sEmailIcon
			.Write "</A>&nbsp;"
		End If
		' .Write "</A>&nbsp;<a href=""private.asp?username="
		' .Write Server.URLEncode(sUsername)
		' .Write "&messageid="
		' .Write nMessageID
		' .Write "&threadid="
		' .Write nThreadID
		' .Write """ class=""forumaction"">"
		' .Write sPrivIcon
		.Write "<a href=""edit.asp?threadid="
		.Write nThreadID
		.Write "&messageid="
		.Write nMessageID
		.Write """ class=""forumaction"">"
		.Write sEditIcon
		.Write "</A>&nbsp;<a href=""reply.asp?messageid="
		.Write nMessageID 
		.Write """ class=""forumaction"">"
		.Write sReplyIcon
		.Write "</A>"
	End With
End Sub

'------------------------------------------------------------------
' forMessage
'	Displays an individual message within a thread.

Sub forMessage(rsMessage, nTopicID, nThreadID)
	Dim sIcon, sHomePage, sEmail

	' build image HTML to the member icon (if nec)
	If Not IsNull(rsMessage.Fields("ForumIcon").Value) Then
		If Trim(rsMessage.Fields("ForumIcon").Value) <> "" Then
			sIcon = "<img src=""" & rsMessage.Fields("ForumIcon").Value & """ border=""0"" ALT=""Click to View Member Profile""><BR>"
		End If
	End If
	If sIcon = "" Then
		sIcon = "<div class=""forumnopic"">Sorry,<br>No<br>Picture</div>"
	End If
	' build HTML for the member's home page (if any)
	If Not IsNull(rsMessage.Fields("HomePage").Value) Then
		If Trim(rsMessage.Fields("HomePage").Value) <> "" Then
			sHomePage = "<a href=""" & rsMessage.Fields("HomePage").Value & """ TARGET=""_new"">" &_
				"<img src=""" & Application("frm_HomePageIcon") & """ border=""0""></A>&nbsp;"
		End If
	End If
%>
<tr>
	<td class="threadprofile" ROWSPAN="2" valign="middle" align="center" width="100">
	<% If Not IsNull(rsMessage.Fields("Username").Value) Then %>
		<a href="profile.asp?topicid=<%= nTopicID %>&threadid=<%= nThreadID %>&username=<%= Server.URLEncode(rsMessage.Fields("Username").Value) %>" CLASS="forumprofile">
	<% Else %>
		<a href="javascript:void(0)" class="forumprofile">
	<% End If %>
	<%= sIcon %>
	<% If IsNull(rsMessage.Fields("Username").Value) Then %>Anonymous<% Else %><%= rsMessage.Fields("Username").Value %><% End If %></A>
	</td>
	<td class="threadheader">
		<font class="threadsubject"><%= rsMessage.Fields("Subject").Value %></FONT><br>
		<font class="threaddate"><%= adoFormatDateTime(rsMessage.Fields("Created").Value, vbLongDate) %></FONT>
	</td>
	<td class="threadheader" valign="top" align="right"><%= sHomePage %>
		<% forActionLinks nThreadID, rsMessage.Fields("MessageID").Value, rsMessage.Fields("Username").Value %>
	</td>
</tr><tr>
	<td class="threadbody" COLSPAN="2">
	<%= ConvertUBB(rsMessage.Fields("MessageBody").Value) %>
	</td>
</tr><tr>
	<TD COLSPAN="3" class="threadseparator"><img src="<%= Application("ASPNukeBasePath") %>img/pixel.gif" width=1 height=1></td>
</tr>
<%
End Sub

'------------------------------------------------------------------
' forSingleMessage
'	Retrieve one single message for display to the user

Sub forSingleMessage(nTopicID, nMessageID)
	Dim query, rsMessage
	' retrieve all of the messages in this thread (ordered by post date)
	query = "SELECT	tblMessage.MessageID, tblMessage.Subject, tblMessage.MessageBody, tblMessage.Created, " &_
			"		tblMessage.ModPoints, tblMember.Username, tblMessageProfile.ForumIcon, tblMember.HomePage " &_
			"FROM	tblMessage " &_
			"LEFT JOIN	tblMember ON tblMember.MemberID = tblMessage.MemberID " &_
			"LEFT JOIN	tblMessageProfile ON tblMessageProfile.MemberID = tblMessage.MemberID " &_
			"WHERE	tblMessage.MessageID = " & nMessageID & " " &_
			"AND	tblMessage.Active <> 0 " &_
			"AND	tblMessage.Archive = 0"
	Set rsMessage = adoOpenRecordset(query)
	If Not rsMessage.EOF Then
		forThreadStart
		forMessage rsMessage, nTopicID, 0
		forThreadEnd
	Else
		Response.Write "<P><B CLASS=""error"">Sorry, The message requested could not be found</B></P>"
	End If
End Sub

'------------------------------------------------------------------
' forPreviewMessage
'	Preview a message before it is posted to the forums

Sub forPreviewMessage(nMemberID, nTopicID, sSubject, sBody)
	Dim query, rsMember, sUsername, sFormIcon, sHomePage

	query = "SELECT	tblMember.Username, tblMessageProfile.ForumIcon, tblMember.HomePage " &_
			"FROM	tblMember " &_
			"LEFT JOIN	tblMessageProfile ON tblMessageProfile.MemberID = tblMember.MemberID " &_
			"WHERE	tblMember.MemberID = " & nMemberID
	Set rsMember = adoOpenRecordset(query)
	If Not rsMember.EOF Then
		sUsername = rsMember.Fields("Username").Value
		sForumIcon = rsMember.Fields("ForumIcon").Value
		sHomePage = rsMember.Fields("HomePage").Value
	Else
		sUsername = "Anonymous Coward"
		sForumIcon = Null
		sHomePage = Null
	End If
	rsMember.Close
	rsMember = Empty
	' display the message here
	forThreadStart
	' build image HTML to the member icon (if nec)
	If Not IsNull(sForumIcon) Then
		If Trim(sForumIcon) <> "" Then
			sIcon = "<img src=""" & sForumIcon & """ border=""0""><BR>"
		End If
	End If
	If sIcon = "" Then
		sIcon = "<DIV class=""forumnopic"">Sorry,<BR>No<BR>Picture</DIV>"
	End If
	' build HTML for the member's home page (if any)
	If Not IsNull(sHomePage) Then
		If Trim(sHomePage) <> "" Then
			sHomePage = "<a href=""" & sHomePage & """ TARGET=""_new"">" &_
				"<img src=""" & Application("frm_HomePageIcon") & """ border=""0""></A>&nbsp;"
		End If
	End If
%>
<tr>
	<td class="threadprofile" ROWSPAN="2" valign="middle" align="center" width="100">
	<a href="profile.asp?topicid=<%= nTopicID %>&username=<%= Server.URLEncode(sUsername) %>" CLASS="forumprofile" ALT="Click to View Member Profile">
	<%= sIcon %><%= sUsername %></A>
	</td>
	<td class="threadheader">
		<font class="threadsubject"><%= sSubject %></FONT><br>
		<font class="threaddate"><%= FormatDateTime(Now(), vbLongDate) %></FONT>
	</td>
	<td class="threadheader" valign="top" align="right"><%= sHomePage %>
		<% forActionLinks nThreadID, 0, sUsername %>
	</td>
</tr><tr>
	<td class="threadbody" COLSPAN="2">
	<%= ConvertUBB(sBody) %>
	</td>
</tr><tr>
	<TD COLSPAN="3"><img src="<%= Application("ASPNukeBasePath") %>img/pixel.gif" width=1 height=1></td>
</tr>
<%
	forThreadEnd
End Sub

'------------------------------------------------------------------
' forThread
'	Retrieve all of the messages for a thread from the database.

Sub forThread(nTopicID, nThreadID)
	Dim query, rsMessage

	' retrieve all of the messages in this thread (ordered by post date)
	query = "SELECT	tblMessage.MessageID, tblMessage.Subject, tblMessage.MessageBody, tblMessage.Created, " &_
			"		tblMessage.ModPoints, tblMember.Username, tblMessageProfile.ForumIcon, tblMember.HomePage " &_
			"FROM	tblMessage " &_
			"LEFT JOIN	tblMember ON tblMember.MemberID = tblMessage.MemberID " &_
			"LEFT JOIN	tblMessageProfile ON tblMessageProfile.MemberID = tblMessage.MemberID " &_
			"WHERE	(tblMessage.ThreadID = " & nThreadID & " " &_
			"		OR	tblMessage.MessageID = " & nThreadID & ") " &_
			"AND	tblMessage.Active <> 0 " &_
			"AND	tblMessage.Archive = 0 " &_
			"ORDER BY tblMessage.Created"
	' Response.Write query : Response.End
	Set rsMessage = adoOpenRecordset(query)
	If Not rsMessage.EOF Then
		forThreadStart
		Do Until rsMessage.EOF
			forMessage rsMessage, nTopicID, nThreadID
			rsMessage.MoveNext
		Loop
		forThreadEnd
	Else
		Response.Write "<P><B CLASS=""error"">Sorry, No messages could be found in the thread</B></P>"
	End If
End Sub

'------------------------------------------------------------------
' forThreadList
'	Builds the list of threads from the message forums

Sub forThreadList(nTopicID)
	Dim query, rsList, nThreadID, I

	' retrieve the list of message threads here
	query = "SELECT	MessageID, Subject, Messages, LastPost, Username " &_
			"FROM	tblMessage " &_
			"INNER JOIN	tblMember ON tblMessage.MemberID = tblMember.MemberID " &_
			"WHERE	tblMessage.TopicID = " & nTopicID & " " &_
			"AND	tblMessage.ParentMessageID = 0 " &_
			"AND	tblMessage.Active <> 0 " &_
			"AND	tblMessage.Archive = 0 " &_
			"ORDER BY tblMessage.Created DESC"
	Set rsList = adoOpenRecordset(query)
%>
<TABLE border=0 CELLPADDING=2 CELLSPACING=0 CLASS="list" width="100%">
<TR BGCOLOR="#000080">
	<td class="listhead">Thread</td>
	<td class="listhead">Author</td>
	<td class="listhead" align="center">Messages</td>
	<td class="listhead" align="right">Last Post</td>
</tr>
<%	I = 0
	Do Until rsList.EOF  %>
<TR CLASS="list<%= I mod 2 %>">
	<TD><B><a href="thread.asp?topicid=<%= nTopicID %>&threadid=<%= rsList.Fields("MessageID").Value %>" class="forumtopic"><%= Server.HTMLEncode(rsList.Fields("Subject").Value) %></A></B></td>
	<TD><a href="profile.asp?topicid=<%= nTopicID %>&threadid=<%= rsList.Fields("MessageID").Value %>&username=<%= Server.URLEncode(rsList.Fields("Username").Value) %>" class="forumprofile" ALT="Click to View Member Profile"><%= Server.URLEncode(rsList.Fields("Username").Value) %></A></td>
	<TD align="center"><B><%= rsList.Fields("Messages").Value %></B></td>
	<TD align="right"><%= adoFormatDateTime(rsList.Fields("LastPost").Value, vbShortDate) %></B></td>
</tr>
<%		rsList.MoveNext
		I = I + 1
	Loop %>
</TABLE>
<%
End Sub


'----------------------------------------------------------------------------
' ReplaceEmails
'	Replaces email links added to the text using UBB syntax with the HTML
'	equivalent

Function ReplaceEmails(ByVal strText)
	On Error Resume Next
	
	Dim objRegExp
	Set objRegExp = New RegExp 
	objRegExp.Pattern="\[email]([^\[]*)\[/email]"
	objRegExp.IgnoreCase=True
	objRegExp.Global=True
	strText = objRegExp.Replace(strText,"<a href=""mailto:$1"" class=""linktext"">$1</a>")
	
	If UseErrHandler = 1 then
		If Err.Number <> 0 then
			Call ErrHandler (Err.Description, Err.Number, Request.ServerVariables("Script_Name"), Request.ServerVariables("Remote_Addr"), Err.Source, "ReplaceEmails")
		End If
	End If
	
	ReplaceEmails = strText
End Function

'----------------------------------------------------------------------------
' ReplaceURLs
'	Replaces URL strings added to the text using UBB syntax with the HTML
'	equivalent

Function ReplaceURLs(ByVal strText)
	On Error Resume Next
	
	Dim objRegExp
	Set objRegExp = New RegExp
	objRegExp.Pattern = "\[url=([^]]*)]([^\[]*)\[/url]"
	objRegExp.IgnoreCase = True
	objRegExp.Global = True
	strText = objRegExp.Replace(strText,"<a href=""$1"" class=""linktext"" target=""_blank"">$2</a>")
	
	If UseErrHandler = 1 then
		If Err.Number <> 0 then
			Call ErrHandler (Err.Description, Err.Number, Request.ServerVariables("Script_Name"), Request.ServerVariables("Remote_Addr"), Err.Source, "ReplaceURLs")
		End If
	End If
	
	ReplaceURLs = strText
End Function

'----------------------------------------------------------------------------
' ConvertUBB
'	Convert Ultimate Bulletin Board markup to HTML markup

Function ConvertUBB(ByVal strText)	
	strText = replace(strText, vbcrlf, "<br>", 1, -1, 1)
	strText = replace(strText, "<br><br>[$More]", "", 1, -1, 1)
	strText = replace(strText, "[image]", "<img src=""", 1, -1, 1)
	strText = replace(strText, "[/image]", """ border=""0"">", 1, -1, 1)
	strText = replace(strText, "[b]", "<b>", 1, -1, 1)
	strText = replace(strText, "[/b]", "</b>", 1, -1, 1)
	strText = replace(strText, "[strong]", "<strong>", 1, -1, 1)
	strText = replace(strText, "[/strong]", "</strong>", 1, -1, 1)
	strText = replace(strText, "[s]", "<s>", 1, -1, 1)
    strText = replace(strText, "[/s]", "</s>", 1, -1, 1)
	strText = replace(strText, "[u]", "<u>", 1, -1, 1)
	strText = replace(strText, "[/u]", "</u>", 1, -1, 1)
	strText = replace(strText, "[i]", "<i>", 1, -1, 1)
	strText = replace(strText, "[/i]", "</i>", 1, -1, 1)
	strText = replace(strText, "[hr]", "<hr size=""1"" color=""#000000"">", 1, -1, 1)
	strText = replace(strText, "[:(!]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_angry.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[B)]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_blackeye.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[xx(]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_dead.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[XX(]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_dead.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:O]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_shock.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:o]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_shock.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:0]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_shock.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:I]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_blush.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:(]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_sad.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[8)]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_shy.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:)]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[}:)]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_evil.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:D]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_big.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[8D]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_cool.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[|)]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_sleepy.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:o)]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_clown.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:O)]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_clown.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:0)]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_clown.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:P]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_tongue.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:p]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_tongue.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[;)]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_wink.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[8]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_8ball.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[?]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_question.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[^]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_approve.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[V]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_dissapprove.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:X]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/icon_smile_kisses.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:boxing:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/boxing.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:crash:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/crash.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:drool:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/drool.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:drunk:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/drunk.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:mwink:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/mwink.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:nono:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/nono.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:pimp:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/pimp.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:spank:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/spank.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:sweat:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/sweat.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:thefinger:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/thefinger.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:2gunsfiring:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/2gunsfiring.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:angel:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/angel.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:angry2:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/angry2.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:banana:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/banana.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:beerchug:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/beerchug.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:birthday:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/birthday.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:square:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/square.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:bigeyes:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/bigeyes.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:waving:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/waving.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:eek:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/eekr.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:finger:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/finger.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:freak:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/freak.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:frustrated:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/frustrated.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:hammer:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/hammer.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:idea:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/idea.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:looney:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/looney.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:machinegun:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/machinegun.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:newconfuse:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/newconfuse.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:nut:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/nut.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:peek:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/peek.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:pukey:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/pukey.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:rocketlauncher:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/rocketlauncher.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:rolleyes2:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/rolleyes2.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:s:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/s.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:scared:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/scared.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:sleep:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/sleep.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:swear:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/swear.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = replace(strText, "[:what:]", "<img src=""" & Application("ASPNukeBasePath") & "img/graemlins/what.gif"" border=""0"" align=""middle"">", 1, -1, 1)
	strText = ReplaceEmails(strText)
	strText = ReplaceURLS(strText)
	
	ConvertUBB = strText
End Function
%>