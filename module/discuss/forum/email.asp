<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/class/nukemail.asp"-->
<%
'--------------------------------------------------------------------
' email.asp
'	Allows one user to send an e-mail to another user of the forums
'	without disclosing that member's e-mail address.
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

Dim query			' SQL query to run
Dim nMessageID		' message user clicked on "mail link"
Dim sAction			' action performed by this script
Dim sUsername		' username to send message to
Dim sSubject		' subject of the mail message
Dim sBody			' body for the mail message
Dim rsMember		' information about the member
Dim nThreadID		' message thread
Dim sErrorMsg		' error message to report to user

sUsername = Request("Username")
nMessageID = Request("MessageID")
nThreadID = Request("ThreadID")
sAction = Request.Form("Action")
sSubject = Request.Form("Subject")
sBody = Request.Form("Body")

If IsNumeric(nMessageID) And CStr(nMessageID) <> "" Then nMessageID = CInt(nMessageID) Else nMessageID = 0
If IsNumeric(nThreadID) And CStr(nThreadID) <> "" Then nThreadID = CInt(nThreadID) Else nThreadID = 0

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
End If

' retrieve the information about the recipient
query = "SELECT	MemberID, Firstname, Lastname, EmailAddress " &_
		"FROM	tblMember " &_
		"WHERE	Username = '" & Replace(sUsername, "'", "''") & "'"
Set rsMember = adoOpenRecordset(query)

' send the e-mail here if necessary
If UCase(sAction) = "SEND" Then
	If Trim(sSubject) = "" Then
		sErrorMsg = steGetText("Please enter the subject for your mail message")
	ElseIf Trim(sBody) = "" Then
		sErrorMsg = steGetText("Please enter the body for your mail message")
	Else
		' retrieve information about the sender
		query = "SELECT Firstname, Lastname, EmailAddress " &_
				"FROM	tblMember " &_
				"WHERE	MemberID = " & Request.Cookies("MemberID")
		Set rsSender = adoOpenRecordset(query)
		If Not rsSender.EOF Then
			' create a log of the mail message here
			query = "INSERT INTO tblMessageEmail (" &_
					"		MessageID, FromMemberID, ToMemberID, Subject, Body" &_
					") VALUES (" &_
							nMessageID & "," & Request.Cookies("MemberID") & "," & rsMember.Fields("MemberID").Value &_
							",'" & Replace(sSubject, "'", "''") & "','" & Replace(sBody, "'", "''") & "'" &_
					")"
			Call adoExecute(query)

			' send out the e-mail to the member
			Set oMail = New NukeMail
			oMail.FromAddress = rsSender.Fields("EmailAddress").Value
			oMail.FromName = rsSender.Fields("Firstname").Value & " " & rsSender.Fields("Lastname").Value
			oMail.ToAddress = rsMember.Fields("EmailAddress").Value
			oMail.ToName = rsMember.Fields("Firstname").Value & " " & rsMember.Fields("Lastname").Value
			oMail.Subject = sSubject
			oMail.TextBody = sBody
			If Not oMail.Send Then
				Response.Write "<p><b class=""error"">" & oMail.ErrorMsg & "</b></p>"
			End If
		Else
			sErrorMsg = steGetText("Unable to determine identity of sender") & " (ID = " & Request.Cookies("MemberID") & ")"
		End If
	End If
End If

%>
<!-- #include file="../../../header.asp" -->

<% If rsMember.EOF Then %>
<H3><% steTxt "Member Not Found" %></H3>

<P><B CLASS="error">
<% steTxt "Sorry, but the member you wish to send an e-mail to does not allow e-mail to be sent to them." %>&nbsp;
<% steTxt "The reason for this could be that the member account has been disabled, or the member simply refused to accept e-mail." %>&nbsp;
<% steTxt "We regret any inconvenience this may cause you." %>
</B></P>

<% ElseIf UCase(sAction) <> "SEND" Or sErrorMsg <> "" Then %>
<H4><% steTxt "Send E-Mail to Member" %> <%= sUsername %></H4>

<P>
<% steTxt "Please enter the subject and body of your mail message in the form below." %>&nbsp;
<% steTxt "This message will be e-mailed to the member." %>&nbsp;
<% steTxt "The member will be given your e-mail address so they may reply to you directly." %>&nbsp;
<% steTxt "For security reasons, we do not give out the e-mail addresses of our members." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="email.asp">
<INPUT TYPE="hidden" NAME="ThreadID" VALUE="<%= nThreadID %>">
<INPUT TYPE="hidden" NAME="MessageID" VALUE="<%= nMessageID %>">
<INPUT TYPE="hidden" NAME="Username" VALUE="<%= Server.HTMLEncode(sUsername) %>">
<INPUT TYPE="hidden" NAME="Action" VALUE="send">

<TABLE BORDER=0 CELLPADDING="2" CELLSPACING="0">
<TR>
	<TD><b class="forml"><% steTxt "From User" %></b></td><td></td>
	<TD CLASS="formd">
	<% If Request.Cookies("Username") <> "" Then %>
	<%= Request.Cookies("Username") %>
	<% Else %>
	<% steTxt "Anonymous Coward" %><br><br>
	<p>
	<% steTxt "We noticed that you are not logged in." %>
	<% steTxt "Please enter your member login to identify who is sending the message." %>
	<% steTxt "Or you may" %> <a href="<%= Application("ASPNukeBasePath") %>module/account/register/register.asp"><% steTxt "register now for a free account" %></A>.
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
	<TD CLASS="forml"><% steTxt "To" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD CLASS="formd"><%= sUsername %></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Subject" %></TD><TD></TD>
	<TD>
	<INPUT TYPE="text" NAME="subject" VALUE="<%= Server.HTMLEncode(Request.Form("Subject")) %>" SIZE="46" MAXLENGTH="100" class="form" style="width:380px">
	</TD>
</TR><TR>
	<TD CLASS="formd" COLSPAN="3"><B class="forml"><% steTxt "Body" %></B> <I>(<% steTxt "enter plain text only" %>)</I><BR>
	<TEXTAREA NAME="body" COLS="55" ROWS="12" WRAP="Virtual" CLASS="form" style="width:500px"><%= Server.HTMLEncode(Request.Form("Body")) %></TEXTAREA>
	</TD>
</TR><TR>
	<TD colspan="3" ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Send" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>
<H3><% steTxt "Message Sent to" %>&nbsp;<%= sUsername %></H3>

<P>
<% steTxt "Your e-mail message was successfully sent to the specified user." %>&nbsp;
<% steTxt "The member may respond to your message at their discretion." %>&nbsp;
<% steTxt "Please respect the privacy of others and do not send offensive or abusive messages." %>
</P>

<P ALIGN="center">
	<A href="thread.asp?threadid=<%= nThreadID %>" class="actionlink">
</P>
<% End If %>

<P ALIGN="center">
	<A HREF="thread.asp?threadid=<%= nThreadID %>" CLASS="footerlink"><% steTxt "Back to Thread" %></A> &nbsp;
	<A HREF="index.asp" CLASS="footerlink"><% steTxt "Forum Overview" %></A>
</P>

<!-- #include file="../../../footer.asp" -->