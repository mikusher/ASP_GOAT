<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/bbforum_lib.asp" -->
<%
'--------------------------------------------------------------------
' edit.asp
'	Ability for author to edit a message post (& add remarks).
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
Dim nTopicID		' topic that the message is in
Dim nMessageID		' message the user is editting
Dim nThreadID		' thread that we are working with
Dim nMemberID		' member who owns the message in question
Dim sErrorMsg		' error message to display to user
Dim sStatusMsg 		' status message to display to user

' check for user login here
If Request.Cookies("MemberID") <> "" Then
	nMemberID = Request.Cookies("MemberID")
Else
	nMemberID = 0
End If

sAction = Trim(LCase(steForm("Action")))
nTopicID = steNForm("TopicID")
nThreadID = steNForm("ThreadID")
nMessageID = steNForm("MessageID")
sBody = steStripForm("htmledit")

If UCase(sAction) = "EDIT" Then
	If Trim(Replace(sBody, vbCrLf, "")) = "" Then
		sErrorMsg = steGetText("Please enter the changes to the message below")
	Else
		' append the "additional comments section"
		sBody = vbCrLf & vbCrLf &_
			"<DIV CLASS=""forumnote"">" & steGetText("Additional Remarks added") & " " & FormatDateTime(Now(), vbGeneralDate) & "</DIV>" & vbCrLf &_
			sBody

		' retrieve the existing message body here
		query = "SELECT	MessageBody " &_
				"FROM	tblMessage " &_
				"WHERE	MessageID = " & nMessageID
		Set rsMessage = adoOpenRecordset(query)
		If Not rsMessage.EOF Then
			sBody = rsMessage.Fields("MessageBody") & sBody

			' append the changes to the record
			query = "UPDATE	tblMessage " &_
					"SET	MessageBody = '" & Replace(sBody, "'", "''") & "', " &_
					"		Modified = " & adoGetDate & " " &_
					"WHERE	MessageID = " & nMessageID & " " &_
					"AND	MemberID = " & nMemberID
			Call adoExecute(query)
		Else
			sErrorMsg = steGetText("Unable to retrieve original message") & " (ID = " & nMessageID & ")"
		End If
		rsMessage.Close
		rsMessage = Empty
	End If
ElseIf UCase(sAction) = "DELETE" Then
	' make sure the message exists first
	query = "SELECT ParentMessageID FROM tblMessage " &_
			"WHERE	MessageID = " & nMessageID & " " &_
			"AND	MemberID = " & nMemberID
	Set rsMessage = adoOpenRecordset(query)
	If Not rsMessage.EOF Then
		Dim sThreadUpdate

		If rsMessage.Fields("ParentMessageID").Value > 0 Then
			sThreadUpdate = ", Threads = Threads - 1"
		End If
		rsMessage.Close
		Set rsMessage = Nothing

		' delete the message here
		query = "DELETE FROM tblMessage " &_
				"WHERE	MessageID = " & nMessageID & " " &_
				"AND	MemberID = " & nMemberID
		Call adoExecute(query)

		' update the message count and last post for the topic
		query = "UPDATE	tblMessageTopic " &_
				"SET	Messages = Messages - 1" & sThreadUpdate & ", Modified = " & adoGetDate & " " &_
				"WHERE	TopicID = " & nTopicID
		Call adoExecute(query)
		sStatusMsg = steGetText("The forum message post has been removed")
	Else
		rsMessage.Close
		Set rsMessage = Nothing
		sStatusMsg  = steGetText("Message no longer exists in the database")
	End If

End If


' retrieve the message information
If nMessageID > 0 Then
	query = "SELECT	TopicID, ThreadID, MemberID, Subject " &_
			"FROM	tblMessage " &_
			"WHERE	MessageID = " & nMessageID & " " &_
			"AND	Active <> 0 " &_
			"AND	Archive = 0"
	Set rsMessage = adoOpenRecordset(query)
	If Not rsMessage.EOF Then
		nTopicID = rsMessage.Fields("TopicID").Value
		' nThreadID = rsMessage.Fields("ThreadID").Value
		nMemberID = rsMessage.Fields("MemberID").Value
		sSubject = rsMessage.Fields("Subject").Value
	End If
End If
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/htmledit_lib.asp" -->

<script language="javascript" type="text/javascript">
<!-- // hide
function confirmClick(sActionType, sItemName, sLink)
{
	if (confirm('<% steTxt "Are you sure you want to" %> ' + sActionType + ': ' + sItemName + '?\n'))
	{
		window.open(sLink, "_self");
		return true;
	}
	return;
}
// unhide -->
</script>

<% If nMemberID = 0 Then %>

<H3><% steTxt "Login Required" %></H3>

<p>
<% steTxt "Before you can proceed editing your messages, you will need to identify yourself by using the member login that appears on this page to the left." %>&nbsp;
<% steTxt "You will only be allowed to edit messages that were originally posted by you." %>
</p>

<% ElseIf CStr(Request.Cookies("MemberID")) <> CStr(nMemberID) Then %>

<H3><% steTxt "Access Forbidden" %></H3>

<P>
<% steTxt "The message you selected to edit was not originally posted by you." %>&nbsp;
<% steTxt "For this reason, you are not permitted to edit the message." %>
</P>

<% ElseIf UCase(sAction) = "EDIT" Or UCase(sAction) = "" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Message" %></H3>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM NAME="formedit" METHOD="post" ACTION="edit.asp">
<INPUT TYPE="hidden" NAME="Action" VALUE="edit">
<INPUT TYPE="hidden" NAME="TopicID" VALUE="<%= nTopicID %>">
<INPUT TYPE="hidden" NAME="ThreadID" VALUE="<%= nThreadID %>">
<INPUT TYPE="hidden" NAME="MessageID" VALUE="<%= nMessageID %>">

<% forSingleMessage nTopicID, nMessageID %>

<P>
<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 width="500">
<TR>
	<TD>
	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 WIDTH="100%">
	<TR>
		<TD class="forml"><% steTxt "Additional Comments" %></TD>
		<TD ALIGN="right" VALIGN="top"><% HTMLCommandButtons %></TD>
	</TR>
	</TABLE><BR>
	<TEXTAREA NAME="htmledit" COLS="58" ROWS="10" WRAP="Virtual" class="form" style="width:500px"></TEXTAREA>
	</TD>
</TR><TR>
	<TD ALIGN="RIGHT"><br>
		<INPUT TYPE="button" NAME="_delete" VALUE=" <% steTxt "Delete" %> " class="form" onClick="confirmClick('<% steTxt "DELETE" %>', '<%= Replace(Server.HTMLEncode(sSubject), "'", "\'") %>', 'edit.asp?topicid=<%= nTopicID %>&threadid=<%= nThreadID %>&messageid=<%= nMessageID %>&action=delete')">
		<INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Edit" %> " class="form">
	</TD>
</TR>
</TABLE>
</P>

</FORM>

<% ElseIf UCase(sAction) = "DELETE" Then %>

<H3><% steTxt "Delete Message Successful" %></H3>

<% If sStatusMsg <> "" Then %>
<p><b class="error"><%= sStatusMsg %></b></p>
<% End If %>
<P>
<% steTxt "The message you selected has been removed from the forum." %>&nbsp;
<% steTxt "Use the buttons below to return to the discussion forum." %>
</P>

<% Else %>

<H3><% steTxt "Edit Message Successful" %></H3>

<P>
<% steTxt "The changes to your message were successful." %>&nbsp;
<% steTxt "The new message is shown below." %>
</P>

<% forSingleMessage nTopicID, nMessageID %>

<% End If %>

<P ALIGN="center">
	<% If nThreadID > 0 Then %>
	<A HREF="thread.asp?threadid=<%= nThreadID %>" CLASS="footerlink"><% steTxt "Back to Thread" %></A> &nbsp;
	<% End If %>
	<A HREF="topic.asp?topicid=<%= nTopicID %>" CLASS="footerlink"><% steTxt "Topic Overview" %></A> &nbsp;
	<A HREF="index.asp" CLASS="footerlink"><% steTxt "Forum Overview" %></A>
</P>

<!-- #include file="../../../footer.asp" -->