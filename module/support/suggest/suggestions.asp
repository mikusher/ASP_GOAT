<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/class/nukemail.asp"-->
<%
'--------------------------------------------------------------------
' suggestions.asp
'	Submit suggestions on how this site may be improved.
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
Dim sBody
Dim sErrorMsg

If steForm("action") = "send" Then
	' check for required fields
	If Trim(steForm("subject")) = "" Then
		sErrorMsg = sErrorMsg & "Please enter the subject of your message."
	End If
	If sErrorMsg = "" Then
		' insert the suggestion
		sStat = "INSERT INTO tblSuggestion (" &_
				"	FromName, FromEmail, Subject, Body" &_
				") VALUES (" &_
				"'" & Replace(steForm("FromName"), "'", "''") & "'," &_
				"'" & Replace(steForm("FromEmail"), "'", "''") & "'," &_
				"'" & Replace(steForm("Subject"), "'", "''") & "'," &_
				"'" & Replace(steForm("Body"), "'", "''") & "'" &_
				")"
		Call adoExecute(sStat)

		' send out the e-mail for this suggestion
		sBody = Application("CompanyName") & " suggestion received: " & Now() & vbCrLf & vbCrLf &_
			"From:    " & steForm("FromName") & " <" & steForm("FromEmail") & ">" & vbCrLf &_
			"Subject: " & steForm("Subject") & vbCrLf & vbCrLf &_
			steForm("Body")

		Set oMail = New NukeMail
		oMail.FromAddress = Application("Suggestion_FromAddress")
		oMail.FromName = Application("Suggestion_FromName")
		oMail.ToAddress = Application("Suggestion_ToAddress")
		oMail.ToName = Application("Suggestion_ToName")
		oMail.Subject = steForm("Subject")
		oMail.TextBody = sBody
		If Not oMail.Send Then
			Response.Write "<p><b class=""error"">" & oMail.ErrorMsg & "</b></p>"
		End If
	End If
End If
%>
<!-- #include file="../../../header.asp" -->

<% If steForm("action") <> "send" Or sErrorMsg <> "" Then %>

<h3>Suggestions</h3>

<p>
Please submit your comments, questions and/or suggestions on how we can 
improve this site and the ASPNuke web content management system.  We
appreciate your comments and will respond to all your inquiries.
</p>

<% If sErrorMsg <> "" Then %>
<p><b class="error"><%= sErrorMsg %></b></p>
<% End If %>

<form method="post" action="suggestions.asp">
<input type="hidden" name="action" value="send">
<table border=0 cellpadding=2 cellspacing=0 width="444">
<tr>
	<td><b class="forml">Your Name</b><br>
	<input type="text" name="fromname" value="<%= steEncForm("fromname") %>" size="16" maxlength="100" class="form" style="width:218px">
	</td>
	<td><b class="forml">Your E-Mail Address</b><br>
	<input type="text" name="fromemail" value="<%= steEncForm("fromemail") %>" size="16" maxlength="100" class="form" style="width:218px">
	</td>
</tr><tr>
	<td colspan=2><b class="forml">Subject</b><br>
	<input type="text" name="subject" value="<%= steEncForm("subject") %>" size="32" maxlength="100" class="form" style="width:440px">
	</td>
</tr><tr>
	<td colspan=2><b class="forml">Body</b><br>
	<textarea name="body" cols="80" rows="10" class="form" style="width:440px"><%= steEncForm("Body") %></textarea>
	</td>
</tr><tr>
	<td colspan=2><br>
	<font class="tinytext">
	By submitting this form, you agree that <%= Application("CompanyName") %> may use
	your suggestions anyway they see fit including, but not limited to, publishing your
	comments on this site.
	</font><br><br>
	</td>
</tr><tr>
	<td colspan=2 align="center">
	<input type="submit" name="_submit" value=" Send Message " class="form">
	</td>
</tr>
</table>
</form>

<% Else %>

<h3>Suggestion Received</h3>

<p>
Thank you for taking the time to send us your suggestions.  Your suggestions
help us to build a better product which benefits the entire world.  We hope
that you enjoy <%= Application("CompanyName") %> and tell your friends about
us!
</p>

<% End If %>
<!-- #include file="../../../footer.asp" -->