<!-- #include file="../../../lib/site_lib.asp"-->
<%
'--------------------------------------------------------------------
' activate.asp
'	Activates a new member account from the registration e-mail
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
Dim rsMember
Dim sAuthCode
Dim sErrorMsg
Dim sStatusMsg

sAuthCode = Request.QueryString("authcode")

If sAuthCode <> "" Then
	query = "SELECT MemberID, Username, AuthCode, Active " &_
			"FROM	tblMember " &_
			"WHERE	AuthCode = '" & Replace(sAuthCode, "'", "''") & "'"
	Set rsMember = adoOpenRecordset(query)
	If Not rsMember.EOF Then
		If steRecordBoolValue(rsMember, "Active") Then
			sErrorMsg = "Your account has already been activated"
		Else
			' set the account active
			query = "UPDATE tblMember " &_
					"SET	AuthCode = '', Active = 1 " &_
					"WHERE	AuthCode = '" & Replace(sAuthCode, "'", "''") & "'"
			Call adoExecute(query)

			' set the cookies to log this member in
			Response.Cookies("MemberID") = rsMember.Fields("MemberID").Value
			Response.Cookies("AuthCode") = rsMember.Fields("AuthCode").Value
			Response.Cookies("Username") = rsMember.Fields("Username").Value
			sStatusMsg = "<P>You are now logged in as member: <B>" & rsMember.Fields("Username").Value & "</B></P>"
		End If
	End If
Else
	sErrorMsg = "No valid authorization code could be found"
End If
%>
<!-- #include file="../../../header.asp"-->

<h3><% steTxt "Account Activation" %></h3>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>

<P>
If you are having trouble activating your account, you may reach our support
staff by using the <A HREF="<%= Application("ASPNukeBasePath") %>contact/index.asp">Contact Us</A> form.  Or you
may simply respond to the confirmation e-mail you were sent.
</P>

<% Else %>

<%= sStatusMsg %>

<P>
Thank you for registering with <%= Application("CompanyName") %>.  You are
currently logged into our site.  We hope that you find our site useful and
look forward to serving you in the future.
</P>

<% End If %>

<!-- #include file="../../../footer.asp"-->