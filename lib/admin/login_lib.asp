<!-- #include file="../../lib/sha256.asp" -->
<%
'--------------------------------------------------------------------
' login_lib.asp
'	Manages the logins to the admin area of the site.  Include this
'	file before any admin script that you want password-protected.
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

Dim sAdmLStat		' SQL query statement
Dim rsAdmLUser
Dim sAdmLErrorMsg
Dim bAdmLLoginSuccess	' user is logged in successfully
Dim nAdmUserID

bAdmLLoginSuccess = true

' the following code creates a login form if user is not logged in
If Request.Cookies("AdminUsername") = "" Then
	bAdmLLoginSuccess = false

	' check for login attempt here
	If steForm("adm_action") <> "" Then
		' check the form variables
		If Trim(steForm("adm_username")) = "" Then
			sAdmLErrorMsg = "Please enter your username"
		ElseIf Trim(steForm("adm_password")) = "" Then
			sAdmLErrorMsg = "Plesae enter your password"
		Else
			' retrieve the user information here
			sAdmLStat = "SELECT	UserID, FirstName, LastName, Username " & _
					"FROM	tblUser " &_
					"WHERE	Username = " & steQForm("adm_username") & " " &_
					"AND	Password = '" & SHA256(steForm("adm_password")) & "'"
			Set rsAdmLUser = adoOpenRecordset(sAdmLStat)
			If Not rsAdmLUser.EOF Then
				' login the user by setting cookies
				nAdmUserID = rsAdmLUser.Fields("UserID").Value
				Response.Cookies("AdminUserID") = nAdmUserID
				Response.Cookies("AdminUsername") = rsAdmLUser.Fields("Username").Value
				Response.Cookies("AdminFullname") = rsAdmLUser.Fields("Firstname").Value & " " & rsAdmLUser.Fields("LastName").Value
				rsAdmLUser.Close
				Set rsAdmLUser = Nothing
				' set the navigation login keys
				Call navLogin(nAdmUserID)
				bAdmLLoginSuccess = True
			Else
				' display error message to user here
				sAdmLErrorMsg = "Username and Password are invalid"
				Set rsAdmLUser = Nothing
			End If		
		End If
	End If
	%>
	
	<% If Not bAdmLLoginSuccess Then %>
	
	<H3>Admin Login Required</H3>
	
	<P>
	In order to proceed any further into the admin area of the site, we first
	need to verify your identity.  Please enter your registered admin username
	and password in the form below.
	</P>
	
	<% If sAdmLErrorMsg <> "" Then %>
	<P><B CLASS="error"><%= sAdmLErrorMsg %></B></P>
	<% End If %>
	
	<FORM METHOD="post" ACTION="<%= Request.ServerVariables("SCRIPT_NAME") %>">
	<INPUT TYPE="hidden" NAME="adm_action" VALUE="login">
	<% Call locPassQueryStringVars %>
	<% Call locPassFormVars %>
	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
	<TR>
		<TD class="forml">Username<BR>
		<INPUT TYPE="text" NAME="adm_username" VALUE="<%= steEncForm("adm_username") %>" SIZE="16" MAXLENGTH="16" class="form">
		</TD>
	</TR><TR>
		<TD class="forml">Password<BR>
		<INPUT TYPE="password" NAME="adm_password" VALUE="" SIZE="16" MAXLENGTH="16" class="form">
		</TD>
	</TR><TR>
		<TD ALIGN="center"><INPUT TYPE="submit" NAME="adm_submit" VALUE=" Login " class="form">
		</TD>
	</TR>
	</TABLE>
	</FORM>
	
	<!-- #include file="../../footer.asp" -->
	<%	Response.End
	End If
End If ' Request.Cookies("Username") = ""

' convert all of the querystring variables to form variables so they
' will be available in the intended script.

Sub locPassQueryStringVars
	Dim sKey

	For Each sKey In Request.QueryString
		With Response
			.Write "<INPUT TYPE=""hidden"" NAME="""
			.Write sKey
			.Write """ VALUE="""
			.Write Server.HTMLEncode(Request.QueryString(sKey))
			.Write """>"
			.Write vbCrLf
		End With
	Next
End Sub

Sub locPassFormVars
	Dim sKey

	For Each sKey In Request.Form
		If Left(sKey, 4) <> "adm_" Then
			With Response
				.Write "<INPUT TYPE=""hidden"" NAME="""
				.Write sKey
				.Write """ VALUE="""
				.Write Server.HTMLEncode(Request.Form(sKey))
				.Write """>"
				.Write vbCrLf
			End With
		End If
	Next
End Sub
%>
<!-- #include file="../../lib/admin/nav_lib.asp" -->