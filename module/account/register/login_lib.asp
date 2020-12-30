<!-- #include file="../../../lib/sha256.asp" -->
<%
'--------------------------------------------------------------------
' login_lib.asp
'	Manages the login for members =of the site.  Include this
'	file before any script that you want password-protected.
'	REQUIRES: /lib/site_lib.asp
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

Dim sMbrStat		' SQL query statement
Dim rsMbrMbr
Dim sMbrErrorMsg
Dim bMbrLoginSuccess	' user is logged in successfully
Dim nMemberID

bMbrLoginSuccess = true

' the following code creates a login form if user is not logged in
If Request.Cookies("Username") = "" Then
	bMbrLoginSuccess = false

	' check for login attempt here
	If steForm("mbr_action") <> "" Then
		' check the form variables
		If Trim(steForm("mbr_username")) = "" Then
			sMbrErrorMsg = steGetText("Please enter your Username")
		ElseIf Trim(steForm("mbr_password")) = "" Then
			sMbrErrorMsg = steGetText("Plesae enter your Password")
		Else
			' retrieve the user information here
			sMbrStat = "SELECT	MemberID, FirstName, LastName, Username " & _
					"FROM	tblMember " &_
					"WHERE	Username = " & steQForm("mbr_username") & " " &_
					"AND	Password = '" & SHA256(steForm("mbr_password")) & "'"
			Set rsMbrMbr = adoOpenRecordset(sMbrStat)
			If Not rsMbrMbr.EOF Then
				' login the user by setting cookies
				nMemberID = rsMbrMbr.Fields("MemberID").Value
				Response.Cookies("MemberID") = nMemberID
				Response.Cookies("Username") = rsMbrMbr.Fields("Username").Value
				Response.Cookies("Fullname") = rsMbrMbr.Fields("Firstname").Value & " " & rsMbrMbr.Fields("LastName").Value
				rsMbrMbr.Close
				Set rsMbrMbr = Nothing
				bMbrLoginSuccess = True
			Else
				' display error message to user here
				sMbrErrorMsg = steGetText("Username and Password are invalid")
				Set rsMbrMbr = Nothing
			End If		
		End If
	End If
	%>
	
	<% If Not bMbrLoginSuccess Then %>
	
	<H3><% steTxt "Member Login Required" %></H3>
	
	<P>
	<% steTxt "In order to proceed any further into the admin area of the site, we first need to verify your identity." %>
	<% steTxt "Please enter your registered admin username and password in the form below." %>
	</P>
	
	<% If sMbrErrorMsg <> "" Then %>
	<P><B CLASS="error"><%= sMbrErrorMsg %></B></P>
	<% End If %>
	
	<FORM METHOD="post" ACTION="<%= Request.ServerVariables("SCRIPT_NAME") %>">
	<INPUT TYPE="hidden" NAME="mbr_action" VALUE="login">
	<% Call locPassQueryStringVars %>
	<% Call locPassFormVars %>
	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
	<TR>
		<TD class="forml"><% steTxt "Username" %><BR>
		<INPUT TYPE="text" NAME="mbr_username" VALUE="<%= steEncForm("mbr_username") %>" SIZE="16" MAXLENGTH="16" class="form">
		</TD>
	</TR><TR>
		<TD class="forml"><% steTxt "Password" %><BR>
		<INPUT TYPE="password" NAME="mbr_password" VALUE="" SIZE="16" MAXLENGTH="16" class="form">
		</TD>
	</TR><TR>
		<TD ALIGN="center"><INPUT TYPE="submit" NAME="mbr_submit" VALUE=" <% steTxt "Login" %> " class="form">
		</TD>
	</TR>
	</TABLE>
	</FORM>
	
	<!-- #include file="../../../footer.asp" -->
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
		If Left(sKey, 4) <> "mbr_" Then
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