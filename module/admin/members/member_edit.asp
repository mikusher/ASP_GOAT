<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' member_edit.asp
'	Edit an existing member in the database.
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
Dim sErrorMsg
Dim sUsername
Dim sPassword
Dim sConfirm
Dim sFirstName
Dim sLastName
Dim sAddress1
Dim sAddress2
Dim sCity
Dim sStateCode
Dim sZipCode
Dim nCountryID
Dim sEmail
Dim sCode			' authorization code
Dim rsState
Dim rsCountry
Dim rsEdit
Dim sPassUpdate

sAction = LCase(steForm("Action"))
sUsername = steForm("Username")
sPassword = steForm("Password")
sConfirm = steForm("Confirm")
sFirstName = steForm("FirstName")
sLastName = steForm("LastName")
sAddress1 = steForm("Address1")
sAddress2 = steForm("Address2")
sCity = steForm("City")
sStateCode = steForm("StateCode")
sZipCode = steForm("ZipCode")
nCountryID = steNForm("CountryID")
sEmail = steForm("EmailAddress")

If sAction = "edit" Then
	' check for required fields
	If (Trim(sFirstName) = "") Then
		sErrorMsg = sErrorMsg & steGetText("Please Enter your First Name") & "<BR>"
	End If
	If (Trim(sLastName) = "") Then
		sErrorMsg = sErrorMsg & steGetText("Please Enter your Last Name") & "<BR>"
	End If
	If (Trim(sUsername) = "") Then
		sErrorMsg = sErrorMsg & steGetText("Please Enter your Username") & "<BR>"
	End If
	If Len(Trim(sUsername)) < 6 Then
		sErrorMsg = sErrorMsg & steGetText("Your username must be at least 6 characters long") & "<BR>"
	End If
	If (Trim(sPassword) = "") And (Trim(sConfirm) <> "") Then
		sErrorMsg = sErrorMsg & steGetText("Please Enter your Password") & "<BR>"
	End If
	If (Trim(sPassword) <> "") And Len(Trim(sPassword)) < 6 Then
		sErrorMsg = sErrorMsg & steGetText("Your password must be at least 6 characters long") & "<BR>"
	End If
	If (Trim(sConfirm) = "") And (Trim(sPassword) <> "") Then
		sErrorMsg = sErrorMsg & steGetText("Please confirm the password that you entered") & "<BR>"
	ElseIf sPassword <> sConfirm Then
		sErrorMsg = sErrorMsg & steGetText("Confirmation password does not match password") & "<BR>"
	End If
	If (Trim(sEmail) = "") Then
		sErrorMsg = sErrorMsg & steGetText("Please Enter your E-Mail Address") & "<BR>"
	End If
	If sErrorMsg = "" Then
		' check to see if the username is already in use
		query = "SELECT MemberID " &_
				"FROM	tblMember " &_
				"WHERE	Username = '" & sUsername & "' " &_
				"AND	MemberID <> " & steForm("memberid")
		Set rsMember = adoOpenRecordset(query)
		If Not rsMember.EOF Then
			sErrorMsg = sErrorMsg & steGetText("The username you entered is already taken, please choose another") & "<BR>"
		Else
			' no errors occurred - add this member to the database
			If Trim(sPassword) <> "" Then
				sPassUpdate = "		Password = '" & SHA256(sPassword) & "',"
			End If
			query = "UPDATE tblMember SET " &_
					"		Username = '" & Replace(sUsername, "'", "''") & "'," &_
					sPassUpdate &_
					"		FirstName = '" & Replace(sFirstName, "'", "''") & "'," &_
					"		LastName = '" & Replace(sLastName, "'", "''") & "'," &_
					"		Address1 = '" & Replace(sAddress1, "'", "''") & "'," &_
					"		Address2 = '" & Replace(sAddress2, "'", "''") & "'," &_
					"		City = '" & Replace(sCity, "'", "''") & "'," &_
					"		StateCode = '" & sStateCode & "'," &_
					"		ZipCode = '" & Replace(sZipCode, "'", "''") & "'," &_
					"		CountryID = " & nCountryID & "," &_
					"		EMailAddress = '" & Replace(sEmail, "'", "''") & "' " &_
					"WHERE	MemberID = " & steForm("memberid")
			Call adoExecute(query)

		End If
	End IF
End If

' retrieve the member record to edit here
If steNForm("memberid") <> 0 Then
	query = "SELECT	* FROM	tblMember WHERE MemberID = " & steForm("memberid")
	Set rsEdit = adoOpenRecordset(query)
End If

' build the selection list for the states
query = "SELECT StateCode, StateName " &_
		"FROM	tblState " &_
		"WHERE	Active <> 0 " &_
		"AND	Archive = 0 " &_
		"ORDER BY StateName"
Set rsState = adoOpenRecordset(query)

' build the selection ist for the countries
query = "SELECT CountryID, CountryName " &_
		"FROM	tblCountry " &_
		"WHERE	Active <> 0 " &_
		"AND	Archive = 0 " &_
		"ORDER BY CountryName"
Set rsCountry = adoOpenRecordset(query)
%>
<!-- #include file="../../../header.asp"-->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Members" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If sAction <> "edit" Or sErrorMsg <> "" Then %>
	<H3><% steTxt "Edit Member Registration" %></H3>

	<P>
	<% steTxt "Use the form below to make changes to a member account." %>
	<% steTxt "These changes will take effect immediately." %>
	<% steTxt "If you leave the password and confirm password blank, no change will be made to the password." %>
	</P>

	<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
	<% End If %>

	<FORM METHOD="post" ACTION="member_edit.asp">
	<INPUT TYPE="hidden" NAME="memberid" VALUE="<%= steEncForm("memberid") %>">

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
	<TR>
		<TD class="forml"><% steTxt "First Name" %> (req)<BR>
		<INPUT TYPE="text" NAME="FirstName" VALUE="<%= steRecordEncValue(rsEdit, "FirstName") %>" SIZE="22" MAXLENGTH="32" class="form">
		</TD>
		<TD class="forml"><% steTxt "Last Name" %> (req)<BR>
		<INPUT TYPE="text" NAME="LastName" VALUE="<%= steRecordEncValue(rsEdit, "LastName") %>" SIZE="22" MAXLENGTH="32" class="form">
		</TD>
	</TR>
	<TR>
		<TD class="forml"><% steTxt "Username" %> (req)<BR>
		<INPUT TYPE="text" NAME="Username" VALUE="<%= steRecordEncValue(rsEdit, "UserName") %>" SIZE="22" MAXLENGTH="32" class="form">
		</TD>
	</TR><TR>
		<TD class="forml"><% steTxt "Password" %> (req)<BR>
		<INPUT TYPE="password" NAME="Password" VALUE="" SIZE="22" MAXLENGTH="32" class="form">
		</TD>
		<TD class="forml"><% steTxt "Confirm Password" %> (req)<BR>
		<INPUT TYPE="password" NAME="Confirm" VALUE="" SIZE="22" MAXLENGTH="32" class="form">
		</TD>
	</TR>
	<TR>
		<TD class="forml" COLSPAN=2><% steTxt "Address (Line 1)" %><BR>
		<INPUT TYPE="text" NAME="Address1" VALUE="<%= steRecordEncValue(rsEdit, "Address1") %>" SIZE="52" MAXLENGTH="40" class="form">
		</TD>
	</TR>
	<TR>
		<TD class="forml" COLSPAN=2><% steTxt "Address (Line 2)" %><BR>
		<INPUT TYPE="text" NAME="Address2" VALUE="<%= steRecordEncValue(rsEdit, "Address2") %>" SIZE="52" MAXLENGTH="40" class="form">
		</TD>
	</TR>
		<TR>
		<TD COLSPAN=2>
		<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 WIDTH="100%">
		<TR>
			<TD class="forml"><% steTxt "City" %><BR>
			<INPUT TYPE="text" NAME="City" VALUE="<%= steRecordEncValue(rsEdit, "City") %>" SIZE="14" MAXLENGTH="32" class="form">
			</TD>
			<TD class="forml"><% steTxt "State" %><BR>
			<SELECT NAME="StateCode" class="form">
			<OPTION VALUE=""> -- Choose --
			<% Do Until rsState.EOF %>
			<OPTION VALUE="<%= rsState.Fields("StateCode").Value %>"<% If steRecordValue(rsEdit, "StateCode") = rsState.Fields("StateCode").Value Then Response.Write(" SELECTED") %>> <%= rsState.Fields("StateName").Value %>
			<%	rsState.MoveNext
			   Loop %>
			</SELECT>
			</TD>
			<TD class="forml"><% steTxt "Zip Code" %><BR>
			<INPUT TYPE="text" NAME="ZipCode" VALUE="<%= steRecordEncValue(rsEdit, "ZipCode") %>" SIZE="10" MAXLENGTH="10" class="form">
			</TD>
		</TR>
		</TABLE>
		</TD>
	</TR>
	<TR>
		<TD class="forml" COLSPAN=2><% steTxt "Country" %><BR>
		<SELECT NAME="CountryID" class="form">
		<OPTION VALUE="0"> -- Choose --
		<% Do Until rsCountry.EOF %>
		<OPTION VALUE="<%= rsCountry.Fields("CountryID").Value %>"<% If steRecordEncValue(rsEdit, "CountryID") = CStr(rsCountry.Fields("CountryID").Value) Then Response.Write(" SELECTED") %>> <%= rsCountry.Fields("CountryName").Value %>
		<%	rsCountry.MoveNext
		   Loop %>
		</SELECT>
		</TD>
	</TR>
	<TR>
		<TD class="forml" COLSPAN=2><% steTxt "E-Mail Address" %> (req)<BR>
		<INPUT TYPE="text" NAME="EmailAddress" VALUE="<%= steRecordEncValue(rsEdit, "EmailAddress") %>" SIZE="52" MAXLENGTH="80" class="form">
		</TD>
	</TR>
	<TR>
		<TD COLSPAN=2 ALIGN="center"><br>
		<INPUT TYPE="hidden" NAME="Action" VALUE="edit">
		<INPUT TYPE="submit" NAME="_dummy" VALUE=" <% steTxt "Update Member" %> " class="form">
		</TD>
	</TR>
	</TABLE>
	</P>

	</FORM>

	<P>
	<FONT CLASS="small"><% steTxt "REQ - Indicates that a form field is required" %></FONT>
	</P>
<% Else %>
	<H3><% steTxt "Member Updated" %></H3>

	<P>
	<% steTxt "The member account was successfully updated in the database." %>
	<% steTxt "The changes will take effect immediately." %>
	<% steTxt "Please use the admin menu at the top of the screen to continue administering the site." %>
	</P>
<% End If %>

<p align="center">
	<a href="member_list.asp?pageno=<%= steNForm("pageno") %>" class="adminlink"><% steTxt "Member List" %></a>
</p>

<!-- #include file="../../../footer.asp"-->