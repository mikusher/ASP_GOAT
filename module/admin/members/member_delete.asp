<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' member_delete.asp
'	Delete an existing member in the database.
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

If sAction = "delete" Then
	' check for required fields
	If steNForm("Confirm") <> 1 Then
		sErrorMsg = sErrorMsg & steGetText("You must confirm the deletion of this member") & "<BR>"
	End If
	If sErrorMsg = "" Then
		' check to see if the username is already in use
		query = "DELETE " &_
				"FROM	tblMember " &_
				"WHERE	MemberID = " & steNForm("memberid")
		Call adoExecute(query)

		' TODO - delete all associated records (comments, etc)
	End If
End If

' retrieve the member record to edit here
If steNForm("memberid") <> 0 Then
	query = "SELECT	* FROM	tblMember WHERE MemberID = " & steForm("memberid")
	Set rsEdit = adoOpenRecordset(query)
End If

Dim sStateName, sCountryName
If Not rsEdit.EOF Then
	' build the selection list for the states
	query = "SELECT StateCode, StateName " &_
			"FROM	tblState " &_
			"WHERE	StateCode = '" & rsEdit.Fields("StateCode").Value & "' " &_
			"AND	Active <> 0 " &_
			"AND	Archive = 0 " &_
			"ORDER BY StateName"
	Set rsState = adoOpenRecordset(query)
	If Not rsState.EOF Then sStateName = rsState.Fields("StateName").Value
	rsState.Close
	rsState = Empty

	' build the selection ist for the countries
	If Not IsNull(rsEdit.Fields("CountryID").Value) Then
		query = "SELECT CountryID, CountryName " &_
				"FROM	tblCountry " &_
				"WHERE	CountryID = " & rsEdit.Fields("CountryID").Value & " " &_
				"AND	Active <> 0 " &_
				"AND	Archive = 0 " &_
				"ORDER BY CountryName"
		Set rsCountry = adoOpenRecordset(query)
		If Not rsCountry.EOF Then sCountryName = rsCountry.Fields("CountryName").Value
		rsCountry.Close
		rsCountry = Empty
	End If
End If
%>
<!-- #include file="../../../header.asp"-->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Members" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If sAction <> "delete" Or sErrorMsg <> "" Then %>
	<H3><% steTxt "Delete Member Registration" %></H3>

	<P>
	<% steTxt "Please confirm that you would like to permanently delete the member shown below." %>
	<% steTxt "Once this action has been completed, there will be no way to recover the member information." %>
	</P>

	<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
	<% End If %>

	<FORM METHOD="post" ACTION="member_delete.asp">
	<INPUT TYPE="hidden" NAME="memberid" VALUE="<%= steEncForm("memberid") %>">

	<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
	<TR>
		<TD class="forml"><% steTxt "First Name" %><BR>
		<font class="formd"><%= steRecordEncValue(rsEdit, "FirstName") %><font>
		</TD>
		<TD class="forml"><% steTxt "Last Name" %><BR>
		<font class="formd"><%= steRecordEncValue(rsEdit, "LastName") %><font>
		</TD>
	</TR>
	<TR>
		<TD class="forml"><% steTxt "Username" %><BR>
		<font class="formd"><%= steRecordEncValue(rsEdit, "UserName") %></font>
		</TD>
	</TR><TR>
		<TD class="forml"><% steTxt "Password" %><BR>
		<font class="formd"><i>n/a</i></font>
		</TD>
		<TD class="forml"><% steTxt "Confirm Password" %><BR>
		<font class="formd"><i>n/a</i></font>
		</TD>
	</TR>
	<TR>
		<TD class="forml" COLSPAN=2><% steTxt "Address (Line 1)" %><BR>
		<font class="formd"><%= steRecordEncValue(rsEdit, "Address1") %>&nbsp;<font>
		</TD>
	</TR>
	<TR>
		<TD class="forml" COLSPAN=2><% steTxt "Address (Line 2)" %><BR>
		<font class="formd"><%= steRecordEncValue(rsEdit, "Address2") %>&nbsp;</font>
		</TD>
	</TR>
		<TR>
		<TD COLSPAN=2>
		<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 WIDTH="100%">
		<TR>
			<TD class="forml"><% steTxt "City" %><BR>
			<font class="formd"><%= steRecordEncValue(rsEdit, "City") %>&nbsp;</font>
			</TD>
			<TD class="forml"><% steTxt "State" %><BR>
			<font class="formd"><%= sStateName %>&nbsp;</font>
			</TD>
			<TD class="forml"><% steTxt "Zip Code" %><BR>
			<font class="formd"><%= steRecordEncValue(rsEdit, "ZipCode") %>&nbsp;<font>
			</TD>
		</TR>
		</TABLE>
		</TD>
	</TR>
	<TR>
		<TD class="forml" COLSPAN=2><% steTxt "Country" %><BR>
		<font class="formd"><%= sCountryName %></font>
		</TD>
	</TR>
	<TR>
		<TD class="forml" COLSPAN=2><% steTxt "E-Mail Address" %> (req)<BR>
		<font class="formd"><%= steRecordEncValue(rsEdit, "EmailAddress") %></font>
		</TD>
	</TR>
	<TR>
		<TD CLASS="forml" VALIGN="top"><B><% steTxt "Confirm Delete" %></B></TD><TD></TD>
		<TD><INPUT TYPE="radio" NAME="confirm" VALUE="1" class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="confirm" VALUE="0" class="formradio"> <% steTxt "No" %>
		</TD>
	</TR>
	<TR>
		<TD COLSPAN=2 ALIGN="right"><BR>
		<INPUT TYPE="hidden" NAME="Action" VALUE="delete">
		<INPUT TYPE="submit" NAME="_dummy" VALUE=" <% steTxt "Delete Member" %> " class="form">
		</TD>
	</TR>
	</TABLE>
	</P>

	</FORM>
<% Else %>
	<H3><% steTxt "Member Deleted" %></H3>

	<P>
	<% steTxt "The member account was successfully deleted from the database." %>
	<% steTxt "Please use the admin menu at the top of the screen to continue administering the site." %>
	</P>
<% End If %>

<p align="center">
	<a href="member_list.asp?pageno=<%= steNForm("pageno") %>" class="adminlink"><% steTxt "Member List" %></a>
</p>

<!-- #include file="../../../footer.asp"-->