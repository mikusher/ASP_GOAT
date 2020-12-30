<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' lang_edit.asp
'	Edit an existing language from the database
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

Dim sErrorMsg
Dim sStat
Dim rsLanguage
Dim sLangCode
Dim nUserID

sLangCode = steForm("LangCode")

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("langcode")) = ""	Then
		sErrorMsg = steGetText("Please enter the Subject for this language")
	ElseIf Trim(steForm("countryname")) = "" Then
		sErrorMsg = steGetText("Please enter the Country Name for this language")
	ElseIf Trim(steForm("nativelanguage")) = "" Then
		sErrorMsg = steGetText("Please enter the Native Language for this language")
	ElseIf Trim(steForm("PctComplete")) = "" Or Not IsNumeric(steForm("PctComplete")) Then
		sErrorMsg = steGetText("Please enter a numeric value for Percent Complete")
	Else
		' determine the user (for username)
		nUserID = 0
		If steForm("Username") <> "" Then
			Dim rsUser

			sStat = "SELECT	UserID FROM tblUser WHERE Username = " & steQForm("Username") & " AND Archive = 0"
			Set rsUser = adoOpenRecordset(sStat)
			If Not rsUser.EOF Then nUserID = rsUser.Fields("UserID").Value
			rsUser.Close
			Set rsUser = Nothing
		End If

		' create the author in the database
		If nUserID > 0 Then
			sStat = "UPDATE tblLang " &_
					"SET	LangCode = " & steQForm("LangCode") & "," &_
					"		CountryName = " & steQForm("CountryName") & "," &_
					"		NativeLanguage = " & steQForm("NativeLanguage") & "," &_
					"		FlagIcon = " & steQForm("FlagIcon") & "," &_
					"		UserID = " & nUserID & "," &_
					"		PctComplete = " & steFForm("PctComplete") & "," &_
					"		Published = " & steNForm("Published") & " " &_
					"WHERE	LangCode = " & steQForm("OldLangCode")
			Call adoExecute(sStat)
		Else
			sErrorMsg = steGetText("Please enter a valid Maintained by User for this language")
		End If
	End If
End If

' retrieve the language to edit
sStat = "SELECT	lang.*, u.Username " &_
		"FROM	tblLang lang " &_
		"LEFT JOIN	tblUser u on u.UserID = lang.UserID " &_
		"WHERE lang.LangCode = '" & Replace(sLangCode, "'", "''") & "'"
Set rsLanguage = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Language" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Language" %></H3>

<P>
<% steTxt "Please make your changes to the language using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="lang_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="OldLangCode" VALUE="<%= sLangCode %>">
<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" nowrap><% steTxt "Language Code" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="LangCode" VALUE="<%= steRecordEncValue(rsLanguage, "LangCode") %>" SIZE="8" MAXLENGTH="4" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Country Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="CountryName" VALUE="<%= steRecordEncValue(rsLanguage, "CountryName") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Native Language" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="NativeLanguage" VALUE="<%= steRecordEncValue(rsLanguage, "NativeLanguage") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Flag Icon" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="FlagIcon" VALUE="<%= steRecordEncValue(rsLanguage, "FlagIcon") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Percent Complete" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="PctComplete" VALUE="<%= steRecordEncValue(rsLanguage, "PctComplete") %>" SIZE="8" MAXLENGTH="7" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Published" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="Published" VALUE="1"<% If steRecordBoolValue(rsLanguage, "Published") Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="Published" VALUE="0"<% If Not steRecordBoolValue(rsLanguage, "Published") Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Maintained by User" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Username" VALUE="<%= steRecordEncValue(rsLanguage, "Username") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 align="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt("Update Language") %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Language Updated" %></H3>

<P>
<% steTxt "The language was successfully updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<!-- #include file="../../../../footer.asp" -->
