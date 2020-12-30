<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' lang_delete.asp
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

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If steNForm("confirm") = 0	Then
		sErrorMsg = steGetText("Please confirm the deletion of this language")
	Else
		' create the author in the database
		sStat = "DELETE FROM tblLang WHERE LangCode = " & steQForm("LangCode")
		Call adoExecute(sStat)
	End If
End If

' retrieve the language to delete
sStat = "SELECT	lang.*, usr.Username " &_
		"FROM	tblLang lang " &_
		"INNER JOIN	tblUser usr ON usr.UserID = lang.UserID " &_
		"WHERE LangCode = '" & Replace(sLangCode, "'", "''") & "'"
Set rsLanguage = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Language" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Language" %></H3>

<P>
<% steTxt "Please confirm that you would like to delete the language shown below." %>&nbsp;
<% steTxt "Once a language has been deleted, it will be lost forever." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="lang_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="LangCode" VALUE="<%= sLangCode %>">
<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" nowrap><% steTxt "Language Code" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsLanguage, "LangCode") %></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Country Name" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsLanguage, "CountryName") %></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Native Language" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsLanguage, "NativeLanguage") %></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Flag Icon" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsLanguage, "FlagIcon") %></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Percent Complete" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordValue(rsLanguage, "PctComplete") %></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Published" %></TD><TD></TD>
	<TD CLASS="formd"><% If steRecordBoolValue(rsLanguage, "Published") Then Response.Write "Yes" Else Response.Write "No" %></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Maintained by User" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsLanguage, "Username") %></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Confirm Delete?" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="Confirm" VALUE="1"<% If steNForm("Confirm") = 1 Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="Confirm" VALUE="0"<% If steNForm("Confirm") = 0 Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 align="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Language" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Language Deleted" %></H3>

<P>
<% steTxt "The language was successfully deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<!-- #include file="../../../../footer.asp" -->
