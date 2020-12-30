<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' suggestion_edit.asp
'	Edit an existing suggestion from the database
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
Dim rsSuggestion
Dim nSuggestionID

nSuggestionID = steNForm("SuggestionID")

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("subject")) = ""	Then
		sErrorMsg = steGetText("Please enter the subject for this suggestion")
	ElseIf Trim(steForm("body")) = "" Then
		sErrorMsg = steGetText("Please enter the body for this suggestion")
	Else
		' create the author in the database
		sStat = "UPDATE tblSuggestion " &_
				"SET	FromName = " & steQForm("FromName") & "," &_
				"		FromEmail = " & steQForm("FromEmail") & "," &_
				"		Subject = " & steQForm("Subject") & "," &_
				"		Body = " & steQForm("Body") & " " &_
				"WHERE	SuggestionID = " & nSuggestionID
		Call adoExecute(sStat)
	End If
End If

' retrieve the suggestion to edit
sStat = "SELECT	* FROM tblSuggestion WHERE SuggestionID = " & nSuggestionID
Set rsSuggestion = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Suggestions" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Suggestion" %></H3>

<P>
<% steTxt "Please make your changes to the suggestion using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="suggestion_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="SuggestionID" VALUE="<%= nSuggestionID %>">
<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" nowrap><% steTxt "From Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="FromName" VALUE="<%= steRecordEncValue(rsSuggestion, "FromName") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "From E-Mail" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="FromEmail" VALUE="<%= steRecordEncValue(rsSuggestion, "FromEmail") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Subject" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Subject" VALUE="<%= steRecordEncValue(rsSuggestion, "Subject") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top" nowrap><% steTxt "Body" %></TD><TD></TD>
	<TD><textarea NAME="Body" cols="80" rows="10" class="form" style="width:440px"><%= steRecordEncValue(rsSuggestion, "Body") %></textarea></TD>
</TR><TR>
	<TD COLSPAN=3 align="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update Suggestion" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Suggestion Updated" %></H3>

<P>
<% steTxt "The suggestion was successfully updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align="center">
	<a href="suggestion_list.asp" class="adminlink">Suggestion List</a>
</p>

<!-- #include file="../../../../footer.asp" -->
