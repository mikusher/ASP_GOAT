<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' suggestion_delete.asp
'	Delete an existing suggestion from the database
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
Dim rsSuggest
Dim nSuggestionID

nSuggestionID = steNForm("SuggestionID")

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If steNForm("confirm") <> 1 Then
		sErrorMsg = steGetText("Please confirm the deletion of this suggestion")
	Else
		' delete the suggestion from the database
		sStat = "DELETE FROM tblSuggestion " &_
				"WHERE	SuggestionID = " & nSuggestionID
		Call adoExecute(sStat)
	End If
End If

If nSuggestionID > 0 Then
	sStat = "SELECT * FROM tblSuggestion " &_
			"WHERE SuggestionID = " & nSuggestionID
	Set rsSuggest = adoOpenRecordset(sStat)
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Suggestions" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Suggestion" %></H3>

<P>
<% steTxt "Please confirm the deletion of the suggestion by clicking the yes button next to <I>Confirm</I>." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="suggestion_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="SuggestionID" VALUE="<%= nSuggestionID %>">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD CLASS="forml" nowrap><% steTxt "From Name" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsSuggest, "FromName") %></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "From E-Mail" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsSuggest, "FromEmail") %></TD>
</TR><TR>
	<TD CLASS="forml"nowrap><% steTxt "Subject" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsSuggest, "Subject") %></TD>
</TR><TR>
	<TD CLASS="forml" valign="top" nowrap><% steTxt "Body" %></TD><TD></TD>
	<TD CLASS="formd"><%= Replace(steRecordEncValue(rsSuggest, "Body"), vbCrLf, "<BR>") %></TD>
</TR><TR>
	<TD CLASS="forml" nowrap>Confirm Delete</TD><TD></TD>
	<TD><INPUT TYPE="radio" NAME="confirm" VALUE="1" class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="confirm" VALUE="0" checked class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 align=right><br>
		<INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Suggestion" %> " class="form">
	</TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Suggestion Deleted" %></H3>

<P>
<% steTxt "The suggestion was successfully deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align=center>
	<a href="suggestion_list.asp" class="adminlink"><% steTxt "Suggestion List" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->
