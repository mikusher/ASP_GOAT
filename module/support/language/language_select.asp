<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' language_select.asp
'	Select a language to edit
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
Dim nPollID
Dim rsLang

nPollID = steNForm("pollid")
 
If steForm("action") = "go" Then
	' make sure the required fields are present
	If Trim(steForm("LangCode")) = "" Then
		sErrorMsg = steGetText("Please select a language from the list below")
	Else
		' redirect to the language administration
		Response.Redirect "tran_list.asp?langcode=" & steEncForm("LangCode")
	End If
End If

' get the language list to select from
sStat = "SELECT	LangCode, NativeLanguage, CountryName " &_
		"FROM	tblLang " &_
		"WHERE	Archive = 0 " &_
		"ORDER BY CountryName"
Set rsLang =adoOpenRecordset(sStat)

%>
<!-- #include file="../../../header.asp" -->

<H3><% steTxt "Language Translations" %></H3>

<P>
<% steTxt "To view or contribute language translations, please select a language from the list below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B class="error"><%= sErrorMsg %></B></P>
<% ElseIf steForm("error") <> "" Then %>
<p><b class="error"><%= steForm("error") %></b></p>
<% End If %>

<form method="post" action="language_select.asp">
<input type="hidden" name="action" value="go">

<P>
<table border="0" cellpadding="2" cellspacing="0">
<tr>
	<td class="forml">Language</td><td>&nbsp;&nbsp;</td>
	<td>
	<select name="LangCode" class="form">
	<option value=""> -- Choose --
	<% Do Until rsLang.EOF %>
	<option value="<%= rsLang.Fields("LangCode").Value %>"> <%= rsLang.Fields("NativeLanguage").Value %>&nbsp;(<%= rsLang.Fields("CountryName").Value %>)
	<%	rsLang.MoveNext
	   Loop
		rsLang.Close
		Set rsLang = Nothing %>
	</select>
	</td><td>&nbsp;&nbsp;</td>
	<td><input type="submit" name="_submit" value="GO" class="form"></td>
</tr>
</table>
</P>
</form>

<!-- #include file="../../../footer.asp" -->
