<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/class/adminlist.asp" -->
<%
'--------------------------------------------------------------------
' tran_list.asp
'	Display a list of all of the english text words/phrases that
'	require translation.
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

Dim rsLang
Dim sLangCode

sLangCode = steForm("langcode")

' build the drop-list for the languages
sStat = "select	LangCode, NativeLanguage, CountryName " &_
		"from	tblLang " &_
		"where	Archive = 0 " &_
		"order by NativeLanguage"
Set rsLang = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Translation" %>
<!-- #include file="pagetabs_inc.asp" -->

<script language="Javascript" type="text/javascript">
<!-- // hide
function pickLang(sLangCode) {
	if (sLangCode != '') {
		location.href='tran_list.asp?langcode=' + sLangCode;
	}
}
// unhide -->
</script>

<h3><% steTxt "Translation List" %></h3>

<p>
<table border=0 cellpadding=2 cellspacing=0>
<tr>
	<td class="forml"><% steTxt "Language Shown" %></td><td>&nbsp;&nbsp;</td>
	<td>
	<select name="langcode" onChange="pickLang(this.options[this.selectedIndex].value)">
	<option value=""> -- <% steTxt "Choose" %> --
	<%	Do Until rsLang.EOF %>
	<option value="<%= rsLang.Fields("LangCode").Value %>"<% If sLangCode = rsLang.Fields("LangCode").Value Then Response.Write " SELECTED" %>> <%= rsLang.Fields("NativeLanguage").Value & " (" & rsLang.Fields("CountryName").Value & ")" %>
	<%		rsLang.MoveNext
		Loop
		rsLang.Close
		Set rsLang = Nothing %>
	</select>
	</td>
</tr>
</table>
</p>

<%
If sLangCode <> "" Then
	Dim oList
	Set oList = New clsAdminList
	oList.Query = "SELECT	t2.TranslationID, t1.TextID, t1.EnglishText, t2.Translation, t2.Modified " &_
				"FROM	tblLangText t1 " &_
				"LEFT JOIN	tblLangTranslation t2 ON t2.TextID = t1.TextID " &_
				"AND	t2.LangCode = '" & Replace(sLangCode, "'", "''") & "' " &_
				"WHERE	t1.Archive = 0 " &_
				"AND	t2.Archive = 0 " &_
				"ORDER BY t1.EnglishText"
	Call oList.AddColumn("EnglishText", steGetText("English Text"), "")
	Call oList.AddColumn("Translation", steGetText("Translation"), "")
	Call oList.AddColumn("Modified", steGetText("Modified"), "right")
	oList.ActionLink = "<a href=""tran_edit.asp?langcode=" & sLangCode &_
		"&textid=##textid##&translationid=##translationid##"" class=""actionlink"">" & steGetText("edit") & "</a> . <a href=""tran_delete.asp?langcode=" &_
		sLangCode & "&textid=##textid##&translationid=##translationid##"" class=""actionlink"">" & steGetText("delete") & "</a>"
	
	' show the list here
	Call oList.Display
Else
	Response.Write "<p><b class=""error"">" & steGetText("Please select a language to work with") & "</b></p>"
End If
%>

<!-- #include file="../../../../footer.asp" -->
