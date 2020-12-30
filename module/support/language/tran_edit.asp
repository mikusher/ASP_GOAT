<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' tran_edit.asp
'	Edit a translation for the specified language / text phrase
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
Dim rsLang
Dim rsTran
Dim sLangCode
Dim sLanguage
Dim nTotalRecords
Dim nTextID

sLangCode = steForm("langcode")
nTextID = steNForm("TextID")
nTotalRecords = steNForm("TotalRecords")

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("Translation")) = "" Then
		sErrorMsg = steGetText("Please enter the Translation for the English text")
	Else
		sStat = "UPDATE tblLangTranslation SET " &_
				"	Translation = " & steQForm("Translation") & ", " &_
				"	MemberID = " & Request.Cookies("MemberID") & ", " &_
				"	Modified = " & adoGetDate & " " &_
				"WHERE	LangCode = " & steQForm("langcode") & " " &_
				"AND	TextID = " & nTextID
		Call adoExecute(sStat)

		' redirect to the language administration
		Response.Redirect "tran_list.asp?translated=1&langcode=" & steEncForm("LangCode") &_
			"&pageno=" & steNForm("pageno") & "&totalrecords=" & CStr(nTotalRecords - 1) &_
			"&statusmsg=" & Server.URLEncode(steGetText("Translation updated - Thank you for your contribution"))
	End If
End If

' retrieve the language name
sStat = "SELECT	NativeLanguage, CountryName " &_
		"FROM	tblLang " &_
		"WHERE	LangCode = " & steQForm("langcode")
Set rsLang = adoOpenRecordset(sStat)
If Not rsLang.EOF Then
	sLanguage = rsLang.Fields("NativeLanguage").Value & " (" & rsLang.Fields("CountryName").Value & ")"
Else
	sLanguage = "<I>*Unknown*</I>"
End If
rsLang.Close
Set rsLang = Nothing

' get the language translations to update
sStat = "SELECT	txt.TextID, txt.EnglishText, trn.TranslationID, trn.Translation " &_
		"FROM	tblLangText txt " &_
		"LEFT JOIN tblLangTranslation trn on trn.TextID = txt.TextID " &_
		"	AND	trn.LangCode = " & steQForm("langcode") & " " &_
		"WHERE	txt.TextID = " & nTextID & " " &_
		"AND	txt.Archive = 0"
Set rsTran = adoOpenRecordset(sStat)
If Not rsTran.EOF Then
	If IsNull(rsTran.Fields("TranslationID").Value) Then
		Response.Redirect "tran_add.asp?langcode=" & sLangCode & "&textid=" & nTextID &_
			"&pageno=" & steNForm("pageno") & "&totalrecords=" & nTotalRecords &_
			"&error=" & Server.URLEncode("Translation was not found for the selected text")
	End if
End If

%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../account/register/login_lib.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<h3><%= sLanguage %> - <% steTxt "Edit Translation" %></h3>

<p>
<% steTxt "Please modify the language translation for the English text below." %>
</p>

<% If sErrormsg <> "" Then %>
<P><b class="error"><%= sErrorMsg %></b></p>
<% End If %>

<form method="post" action="tran_edit.asp">
<input type="hidden" name="action" value="edit">
<input type="hidden" name="langcode" value="<%= sLangCode %>">
<input type="hidden" name="textid" value="<%= nTextID %>">
<input type="hidden" name="pageno" value="<%= steNForm("pageno") %>">
<input type="hidden" name="totalrecords" value="<%= nTotalRecords %>">

<table border="0" cellpadding="2" cellspacing="0">
<tr>
	<td class="forml"><% steTxt "English Text" %></td><td>&nbsp;&nbsp;</td>
	<td class="formd"><%= steRecordEncValue(rsTran, "EnglishText") %></td>
</tr><tr>
	<td class="forml"><% steTxt "Translation" %></td><td></td>
	<td><input type="Text" name="translation" value="<%= Replace(Replace(steRecordValue(rsTran, "Translation")&"", "<", "&lt;"), ">", "&gt;") %>" size="32" maxlength="255" class="form"></td>
</tr><tr>
	<td colspan="3" align="Right"><br>
		<input type="submit" name="_submit" value=" <% steTxt "Update Translation" %> ">
	</td>
</tr>
</table>
</form>

<% Else %>

<h3><%= sLanguage %> - <% steTxt "Translation Updated" %></h3>

<p>
<% steTxt "Thank you for submitting your language translation to the ASP Nuke project." %>
</p>

<% End If %>

<p align="center">
	<a href="tran_list.asp?langcode=<%= sLangCode %>&pageno=<%= steNForm("pageno") %>" class="footerlink"><% steTxt "Back to Translation List" %></a>
</p>

<!-- #include file="../../../footer.asp" -->
