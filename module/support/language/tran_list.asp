<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' tran_list.asp
'	List all of the words / phrases needed for translation
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
Dim aTran
Dim nPageNo
Dim nTotalRecords
Dim nTotalPages
Dim rsLang
Dim sLanguage
Dim sLangCode
Dim nTotal			' total language elements
Dim nTranslated		' total translated elements
Dim nUntranslated	' total untranslated elements

Const DEF_PAGESIZE = 25

sLangCode = steForm("langcode")
nPageNo = steNForm("pageno")
nTotalRecords = steNForm("TotalRecords")
nTotalPages = steNForm("TotalPages")

If steForm("action") = "go" Then
	' make sure the required fields are present
	If Trim(steForm("LangCode")) = "" Then
		sErrorMsg = steGetText("Please select a language from the list below")
	Else
		' redirect to the language administration
		Response.Redirect "language_select.asp?langcode=" & steEncForm("LangCode") &_
			"&error=" & steGetText("Language code missing in tran_list.asp")
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
	Response.Redirect "language_select.asp?error=" & Server.URLEncode("Please choose a language to work with from the drop-list")
End If
rsLang.Close
Set rsLang = Nothing

' get the list of language translations for this language
sStat = "SELECT	txt.TextID, txt.EnglishText, trn.TranslationID, trn.Translation " &_
		"FROM	tblLangText txt " &_
		"LEFT JOIN tblLangTranslation trn on trn.TextID = txt.TextID " &_
		"	AND	trn.LangCode = " & steQForm("langcode") & " " &_
		"WHERE	txt.Archive = 0 "
If steNForm("Translated") = 0 Then
	sStat = sStat & "AND	trn.TranslationID IS NULL "
Else
	sStat = sStat & "AND	trn.TranslationID IS NOT NULL "
End If
sStat = sStat & "ORDER BY txt.EnglishText"
Set rsTran = adoOpenRecordset(sStat)
If Not rsTran.EOF Then
	' count the total number of results (for page nav)
	If nTotalRecords = 0 Then
		Do Until rsTran.EOF
			nTotalRecords = nTotalRecords + 1
			rsTran.MoveNext
		Loop
		rsTran.MoveFirst
		If steNForm("translated") = 0 Then
			Application("TRANSUNTRANSLATED" & sLangCode) = nTotalRecords
		Else
			Application("TRANSTRANSLATED" & sLangCode) = nTotalRecords
		End If
		
	End If
	If nTotalRecords > 0 Then nTotalPages = ((nTotalRecords - 1) \ DEF_PAGESIZE + 1)
	If nPageNo > nTotalPages - 1 Then nPageNo = nTotalPages - 1
	If nPageNo < 0 Then nPageNo = 0
	' grab the page of results here
	If nPageNo > 0 Then rsTran.Move nPageNo * DEF_PAGESIZE
	aTran = rsTran.GetRows(DEF_PAGESIZE)
End If

' get the total counts here
locTranslateCount nTotal, nTranslated, nUntranslated
%>
<!-- #include file="../../../header.asp" -->

<table border=0 cellpadding=4 cellspacing=0 width="100%">
<tr>
	<td><h3><%= sLanguage %> - <% If steNForm("translated") = 0 Then Response.Write steGetText("Untranslated Text") Else Response.Write steGetText("Translated Text") %></h3></td>
	<td align="right" valign="top">
	<% If steNForm("translated") = 0 Then %>
	<input type="button" name="_switch1" value="View <%= nTranslated %> Translated" onclick="location.href='tran_list.asp?langcode=<%= steForm("langcode") %>&translated=1'" class="form">
	<% Else %>
	<input type="button" name="_switch2" value="View <%= nUntranslated %> Untranslated" onclick="location.href='tran_list.asp?langcode=<%= steForm("langcode") %>&translated=0'" class="form">
	<% End If %>
	</td>
</tr>
</table>

<% If steForm("statusmsg") <> "" Then %>
<p align="center"><b class="error"><%= steForm("statusmsg") %></b></p>
<% End If %>

<% locPageNav nPageNo, nTotalRecords %>

<table border="0" cellpadding="2" cellspacing="0" class="list">
<tr>
	<td class="listhead"><% steTxt "English Text / Translation" %></td>
	<!-- td class="listhead">Action</td -->
</tr>
<%
If IsArray(aTran) Then
	For I = 0 To UBound(aTran, 2) %>
<tr class="list<%= I mod 2 %>">
	<td><a href="<%	If IsNull(aTran(2, I)) Then Response.Write "tran_add.asp" Else Response.Write "tran_edit.asp" %>?translated=<%= steNForm("translated") %>&langcode=<%= sLangCode %>&textid=<%= aTran(0, I) %>&pageno=<%= nPageNo %>&totalrecords=<%= nTotalRecords %>"><%= aTran(1, I) %><br><% If IsNull(aTran(3, I)) Then %>&nbsp;<% Else %><%= Replace(Replace(aTran(3, I), "<", "&lt;"), ">", "&gt;") %><% End If %></a></td>
</tr>
<%	Next
End If
%>
</table>

<p>
<% steTxt "The list above contains the translations for the selected language." %>&nbsp;
<% If steNForm("translated") = 0 Then %>
<% steTxt "Please click the text to enter a translation." %>&nbsp;
<% steTxt "You will be asked to login with your member account so that you receive credit for your translation work." %>&nbsp;
<% steTxt "Please use the" %>&nbsp;
<a href="http://www.aspnuke.com/suggestions.asp"><% steTxt "Contact Us" %></a>&nbsp;
<% steTxt "if you encounter any difficulties." %>

<% Else %>
<% steTxt "Please verify the existing translations are correct and edit any incorrect translations by clicking on the text." %>
<% End If %>
</p>

<!-- #include file="../../../footer.asp" -->
<%
Sub locPageNav(nPageNo, nTotalRecords)
	Dim I, nTotalPages, nLast

	nTotalPages = ((nTotalRecords - 1) \ DEF_PAGESIZE + 1)
	nLast = (nPageNo + 1) * DEF_PAGESIZE
	If nLast > nTotalRecords Then nLast = nTotalRecords

	With Response
	.Write "<p align=""center""><b>Displaying Records "
	.Write (nPageNo * DEF_PAGESIZE + 1)
	.Write " to "
	.Write nLast
	.Write " of "
	.Write nTotalRecords
	.Write "</b><br>"
	.Write "<I>Page No:</I> "
	For I = 0 To (nTotalPages - 1)
		If I > 0 Then .Write " | "
		If nPageNo = I Then
			.Write "<b style=""color:red"">" & (I + 1) & "</b>"
		Else
			.Write "<a href=""tran_list.asp?translated="
			.Write steNForm("translated")
			.Write "&langcode="
			.Write steForm("langcode")
			.Write "&pageno=" 
			.Write I
			.Write "&totalrecords="
			.Write nTotalRecords
			.Write """>"
			.Write (I + 1) 
			.Write "</a>"
		End If
	Next
	.Write "</p>" & vbCrLf
	End With
End Sub

' retrieve the total count of items / translated / untranslated
Sub locTranslateCount(nTotal, nTranslated, nUntranslated)
	Dim sStat, rsTran, bUpdate

	bUpdate = False
	If IsDate(Application("TRANSUPDATED" & sLangCode)) Then
		If DateDiff("h", Application("TRANSUPDATED" & sLangCode), Now()) > 24 Then bUpdate = True
	End If

	If Application("TRANSTOTAL" & sLangCode) = "" Or bUpdate Then
		' retrieve the total count from the database
		sStat = "SELECT	Count(*) AS TotalCount " &_
				"FROM	tblLangText txt " &_
				"LEFT JOIN tblLangTranslation trn on trn.TextID = txt.TextID " &_
				"	AND	trn.LangCode = " & steQForm("langcode") & " " &_
				"WHERE	txt.Archive = 0"
		Set rsTran = adoOpenRecordset(sStat)
		If Not rsTran.EOF Then
			nTotal = rsTran.Fields("TotalCount").Value
		End If

		' retrieve the total count from the database
		sStat = "SELECT	Count(*) AS Untranslated " &_
				"FROM	tblLangText txt " &_
				"LEFT JOIN tblLangTranslation trn on trn.TextID = txt.TextID " &_
				"	AND	trn.LangCode = " & steQForm("langcode") & " " &_
				"WHERE	txt.Archive = 0 " &_
				"AND	trn.TranslationID IS NULL"
		Set rsTran = adoOpenRecordset(sStat)
		If Not rsTran.EOF Then
			nUntranslated = rsTran.Fields("Untranslated").Value
		End If

		nTranslated = nTotal - nUntranslated
		Application("TRANSTOTAL" & sLangCode) = nTotal
		Application("TRANSTRANSLATED" & sLangCode) = nTranslated
		Application("TRANSUNTRANSLATED" & sLangCode) = nUntranslated
		Application("TRANSUPDATED" & sLangCode) = Now()
	Else
		' pull the totals from the cache
		nTotal = Application("TRANSTOTAL" & sLangCode)
		nTranslated = Application("TRANSTRANSLATED" & sLangCode)
		nUntranslated = Application("TRANSUNTRANSLATED" & sLangCode)
	End If
End Sub
%>