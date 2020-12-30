<%
'--------------------------------------------------------------------
' lang_lib.asp
'	This library is used to manage multiple language support.  It
'	used to be part of site_lib.asp (and is included via SSI in
'	that file) but has been separated for pages that require
'	language support but not all of site_lib.asp.  An example of
'	this are the module capsules.)
'
' AUTH:	Ken Richards
' DATE:	07/25/2001
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

Dim steLanguage			' language (2-letter country code like US or CA)

'------------------------------------------------------------------
' Retrieve a translated piece of text from the database / app.cache

Function steGetText(sEnglishText)
	Dim sStat, rsTran

	If steLanguage = "" Then
		If Request.Cookies("LANGUAGE") <> "" Then
			steLanguage = Request.Cookies("LANGUAGE")
		Else
			steLanguage = "US"
		End If
	End If
	' regurgitate the english (US) language as needed
	If steLanguage = "US" Then
		steGetText = sEnglishText
		Exit Function
	End If
	' rebuild the cache for the language translation (if nec)
	If Not Application("LANG" & steLanguage) Then
		' rebuild the cache for this language
		sStat = "SELECT	txt.EnglishText, trn.Translation " &_
				"FROM	tblLangText txt " &_
				"INNER JOIN	tblLangTranslation trn ON txt.TextID = trn.TextID " &_
				"WHERE	trn.LangCode = '" & steLanguage & "' " &_
				"AND	txt.Archive = 0 " &_
				"AND	trn.Archive = 0"
		Set rsTran = adoOpenRecordset(sStat)
		Do Until rsTran.EOF
			Application("LANG" & steLanguage & rsTran.Fields("EnglishText").Value) = rsTran.Fields("Translation").Value
			rsTran.MoveNext
		Loop
		rsTran.Close
		Set rsTran = Nothing
		Application("LANG" & steLanguage) = True
	End If
	' write the translation (or the english text if none exists)
	If Application("LANG" & steLanguage & sEnglishText) <> "" Then
		steGetText = Application("LANG" & steLanguage & sEnglishText)
	Else
		steGetText = sEnglishText
	End If
End Function

'------------------------------------------------------------------
' Retrieve a translated piece of text from the database / app.cache

Sub steTxt(sEnglishText)
	Response.Write steGetText(sEnglishText)
End Sub
%>