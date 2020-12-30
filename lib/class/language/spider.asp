<%
'--------------------------------------------------------------------
' spider.asp
'	A spider for finding labels in web pages (the source files)
'	That need to be translated for internationalization.
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

Class clsSpider
	Private objDict
	Private strExtList
	Private arrPattern()		' patterns for extracting phrases
	Private intPatterns			' total number of patterns defined
	Private strErrorMsg			' error message to report
	Private boolDebug			' output debug information

	Private FSO_FOR_READING

	Private Sub Class_Initialize
		Set objDict = Server.CreateObject("Scripting.Dictionary")
		strExtList = ".asp,.js,.htm,.html"
		' define the default patterns for matching text for translation
		ReDim arrPattern(10)
		arrPattern(0) = "steTxt\s+""([^""]+)"""
		arrPattern(1) = "steGetText\(""([^""]+)""\)"
		intPatterns = 2
		'arrPattern(0) = "<\s*td\s[^>]*class=""forml""\s?[^>]*>([^<])+</td>"
		'arrPattern(1) = "<input\s+type=""button""\s+[^>]*value=""([^""]+)""\s+[^>]*>"
		'arrPattern(2) = "<input\s+type=""submit""\s+[^>]*value=""([^""]+)""\s+[^>]*>"
		'arrPattern(3) = "<input\s+type=""radio""\s+[^>]*> ([^<]+)"
		'arrPattern(4) = "<h3[^>]*>([^<]+)</h3>"
		'arrPattern(5) = "<h4[^>]*>([^<]+)</h4>"
		'intPatterns = 6
		FSO_FOR_READING = 1
		boolDebug = False
	End Sub

	'--------------------------------------------------------------
	' Strip out HTML comments from contents of a file

	Private Function StripHTMLComments(sContents)
		Dim oRE
		Set oRE = New RegExp
		oRE.IgnoreCase = True
		oRE.Global = True
		oRE.Pattern = "<!--[\s\S]*-->"
		sContents = oRE.Replace(sContents, "")
	End Function

	'--------------------------------------------------------------
	' Strip out script blocks from contents of a file

	Private Function StripScriptBlocks(sContents)
		Dim oRE
		Set oRE = New RegExp
		oRE.IgnoreCase = True
		oRE.Global = True
		oRE.Pattern = "<script(\s[^>]*)?>[\s\S]*</script>"
		sContents = oRE.Replace(sContents, "")
	End Function

	'--------------------------------------------------------------
	' Strip out the HTML head from contents of a file

	Private Function StripHTMLHead(sContents)
		Dim oRE
		Set oRE = New RegExp
		oRE.IgnoreCase = True
		oRE.Global = True
		oRE.Pattern = "<head(\s[^>]*)?>[\s\S]*</head>"
		sContents = oRE.Replace(sContents, "")
	End Function

	'--------------------------------------------------------------
	' Strip out the HTML includes from contents of a file

	Private Function StripIncludes(sContents)
		Dim oRE
		Set oRE = New RegExp
		oRE.IgnoreCase = True
		oRE.Global = True
		oRE.Pattern = "<include\s[^>]*>"
		sContents = oRE.Replace(sContents, "")
	End Function

	'--------------------------------------------------------------
	' Read all of the contents of a file

	Private Function GetFile(sFolder, sFile)
		Dim oFSO, oFile
		Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
		On Error Resume Next
		Set oFile = oFSO.OpenTextFile(Server.MapPath(oFSO.BuildPath(sFolder, sFile)), FSO_FOR_READING)
		If Err.Number <> 0 Then
			Response.Write "Unable to open file: " & Server.MapPath(oFSO.BuildPath(sFolder, sFile)) & "<br>"
			Response.Write Err.Description & " (" & Err.Number & ")<br>"
			Response.End
		End If
		GetFile = oFile.ReadAll
		On Error Goto 0
	End Function

	'--------------------------------------------------------------
	' Parse the individual labels within a folder

	Private Function ParseFile(sFolder, sFile)
		Dim sContents, oRE, oMatches, oMatch, sText

		sContents = GetFile(sFolder, sFile)
		'Call StripHTMLHead(sContents)
		'Call StripScriptBlocks(sContents)
		'Call StripIncludes(sContents)
		'Call StripHTMLComments(sContents)
		Set oRE = New RegExp
		oRE.IgnoreCase = True
		oRE.Global = True
		' parse all of the regular expressions to pull out labels
		For I = 0 To intPatterns - 1
			oRE.Pattern = arrPattern(I)
			Set oMatches = oRE.Execute(sContents)
			For Each oMatch In oMatches
				sText = oMatch.SubMatches(0)
				' Response.Write "Found match: *" & sText & "*" : Response.End
				If Not IsNumeric(sText) And Trim(sText) <> "" Then
					' add this word/phrase to the dictionary
					objDict.Item(sText) = True
				End If
			Next
		Next
		ParseFile = True
	End Function

	'--------------------------------------------------------------
	' Retrieve all of the ASP scripts in a folder

	Private Function ParseFilesInFolder(sFolder)
		Dim oFSO, oFolder, oFiles, oFile

		Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
		Set oFolder = oFSO.GetFolder(Server.MapPath(sFolder))
		Set oFiles = oFolder.Files
		For Each oFile in oFiles
			If InStr(1, oFile.Name, ".") Then
				sExt = Mid(oFile.Name, InStrRev(oFile.Name, "."))
				If InStr(1, ","&strExtList&",", ","&sExt&",") > 0 Then
					If ParseFile(sFolder, oFile.Name) Then
						If boolDebug Then Response.Write "File: " & oFSO.BuildPath(sFolder, oFile.Name) & "... success!<br>"
					Else
						' report all failures
						Response.Write "File: " & oFSO.BuildPath(sFolder, oFile.Name) & "... <b style=""{color:red}"">failure!</b><br>"
					End If
				End If
			End If
		Next
		ParseFilesInFolder = True
	End Function
	
	'--------------------------------------------------------------
	' Parse all of the files and sub-folders under 'sFolder'

	Public Function ParseFolder(sFolder)
		Dim oFSO, oFolder, oFolders

		If boolDebug Then Response.Write "Folder: " & sFolder & "<br>"
		If Not ParseFilesInFolder(sFolder) Then
			ParseFolder = False
			Exit Function
		End If
		' process all of the sub-folders
		Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
		Set oFolder = oFSO.GetFolder(Server.MapPath(sFolder))
		Set oFolders = oFolder.SubFolders
		For Each oFolder In oFolders
			If oFolder.Name <> "." And oFolder.Name <> ".." Then
				If Not ParseFolder(oFSO.BuildPath(sFolder, oFolder.Name)) Then
					ParseFolder = False
					Exit Function
				End If
			End If
		Next
		ParseFolder = True
	End Function

	'--------------------------------------------------------------
	' update the english phrases (requiring translation in the DB)

	Public Sub UpdateText
		Dim nCount, sKey, sStat

		nCount = 0
		For Each sKey In objDict
			sStat = sStat & "IF NOT EXISTS (select * from tblLangText where EnglishText = '" & Replace(sKey, "'", "''") & "') insert into tblLangText (EnglishText) values ('" & Replace(sKey, "'", "''") & "');" & vbCrLf
			nCount = nCount + 1
		Next
		' perform all of the inserts that need to be done
		Call adoExecute(sStat)
		Response.Write "<p><b>A total of " & nCount & " words/phrases were updated in the database</b></p>"
	End Sub

	'--------------------------------------------------------------
	' dump all of the words/phrases parsed from the web site

	Public Sub DumpWords
		Dim nCount, sKey

		With Response
		nCount = 0
		For Each sKey In objDict
			.Write "insert into tblLangText (Phrase) values ('" & Replace(sKey, "'", "''") & "');<br>"
			nCount = nCount + 1
		Next
		.Write "<p><b>A total of " & nCount & " words/phrases were found</b></p>"
		End With
	End Sub

	'--------------------------------------------------------------
	' PROPERTY: Debug

	Public Property Get ShowDebug
		ShowDebug = boolDebug
	End Property

	Public Property Let ShowDebug(bValue)
		boolDebug = bDebug
	End Property

	'--------------------------------------------------------------
	' PROPERTY: ErrorMsg

	Public Property Get ErrorMsg
		ErrorMsg = strErrorMsg
	End Property
End Class
%>