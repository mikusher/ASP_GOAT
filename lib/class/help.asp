<%
'--------------------------------------------------------------------
' help.asp
'	Class for reading and parsing the integrated help files
' REQUIRES
'	tab_lib.asp - if you call the function TabControl
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

Class clsHelp
	Private mstrError		' error message
	Private mstrPathInfo	' path to the page help was requested on
	Private mstrContents	' contents of the help file
	Private mstrSectionContents	' contents of section requested

	Private Sub Class_Initialize
	End Sub

	'--------------------------------------------------------------
	' Find the integrated help file on the local filesystem

	Private Function FindHelp(bTraverseUp)
		Dim oFSO, sRelPath, sPhysPath, sLastPath

		If mstrPathInfo <> "" Then
			sRelPath = mstrPathInfo
		Else
			sRelPath = "."
		End If

		Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
		sLastPath = Server.MapPath("/")
		Do Until sLastPath = sPhysPath
			sPhysPath = Server.MapPath(sRelPath)
			' mstrError = "Check for file = " & oFSO.BuildPath(sPhysPath, "help.asp") : FindHelp = "" : Exit Function
			' check to see if help file was found
			If oFSO.FileExists(oFSO.BuildPath(sPhysPath, "help.asp")) Then
				FindHelp = oFSO.BuildPath(sPhysPath, "help.asp")
				Exit Function
			End If

			' abort if we don't want to traverse upwards
			If Not bTraverseUp Then Exit Do

			' move up to the next parent level (in folder structure)
			If sRelPath = "." then sRelPath = ".." Else sRelPath = sRelPath & "../"
			On Error Resume Next
				sPhysPath = Server.MapPath(sRelPath)
			If Err.Number <> 0 Then
				mstrError = "clsHelp::FindHelp - Invalid relative path (" & sRelPath & ")"
				FindHelp = ""
				Exit Function
			End If
			On Error Goto 0
		Loop
		mstrError = "Unable to find help document (help.asp) to display"
		FindHelp = ""
	End Function

	'--------------------------------------------------------------
	' Retrieve the section (sSectionName) from the help contents

	Private Function SectionContents(sSectionName)
		Dim oRE, oMatches, oMatch

		Set oRE = New RegExp
		oRE.Pattern = "<!--\s*SECTION_START:\s*" & sSectionName & "\s*-->" &_
			"[\s\S]*?" &_
			"<!--\s*SECTION_END:\s*" & sSectionName & "\s*-->"
		oRE.IgnoreCase = True
		oRE.Global = True
		Set oMatches = oRE.Execute(mstrContents)
		For Each oMatch In oMatches
			' just return the first match - don't care about the rest
			SectionContents = oMatch.Value
			Exit Function
		Next
		mstrError = "Unable to locate help section ""<B>" & sSectionName & "</b>"" in help"
		SectionContents = ""
	End Function

	'--------------------------------------------------------------
	' Display the tab control for all of the sections in the help

	Private Function ParseContents(sSearchText, sAfterText, sBeforeText)
		Dim nStart, nEnd

		nStart = InStr(1, sSearchText, sAfterText, vbTextCompare)
		If nStart > 0 Then
			nEnd = InStr(nStart, sSearchText, sBeforeText, vbTextCompare)
			If nEnd > nStart Then
				ParseContents = Trim(Mid(sSearchText, nStart + Len(sAfterText), nEnd - (nStart + Len(sAfterText))))
				Exit Function
			End If
		End If
		ParseContents = ""
	End Function

	'--------------------------------------------------------------
	' Display the tab control for all of the sections in the help
	' RETURNS: True on success, False otherwise

	Public Function TabControl(sCurrentSection)
		Dim oRE, oMatches, oMatch, sSectionName
		Dim sLabelList, sURLList

		Set oRE = New RegExp
		oRE.Pattern = "<!--\s*SECTION_START:\s*(.*)\s*-->" &_
			"[\s\S]*" &_
			"<!--\s*SECTION_END:\s*\1\s*-->"
		oRE.IgnoreCase = True
		oRE.Global = True
		Set oMatches = oRE.Execute(mstrContents)

		bCurrentFound = False
		For Each oMatch In oMatches
			' parse the section name from the match
			sSectionName = ParseContents(oMatch.Value, "SECTION_START:", "-->")
			If sSectionName <> "" Then
				If sLabelList <> "" Then
					sLabelList = sLabelList & ","
					sURLList = sURLList & ","
				End If
				sLabelList = sLabelList & sSectionName
				sURLList = sURLList & Request.ServerVariables("SCRIPT_NAME") &_
					"?pathinfo=" & Server.URLEncode(mstrPathInfo) &_
					"&section=" & Server.URLEncode(sSectionName)
			End If
		Next
		' display the tab control and return success code
		If sLabelList <> "" Then
			tabShow sLabelList, sURLLIst, sCurrentSection
			TabControl = True
		Else
			mstrError = "Unable to parse sections from help document (help.asp)"
			TabControl = False
		End If
	End Function

	'--------------------------------------------------------------
	' Retrieve the help section (sSectionName) from the filesystem

	Public Function RetrieveHelp(sSectionName, bTraverseUp)
		Dim sFile, oFSO, oFile

		sFile = FindHelp(bTraverseUp)
		If sFile = "" Then
			RetrieveHelp = False
			Exit Function
		End If
		Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
		Set oFile = oFSO.OpenTextFile(sFile, 1) ' ForReading
		mstrContents = oFile.ReadAll
		oFile.Close
		Set oFile = Nothing
		Set oFSO = Nothing

		' find the named section (if any)
		mstrSectionContents = SectionContents(sSectionName)
		If mstrSectionContents = "" Then
			RetrieveHelp = False
		Else
			RetrieveHelp = True
		End If
	End Function

	'--------------------------------------------------------------
	' Property Get/Let for the relative path to help page

	Public Property Let PathInfo(strValue)
		mstrPathInfo = strValue
	End Property

	Public Property Get PathInfo
		PathInfo = mstrPathInfo
	End Property


	'--------------------------------------------------------------
	' Property Get for the Error message

	Public Property Get ErrorMsg
		ErrorMsg = mstrError
	End Property

	'--------------------------------------------------------------
	' Property Get for the help section

	Public Property Get HelpSection
		HelpSection = mstrSectionContents
	End Property
End Class
%>