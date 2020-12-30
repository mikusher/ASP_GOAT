<%
'--------------------------------------------------------------------
' update_lib.asp
'	Add a new module to the database
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

Const UPD_FORREADING = 1
Const UPD_FORWRITING = 2
Dim updError		' error message

' ---------------------------------------------------------------
' check to see if a file exists at the given web path (sPath)

Function updFileExists(sPath)
	Dim oFSO
	
	' check to see if history file already exists
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	If oFSO.FileExists(Server.MapPath(sPath)) Then
		updFileExists = False
	Else
		updFileExists = True
	End If
	Set oFSO = Nothing
End Function		

' ---------------------------------------------------------------
' create a folder structure based on the path name passed
' arg should be in the form /folder1/folder2/folder3/filename.ext
' RETURNS: true if operation was successful, false otherwise

Function updBuildPath(sPathName)
	Dim oFSO, nPos, sCheckPath
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	nPos = InStr(2, sPathName, "/")
	Do While nPos > 0
		sCheckPath = Left(sPathName, nPos - 1)
		If Not oFSO.FolderExists(Server.MapPath(sCheckPath)) Then
			On Error Resume Next
			Set oFolder = oFSO.CreateFolder(Server.MapPath(sCheckPath))
			If Err.Number <> 0 Then
				updError = "updBuildPath - Unable to create folder: " & sCheckPath & "<BR>" &_
					Err.Number & " - " & Err.Description
				updBuildPath = False
				Exit Function
			End If
			On Error Goto 0
		End If
		nPos = InStr(nPos + 1, sPathName, "/")
	Loop
	updBuildPath = True
End Function

' ---------------------------------------------------------------
' read an entire text file and puts the contents in arg 2 (sContents)
' RETURNS: true on success, false otherwise

Function updRetrieveFile(sPathName, sContents, dtModified)
	Dim oFSO, oFile

	Set oFSO = CreateObject("Scripting.FileSystemObject")
	If (oFSO.FileExists(Server.MapPath(sPathName))) Then
		On Error Resume Next
		Set oFile = oFSO.GetFile(Server.MapPath(sPathName))
		dtModified = oFile.DateLastModified
		If Err.Number <> 0 Then
			updError = "updRetrieveFile - " & Err.Number & " - " & Err.Description
			updRetrieveFile = False
			Exit Function
		End If		
		Set oFile = oFSO.OpenTextFile(Server.MapPath(sPathName), UPD_FORREADING)
		If Err.Number <> 0 Then
			updError = "updRetrieveFile - " & Err.Number & " - " & Err.Description
			updRetrieveFile = False
			Exit Function
		End If
		sContents = oFile.ReadAll
		If Err.Number <> 0 Then
			updError = "updRetrieveFile - " & Err.Number & " - " & Err.Description
			updRetrieveFile = False
			Exit Function
		End If
		On Error Goto 0
	Else
		dtModified = Now()
		sContents = ""
	End If
	Set oFSO = Nothing
	updRetrieveFile = True
End Function

' ---------------------------------------------------------------
' store an entire text file and return it

Function updStoreFile(sPathName, sContents)
	Const bOverwrite = True
	Dim oFSO, oFile

	' make sure the path exists first
	If Not updBuildPath(sPathName) Then
		updStoreFile = False
		Exit Function
	End If
	' now try storing the file
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	On Error Resume Next
	Set oFile = oFSO.CreateTextFile(Server.MapPath(sPathName), bOverwrite)
	If Err.Number <> 0 Then
		updError = "updStoreFile 1 - " & Err.Number & " - " & Err.Description
		updStoreFile = False
		Exit Function
	End If
	oFile.Write(sContents)
	If Err.Number <> 0 Then
		updError = "updStoreFile 2 - " & Err.Number & " - " & Err.Description
		updStoreFile = False
		Exit Function
	End If
	oFile.Close
	If Err.Number <> 0 Then
		updError = "updStoreFile 3 - " & Err.Number & " - " & Err.Description
		updStoreFile = False
		Exit Function
	End If
	On Error Goto 0
	updStoreFile = True
End Function


' ---------------------------------------------------------------
' delete the file at the given web path (sPath)

Function updDeleteFile(sPath)
	Dim oFSO, sFile
	
	' make sure the filename has an extension
	If InStrRev(sPath, "/") > 0 Then
		sFile = Mid(sPath, InStrRev(sPath, "/"))
	Else
		sFile = sPath
	End If
	If Not (InStr(1, sFile, ".") > 0) Then
		updError = "updDeleteFile - Filename to delete doesn't contain extension (" & sPath & ")"
		updDeleteFile = False
		Exit Function
	End If

	' check to see if file exists
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	If oFSO.FileExists(Server.MapPath(sPath)) Then
		On Error Resume Next
		oFSO.Delete(Server.MapPath(sPath))
		If Err.Number <> 0 Then
			updError = "updDeleteFile - " & Err.Number & " - " & Err.Description
			updDeleteFile = False
			Exit Function
		End If
		On Error Goto 0
		updDeleteFile = True
	Else
		' assume file deleted successfully if it doesn't exist
		updDeleteFile = True
	End If
	Set oFSO = Nothing
End Function		

' ---------------------------------------------------------------
' Attempt to move a file from one location to another, optionally
' renaming the file.
' RETURNS: True on success, False otherwise

Function updMoveFile(srcPath, destPath)
	Dim sContents, dtModified

	' retrieve the contents of the source file
	updMoveFile = True
	If Not updRetrieveFile(srcPath, sContents, dtModified) Then
		updMoveFile = False
		Exit Function
	End If
	If Not updStoreFile(destPath, sContents) Then updMoveFile = False
End Function

' ---------------------------------------------------------------
' Attempt to execute the contents of a Transact-SQL database
' script file.
' RETURNS: True on success, False otherwise

Function updExecuteFile(srcPath)
	Dim sContents, dtModified

	' retrieve the contents of the source file
	updExecuteFile = True
	If Not updRetrieveFile(srcPath, sContents, dtModified) Then
		updExecuteFile = False
		Exit Function
	End If
	On Error Resume Next
	Call adoExecute(sContents)
	If Err.Number <> 0 Then
		updMoveFile = False
		Exit Function
	End If
	On Error Goto 0
End Function
' ---------------------------------------------------------------
' Process a site update file located on the local filesystem

Function updProcessControl(sControlPath)
	Dim sContents, dtModified, aLine, aParam, sCmd

	If Not updRetrieveFile(srcPath, sContents, dtModified) Then
		updProcessControl = False
		Exit Function
	End If
	aLine = Split(sContents, vbCrLf)
	For I = 0 To UBound(aLine)
		sCmd = Trim(aLine(I))
		If sCmd <> "" Then
			aParam = Split(sCmd, " ")
			Select Case UCase(aParam(0))
				Case "MOVE" :
					' move a source file to a specific destination
					If Not updMoveFile(sControlPath & aParam(1), aParam(2)) Then
						updProcessControl = False
						Exit Function
					End If
				Case "DELETE" :
					' should we even allow this?
					If Not updDeleteFile(aParam(1)) Then
						updProcessControl = False
						Exit Function
					End If
			End Select
		End If
	Next
End Function

%>