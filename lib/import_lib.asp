<%
' -------------------------------------------------------------------
' import_lib.asp
'	Imports data from a ASP Nuke datafile into the database
'	Format of the file is as follows:
'	IMPORT TABLE [TableName]
'	PRIMARY KEY [KeyName],[KeyName],[KeyName]...
'	FIELD NAMES [Field1],[Field2],[Field3]...
'	ALLOW [ INSERT | UPDATE ]
'	row1field1,row1field2,row1field3,...
'	row2field1,row2field2,row2field3,...
'
'	Data is encrypted using URL encoding with commas replaced by %2C
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

Const FSO_FORREADING = 1
Const FSO_FORWRITING = 2

Dim impError

' -------------------------------------------------------------------------
' Process the CSV data (which is URL encoded) in the data file and insert
' or update it into the specified table

Function impProcessData(ByRef nPos, ByRef aLine, sTableName, sPrimaryKey, sFieldList, sAllow)
	Dim aKey, aKeyPos(), sKeyPosList, aField, aAllow, sWhere, sWhereTemplate
	Dim aValue, sValueList, sSetList
	Dim sStat, bInsert, bUpdate

	aKey = Split(Replace(sPrimaryKey, " ", ""), ",")
	aField = Split(Replace(sFieldList, " ", ""), ",")
	aAllow = Split(Replace(sAllow, " ", ""), ",")
	If InStr(sAllow, "insert", vbTextCompare) > 0 Then bInsert = True
	If InStr(sAllow, "update", vbTextCompare) > 0 Then bUpdate = True
	' find the field positions for the primary keys
	ReDim aKeyPos(UBound(aKey))
	For I = 0 To UBound(aKey)
		For J = 0 To UBound(aField)
			If StrCmp(aKey(I), aField(J), vbTextCompare) = 0 Then
				aKeyPos(I) = J
				sKeyPosList = sKeyPosList & "," & J
				Exit For
			End If
			If bUpdate Then
				impError = "Primary key field """ & aKey(I) & """ not found in FIELD LIST"
				impProcessData = False
				Exit Function
			End If
			sWhereTemplate = sWhereTemplate & "AND " & aKey(I) & " = [PrimaryKey" & I & "]"
		Next
	Next
	sKeyPosList = sKeyPosList & ","
	If sWhereTemplate <> "" Then sWhereTemplate = Mid(sWhereTemplate, 5)
	' parse all of the data lines in the file
	Do Until Trim(aLine(nPos)) = ""
		' split out the field values and perform URL decode
		aValue = Split(aLine(nPos), ",")

		' make sure the right number of fields were passed
		If UBound(aValue) <> UBound(aField) Then
			impError = "Number of field values is incorrect (Got " & (UBound(aValue) + 1) & ", Expected " & (UBound(aField) + 1) & ") at line " & nPos
			impProcessData = False
			Exit Function
		End If

		' perform a URL decode on the values
		For I = 0 to UBound(aValue)
			aValue(I) = impURLDecode(aValue(I))
			sValueList = sValueList & ",'" & Replace(aValue(I), "'", "''") & "'"
			If Not (InStr(1, sKeyPosList, "," & I & ",") > 0) Then
				sSetList = sSetList & ", " & aField(I) & " = '" & Replace(aValue(I), "'", "''") & "'"
			End If
		Next
		If sValueList <> "" Then sValueList = Mid(sValueList, 2)
		If sSetList <> "" Then sSetList = Mid(sSetList, 2)

		' build the where clause for update / insert
		sWhere = sWhereTemplate
		For I = 0 To UBound(aKeyPos)
			sWhere = Replace(sWhere, "[PrimaryKey" & I & "]", "'" & Replace(aValue(aKeyPos(I)), "'", "''") & "'")
		Next
		If bInsert And bUpdate Then
			' perform insert or update
			sStat = sStat & "IF NOT EXISTS(SELECT * FROM [" & sTableName & "] WHERE " & sWhere) " & vbCrLf &_
				"	INSERT INTO [" & sTableName & "] (" & sFieldList & ") VALUES (" & sValueList & ")" & vbCrLf &_
				"ELSE" & vbCrLf &_
				"	UPDATE [" & sTableName & "] SET " & sSetList & " WHERE " & sWhere & vbCrLf
		ElseIf bInsert Then
			sStat = sStat & "IF NOT EXISTS(SELECT * FROM [" & sTableName & "] WHERE " & sWhere) " & vbCrLf &_
				"	INSERT INTO [" & sTableName & "] (" & sFieldList & ") VALUES (" & sValueList & ")" & vbCrLf
		ElseIf bUpdate Then
			sStat = "UPDATE [" & sTableName & "] SET " & sSetList & " WHERE " & sWhere & vbCrLf
		Else
			impError = "Allow must specify INSERT and UPDATE at line " & nPos - 1
			impProcessDate = False
			Exit Function
		End If
		nPos = nPos + 1
	Loop
	' execute all of the queries necessary for this table
	If sStat <> "" Then
		Call adoExecute(sStat)
	End If
	impProcessData = True
End Function

' -------------------------------------------------------------------------
' Import a data file into the database.  This will handle multiple
' tables

Function impImport(sPathName)
	Dim sContents, dtModified, aLine, oReg, I
	Dim sTableName, sPrimaryKey, sFieldList, sAllow
	Dim oMatch, oMatches

	If impRetrieveFile(sPathName, sContents, dtModified) Then
		' we have the data file, begin the import process
	Else
		impImport = False
		Set oReg = New RegExp
		oReg.Pattern = "\n\s*IMPORT\s*TABLE\s*(\w+)\s*\n"
		oReg.Global = True
		oReg.Multiline = True
		aLine = Split(sContents, vbCrLf)
		I = 0
		Do Until I > UBound(aLine)
			' parse the table name to import to
			sTableName = ""
			Set oMatches = regEx.Execute(aLine(I))
			If oMatches.Count > 0 Then
				sTableName = oMatches(0).SubMatches(0)
				I = I + 1
				If I > UBound(aLine) Then
					impError = "Expected ""PRIMARY KEY"" at Line: " & I
					impImport = False
					Exit Function
				End If
				' parse the primary keys to import to
				oReg.Pattern = "\n\s*PRIMARY\s*KEY\s*(\w.*?\w)\s*\n"
				Set oMatches = regEx.Execute(aLine(I))
				If oMatches.Count > 0 Then
					sPrimaryKey= oMatches(0).SubMatches(0)
					I = I + 1
					If I > UBound(aLine) Then
						impError = "Expected ""FIELD NAMES"" at Line: " & I
						impImport = False
						Exit Function
					End If
					' parse the primary keys to import to
					oReg.Pattern = "\n\s*FIELD\s*NAMES\s*(\w.*?\w)\s*\n"
					Set oMatches = regEx.Execute(aLine(I))
					If oMatches.Count > 0 Then
						sFieldNames = oMatches(0).SubMatches(0)
						I = I + 1
						If I > UBound(aLine) Then
							impError = "Expected ""ALLOW"" at Line: " & I
							impImport = False
							Exit Function
						End If
						' parse the operations to "allow"
						oReg.Pattern = "\n\s*ALLOW\s*(\w.*?\w)\s*\n"
						Set oMatches = regEx.Execute(aLine(I))
						If oMatches.Count > 0 Then
							sAllow = oMatches(0).SubMatches(0)
							I = I + 1
							If I > UBound(aLine) Then
								impError = "Expected CSV data at Line: " & I
								impImport = False
								Exit Function
							End If
							' now we can process the data
							If Not impProcessData(I, aLine, sTableName, sPrimaryKey, sFieldNames, sAllow) Then
								impImport = False
								Exit Function
							End If
						End If
					End If
				End If			
			End If
			I = I + 1
		Loop
		Exit Function
	End If
	impImport = True
End Function

' -------------------------------------------------------------------------
' perform a URL decode to retrieve the original value

Function impURLDecode(sConvert)
	Dim aSplit
	Dim sOutput
	Dim I
	If IsNull(sConvert) Then
	   lgnURLDecode = ""
	   Exit Function
	End If
	
	' convert all pluses to spaces
	sOutput = Replace(sConvert, "+", " ")
	
	' next convert %hexdigits to the character
	If InStr(1, sOutput, "%") > 0 Then
		aSplit = Split(sOutput, "%")
		
		If IsArray(aSplit) Then
		   sOutput = aSplit(0)
		   For I = LBound(aSplit) to UBound(aSplit) - 1
		      sOutput = sOutput & Chr("&H" & Left(aSplit(i + 1), 2)) & Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
		   Next
		End If
	End If
	
	impURLDecode = sOutput
End Function

' -------------------------------------------------------------------
' read an entire text file and puts the contents in arg 2 (sContents)
' RETURNS: true on success, false otherwise

Function impRetrieveFile(sPathName, sContents, dtModified)
	Dim oFSO, oFile

	Set oFSO = CreateObject("Scripting.FileSystemObject")
	If (oFSO.FileExists(Server.MapPath(sPathName))) Then
		On Error Resume Next
		Set oFile = oFSO.GetFile(Server.MapPath(sPathName))
		dtModified = oFile.DateLastModified
		If Err.Number <> 0 Then
			impError = "impRetrieveFile - " & Err.Number & " - " & Err.Description
			impRetrieveFile = False
			Exit Function
		End If		
		Set oFile = oFSO.OpenTextFile(Server.MapPath(sPathName), FSO_FORREADING)
		If Err.Number <> 0 Then
			impError = "impRetrieveFile - " & Err.Number & " - " & Err.Description
			impRetrieveFile = False
			Exit Function
		End If
		sContents = oFile.ReadAll
		If Err.Number <> 0 Then
			impError = "impRetrieveFile - " & Err.Number & " - " & Err.Description
			impRetrieveFile = False
			Exit Function
		End If
		On Error Goto 0
	Else
		dtModified = Now()
		sContents = ""
	End If
	Set oFSO = Nothing
	impRetrieveFile = True
End Function
%>