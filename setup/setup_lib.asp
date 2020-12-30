<%
' -------------------------------------------------------------------
' setup_lib.asp
'	Support routines for building the dynamic ASP Nuke application
'	configuration forms (/admin/configure.asp)
'
' AUTH:	Ken Richards
' DATE:	10/13/03
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

Dim setStep
Dim setDatabaseType		' type ("sqlserver7" or "sqlserver2000")
Dim setErrorMsg
Dim setStatusMsg
Dim setWasError

' define the scripts which makeup the steps of the wizard
setStep = Array("setup1.asp", "setup2.asp", "setup3.asp")
setStepButton = Array("Step 1: Database Connection", "Step 2: Database Schema", "Step 3: Configuration")

'--------------------------------------------------------------------
' Display the status and/or error messages to the user 
Sub setDisplayStatus
	' check to see if the database connection string is valid
	If setErrorMsg = "" Then
		' database connection is valid
		Response.Write "<p><b>Success!!!</b></p>" & vbCrLf
	Else
		Response.Write "<p><b class=""error"">FAILURE!!!</b></p>" & vbCrLf
	End If
	If setStatusMsg <> "" Then
		Response.Write "<p><kbd>" & Replace(setStatusMsg, "<br>", "<br><br>") & "</kbd></p>" & vbCrLf
	End If
	If setErrorMsg <> "" Then
		Response.Write "<p class=""error"">" & Replace(setErrorMsg, "<br>", "<br><br>") & "</p>" & vbCrLf
	End If	
	' clear the variables to prepare for the next test
	setStatusMsg = ""
	setErrorMsg = ""
End Sub

'--------------------------------------------------------------------
' check to see if the database connection can be parsed from the
' global.asa file
' setup application variables needed for ado_lib.asp on success

Function setDBConnTest
	Dim oFSo, oFile, sRootPath, sPath, sContents, nPos, sConnStr 

	' check to see if connection string is already defined
	If Application("adoConn_ConnectionString") <> "" Then
		setStatusMsg = setStatusMsg & "Your connection string has already defined (Application(""adoConn_ConnectionString""))<br>" &_
			"Resave your root-level web configuration file (global.asa) to clear the cache<br>"
	End If

	' attempt to locate the global.asa configuration file
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	sRootPath = Server.MapPath("/")
	sPath = Server.MapPath(".")
	Do Until oFSO.FileExists(oFSO.BuildPath(sPath, "global.asa"))
		nPos = InStrRev(sPath, "\", Len(sPath) - 1)
		If nPos < 1 Then Exit Do
		sPath = Left(sPath, nPos)
	Loop
	' abort if the global.asa file was not found
	If Not oFSO.FileExists(oFSO.BuildPath(sPath, "global.asa")) Then
		setErrorMsg = setErrorMsg & "<P><B class=""error"">Unable to locate global.asa configuration file... " & oFSO.BuildPath(sPath, "global.asa") & "</B></P>"
		setWasError = True
		setDBConnTest = False
		Exit Function
	End If
	' read the contents of the file
	Set oFile = oFSO.OpenTextFile(oFSO.BuildPath(sPath, "global.asa"), 1)
	sContents = oFile.ReadAll
	oFile.Close
	' find the connection string variable
	Set oRegex = New RegExp
	oRegex.Pattern = "asaConnectionString\s*=\s*""([^""]*)"""
	' oRegex.Multiline = True
	oRegex.IgnoreCase = True
	Set oMatches = oRegex.Execute(sContents)
	For Each oMatch In oMatches
		sConnStr = oMatch.SubMatches(0)
		Exit For
	Next
	' abort if the connection string could not be parsed
	If sConnStr = "" Then
		setErrorMsg = setErrorMsg & "Unable to parse the asaConnectionString value from: """ & sPath & "global.asa"".<br>"
		setWasError = True
		setDBConnTest = False
	ElseIf sConnStr = "Provider=SQLOLEDB;server=127.0.0.1;driver={SQL Server};uid=dbuserid;pwd=dbpassword;database=dbname;" Then
		setErrorMsg = setErrorMsg & "Found default database connection string: """ & Server.HTMLEncode(sConnStr) & """<BR>"
		setStatusMsg = setStatusMsg & "You need to modify your global.asa file on your server located at: " &_
			oFSO.BuildPath(sPath, "global.asa") & " and replace the <kbd>asaConnectionString</kbd> constant with your database connection string."
		setWasError = True
		setDBConnTest = False
	Else ' everything is setup properly
		setStatusMsg = setStatusMsg & "Found database connection string in global.asa file!<BR>"
		setWasError = False
		setDBConnTest = True
		' set the application variables for database connectivity
		Application("adoConn_ConnectionString") = sConnStr
		Application("adoConn_CommandTimeout") = 1000
		Application("adoConn_ConnectionTimeout") = 500
	End If
End Function

'--------------------------------------------------------------------
' make sure the database is empty and the connection string is valid

Function setDBConnTest2(bTablesExist)
	Dim oConn, oRS, sQuery, rsTest, I
	Const adOpenForwardOnly = 0
	Const adOpenKeySet = 1
	Const adLockReadOnly = 1
	Const adLockOptimistic = 3
	Const adCmdText = 1					' command is SQL text
	Const adExecuteNoRecords = 128		' indicate to ADO that no recordset is returned

	On Error Resume Next
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open Application("adoConn_ConnectionString")
	If Err.Number <> 0 Then
		setErrorMsg = "Database connection string is invalid, database error is<br>" &_
			Err.Source & ": " & Err.Description & " (" & Err.Number & ")<br>" &_
			"value: """ & Application("adoConn_ConnectionString") & """"
		setWasError = True
		setDBConnTest2 = False
		Exit Function
	End If

	Set rs = Server.CreateObject("ADODB.Recordset")
	' check for any existing user tables
	If InStr(1, Application("adoConn_ConnectionString"), "MySQL", vbTextCompare) > 0 Then
		sQuery = "show tables;"
	Else
		sQuery = "select * from dbo.sysobjects " &_
				"where	xtype= 'U' " &_
				"and	OBJECTPROPERTY(id, N'IsUserTable') = 1 " &_
				"and	name <> 'dtproperties'"
		'		"and	uid = user_id('dbo') " &_
	End If
	rs.Open sQuery, oConn, adOpenKeySet, adLockReadOnly, adCmdText	
	If Err.Number <> 0 Then
		setErrorMsg = "Unable to open recordset, database error is<br>" &_
			Err.Source & ": " & Err.Description & " (" & Err.Number & ")"
		setWasError = True
		setDBConnTest2 = False
		Exit Function
	End If
	' make sure no tables exist in this table
	If Not rs.EOF Then
		setStatusMsg = setStatusMsg & "<B>!!!! WARNING !!!</B><BR>The following tables are already defined in the database<BR><BR>" & vbCrlf
		I = 0
		Do Until rs.EOF
			If I > 0 Then setStatusMsg = setStatusMsg & " "
			setStatusMsg = setStatusMsg & rs.Fields(0).Value
			rs.MoveNext
			I = I + 1
		Loop
		rs.Close
		Set rs = Nothing
		bTablesExist = True
	Else
		bTablesExist = False
	End If
	On Error Goto 0
	setWasError = False
	setDBConnTest2 = True
End Function

' -------------------------------------------------------------------------
' Run the initial setup script for the Microsoft SQL database schema
' RETURNS: true on success, false otherwise

Function setSetupSchemaSQL
	Dim dtModified, sContents, I

	If InStr(1, Application("adoConn_ConnectionString"), "MySQL", vbTextCompare) > 0 Then
		sSchemaFile = "schema_mysql.sql"
	Else
		sSchemaFile = "schema.sql"
	End If
	If Not setRetrieveFile(sSchemaFile, sContents, dtModified) Then
		setSetupSchemaSQL = False
		Exit Function
	End If

	' modify the script for SQL Server 7 (if nec)
	If setDatabaseType = "sqlserver7" Then
		sContents = Replace(sContents, "COLLATE SQL_Latin1_General_CP1_CI_AS", "")
	End If

	On Error Resume Next
	If setDatabaseType = "MySQL" Then
		Dim aSQL
		aSQL = Split(sContents, ";" & vbCrLf)
		For I = 0 To UBound(aSQL)
			If Trim(Replace(Replace(aSQL(I), vbCr, ""), vbLf, "")) <> "" Then
				Call adoExecute(aSQL(I))
				If Err.Number <> 0 Then
					setWasError = True
					setErrorMsg = setErrorMsg & "Database error running SQL Setup Script (setup/" & sSchemaFile & ")<BR>" &_
						Err.Source & ": " & Err.Description & " (" & Err.Number & ")"
					setSetupSchemaSQL = False
					Exit Function
				End If
			End If
		Next
	Else
		Call adoExecute(sContents)
		If Err.Number <> 0 Then
			setWasError = True
			setErrorMsg = setErrorMsg & "Database error running SQL Setup Script (setup/schema.sql)<BR>" &_
				Err.Source & ": " & Err.Description & " (" & Err.Number & ")"
			setSetupSchemaSQL = False
			Exit Function
		End If
	End If
	On Error Goto 0
	setSetupSchemaSQL = True
End Function


' -------------------------------------------------------------------------
' Run the initial setup script for the Microsoft SQL database data import
' RETURNS: true on success, false otherwise

Function setSetupDataSQL
	Dim sDataFile, dtModified, sContents

	If setDatabaseType = "MySQL" Then
		sDataFile = "data_mysql.sql"
	Else
		sDataFile = "data.sql"
	End If

	If Not setRetrieveFile(sDataFile, sContents, dtModified) Then
		setSetupDataSQL = False
		Exit Function
	End If

	' modify the script for SQL Server 7 (if nec)
	If setDatabaseType = "sqlserver7" Then
		sContents = Replace(sContents, "COLLATE SQL_Latin1_General_CP1_CI_AS", "")
	End If

	On Error Resume Next
	If setDatabaseType = "MySQL" Then
		Dim aSQL
		aSQL = Split(sContents, ";" & vbCrLf)
		For I = 0 To UBound(aSQL)
			If Trim(Replace(Replace(aSQL(I), vbCr, ""), vbLf, "")) <> "" Then
				Call adoExecute(aSQL(I))
				If Err.Number <> 0 Then
					setWasError = True
					setErrorMsg = setErrorMsg & "Database error running SQL Setup Script (setup/" & sDataFile & ")<BR>" &_
						Err.Source & ": " & Err.Description & " (" & Err.Number & ")"
					setSetupDataSQL = False
					Exit Function
				End If
			End If
		Next
	Else
		Call adoExecute(sContents)
		If Err.Number <> 0 Then
			setWasError = True
			setErrorMsg = setErrorMsg & "Database error running SQL Setup Script (setup/" & sDataFile & ")<BR>" &_
				Err.Source & ": " & Err.Description & " (" & Err.Number & ")"
			setSetupDataSQL = False
			Exit Function
		End If
	End If
	On Error Goto 0
	setSetupDataSQL = True
End Function

' -------------------------------------------------------------------------
' read an entire text file and puts the contents in arg 2 (sContents)
' RETURNS: true on success, false otherwise

Function setRetrieveFile(sPathName, sContents, dtModified)
	Dim oFSO, oFile

	Set oFSO = CreateObject("Scripting.FileSystemObject")
	If (oFSO.FileExists(Server.MapPath(sPathName))) Then
		On Error Resume Next
		Set oFile = oFSO.GetFile(Server.MapPath(sPathName))
		dtModified = oFile.DateLastModified
		If Err.Number <> 0 Then
			setErrorMsg = setErrorMsg & "setRetrieveFile - " & Err.Number & " - " & Err.Description
			setRetrieveFile = False
			setWasError = True
			Exit Function
		End If		
		Set oFile = oFSO.OpenTextFile(Server.MapPath(sPathName), FSO_FORREADING)
		If Err.Number <> 0 Then
			setErrorMsg = setErrorMsg  & "setRetrieveFile - " & Err.Number & " - " & Err.Description
			setRetrieveFile = False
			setWasError = True
			Exit Function
		End If
		sContents = oFile.ReadAll
		If Err.Number <> 0 Then
			setErrorMsg = setErrorMsg & "setRetrieveFile - " & Err.Number & " - " & Err.Description
			setRetrieveFile = False
			setWasError = True
			Exit Function
		End If
		On Error Goto 0
	Else
		dtModified = Now()
		sContents = ""
	End If
	Set oFSO = Nothing
	setRetrieveFile = True
End Function

'--------------------------------------------------------------------
' Build the wizard navigation buttons to go to previous screen
' or the next screen

Sub setWizardButtons(bAllowForward)
	Dim sPage, I, nCurrPos

	' determine the current page we are on
	sPage = Request.ServerVariables("SCRIPT_NAME")
	If InStrRev(sPage, "/") > 0 Then sPage = Mid(sPage, InStrRev(sPage, "/")+ 1)
	If InStr(sPage, "?") > 0 Then sPage = Left(sPage, InStr(1, sPage, "?")- 1)
	If InStr(sPage, ":") > 0 Then sPage = Left(sPage, InStr(1, sPage, ":")- 1)
	If InStr(sPage, "#") > 0 Then sPage = Left(sPage, InStr(1, sPage, "#")- 1)
	' find the position within the wizard path
	nCurrPos = -1
	For I = 0 To UBound(setStep)
		If sPage = setStep(I) Then
			nCurrPos = I
			Exit For
		End If
	Next
	With Response
	If nCurrPos >= 0 Then
		.Write "<p align=""center"">" & vbCrLf
		If nCurrPos > 0 Then
			.Write "<input type=""button"" name=""_prev"" value=""&lt;&lt; "
			.Write Server.HTMLEncode(setStepButton(nCurrPos - 1))
			.Write """ onClick=""location.href='"
			.Write setStep(nCurrPos - 1)
			.Write "'"" class=""form"">" & vbCrLf
			If nCurrPos < UBound(setStep) Then .Write "&nbsp;&nbsp;"
		End If
		If nCurrPos < UBound(setStep) And bAllowForward Then
			.Write "<input type=""button"" name=""_next"" value="""
			.Write Server.HTMLEncode(setStepButton(nCurrPos + 1))
			.Write " &gt;&gt;"" onClick=""location.href='"
			.Write setStep(nCurrPos + 1)
			.Write "'"" class=""form"">" & vbCrLf
		End If
	End If
	End With
End Sub

%>