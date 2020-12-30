<%
'--------------------------------------------------------------------
' ado_lib.asp
'	This is a library for managing connections and queries run on
'	a SINGLE database.  The database connection variables are stored
'	in application variables which should be defined in the global.asa
'	Provides debug information which may be displayed on a web page
'	by setting the variable adoDebug to True after you include this
'	file on your web page.
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

Dim adoConnection		' connection to the database
Dim adoLib				' determines if this library has been included
Dim adoDebug			' set to true to enable debug output

adoLib = true

'ADO Database constants
Const adOpenForwardOnly = 0
Const adOpenKeySet = 1

Const adLockReadOnly = 1
Const adLockOptimistic = 3

Const adCmdText = 1					' command is SQL text
Const adExecuteNoRecords = 128		' indicate to ADO that no recordset is returned

'Stored Procedures
Const adParamInput = 1
Const adParamOutput = 2
Const adParamInputOutput = 3
Const adParamReturnValue = 4
Const adVarChar = 200
Const adChar = 129
Const adInteger = 3
Const adCurrency = 6

'Data type constants
Const VariantDate = 135

'--------------------------------------------------------------------
' adoRecordsetErrors
' Displays all of the errors in the Errors collection of the
' recordset.  Call this function when opening a recordset creates
' an error.

Sub adoRecordsetErrors(recordset)
	Dim sError
	
	For Each sError In recordset.Errors
		Response.Write("<B CLASS=""error"">" & UCase(TypeName(sError)) & "</B><BR>")
	Next
End Sub

'--------------------------------------------------------------------
' adoConnect
' Opens a new connection to the database (configured in global.asa)
' unless a connection has already been opened.  No need to do pooling
' since that is handled by IIS internally.

Sub adoConnect()
	' open a connection to the database
	Set adoConnection = Server.CreateObject("ADODB.Connection")
	If adoDebug Then Response.Write("CONNECTION: " & Application("adoConn_ConnectionString") & "<BR><BR>")
	' connection timeout and command timeout are not supported for OLEDB
	' adoConnection.ConnectionTimeout = Application("adoConn_ConnectionTimeout")
	' adoConnection.CommandTimeout = Application("adoConn_CommandTimeout")
	' Response.Write "Connecting to database: *" & Application("adoConn_ConnectionString") & "*<BR>" : Response.End
	adoConnection.Open Application("adoConn_ConnectionString")	
End Sub

'--------------------------------------------------------------------
' adoExecute
' Executes a query without returning a recordset.  This method will
' return a number indicating the number of rows that were affected
' by the query.

Function adoExecute(sQuery)
	Dim nRecordsAffected		' number of records affected
	
	If UCase(TypeName(adoConnection)) <> "CONNECTION" Then adoConnect()
	
	If adoDebug Then Response.Write(sQuery & "<BR>")
	On Error Resume Next

	adoConnection.Execute sQuery, nRecordsAffected, adCmdText + adExecuteNoRecords
	If Err.Number <> 0 Then
		Response.Write("<P><B CLASS=""error"">Error # " & CStr(Err.Number) & " (0x" & Hex(Err.Number) & ")<BR>" & Err.Description & "</B><BR>" & sQuery & "</P>")
		Err.Clear   ' Clear the error.
		Response.End
	End If

	On Error Goto 0
	If adoDebug Then Response.Write(CStr(nRecordsAffected) & " Record(s) Affected<BR><BR>")

	' Set rs = Nothing
	adoExecute = nRecordsAffected
End Function

'--------------------------------------------------------------------
' adoOpenRecordset
' Opens a forward-only recordset from the database using the
' supplied query (sQuery).

Function adoOpenRecordset(sQuery)
	Dim rs						' recordset for query

	' open a connection to the database
	If UCase(TypeName(adoConnection)) <> "CONNECTION" Then adoConnect()

	If adoDebug Then Response.Write(sQuery & "<BR><BR>")
	On Error Resume Next

	' open the recordset
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sQuery, adoConnection, adOpenKeySet, adLockReadOnly, adCmdText
	If Err.Number <> 0 Then
		Response.Write("<P><B CLASS=""error"">Error # " & CStr(Err.Number) & " (0x" & Hex(Err.Number) & ")<BR>" & Err.Description & "</B><BR>" & sQuery & "</P>")
		' Response.Write sQuery & ", adoConnection, " & adOpenForwardOnly & ", " & adLockReadOnly & ", " & adCmdText & "<BR><BR>"
		adoRecordsetErrors(rs)
		Err.Clear   ' Clear the error.
		Response.End
	End If
	
	On Error Goto 0
	If adoDebug Then Response.Write(CStr(rs.RecordCount) & " Record(s) in Recordset<BR><BR>")
	
	' return the result
	Set adoOpenRecordset = rs
End Function

'--------------------------------------------------------------------
' adoDisconnect
' disconnect from the database here

Sub adoDisconnect()
	' close cursor and database connection
	If UCase(TypeName(adoConnection)) = "CONNECTION" then
		adoConnection.Close
		Set adoConnection = Nothing
	End If
End Sub

'--------------------------------------------------------------------
' adoQuoteFields
' special quoting of reserved words in field lists

Sub adoQuoteFields(aFields)
	Dim I
	Dim sField			' individual field from fields array
	
	If InStr(1, Application("adoConn_ConnectionString"), "Provider=Microsoft.Jet") > 0 Then
		' database is access, quote reserved words
		If IsArray(aFields) Then
			For I = 0 To UBound(aWords)
				Select Case UCASE(aWords(I))
					Case "PASSWORD" : aWords(I) = "[" & aWords(I) & "]"
				End Select
			Next
		End If
	End If
End Sub

'--------------------------------------------------------------------
' adoDetermineType
' determine the database type that we are connecting to

Sub adoDetermineType
	If InStr(1, Application("adoConn_ConnectionString"), "MySQL") > 0 Then
		Application("ADO_DATABASETYPE") = "MySQL"
	ElseIf InStr(1, Application("adoConn_ConnectionString"), "Provider=Microsoft.Jet") > 0 Then
		Application("ADO_DATABASETYPE") = "Access"
	Else
		Application("ADO_DATABASETYPE") = "sqlserver2000"
	End If		
End Sub

'--------------------------------------------------------------------
' adoGetDate
' database function returning the current date and time

Function adoGetDate
	If Application("ADO_DATABASETYPE") = "" Then Call adoDetermineType
	Select Case Application("ADO_DATABASETYPE")
		Case "MySQL" : adoGetDate = "CURRENT_DATE()"
		Case "Access" : adoGetDate = "Now()"
		Case Else : adoGetDate = "GetDate()"
	End Select
End Function

'--------------------------------------------------------------------
' adoTop
' return SQL for "Top X" after the "SELECT" keyword

Function adoTop(x)
	If Application("ADO_DATABASETYPE") = "" Then Call adoDetermineType
	Select Case Application("ADO_DATABASETYPE")
		Case "MySQL" : adoTop = ""
		Case "Access" : adoTop = "TOP " & x
		Case Else : adoTop = "TOP " & x
	End Select
End Function

'--------------------------------------------------------------------
' adoTop2
' return SQL for "LIMIT X" after the query statement (if nec)

Function adoTop2(x)
	If Application("ADO_DATABASETYPE") = "" Then Call adoDetermineType
	Select Case Application("ADO_DATABASETYPE")
		Case "MySQL" : adoTop2 = " LIMIT " & x
		Case "Access" : adoTop2 = ""
		Case Else : adoTop2 = ""
	End Select
End Function

'--------------------------------------------------------------------
' adoConcat
' return SQL to convert expressions to strings and concatenate them

Function adoConcat(arrParams)
	Dim sExpr, I
	If Not IsArray(arrParams) Then
		Response.Write("<P><B CLASS=""error"">adoConcat Error: Expected array parameter, got """ & TypeName(arrParams) & """<BR>" & Err.Description & "</B></P>")
		adoConcat = ""
		Exit Function
	End If
	If Application("ADO_DATABASETYPE") = "" Then Call adoDetermineType
	Select Case Application("ADO_DATABASETYPE")
		Case "MySQL" : sExpr = "CONCAT("
			For I = 0 To UBound(arrParams)
				If I > 0 Then sExpr = sExpr & ","
				sExpr = sExpr & arrParams(I)
			Next
			adoConcat = sExpr & ")"
		' Case "Access" : ' not implemented
		Case Else : 
			For I = 0 To UBound(arrParams)
				If I > 0 Then sExpr = sExpr & " + "
				sExpr = sExpr & arrParams(I)
			Next
			adoConcat = sExpr
	End Select
End Function

'--------------------------------------------------------------------
' adoFormatDateTime
' return SQL for "LIMIT X" after the query statement (if nec)

Function adoFormatDateTime(vValue, nFmt)

	On Error Resume Next
	adoFormatDateTime = FormatDateTime(vValue, nFmt)
	If Err.Number <> 0 Then
		adoFormatDateTime = "<i>n/a</i>"
	End If
	On Error Goto 0
End Function
%>