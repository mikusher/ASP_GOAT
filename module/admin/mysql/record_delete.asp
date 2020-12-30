<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' record_delete.asp
'	Delete a mysql data row from the database
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

Dim sAction
Dim sTable		' name of table containing record to delete
Dim oCol		' column definitions for the table
Dim sWhere		' where clause for selecting / updating
Dim sErrorMsg
Dim aType
Dim aQuote
Dim aSize
Dim aRecord

aType = Array("bigint", "blob", "char", "date", "datetime", "decimal", "double", "enum", "float", "int", "mediumint", "numeric", "real", "set", "smallint", "text", "time", "timestamp", "tinyint", "varchar", "year")
aQuote = Array("",		"'", 	"'",	"'",	"'",		"",			"",		"",		"",			"",		"",			"",			"",		"",	"",			"'",	"'",	"'",				"",			"'",	"")
aSize = Array(8,		0,		0,		3,		8,			0,			8,			0,		4,		4,		4,			0,			8,		0,		2,		0,		3,		4,				1,		0,			1)
'aUseSize = Array(0,		0,		1,		0,		0,			1,			0,			0,		0,		0,		0,			0,			0,		0,		0,		0,		0,		0,				0,		1,			0)
'aUsePrec = Array(0,		0,		0,		0,		0,			1,			0,			0,		0,		0,		0,			0,			0,		0,		0,		0,		0,		0,				0,		0,			0)
'aAutoInc = Array(1,		0,		0,		0,		0,			0,			0,			0,		0,		1,		1,			0,			0,		0,		1,		0,		0,		0,				1,		0,			0)

sAction = steForm("action")
sTable = steForm("table")
sErrorMsg = ""

If sAction = "delete" Then
	' perform the delete
	Call locGetTableDef(sTable, sErrorMsg)
	If sErrorMsg = "" Then
		If steNForm("Confirm") <> 1 Then
			sErrorMsg = sErrorMsg & "You must confirm the deletion of this record<br>"
			sAction = ""
		Else
			For I = 0 To oCol("Total") - 1
				If oCol("Key"&I) = 1 Then
					If sWhere <> "" Then sWhere = sWhere & " AND "
					sWhere = sWhere & oCol("Field"&I) & "="
					If aQuote(oCol("Type"&I)) <> "" Then
						sWhere = sWhere & aQuote(oCol("Type"&I)) & Replace(steForm(oCol("Field"&I)), aQuote(oCol("Type"&I)), "\" & aQuote(oCol("Type"&I))) & aQuote(oCol("Type"&I))
					Else
						sWhere = sWhere & steForm(oCol("Field"&I))
					End If
				End If
			Next
			' make sure we don't update the whole table
			If sWhere <> "" Then
				Call adoExecute("DELETE FROM " & sTable & " WHERE " & sWhere)
			Else
				sErrorMsg = sErrorMsg & "The where clause for the DELETE statement could not be built<br>"
			End If
		End If
	End If
End If

If (sAction <> "delete" Or sErrorMsg <> "") And sTable <> "" Then
	' first get the table definition
	If Not IsObject(oCol) Then Call locGetTableDef(sTable, sErrorMsg)
	' build the where clause for retrieving the row
	sWhere = locBuildWhere(oCol, sTable, sErrorMsg)
	' retrieve the row
	aRecord = locGetRow(sTable, sWhere, sErrorMsg)
End If
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Tables" %>
<!-- #include file="pagetabs_inc.asp" -->

<h4><%= sTable %> - Delete Record</h4>

<% If sAction <> "delete" Or sErrorMsg <> "" Then %>

<p>
<% steTxt "Please enter the changes to the MySQL record in the form below and click ""delete"" when you are done." %>
</p>

<% If sErrorMsg <> "" Then %>
<p><b class="error">ERROR - <%= sErrorMsg %></b></p>
<% End If %>

<% If IsArray(aRecord) And CStr(oCol("Total")) <> "" Then %>

<form method="post" action="record_delete.asp">
<input type="hidden" name="action" value="delete">
<input type="hidden" name="table" value="<%= sTable %>">

<table border=0 cellpadding=2 cellspacing=0 class="list">
<%	For I = 0 To oCol("Total") - 1
	If oCol("Key"&I) = "1" Then %>
	<input type="hidden" name="<%= oCol("Field"&I) %>" value="<%= Server.HTMLEncode(aRecord(I, 0)&"") %>">
<%	End If %>
<tr class="list<%= I mod 2 %>">
	<td valign="top" class="forml"><%= oCol("Field"&I) %></td><td>&nbsp;&nbsp;</td>
	<td valign="top" class="formd"><%= Server.HTMLEncode(aRecord(I, 0)&"") %></td>
</tr>
<%	Next %>
<tr class="list<%= I mod 2 %>">
	<td valign="top" class="forml">Are you sure?</td><td>&nbsp;&nbsp;</td>
	<td valign="top" class="formd">
		<input type="radio" name="confirm" value="1" class="formradio"> Yes
		<input type="radio" name="confirm" value="0" class="formradio"> No
	</td>
</tr>
</table>

<p align="center">
	<input type="submit" class="form" name="_delete" value="Delete">
</p>
</form>

<% Else %>
<p><b class="error">Record delete form could not be displayed here</b></p>
<% End If %>

<% Else %>

<p>
<% steTxt "Your deletion of the record was completed successfully." %>
</p>

<% End If %>

<p align="center">
	<a href="table_list.asp" class="adminlink"><% steTxt "Table List" %></A>&nbsp;
	<a href="table_browse.asp?table=<%= Server.URLEncode(sTable) %>" class="adminlink"><% steTxt "Table" %>&nbsp;<%= sTable %></a>
</p>

<!-- #include file="../../../footer.asp" -->
<%
Function locGetRow(sTable, sWhere, sErrorMsg)
	Dim rsResult, query

	query = "SELECT * FROM " & sTable & " WHERE " & sWhere
	On Error Resume Next
	Set rsResult = adoOpenRecordset(query)
	If Err.Number <> 0 Then
		sErrorMsg = sErrorMsg & "Error retrieving record to delete:<br>" & Err.Description & "<br>"
		locGetRow = Nothing
		Exit Function
	End If
	On Error Goto 0
	If rsResult.EOF Then
		sErrorMsg = sErrorMsg & "The record to delete could not be found:<br>" & query & "<br>"
		locGetRow = Nothing
	Else
		locGetRow = rsResult.GetRows
	End If
End Function
		
' build WHERE to pull the record based on primary keys
Function locBuildWhere(oCol, sTable, sErrorMsg)
	Dim sWhere, I

	' make sure the field definition exists
	If CStr(oCol("Total")) = "" Or Not IsNumeric(oCol("Total")) Then
		sErrorMsg = sErrorMsg & "No field definitions found for table """ & sTable & """<br>"
		locBuildWhere = ""
		Exit Function
	End If
	For I = 0 To oCol("Total") - 1
		If oCol("Key"&I) = 1 Then
			' check to see if the key value was passed for this field
			If steForm(oCol("Field"&I)) <> "" Then
				If sWhere <> "" Then sWhere = sWhere & " AND "
				sWhere = sWhere & oCol("Field"&I) & " = "
				If CStr(oCol("Type"&I)) <> "" Then
					If aQuote(oCol("Type"&I)) <> "" Then
						sWhere = sWhere & aQuote(oCol("Type"&I)) & Replace(steForm(oCol("Field"&I)), aQuote(oCol("Type"&I)), "\" & aQuote(oCol("Type"&I))) & aQuote(oCol("Type"&I))
					Else
						sWhere = sWhere & steForm(oCol("Field"&I))
					End If
				Else
					sWhere = sWhere & steForm(oCol("Field"&I))
				End If
			Else
				sErrorMsg = sErrorMsg & "No value passed for primary key field: """ & oCol("Field"&I) & """<br>"
				locBuildWhere = ""
				Exit Function
			End If
		End If
	Next
	If sWhere = "" Then
		sErrorMsg = sErrorMsg & "No primary key fields found in the table: """ & sTable & """<br>"
	End If
	locBuildWhere = sWhere
End Function

' retrieve the table definition and set the form variables
Function locLoadTableDef(rsCol, sField, nTypeID, nSize, nPrec, nNull, nInc, nKey, sDefault, sErrorMsg)
	Dim sDatatype, aParam, I

	sField = rsCol.Fields("Field").Value
	' determine the data type (size , precision)
	If InStr(1, rsCol.Fields("Type").Value, "(") > 0 Then
		sDatatype = Left(rsCol.Fields("Type").Value, InStr(1, rsCol.Fields("Type").Value, "(") - 1)
		aParam = Split(Replace(Mid(rsCol.Fields("Type").Value, InStr(1, rsCol.Fields("Type").Value, "(") + 1), ")", ""), ",")
		If UBound(aParam) = 0 Then
			nSize = aParam(0)
			nPrec = 0
		ElseIf UBound(aParam) = 1 Then
			nSize = aParam(0)
			nPrec = aParam(1)
		Else
			sErrorMsg = "locLoadTableDef - Error splitting the parameters for the datatype """ & Mid(rsCol.Fields("Type").Value, InStr(1, rsCol.Fields("Type").Value, "(")) & """"
			locLoadTableDef = False
			Exit Function
		End If
	Else
		sDatatype = rsCol.Fields("Type").Value
		nSize = 0
		nPrec = 0
	End If
	nTypeID = -1
	For I = 0 To UBound(aType)
		If StrComp(aType(I), sDatatype, vbTextCompare) = 0 Then nTypeID = CStr(I)
	Next

	If rsCol.Fields("Null").Value = "YES" Then nNull = "1" Else nNull = "0"
	If rsCol.Fields("Key").Value = "PRI" Then nKey = "1" Else nKey = "0"
	If rsCol.Fields("Extra").Value = "auto_increment" Then nInc = "1" Else nInc = "0"
	If Not IsNull(rsCol.Fields("Default").Value) Then
		sDefault = rsCol.Fields("Default").Value
	Else
		sDefault = ""
	End If
	locLoadTableDef = True
End Function

' load the table definition
Function locGetTableDef(sTable, sErrorMsg)
	Dim rsCol, sField, nTypeID, nSize, nPrec, nNull, nInc, nKey, sDefault

	Set rsCol = adoOpenRecordset("describe " & sTable)
	Set oCol = Server.CreateObject("Scripting.Dictionary")
	I = 0
	Do Until rsCol.EOF
		If Not locLoadTableDef(rsCol, sField, nTypeID, nSize, nPrec, nNull, nInc, nKey, sDefault, sErrorMsg) Then
			locBuildSQL = False
			Exit Function
		End If
		oCol("Field"&I) = sField
		oCol("Type"&I) = nTypeID
		oCol("Size"&I) = nSize
		oCol("Prec"&I) = nPrec
		oCol("Null"&I) = nNull
		oCol("Inc"&I) = nInc
		oCol("Key"&I) = nKey
		oCol("Default"&I) = sDefault
		rsCol.MoveNext
		I = I + 1
	Loop
	rsCol.Close
	Set rsCol = Nothing
	oCol("Total") = I
End Function
%>