<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' record_add.asp
'	Add a mysql data row from the database
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
Dim sTable		' name of table containing record to add
Dim oCol		' column definitions for the table
Dim sWhere		' where clause for selecting / updating
Dim sErrorMsg
Dim aType
Dim aQuote
Dim aSize

aType = Array("bigint", "blob", "char", "date", "datetime", "decimal", "double", "enum", "float", "int", "mediumint", "numeric", "real", "set", "smallint", "text", "time", "timestamp", "tinyint", "varchar", "year")
aQuote = Array("",		"'", 	"'",	"'",	"'",		"",			"",		"",		"",			"",		"",			"",			"",		"",	"",			"'",	"'",	"'",				"",			"'",	"")
aSize = Array(8,		0,		0,		3,		8,			0,			8,			0,		4,		4,		4,			0,			8,		0,		2,		0,		3,		4,				1,		0,			1)
'aUseSize = Array(0,		0,		1,		0,		0,			1,			0,			0,		0,		0,		0,			0,			0,		0,		0,		0,		0,		0,				0,		1,			0)
'aUsePrec = Array(0,		0,		0,		0,		0,			1,			0,			0,		0,		0,		0,			0,			0,		0,		0,		0,		0,		0,				0,		0,			0)
'aAutoInc = Array(1,		0,		0,		0,		0,			0,			0,			0,		0,		1,		1,			0,			0,		0,		1,		0,		0,		0,				1,		0,			0)

sAction = steForm("action")
sTable = steForm("table")
sErrorMsg = ""

If sAction = "add" Then
	' perform the add
	Call locGetTableDef(sTable, sErrorMsg)
	If sErrorMsg = "" Then
		Dim sInsert, sValues

		For I = 0 To oCol("Total") - 1
			If oCol("Key"&I) = "1" And oCol("Inc"&I) <> "1" And steForm(oCol("Field"&I)) = "" Then
				sErrorMsg = sErrorMsg = "You must enter a value for the field """ & oCol("Field"&I) & """<br>"
			ElseIf Not (oCol("Key"&I) = "1" And oCol("Inc"&I) = "1") Then
				' add the insert field to the list
				If sInsert <> "" Then sInsert = sInsert & ", "
				sInsert = sInsert & oCol("Field"&I)
				' add the value to the list
				If sValues <> "" Then sValues = sValues & ", "
				If oCol("Null"&I) = "1" And steForm(oCol("Field"&I)) = "" Then
					sValues = sValues & "NULL"
				ElseIf aQuote(oCol("Type"&I)) <> "" Then
					sValues = sValues & aQuote(oCol("Type"&I)) & Replace(steForm(oCol("Field"&I)), aQuote(oCol("Type"&I)), "\" & aQuote(oCol("Type"&I))) & aQuote(oCol("Type"&I))
				Else
					sValues = sValues & steForm(oCol("Field"&I))
				End If
			End If
		Next
		' make sure we have a valid list of insert fields / values
		If sInsert <> "" And sValues <> "" Then
			' Response.Write "INSERT INTO " & sTable & " (" & sInsert & ") VALUES (" & sValues & ")" : Response.End
			Call adoExecute("INSERT INTO " & sTable & " (" & sInsert & ") VALUES (" & sValues & ")")
		Else
			sErrorMsg = sErrorMsg & "The field names and values could not be built to insert<br>"
		End If		
	End If
End If

If (sAction <> "add" Or sErrorMsg <> "") And sTable <> "" Then
	' first get the table definition
	If Not IsObject(oCol) Then Call locGetTableDef(sTable, sErrorMsg)
End If
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Tables" %>
<!-- #include file="pagetabs_inc.asp" -->

<h4><%= sTable %> - Add New Record</h4>

<% If sAction <> "add" Or sErrorMsg <> "" Then %>

<p>
<% steTxt "Please enter the new database record in the form below and click ""add"" when you are done." %>
</p>

<% If sErrorMsg <> "" Then %>
<p><b class="error">ERROR - <%= sErrorMsg %></b></p>
<% End If %>

<% If CStr(oCol("Total")) <> "" Then %>

<form method="post" action="record_add.asp">
<input type="hidden" name="action" value="add">
<input type="hidden" name="table" value="<%= sTable %>">

<table border=0 cellpadding=2 cellspacing=0 class="list">
<%	For I = 0 To oCol("Total") - 1 %>
<tr class="list<%= I mod 2 %>">
	<td valign="top" class="forml"><%= oCol("Field"&I) %></td><td>&nbsp;&nbsp;</td>
	<td valign="top" class="formd">
	<% If oCol("Key"&I) = "1" And oCol("Inc"&I) = "1" Then %>
	AUTO_INCREMENT
	<% ElseIf aType(oCol("Type"&I)) = "text" Or aType(oCol("Type"&I)) = "blob" Then %>
	<textarea class="form" name="<%= oCol("Field"&I) %>" cols="42" rows="10"><%= steEncForm(oCol("Field"&I)) %></textarea>
	<% Else
		 %>
	<input type="text" class="form" name="<%= oCol("Field"&I) %>" value="<%= steEncForm(oCol("Field"&I)) %>" size="<% If oCol("Size"&I) > 48 Then Response.Write "48" Else Response.Write oCol("Size"&I) %>" maxlength="<%= oCol("Size"&I) %>">
	<% End If %>
	</td>
</tr>
<%	Next %>
</table>

<p align="center">
	<input type="submit" class="form" name="_add" value="Update">
</p>
</form>

<% Else %>
<p><b class="error">Record add form could not be displayed here</b></p>
<% End If %>

<% Else %>

<p>
<% steTxt "Your new record was added to the database table successfully." %>
</p>

<% End If %>

<p align="center">
	<a href="table_list.asp" class="adminlink"><% steTxt "Table List" %></A>&nbsp;
	<a href="table_browse.asp?table=<%= Server.URLEncode(sTable) %>" class="adminlink"><% steTxt "Table" %>&nbsp;<%= sTable %></a>
</p>

<!-- #include file="../../../footer.asp" -->
<%
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