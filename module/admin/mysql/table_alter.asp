<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' table_alter.asp
'	Alter a mysql table definition in the database
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
Dim sTable		' name of table to alter
Dim aType		' all of the mysql data types
Dim aSize		' fixed size for each data type
Dim aUseSize	' use the size when creating/altering the table
Dim aUsePrec	' use the precision when creating/altering the table
Dim aAutoInc	' allow the auto_increment modifier on the column def
Dim rsTab		' table to alter
Dim sKeyList	' list of existing primary keys
Dim sErrorMsg
Dim sql
Dim I, J

Const DEF_TOTALFIELDS = 25
aType = Array("bigint", "blob", "char", "date", "datetime", "decimal", "double", "enum", "float", "int", "mediumint", "numeric", "real", "set", "smallint", "text", "time", "timestamp", "tinyint", "varchar", "year")
aSize = Array(8,		0,		0,		3,		8,			0,			8,			0,		4,		4,		4,			0,			8,		0,		2,		0,		3,		4,				1,		0,			1)
aUseSize = Array(0,		0,		1,		0,		0,			1,			0,			0,		0,		0,		0,			0,			0,		0,		0,		0,		0,		0,				0,		1,			0)
aUsePrec = Array(0,		0,		0,		0,		0,			1,			0,			0,		0,		0,		0,			0,			0,		0,		0,		0,		0,		0,				0,		0,			0)
aAutoInc = Array(1,		0,		0,		0,		0,			0,			0,			0,		0,		1,		1,			0,			0,		0,		1,		0,		0,		0,				1,		0,			0)

sAction = LCase(steForm("action"))
sTable = steForm("table")
sKeyList = steForm("keylist")

If sAction = "alter" Then
	' make sure the table doesn't already exist
	If Not locTableExists(steForm("table")) Then
		sErrorMsg = "The selected table (""" & steForm("table") & """) does not exist!"
	End If

	If sErrorMsg = "" Then
		' build the SQL that alters the database table
		sql = locBuildSQL(sErrorMsg)
		If sErrorMsg = "" Then
			' Response.Write Replace(Server.HTMLEncode(sql), vbCrLf, "<br>") : Response.End
			Call adoExecute(sql)
		End If
	End If
End If
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Tables" %>
<!-- #include file="pagetabs_inc.asp" -->
<script language="Javascript">
<!-- // hide
function locTypeChange(nVar, nType) {
<% For I = 0 To UBound(aType)
	If aSize(I) <> 0 Then %>
	if (nType == '<%= I %>') { eval('document.all.Size'+nVar).value = '<%= aSize(I) %>'; }
<%	End If
   Next %>
}
// unhide -->
</script>

<% If sAction <> "add" Or sErrorMsg <> "" Then %>

<h4>Alter MySQL Table</h4>

<p>
Pleae enter the new table definition using the form below.  When
you are finished, click the <i>Alter Table</i> button to finalize
your changes.
</p>

<% If sErrorMsg <> "" Then %>
<p><b class="error"><%= sErrorMsg %></b></p>
<% End If %>

<form method="post" action="table_alter.asp">
<input type="hidden" name="action" value="alter">
<input type="hidden" name="table" value="<%= steEncForm("table") %>">

<table border=0 cellpadding=2 cellspacing=0 class="list" align="center">
<tr>
	<td align="right" class="forml">Table Name</td><td>&nbsp;&nbsp;</td>
	<td class="formd"><%= steEncForm("table") %></td>
</tr>
</table><br>

<table border=0 cellpadding=2 cellspacing=0 class="list">
<tr>
	<td class="listhead">Col</td>
	<td class="listhead">Field Name</td>
	<td class="listhead">Data Type</td>
	<td class="listhead">Size</td>
	<td class="listhead">Prec</td>
	<td class="listhead">Null</td>
	<td class="listhead">Inc</td>
	<td class="listhead">Key</td>
	<td class="listhead">Default</td>
</tr>
<%	Set rsCol = adoOpenRecordset("describe " & sTable)
	Dim sField, nTypeID, nSize, nPrec, nNull, nInc, nKey, sDefault
	sKeyList = ""
	For I = 0 To DEF_TOTALFIELDS-1
		' load the form values for this row
		If sAction = "" And Not rsCol.EOF Then
			Call locLoadTableDef(rsCol, sField, nTypeID, nSize, nPrec, nNull, nInc, nKey, sDefault, sErrorMsg)
			If nKey = "1" Then sKeyList = sKeyList & "," & sField
			rsCol.MoveNext
		Else
			sField = steForm("Field"&I)
			nTypeID = steForm("Type"&I)
			nSize = steForm("Size"&I)
			nPrec = steForm("Prec"&I)
			nNull = steForm("Null"&I)
			nInc = steForm("Inc"&I)
			nKey = steForm("Key"&I)
			sDefault = steForm("Default"&I)
		End If %>
<tr class="list<%= I mod 2 %>">
	<td><%= I +1 %></td>
	<td><input type="text" class="form" name="Field<%= I %>" size="20" maxlength="32" value="<%= Server.HTMLEncode(sField) %>" style="width:140px"></td>
	<td>
	<select name="Type<%= I %>" class="form" onchange="locTypeChange(<%= I %>, this.options[this.options.selectedIndex].value)">
	<option value=""> -- Choose One --
	<% For J = 0 To UBound(aType) %>
	<option value="<%= J %>"<% If nTypeID = CStr(J) Then Response.Write " SELECTED" %>><%= aType(J) %>
	<% Next %>
	</select>
	</td>
	<td><input type="text" class="form" name="Size<%= I %>" size="4" maxlength="4" value="<%= nSize %>" style="width:40px"></td>
	<td><input type="text" class="form" name="Prec<%= I %>" size="4" maxlength="2" value="<%= nPrec %>" style="width:40px"></td>
	<td><input type="checkbox" class="form" name="Null<%= I %>" value="1"<% If nNull = "1" Then Response.Write " CHECKED" %>></td>
	<td><input type="checkbox" class="form" name="Inc<%= I %>" value="1"<% If nInc = "1" Then Response.Write " CHECKED" %>></td>
	<td><input type="checkbox" class="form" name="Key<%= I %>" value="1"<% If nKey = "1" Then Response.Write " CHECKED" %>></td>
	<td><input type="text" class="form" name="Default<%= I %>" size="20" maxlength="32" value="<%= Server.HTMLEncode(sDefault) %>" style="width:140px"></td>
</tr>
<%	Next
	If Not (rsCol Is Nothing) Then
		rsCol.Close
		Set rsCol = Nothing
	End If%>
</table>
<input type="hidden" name="keylist" value="<%= Server.HTMLEncode(sKeyList) %>">

<blockquote><font class="tinytext">
Size - Maximum number of characters or digits the field may hold<br>
Prec - Precision for numeric types (significant places after decimal)<br>
Null - check this box if the column should allow null values<br>
Null - check this box if the column should allow null values<br>
Inc - check this box to make the column an AUTO_INCREMENT value (int only)<br>
Key - check this box if the column makes up the primary key<br>
</blockquote>
<p align="center">
	<input type="submit" name="_submit" value="Alter Table" class="form">
</p>

</form>

<% Else %>

<h4>Table Altered Successfully</h4>

<p>
The MySQL table was successfully altered in the database.  Please click on
the link below to continue administering the MySQL database.
</p>

<% End If %>

<p align="center">
	<a href="table_list.asp" class="adminlink"><% steTxt "Table List" %></A>
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

Function locBuildKeyList(bChanged)
	Dim I, sKeys, aPKey, nKeyCount, nNewKeyCount

	If sKeyList <> "" Then
		aPKey = Split(Mid(sKeyList, 2), ",")
		nKeyCount = UBound(aPKey) + 1
	Else
		nKeyCount = 0
	End If
	bChanged = False
	nNewKeyCount = 0
	For I = 0 To DEF_TOTALFIELDS-1
		If Request("Key" & I) <> "" Then
			If sKeys <> "" Then sKeys = sKeys & ","
			sKeys = sKeys & Request.Form("Field" & I)
			If Not (InStr(1, sKeyList & ",", "," & Request.Form("Field" & I) & ",") > 0) Then bChanged = True
			nNewKeyCount = nNewKeyCount + 1
		End If
	Next
	If nNewKeyCount <> nKeyCount Then bChanged = True
	locBuildKeyList = sKeys
End Function

Function locBuildSQL(sErrorMsg)
	Dim sql, bFirst, bRowErr, sRowSQL, sKeyList, bMultKeys, bChange, bAdd, I
	Dim sField, nTypeID, nSize, nPrec, nNull, nInc, nKey, oCol, sDefault
	Dim rsCol, aCol, nChangeNo, bKeysChanged, bFound

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

	' check to see if the primary keys have changed
	bKeysChanged = False
	For I = 0 To DEF_TOTALFIELDS-1
		If Request.Form("Field" & I) <> "" Then
			bFound = False
			For J = 0 To oCol("Total") - 1
				If StrComp(oCol("Field" & J), Request.Form("Field" & I), vbTextCompare) = 0 Then
					bFound = True
					If oCol("Key" & J) <> Request.Form("Key" & I) Then bKeysChanged = True
				End if
			Next
			If Request.Form("Key" & I) = "1" And Not bFound Then bKeysChanged = True
		End If
	Next
	If Not bKeysChanged Then
		' check for any deleted primary key columns here
		For J = 0 To oCol("Total") - 1
			If oCol("Key" & J) = "1" Then
				bFound = False
				For I = 0 To DEF_TOTALFIELDS-1
					If Request.Form("Field" & I) <> "" Then
						If StrComp(oCol("Field" & J), Request.Form("Field" & I), vbTextCompare) = 0 Then
							bFound = True
							If oCol("Key" & J) <> Request.Form("Key" & I) Then bKeysChanged = True
						End If
					End If
				Next
				If Not bFound Then bKeysChanged = True
			End If
		Next
	End If

	For I = 0 To DEF_TOTALFIELDS-1
		If Request.Form("Field" & I) <> "" Then
			For J = 0 To oCol("Total") - 1
				If StrComp(oCol("Field" & J), Request.Form("Field" & I), vbTextCompare) = 0 Then
					If oCol("Key" & J) <> Request.Form("Key" & I) Then bKeysChanged = True
				End if
			Next
		End If
	Next
		
	sql = "alter table " & steForm("table") & " " & vbCrLf
	bFirst = False
	For I = 0 To DEF_TOTALFIELDS-1
		bRowErr = False
		bAdd = False
		bChange = False
		If Request.Form("Field" & I) <> "" Then
			' check the data type and size parameters
			If Request.Form("Type" & I) = "" Then
				sErrorMsg = sErrorMsg & "Col " & (I+1) & ": Please enter a Data Type for the column: """ & Request.Form("Field" & I) & """<br>"
				bRowErr = True
			Else
				' check to see if anything has changed
				bFound = False
				For J = 0 To oCol("Total") - 1
					If StrComp(oCol("Field" & J), Request.Form("Field" & I), vbTextCompare) = 0 Then
						' found the field - compare the configuration
						bFound = True
						If oCol("Type" & J)&"" <> Request.Form("Type" & I) Or _
							oCol("Size" & J)&"" <> Request.Form("Size" & I) Or _
							oCol("Prec" & J)&"" <> Request.Form("Prec" & I) Or _
							CInt(oCol("Null" & J)) <> steNForm("Null" & I) Or _
							CInt(oCol("Inc" & J)) <> steNForm("Inc" & I) Or _
							oCol("Default" & J)&"" <> Request.Form("Default" & I) Then
							Response.Write "Type (" & (oCol("Type" & J)&"") & "<>" & Request.Form("Type" & I) & ") Size (" & (oCol("Size" & J)&"") & "<>" & Request.Form("Size" & I) & ") Prec (" & (oCol("Prec" & J)&"") & "<>" & Request.Form("Prec" & I) & ") Null (" & (oCol("Null" & J)&"") & "<>" & Request.Form("Null" & I) & ") Inc (" & (oCol("Inc" & J)&"") & "<>" & Request.Form("Inc" & I) & ") Default (" & (oCol("Default" & J)&"") & "<>" & Request.Form("Default" & I) & ")<BR>"
							nChangeNo = J
							bChange = True
						End If
					End If
				Next
				If Not bFound Then bAdd = True

				If bAdd or bChange Then
					' create the row SQL - add the datatype(size,prec)
					nTypeID = steNForm("Type" & I)
					If bAdd Then
						sRowSQL = Request.Form("Field" & I) & " " & aType(CInt(Request.Form("Type" & I)))
					Else
						sRowSQL = Request.Form("Field" & I) & " " & Request.Form("Field" & I) & " " & aType(CInt(Request.Form("Type" & I)))
					End If
					'If (Request.Form("Size" & I) <> "" And aUseSize(nTypeID) = 0) Then
					'	sErrorMsg = sErrorMsg & "Col " & (I+1) & ": The size parameter is not required for column: """ & Request.Form("Field" & I) & """<br>"
					'	bRowErr = True
					'End If
					If (Request.Form("Size" & I) = "" And aUseSize(nTypeID) <> 0) Then
						sErrorMsg = sErrorMsg & "Col " & (I+1) & ": Please enter the size parameter for column: """ & Request.Form("Field" & I) & """<br>"
						bRowErr = True
					End If
					If (Not bRowErr And aUseSize(nTypeID) <> 0) Then
						sRowSQL = sRowSQL & "(" & Request.Form("Size" & I)
					End If
					If (Request.Form("Prec" & I) <> "" And Request.Form("Prec" & I) <> "0" And aUsePrec(nTypeID) = 0) Then
						sErrorMsg = sErrorMsg & "Col " & (I+1) & ": The precision is not required for column: """ & Request.Form("Field" & I) & """<br>"
						bRowErr = True
					End If
					If (Request.Form("Prec" & I) = "" And aUsePrec(nTypeID) <> 0) Then
						sErrorMsg = sErrorMsg & "Col " & (I+1) & ": Please enter the precision parameter for column: """ & Request.Form("Field" & I) & """<br>"
						bRowErr = True
					End If
					If (Not bRowErr And aUsePrec(nTypeID) <> 0) Then
						sRowSQL = sRowSQL & "," & Request.Form("Prec" & I)
					End If
					If (Not bRowErr And aUseSize(nTypeID) <> 0) Then sRowSQL = sRowSQL & ")"
				End If
			End If

			' continue building the SQL if on add or change
			If bAdd Or bChange Then
				' add the null / not null constraint
				If (Request.Form("Null" & I) <> "") Then
					sRowSQL = sRowSQL & " null"
				Else
					sRowSQL = sRowSQL & " not null"
				End If
	
				' add the auto increment parameter (if nec)
				If (Request.Form("Inc" & I) <> "" And aAutoInc(nTypeID) = 0) Then
					sErrorMsg = sErrorMsg & "Col " & (I+1) & ": Datatype " & aType(nTypeID) &  """ does not allow Auto Increment<br>"
					bRowErr = True
				ElseIf Request.Form("Inc" & I) <> "" Then
					sRowSQL = sRowSQL & " auto_increment"
				End If
	
				' add the default value statement here
				If Request.Form("Default" & I) <> "" Then
					sRowSQL = sRowSQL & " default " & Request.Form("Default" & I)
				End If
				If bChange Then
					If Not bRowErr Then
						If bFirst Then
							sql = sql & "," & vbCrLf & vbTab & "change column " & sRowSQL
						Else
							sql = sql & vbTab & "change column " & sRowSQL
							bFirst = True
						End If
					End If
				Else ' bAdd is true
					If Not bRowErr Then
						If bFirst Then
							sql = sql & "," & vbCrLf & vbTab & "add column " & sRowSQL
						Else
							sql = sql & vbTab & "add column " & sRowSQL
							bFirst = True
						End If
					End If
				End If
			End If
		End If
	Next

	' check for dropped columns here
	For I = 0 To oCol("Total")-1
		bFound = False
		For J = 0 To DEF_TOTALFIELDS-1
			If StrComp(oCol("Field"&I), Request.Form("Field" & I), vbTextCompare) = 0 Then
				bFound = True
				Exit For
			End If
		Next
		If Not bFound Then
			If bFirst Then
				sql = sql & "," & vbCrLf & vbTab & "drop column " & oCol("Field"&I)
			Else
				sql = sql & vbTab & "drop column " & oCol("Field"&I)
				bFirst = True
			End If
		End If
	Next

	' add the list of primary keys here
	sKeyList =  locBuildKeyList(bKeysChanged)
	If bKeysChanged Then
		If bFirst Then sql = sql & "," & vbCrLf & vbTab
		sql = sql & "drop primary key"
		If sKeyList <> "" Then sql = sql & "," & vbCrLf & vbTab & "add primary key (" & sKeyList & ")"
	End If
	locBuildSQL = sql & vbCrLf & ";" & vbCrLf
End Function

Function locTableExists(sTable)
	Dim rsTab
	Set rsTab = adoOpenRecordset("show tables;")
	Do Until rsTab.EOF
		If StrComp(rsTab.Fields(0).Value, sTable, vbTextCompare) = 0 Then
			locTableExists = True
			rsTab.Close
			Exit Function
		End If
		rsTab.MoveNext
	Loop
	rsTab.Close
	locTableExists = False
End Function
%>
