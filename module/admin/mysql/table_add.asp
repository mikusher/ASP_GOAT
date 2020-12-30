<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' table_add.asp
'	Create a new mysql table in the database
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
Dim aType		' all of the mysql data types
Dim aSize		' fixed size for each data type
Dim aUseSize	' use the size when creating/altering the table
Dim aUsePrec	' use the precision when creating/altering the table
Dim aAutoInc	' allow the auto_increment modifier on the column def
Dim rsTab		' table to create
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

If sAction = "add" Then
	If Trim(steForm("tablename")) = "" Then
		sErrorMsg = "Please enter a name for the table"
	Else
		' make sure the table doesn't already exist
		If locTableExists(steForm("tablename")) Then
			sErrorMsg = "The Table Name you entered: """ & steForm("tablename") & """ already exists!"
		End If
	End If

	If sErrorMsg = "" Then
		' create the SQL that creates the database table
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

<h4>Create MySQL Table</h4>

<p>
Pleae enter the new table definition using the form below.  When
you are finished, click the <i>Create Table</i> button to finalize
your changes.
</p>

<% If sErrorMsg <> "" Then %>
<p><b class="error"><%= sErrorMsg %></b></p>
<% End If %>

<form method="post" action="table_add.asp">
<input type="hidden" name="action" value="add">

<table border=0 cellpadding=2 cellspacing=0 class="list" align="center">
<tr>
	<td align="right" class="forml">Table Name</td><td>&nbsp;&nbsp;</td>
	<td><input type="text" name="tablename" value="<%= steEncForm("tablename") %>" class="form" style="width:160px"></td>
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
<% For I = 0 To DEF_TOTALFIELDS-1 %>
<tr class="list<%= I mod 2 %>">
	<td><%= I +1 %></td>
	<td><input type="text" class="form" name="Field<%= I %>" size="20" maxlength="32" value="<%= steEncForm("Field"&I) %>" style="width:140px"></td>
	<td>
	<select name="Type<%= I %>" class="form" onchange="locTypeChange(<%= I %>, this.options[this.options.selectedIndex].value)">
	<option value=""> -- Choose One --
	<% For J = 0 To UBound(aType) %>
	<option value="<%= J %>"<% If steForm("Type"&I) = CStr(J) Then Response.Write " SELECTED" %>><%= aType(J) %>
	<% Next %>
	</select>
	</td>
	<td><input type="text" class="form" name="Size<%= I %>" size="4" maxlength="4" value="<%= steEncForm("Size"&I) %>" style="width:40px"></td>
	<td><input type="text" class="form" name="Prec<%= I %>" size="4" maxlength="2" value="<%= steEncForm("Prec"&I) %>" style="width:40px"></td>
	<td><input type="checkbox" class="form" name="Null<%= I %>" value="1"<% If steForm("Null"&I) = "1" Then Response.Write " CHECKED" %>></td>
	<td><input type="checkbox" class="form" name="Inc<%= I %>" value="1"<% If steForm("Inc"&I) = "1" Then Response.Write " CHECKED" %>></td>
	<td><input type="checkbox" class="form" name="Key<%= I %>" value="1"<% If steForm("Key"&I) = "1" Then Response.Write " CHECKED" %>></td>
	<td><input type="text" class="form" name="Default<%= I %>" size="20" maxlength="32" value="<%= steEncForm("Default"&I) %>" style="width:140px"></td>
</tr>
<% Next %>
</table>

<blockquote><font class="tinytext">
Size - Maximum number of characters or digits the field may hold<br>
Prec - Precision for numeric types (significant places after decimal)<br>
Null - check this box if the column should allow null values<br>
Null - check this box if the column should allow null values<br>
Inc - check this box to make the column an AUTO_INCREMENT value (int only)<br>
Key - check this box if the column makes up the primary key<br>
</blockquote>
<p align="center">
	<input type="submit" name="_submit" value="Create Table" class="form">
</p>

</form>

<% Else %>

<h4>New Table Created</h4>

<p>
The new table was successfully created in the database.  Please click on
the link below to continue administering the MySQL database.
</p>

<% End If %>

<p align="center">
	<a href="table_list.asp" class="adminlink"><% steTxt "Table List" %></A>
</p>

<!-- #include file="../../../footer.asp" -->
<%

Function locBuildKeyList
	Dim I, sKeys
	For I = 0 To DEF_TOTALFIELDS-1
		If Request("Key" & I) <> "" Then
			If sKeys <> "" Then sKeys = sKeys & ","
			sKeys = sKeys & Request.Form("Field" & I)
		End If
	Next
	locBuildKeyList = sKeys
End Function

Function locBuildSQL(sErrorMsg)
	Dim sql, bFirst, bRowErr, sRowSQL, sKeyList, bMultKeys, I

	sql = "create table " & steForm("tablename") & " (" & vbCrLf
	bFirst = False
	For I = 0 To DEF_TOTALFIELDS-1
		bRowErr = False
		If Request.Form("Field" & I) <> "" Then
			' check the data type and size parameters
			If Request.Form("Type" & I) = "" Then
				sErrorMsg = sErrorMsg & "Col " & (I+1) & ": Please enter a Data Type for the column: """ & Request.Form("Field" & I) & """<br>"
				bRowErr = True
			Else
				' create the row SQL - add the datatype(size,prec)
				nTypeID = steNForm("Type" & I)
				sRowSQL = Request.Form("Field" & I) & " " & aType(CInt(Request.Form("Type" & I)))
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
				If (Request.Form("Prec" & I) <> "" And aUsePrec(nTypeID) = 0) Then
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
			If Not bRowErr Then
				If bFirst Then
					sql = sql & "," & vbCrLf & vbTab & sRowSQL
				Else
					sql = sql & vbTab & sRowSQL
					bFirst = True
				End If
			End If
		End If
	Next

	' add the list of primary keys here
	sKeyList =  locBuildKeyList
	If sKeyList <> "" Then sql = sql & "," & vbCrLf & vbTab & "primary key (" & sKeyList & ")"
	locBuildSQL = sql & vbCrLf & ") Type=MyISAM;" & vbCrLf
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
