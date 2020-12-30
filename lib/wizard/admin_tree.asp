<%
'--------------------------------------------------------------------
' admin_tree.asp
'	This wizard will build an admin for a hierarchical list of
'	items.
'
' AUTH:	Ken Richards
' DATE:	08/02/2003
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

Dim strAction
Dim strTitle
Dim strObjectName
Dim strPrimaryKey
Dim strPrimaryKeyValue
Dim strParentField
Dim strParentLabel
Dim strTableName
Dim strDisplayFields
Dim strHeaderLabels
Dim strEditFields
Dim strEditLabels
Dim strEditSizes
Dim strEditTypes
Dim intArchive		' 1 = show archived, 0 = show unarchived
Dim intSelected		' primary key of selected item
Dim strCriteria		' database criteria to limit results
Dim strStatusMsg
Dim arrField
Dim arrLabel
Dim arrSize
Dim arrType

' populate the fields from the form / query collections
strAction = Trim(UCase(steForm("Action")))
'strTitle = steForm("Title")
'strObjectName = steForm("ObjectName")
'strPrimaryKey = steForm("PrimaryKey")
strPrimaryKeyValue = steForm("PrimaryKeyValue")
'strParentField = steForm("ParentField")
'strParentLabel = steForm("ParentLabel")
'strTableName = steForm("TableName")
'strEditFields = steForm("EditFields")
'strEditLabels = steForm("EditLabels")
'strEditSizes = steForm("EditSizes")
'strEditTypes = steForm("EditTypes")
'strDisplayFields = steForm("DisplayFields")
'strHeaderLabels = steForm("HeaderLabels")

' build the array of values here
arrField = Split(strEditFields, ",")
arrLabel = Split(strEditLabels, ",")
arrSize = Split(strEditSizes, ",")
arrType = Split(strEditTypes, ",")

' perform an action here
If sAction = "DOADD" Then
	Call treDoAdd
	strStatusMsg = "Your new " & strObjectName & " has been added"
ElseIf sAction = "DOUPDATE" Then
	Call treDoUpdate
	strStatusMsg = "The " & strObjectName & " record has been updated"
End If

' display the form to add / edit (if nec)
If sAction = "ADD" Or sAction = "UPDATE" Then
	' display the add / edit form here
	Call treForm
Else
	' display the admin list
	treTreeAdmin strTableName, strDisplayFields, strHeaderLabels, strCriteria, _
		intArchive, intSelected
End If

'--------------------------------------------------------------------
' treTreeOption
'	Build an individual list of options (calls itself recursively
'	to generate all of the child options)

Sub treTreeOption(nLevel, aRecord, nParentID, nSelected)
	Dim I, sHTML

	With Response
		For I = 0 To UBound(aRecord, 2)
			If nParentID = aRecord(2, I) Then
				.Write "<option value=""" & aRecord(0, I) & """"
				If aRecord(0, I) = nSelected Then .Write " SELECTED"
				.Write ">"
				If nLevel > 0 Then .Write String("&nbsp;&nbsp;&nbsp;", nLevel)
				.Write aRecord(1, I)
				.Write vbCrLf

				' check for any child options
				Call treTreeOption(nLevel+1, aRecord, aRecord(0, I), nSelected)
			End If
		Next
	End With
End Sub

'--------------------------------------------------------------------
' Build a hierarchical tree select

Sub treTreeSelect(sCriteria, nSelected)
	Dim sStat, rs, aRecord

	' fix the critera (where clause) for the select
	If Trim(sCriteria) <> "" And Not InStr(1, sCriteria, "AND ") Then
		sCriteria = " AND " & sCriteria
	Else
		sCriteria = " " & sCriteria
	End If

	sStat = "SELECT	" & strPrimaryKey & ", " & strParentLabel & ", " & strParentField & " " &_
			"FROM " & strTableName & " " &_
			"WHERE	Archive = 0 " &_
			"AND	Active = 1 " &_
			sCriteria
	Set rs = adoOpenRecordset(sStat)
	If rs.EOF Then
		Response.Write "<b class=""error"">NO RECORDS FOUND</b>"
		Exit Sub
	End If
	aRecord = rs.GetRows
	rs = ""

	' output the start of the select input
	With Response
		.Write "<select name=""" & strParentField & """>"
		.Write vbCrLf
		.Write "<option value=""""> -- Choose --"
		.Write vbCrLf
		Call treTreeOption(0, aRecord, 0, nSelected)
		.Write "</select>"
		.Write vbCrLf
	End With
End Sub

'--------------------------------------------------------------------
' treTreeAdminHeaders
'	Build the admin headers for the tree admin table

Sub treTreeAdminHeaders(aHeader)
	Dim I, J, sHTML

	With Response
		.Write "<TR BGCOLOR=""#E0C0A0"">" & vbCrLf
		For I = 0 To UBound(aHeader)
			.Write "<TD>"
			.Write aHeader(I)
			.Write "</TD>"
			.Write vbCrLf
		Next
		.Write "</TR>" & vbCrLf
	End With
End Sub

'--------------------------------------------------------------------
' treTreeAdminRows
'	Build the inidividual rows which make up the hierarchical admin

Sub treTreeAdminRows(nLevel, aRecord, aDisplay, nParentID, nRecNo, nSelected)
	Dim I, J, sHTML

	With Response
		For I = 0 To UBound(aRecord, 2)
			If nParentID = aRecord(2, I) Then
				.Write "<tr class=""tree"
				If aRecord(0, I) = nSelected Then
					.Write "sel"
				Else
					.Write nRecNo Mod 2
				End If
				.Write ">"
				.Write vbCrLf
				For J = 0 To UBound(aDisplay)
					' perform indenting on the first column
					If J = 0 Then
						.Write vbTab & "<td nowrap><table border=0 cellpadding=0 cellspacing=0><tr><td nowrap>"
						.Write String("&nbsp;&nbsp;&nbsp;", nLevel)
						.Write "</td><td nowrap>"
						If Not Trim(aRecord(J+1, I) & "") <> "" Then
							.Write Server.HTMLEncode(aRecord(J+1, I))
						Else
							.Write "&nbsp;"
						End If
						.Write "</td></tr></table></td>"
					Else
						.Write vbTab & "<td>"
						If Not Trim(aRecord(J+1, I) & "") <> "" Then
							.Write Server.HTMLEncode(aRecord(J+1, I))
						Else
							.Write "&nbsp;"
						End If
						.Write "</td>"
					End If
				Next
				' write the action links for the record
				.Write "<td><a href=""?ID="
				.Write aRecord(0, I)
				.Write "&action=edit"" class=""actionlink"">edit</A> . <a href=""?ID="
				.Write aRecord(0, I)
				.Write "&action=delete"" class=""actionlink"">delete</A></td>"
				.Write vbCrLf
				' .Write Request.ServerVariables("SCRIPT_NAME")
				.Write "</tr>"
				.Write vbCrLf

				' check for any child options
				nRecNo = nRecNo + 1
				Call treTreeAdminRows(nLevel+1, aRecord, aDisplay, aRecord(0, I), nRecNo, nSelected)
			End If
		Next
	End With
End Sub

'--------------------------------------------------------------------
' Build a hierarchical tree admin page

Sub treTreeAdmin(sTableName, sDisplayFields, sHeaderLabels, sCriteria, _
		nArchive, nSelected)
	Dim sStat, rs, aRecord, aDisplay, aHeader


	' fix the critera (where clause) for the select
	If Trim(sCriteria) <> "" And Not InStr(1, sCriteria, "AND ") Then
		sCriteria = " AND " & sCriteria
	Else
		sCriteria = " " & sCriteria
	End If

	' build the display array and trim whitespace
	aDisplay = Split(sDisplayFields, ",")
	For I = 0 To UBound(aDisplay)
		aDisplay(I) = Trim(aDisplay(I))
	Next

	aHeader = Split(sHeaderLabels, ",")
	For I = 0 To UBound(aHeader)
		aHeader(I) = Trim(aHeader(I))
	Next

	If CStr(nArchive) = "" Then nArchive = 0
	sStat = "SELECT	" & strPrimaryKey & ", " & strDisplayFields & ", " & strParentField & " " &_
			"FROM " & sTableName & " " &_
			"WHERE	Archive = " & nArchive &_
			"AND	Active = 1 " &_
			sCriteria
	Set rs = adoOpenRecordset(sStat)
	If rs.EOF Then
		Response.Write "<b class=""error"">Nothing has been defined</b>"
		Exit Sub
	End If
	aRecord = rs.GetRows
	rs = ""

	' output the tree admin table
	With Response
		' display the table header
		.Write "<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0 BGCOLOR=""#A08040"">"
		.Write vbCrLf
		.Write "<TR><TD>"
		.Write vbCrLf
		.Write "<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 BGCOLOR=""#FFFFFF"">"
		.Write vbCrLf

		Call treTreeAdminHeaders(aHeader)

		Call treTreeAdminRows(0, aRecord, aDisplay, 0, 0, nSelected)

		.Write "</TABLE>"
		.Write vbCrLf
		.Write "</TD></TR>"
		.Write vbCrLf
		.Write "</TABLE>"
		.Write vbCrLf
	End With
End Sub

'--------------------------------------------------------------------
' Process the tree "add form" here

Sub treDoAdd
	Dim query, sValues, bCreated, I

	query = "INSERT INTO " & strTableName & " (" & strParentField

	sValues = CStr(steNForm(strParentField))

	bCreated = False
	For I = 0 To UBound(arrField)
		query = query & ","
		sValues = sValues & ","
		query = query & arrField(I)
		sValues = sValues & "'" & Replace(steForm(arrField(I)), "'", "''") & "'"
		If arrField(I) = "Created" Then bCreated = True
	Next
	' add the created field to the insert query (if nec)
	If Not bCreated Then
		query = query & ",Created"
		sValues = sValues & "," & adoGetDate
	End If
	query = query & ") VALUES (" & sValues & ")"

	' execute the insert statement here
	Call adoExecute(query)
End Sub

'--------------------------------------------------------------------
' Process the tree "update form" here

Sub treDoUpdate
	Dim query, sValues, I

	query = "UPDATE " & strTableName & " SET " &_
		strParentField & " = " & steNForm(strParentField)

	For I = 0 To UBound(arrField)
		query = query & ", " & arrField(I) & " = '" & Replace(steForm(arrField(I)), "'", "''") & "'"
	Next
	query = query & " WHERE " & strPrimaryKey & " = " & steNForm(strPrimaryKey)

	' execute the insert statement here
	Call adoExecute(query)
End Sub
'--------------------------------------------------------------------
' Retrieve the record to edit (for edit mode)
' Assumes an integer primary key

Function treEditRS
	Dim sSelect, I

	sSelect = strParentField
	For I = 0 To UBound(arrField)
		sSelect = sSelect & "," & arrField(I)
	Next
	query = "SELECT " & sSelect & " " &_
			"FROM " & strTableName & " " &_
			"WHERE	" & strPrimaryKey & " = " & strPrimaryKeyValue
	Set treEditRS = adoOpenRecordset(query)
End Function

'--------------------------------------------------------------------
' Build an "Add New Item" to tree form

Sub treForm(sCriteria, nSelected)
	Dim aField, aLabel, aType, aSize, nSize, nCol, I
	Dim rsEdit

	aField = Split(sEditFields, ",")
	aLabel = Split(sEditLabels, ",")
	aType = Split(sEditTypes, ",")
	aSize = Split(sEditSizes, ",")

	' get the record to edit (if nec)
	If strPrimaryKeyValue <> 0 Then Set rsEdit = treEditRS

	With Response
	.Write "<FORM METHOD=""post"" ACTION=""?action=doadd"">" & vbCrLf
	.Write "<INPUT TYPE=""hidden"" NAME=""" & strPrimaryKey & """ VALUE=""" & strPrimaryKeyValue & """>" & vbCrLf
	.Write "<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>" & vbCrLf
	.Write "<TR>" & vbCrLf
	.Write "	<TD COLSPAN=""2""><B>" & sParentLabel & "</B><BR>" & vbCrLf
	Call treTreeSelect(strParentField, sTableName, strPrimaryKey, strParentField, strParentLabel, sCriteria, _
		nSelected)
	.Write "	</TD>" & vbCrLf
	.Write "</TR>" & vbCrLf

	' build all of the individual edit fields
	nCol = 0
	For I = 0 To UBound(aField)
		iF aSize(I) > 44 Then nSize = 44 Else nSize = aSize(I)
		If nCol mod 2 = 0 And nCol > 0 Then
			.Write "</tr><tr>"
		End If
		.Write "<td>"
		.Write "<b>" & aLabel(I) & "</b><br>"
		Select Case aType(I)
			Case "T"
				.Write "<input type=""text"" name=""" & aField(I) & """ value=""" &_
					steRecordEncValue(rsEdit, aField(I)) & """ size=""" &_
					nSize & """ maxlength=""" & aSize(I) & """>"
			Case "A"
				If nCol mod 2 = 1 Then
					.Write "</tr><tr>" & vbCrLf
				Else ' force a new row after this
					nCol = nCol + 1
				End If
				.Write "<textarea name=""" & aField(I) & """ cols=""38"" rows=""10"">" &_
					steRecordEncValue(rsEdit, aField(I)) & "</textarea>"
			Case "C"
				.Write "<input type=""checkbox"" name=""" & aField(I) & """ value=""" &_
					steRecordEncValue(rsEdit, aField(I)) & """>"
			Case "R"
				If steNForm(aField(I)) <> 0 Then sChecked = " CHECKED" Else sChecked = ""
				.Write "<input type=""radio"" name=""" & aField(I) & """ value=""1""" & sChecked & """>"
			Case "Y"
				If steNForm(aField(I)) <> 0 Then sChecked = " CHECKED" Else sChecked = ""
				.Write "<input type=""radio"" name=""" & aField(I) & """ value=""1""" & sChecked & """> Yes "
				If steNForm(aField(I)) = 0 Then sChecked = " CHECKED" Else sChecked = ""
				.Write "<input type=""radio"" name=""" & aField(I) & """ value=""0""" & sChecked & """> No "
		End Select
		.Write "</td>"
		nCol = nCol + 1
	Next
	' build the submit button here
	.Write "</tr><tr>" & vbCrLf
	.Write "	<td colspan=""2"" align=""right"">" & vbCrLf
	.Write "	<input type=""reset"" name=""tre_reset"" value="" Reset ""> &nbsp; " & vbCrLf
	If strPrimaryKeyValue <> 0 Then
		.Write "	<input type=""hidden"" name=""action"" value=""doupdate"">" & vbCrLf
		.Write "	<input type=""submit"" name=""tre_submit"" value="" Update "">" & vbCrLf
	Else
		.Write "	<input type=""hidden"" name=""action"" value=""doadd"">" & vbCrLf
		.Write "	<input type=""submit"" name=""tre_submit"" value="" Add "">" & vbCrLf
	End If
	.Write "</tr>" & vbCrLf
	.Write "</table>" & vbCrLf
	End With
End Sub
%>