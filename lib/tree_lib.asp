<%
'--------------------------------------------------------------------
' tree_lib.asp
'	This library includes functions for working with hierarchical
'	stored in a database such as the table-of-contents for a book.
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

Dim treStatusMsg		' status to display to user
Dim treAddlLinks		' additional footer links to display

'--------------------------------------------------------------------
' treTreeOption
'	Build an individual list of options (calls itself recursively
'	to generate all of the child options)

Sub treTreeOption(nEditID, nLevel, aRecord, nParentID, nSelected)
	Dim I, J, sValue, sHTML

	With Response
		For I = 0 To UBound(aRecord, 2)
			If nParentID = aRecord(2, I) Then
				.Write "<option value=""" & aRecord(0, I) & """"
				If aRecord(0, I) = nSelected Then .Write " SELECTED"
				.Write ">"
				If nLevel > 0 Then
					For J = 1 To nLevel
						.Write "&nbsp;&nbsp;&nbsp;"
					Next
				End If
				.Write aRecord(1, I)
				.Write vbCrLf

				' check for any child options
				Call treTreeOption(nEditID, nLevel+1, aRecord, aRecord(0, I), nSelected)
			End If
		Next
	End With
End Sub

'--------------------------------------------------------------------
' Build a hierarchical tree select

Sub treTreeSelect(sInputName, sTableName, sKeyField, sParentField, sParentChoice, sCriteria, _
		nSelected)
	Dim sStat, rs, aRecord

	' fix the critera (where clause) for the select
	If Trim(sCriteria) <> "" And Not InStr(1, sCriteria, "AND ") Then
		sCriteria = " AND " & sCriteria
	Else
		sCriteria = " " & sCriteria
	End If

	sStat = "SELECT	" & sKeyField & ", " & sParentChoice & ", " & sParentField & " " &_
			"FROM " & sTableName & " " &_
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
		.Write "<select name=""" & sInputName & """ class=""form"">"
		.Write vbCrLf
		.Write "<option value=""""> -- Choose --"
		.Write vbCrLf
		Call treTreeOption(nSelected, 0, aRecord, 0, nSelected)
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
		If IsArray(aHeader) Then
			For I = 0 To UBound(aHeader)
				.Write "<TD CLASS=""listhead"">"
				.Write aHeader(I)
				.Write "</TD>"
				.Write vbCrLf
			Next
		Else
			.Write "<TD>"
			.Write aHeader
			.Write "</TD>"
			.Write vbCrLf
		End If
			.Write "<TD align=""right"" class=""listhead"">Action</TD>"
			.Write vbCrLf
		.Write "</TR>" & vbCrLf
	End With
End Sub

'--------------------------------------------------------------------
' treTreeAdminRows
'	Build the inidividual rows which make up the hierarchical admin

Sub treTreeAdminRows(sKeyField, nLevel, aRecord, aDisplay, nParentID, nRecNo, nSelected)
	Dim I, J, K, sHTML

	With Response
		For I = 0 To UBound(aRecord, 2)
			If nParentID = aRecord(1, I) Then
				.Write "<tr class=""tree"
				If aRecord(0, I) = nSelected Then
					.Write "sel"
				Else
					.Write nRecNo Mod 2
				End If
				.Write """>"
				.Write vbCrLf
				For J = 0 To UBound(aDisplay)
					' perform indenting on the first column
					If J = 0 Then
						.Write vbTab & "<td nowrap><table border=0 cellpadding=0 cellspacing=0><tr><td nowrap>"
						For K = 1 To nLevel
							.Write "&nbsp;&nbsp;&nbsp;"
						Next
						.Write "</td><td nowrap>"
						If Trim(aRecord(J+2, I) & "") <> "" Then
							.Write Server.HTMLEncode(aRecord(J+2, I))
						Else
							.Write "&nbsp;"
						End If
						.Write "</td></tr></table></td>"
					Else
						.Write vbTab & "<td>"
						If Not Trim(aRecord(J+2, I) & "") <> "" Then
							.Write Server.HTMLEncode(aRecord(J+2, I))
						Else
							.Write "&nbsp;"
						End If
						.Write "</td>"
					End If
				Next
				' write the action links for the record
				.Write "<td><a href=""?"
				.Write sKeyField
				.Write "="
				.Write aRecord(0, I)
				.Write "&action=edit"" class=""actionlink"">edit</A> . <a href=""?"
				.Write sKeyField
				.Write "="
				.Write aRecord(0, I)
				.Write "&action=delete"" class=""actionlink"">delete</A></td>"
				.Write vbCrLf
				' .Write Request.ServerVariables("SCRIPT_NAME")
				.Write "</tr>"
				.Write vbCrLf

				' check for any child options
				nRecNo = nRecNo + 1
				Call treTreeAdminRows(sKeyField, nLevel+1, aRecord, aDisplay, aRecord(0, I), nRecNo, nSelected)
			End If
		Next
	End With
End Sub

'--------------------------------------------------------------------
' Build a hierarchical tree admin page

Sub treTreeAdmin(sObjectName, sTableName, sKeyField, sParentField, sParentChoice, _
		sParentLabel, sDisplayFields, sHeaderLabels, sCriteria, _
		nArchive, nSelected, _
		sEditFields, sEditLabels, sEditTypes, sEditSizes)
	Dim sStat, rs, aRecord, aDisplay, aHeader

	If steForm("action") = "add" Then
		Call treForm(sTableName, sKeyField, steNForm(sKeyField), sParentField, sParentChoice, sParentLabel, _
			sEditFields, sEditLabels, sEditTypes, sEditSizes, _
			sCriteria, nSelected)
		Exit Sub
	ElseIf steForm("action") = "doadd" Then
		Call treDoAdd(sTableName, sKeyField, sParentField, sParentLabel, _
			sEditFields, sEditLabels, sEditTypes, sEditSizes)

		treStatusMsg = "The new " & sObjectName & " was added successfully"
	ElseIf steForm("action") = "edit" Then
		Call treForm(sTableName, sKeyField, steNForm(sKeyField), sParentField, sParentChoice, sParentLabel, _
			sEditFields, sEditLabels, sEditTypes, sEditSizes, _
			sCriteria, nSelected)
		Exit Sub
	ElseIf steForm("action") = "doedit" Then
		Call treDoEdit(sTableName, sKeyField, sParentField, sParentLabel, _
			sEditFields, sEditLabels, sEditTypes, sEditSizes)

		treStatusMsg = "The " & sObjectName & " was updated successfully"
	ElseIf steForm("action") = "delete" Then
		Call treDisplay(sTableName, sKeyField, steNForm(sKeyField), sParentField, sParentChoice, sParentLabel, _
			sEditFields, sEditLabels, sEditTypes, sEditSizes, _
			sCriteria, nSelected)
		Exit Sub
	ElseIf steForm("action") = "dodelete" Then
		Call treDoDelete(sTableName, sKeyField)

		treStatusMsg = "The " & sObjectName & " has been deleted"
	End If
	' default action is to show the entire list of items
	Call treTreeAdminList(sObjectName, sTableName, sKeyField, sParentField, sDisplayFields, sHeaderLabels, sCriteria, _
			nArchive, nSelected)
End Sub

'--------------------------------------------------------------------
' Build a hierarchical tree admin list

Sub treTreeAdminList(sObjectName, sTableName, sKeyField, sParentField, sDisplayFields, sHeaderLabels, sCriteria, _
		nArchive, nSelected)
	Dim sStat, rs, aRecord, aDisplay, aHeader, I


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


	sStat = "SELECT	" & sKeyField & ", " & sParentField & ", " & sDisplayFields & " " &_
			"FROM " & sTableName & " " &_
			"WHERE	Archive = " & nArchive & " " &_
			"AND	Active = 1 " &_
			sCriteria
	Set rs = adoOpenRecordset(sStat)
	If rs.EOF Then
		Response.Write "<b class=""error"">Nothing has been defined</b>" & vbCrLf
		Response.Write "<p align=""center""><a href=""" & Request.ServerVariables("SCRIPT_NAME") &_
			"?action=add"" class=""adminlink"">Add " & sObjectName & "</a></p>"
		Exit Sub
	End If
	aRecord = rs.GetRows
	rs = ""

	' output the tree admin table
	With Response
		' display the status message (if nec)
		If treStatusMsg <> "" Then
			.Write "<P><b class=""error"">"
			.Write treStatusMsg
			.Write "</b></P>" & vbCrLf
		End If

		' display the table header
		.Write "<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 CLASS=""list"">"
		.Write vbCrLf

		Call treTreeAdminHeaders(aHeader)

		Call treTreeAdminRows(sKeyField, 0, aRecord, aDisplay, 0, 0, nSelected)

		.Write "</TABLE>"
		.Write vbCrLf
		.Write "<p align=""center""><a href=""" & Request.ServerVariables("SCRIPT_NAME") &_
			"?action=add"" class=""adminlink"">Add " & sObjectName & "</a></p>"
	End With
End Sub

'--------------------------------------------------------------------
' Process the tree "add form" here

Sub treDoAdd(sTableName, sKeyField, sParentField, sParentLabel, _
	sEditFields, sEditLabels, sEditTypes, sEditSizes)
	Dim aField, aLabel, aType, sValues, bCreated, I
	Dim query

	query = "INSERT INTO " & sTableName & " (" & sParentField
	aField = Split(sEditFields, ",")
	aLabel = Split(sEditLabels, ",")
	aType = Split(sEditTypes, ",")
	sValues = CStr(steNForm(sParentField))
	bCreated = False

	For I = 0 To UBound(aField)
		query = query & ","
		sValues = sValues & ","
		query = query & aField(I)
		sValues = sValues & "'" & Replace(steForm(aField(I)), "'", "''") & "'"
		If aField(I) = "Created" Then bCreated = True
	Next
	If Not bCreated Then
		query = query & ",Created"
		sValues = sValues & "," & adoGetDate
	End If
	query = query & ") VALUES (" & sValues & ")"

	' execute the insert statement here
	Call adoExecute(query)
End Sub


'--------------------------------------------------------------------
' Process the tree "edit form" here

Sub treDoEdit(sTableName, sKeyField, sParentField, sParentLabel, _
	sEditFields, sEditLabels, sEditTypes, sEditSizes)
	Dim aField, aLabel, aType, sValues, I
	Dim query

	query = "UPDATE " & sTableName & " SET " & sParentField & " = " & CStr(steNForm(sParentField))
	aField = Split(sEditFields, ",")
	aLabel = Split(sEditLabels, ",")
	aType = Split(sEditTypes, ",")

	For I = 0 To UBound(aField)
		query = query & ", " & aField(I) & " = '" & Replace(steForm(aField(I)), "'", "''") & "'"
	Next
	query = query & "WHERE " & sKeyField & " = " & steNForm(sKeyField)

	' execute the insert statement here
	Call adoExecute(query)
End Sub

'--------------------------------------------------------------------
' Process the tree "delete form" here

Sub treDoDelete(sTableName, sKeyField)
	Dim query

	query = "DELETE FROM " & sTableName & " WHERE " & sKeyField & " = " & CStr(steNForm(sKeyField))
	Call adoExecute(query)
End Sub

'--------------------------------------------------------------------
' Retrieve the record to edit (for edit mode)
' Assumes an integer primary key

Function treEditRS(sTableName, sKeyField, nKeyValue, sParentField, _
	sEditFields)
	Dim rs, query, aField, sSelect, I

	aField = Split(sEditFields, ",")
	sSelect = sParentField
	For I = 0 To UBound(aField)
		sSelect = sSelect & "," & aField(I)
	Next
	query = "SELECT " & sSelect & " " &_
			"FROM " & sTableName & " " &_
			"WHERE	" & sKeyField & " = " & nKeyValue
	Set treEditRS = adoOpenRecordset(query)
End Function

'--------------------------------------------------------------------
' Build an "Add" or "Edit" form for the hierarchical table

Sub treForm(sTableName, sKeyField, nKeyValue, sParentField, sParentChoice, sParentLabel, _
	sEditFields, sEditLabels, sEditTypes, sEditSizes, _
	sCriteria, nSelected)
	Dim aField, aLabel, aType, aSize, nSize, nCol, nParentID, I
	Dim rsEdit, sAction

	aField = Split(sEditFields, ",")
	aLabel = Split(sEditLabels, ",")
	aType = Split(sEditTypes, ",")
	aSize = Split(sEditSizes, ",")

	' get the record to edit (if nec)
	If nKeyValue <> 0 Then
		Set rsEdit = treEditRS(sTableName, sKeyField, nKeyValue, sParentField, _
			sEditFields)
		If Not rsEdit.EOF Then sAction = "doedit" else sAction = "doadd"
	Else
		sAction = "doadd"
	End If

	' get the parent ID and build the form
	nParentID = steRecordValue(rsEdit, sParentField)
	If IsNumeric(nParentID) Then nParentID = CInt(nParentID) Else nParentID = 0
	With Response
	.Write "<FORM METHOD=""post"" ACTION=""?action=" & sAction & """>" & vbCrLf
	.Write "<INPUT TYPE=""hidden"" NAME=""" & sKeyField & """ VALUE=""" & nKeyValue & """>" & vbCrLf
	.Write "<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>" & vbCrLf
	.Write "<TR>" & vbCrLf
	.Write "	<TD COLSPAN=""2""><B>" & sParentLabel & "</B><BR>" & vbCrLf
	Call treTreeSelect(sParentField, sTableName, sKeyField, sParentField, sParentChoice, sCriteria, _
		nParentID)
	.Write "	</TD>" & vbCrLf
	.Write "</TR>" & vbCrLf

	' build all of the individual edit fields
	nCol = 0
	For I = 0 To UBound(aField)
		iF aSize(I) > 44 Then nSize = 44 Else nSize = aSize(I)
		If nCol mod 2 = 0 And nCol > 0 Then
			.Write "</tr><tr>"
		End If
		Select Case aType(I)
			Case "T"
				.Write "<td><b>" & aLabel(I) & "</b><br>"
				.Write "<input type=""text"" name=""" & aField(I) & """ value=""" &_
					steRecordEncValue(rsEdit, aField(I)) & """ size=""" &_
					nSize & """ maxlength=""" & aSize(I) & """ class=""form""></td>"
			Case "A"
				If nCol mod 2 = 1 Then
					.Write "</tr><tr>" & vbCrLf
				Else ' force a new row after this
					nCol = nCol + 1
				End If
				.Write "<td colspan=""2""><b>" & aLabel(I) & "</b><br>"
				.Write "<textarea name=""" & aField(I) & """ cols=""58"" rows=""10"" class=""form"" style=""width:500px"">" &_
					steRecordEncValue(rsEdit, aField(I)) & "</textarea></td>"
			Case "C"
				.Write "<td><b>" & aLabel(I) & "</b><br>"
				.Write "<input type=""checkbox"" name=""" & aField(I) & """ value=""" &_
					steRecordEncValue(rsEdit, aField(I)) & """ class=""form""></td>"
			Case "R"
				.Write "<td><b>" & aLabel(I) & "</b><br>"
				If steNForm(aField(I)) <> 0 Then sChecked = " CHECKED" Else sChecked = ""
				.Write "<input type=""radio"" name=""" & aField(I) & """ value=""1""" & sChecked & """ class=""formradio""></td>"
			Case "Y"
				.Write "<td><b>" & aLabel(I) & "</b><br>"
				If steNForm(aField(I)) <> 0 Then sChecked = " CHECKED" Else sChecked = ""
				.Write "<input type=""radio"" name=""" & aField(I) & """ value=""1""" & sChecked & """ class=""formradio""> Yes "
				If steNForm(aField(I)) = 0 Then sChecked = " CHECKED" Else sChecked = ""
				.Write "<input type=""radio"" name=""" & aField(I) & """ value=""0""" & sChecked & """ class=""formradio""> No "
				.Write "</td>"
		End Select
		nCol = nCol + 1
	Next
	' build the submit button here
	.Write "</tr><tr>" & vbCrLf
	.Write "	<td colspan=""2"" align=""right""><br>" & vbCrLf
	.Write "	<input type=""reset"" name=""tre_reset"" value="" Reset "" class=""form""> &nbsp; " & vbCrLf
	If sAction = "doedit" Then
		.Write "	<input type=""submit"" name=""tre_submit"" value="" Update "" class=""form"">" & vbCrLf
	Else
		.Write "	<input type=""submit"" name=""tre_submit"" value="" Add "" class=""form"">" & vbCrLf
	End If
	.Write "</tr>" & vbCrLf
	.Write "</table>" & vbCrLf
	.Write "</form>" & vbCrLf
	End With
End Sub

'--------------------------------------------------------------------
' Build a "read-only" form for an individual record (delete)

Sub treDisplay(sTableName, sKeyField, nKeyValue, sParentField, sParentChoice, sParentLabel, _
	sEditFields, sEditLabels, sEditTypes, sEditSizes, _
	sCriteria, nSelected)
	Dim aField, aLabel, aType, aSize, nSize, nCol, I
	Dim rs, rsEdit, sAction, sParentValue

	aField = Split(sEditFields, ",")
	aLabel = Split(sEditLabels, ",")
	aType = Split(sEditTypes, ",")
	aSize = Split(sEditSizes, ",")

	' get the record to edit (if nec)
	If nKeyValue <> 0 Then
		Set rsEdit = treEditRS(sTableName, sKeyField, nKeyValue, sParentField, _
			sEditFields)
	End If

	' retrieve the parent field name
	sStat = "SELECT	" & sParentChoice & " " &_
			"FROM " & sTableName & " " &_
			"WHERE	Archive = 0 " &_
			"AND	Active = 1 " &_
			"AND	" & sKeyField & " = " & steRecordEncValue(rsEdit, sParentField)
	Set rs = adoOpenRecordset(sStat)
	If Not rs.EOF Then
		sParentValue = rs.Fields(sParentChoice).Value
	Else
		sParentValue = "TOP-LEVEL ITEM"
	End If
	rs = ""

	With Response
	.Write "<FORM METHOD=""post"" ACTION=""?action=dodelete"">" & vbCrLf
	.Write "<INPUT TYPE=""hidden"" NAME=""" & sKeyField & """ VALUE=""" & nKeyValue & """>" & vbCrLf
	.Write "<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>" & vbCrLf
	.Write "<TR>" & vbCrLf
	.Write "	<TD COLSPAN=""2""><B>" & sParentLabel & "</B><BR>" & vbCrLf
	.Write "	" & sParentValue & vbCrLf
	.Write "	</TD>" & vbCrLf
	.Write "</TR>" & vbCrLf

	' build all of the individual edit fields
	nCol = 0
	For I = 0 To UBound(aField)
		iF aSize(I) > 44 Then nSize = 44 Else nSize = aSize(I)
		If nCol mod 2 = 0 And nCol > 0 Then
			.Write "</tr><tr>"
		End If
		Select Case aType(I)
			Case "T"
				.Write "<td><b>" & aLabel(I) & "</b><br>"
				.Write steRecordEncValue(rsEdit, aField(I))
				.Write "</td>"
			Case "A"
				If nCol mod 2 = 1 Then
					.Write "</tr><tr>" & vbCrLf
				Else ' force a new row after this
					nCol = nCol + 1
				End If
				.Write "<td colspan=""2""><b>" & aLabel(I) & "</b><br>"
				.Write steRecordEncValue(rsEdit, aField(I)) & vbCrLf
				.Write "</td>"
			Case "C"
				.Write "<td><b>" & aLabel(I) & "</b><br>"
				.Write "<input type=""checkbox"" name=""" & aField(I) & """ value=""" &_
					steRecordEncValue(rsEdit, aField(I)) & """ class=""form""></td>"
			Case "R"
				.Write "<td><b>" & aLabel(I) & "</b><br>"
				If steNForm(aField(I)) <> 0 Then sChecked = " CHECKED" Else sChecked = ""
				.Write "<input type=""radio"" name=""" & aField(I) & """ value=""1""" & sChecked & """ class=""formradio""></td>"
			Case "Y"
				.Write "<td><b>" & aLabel(I) & "</b><br>"
				If steNForm(aField(I)) <> 0 Then .Write "Yes" Else .Write "No"
				.Write "</td>"
		End Select
		nCol = nCol + 1
	Next
	' build the submit button here
	.Write "</tr><tr>" & vbCrLf
	.Write "	<td colspan=""2"" align=""right"">" & vbCrLf
	.Write "	<input type=""submit"" name=""tre_submit"" value="" Delete "" class=""form"">" & vbCrLf
	.Write "</tr>" & vbCrLf
	.Write "</table>" & vbCrLf
	.Write "</form>" & vbCrLf
	End With
End Sub

%>