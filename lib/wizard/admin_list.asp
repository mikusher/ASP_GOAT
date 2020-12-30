<%
'--------------------------------------------------------------------
' admin_list.asp
'	This wizard will build an admin for a simple list of
'	items stored in a database.
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
Dim strTableName
Dim strDisplayFields
Dim strDisplayLabels
Dim strEditFields
Dim strEditLabels
Dim strEditSizes
Dim strEditTypes
Dim strErrorMsg
Dim strCriteria
Dim strOrderField		' field to order the records by
Dim intActive			' show active or inactive items?
Dim intArchive			' show archive or unarchived items?
Dim intSelected			' primary key of currently selected item
Dim boolHasActive		' does table have an active bit?

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
'strPrimaryKeyValue = steForm("PrimaryKeyValue")
'strTableName = steForm("TableName")
'strEditFields = steForm("EditFields")
'strEditLabels = steForm("EditLabels")
'strEditSizes = steForm("EditSizes")
'strEditTypes = steForm("EditTypes")
'strDisplayFields = steForm("DisplayFields")
intArchive = 0
intSelected = 0
boolHasActive = False

' build the array of values here
arrField = Split(strEditFields, ",")
arrLabel = Split(strEditLabels, ",")
arrSize = Split(strEditSizes, ",")
arrType = Split(strEditTypes, ",")

' perform an action here
If strAction = "DOADD" Then
	Call lstDoAdd
	strStatusMsg = "Your new " & strObjectName & " has been added"
ElseIf strAction = "DOUPDATE" Then
	Call lstDoUpdate
ElseIf strAction = "DODELETE" Then
	Call lstDoDelete
ElseIf strAction = "MOVEUP" Then
	Call lstMoveUp
ElseIf strAction = "MOVEDOWN" Then
	Call lstMoveDown
End If

' display the form to add / edit (if nec)
If strAction = "ADD" Or strAction = "EDIT" Then
	' display the add / edit form here
	Call lstForm
ElseIf strAction = "DELETE" Then
	' display the delete form here
	Call lstDeleteForm
ElseIf strAction = "" Or strAction = "MOVEUP" Or strAction = "MOVEDOWN" Then
	' display the admin list
	Call lstListAdmin
End If

'--------------------------------------------------------------------
' lstAdminHeaders
'	Build the admin headers for the admin list table

Sub lstAdminHeaders(aHeader)
	Dim I, J, sHTML

	With Response
		.Write "<TR>" & vbCrLf
		For I = 0 To UBound(aHeader)
			.Write "<TD CLASS=""listhead"">"
			.Write aHeader(I)
			.Write "</TD>"
			.Write vbCrLf
		Next
		.Write "<TD CLASS=""listhead"" align=""right"">Action</TD>" & vbCrLf
		.Write "</TR>" & vbCrLf
	End With
End Sub

'--------------------------------------------------------------------
' lstAdminRows
'	Build the inidividual rows which make up the hierarchical admin

Sub lstAdminRows(aRecord, aDisplay, nRecNo, nSelected)
	Dim I, J, sHTML

	With Response
		For I = 0 To UBound(aRecord, 2)
			.Write "<tr class=""list"
			If aRecord(0, I) = nSelected Then
				.Write "sel"
			Else
				.Write nRecNo Mod 2
			End If
			.Write """>"
			.Write vbCrLf
			For J = 0 To UBound(aDisplay)
				.Write vbTab & "<td>"
				If Trim(aRecord(J+1, I) & "") <> "" Then
					.Write Server.HTMLEncode(aRecord(J+1, I))
				Else
					.Write "&nbsp;"
				End If
				.Write "</td>"
			Next
			' write the action links for the record
			.Write "<td>"
			If strOrderField <> "" Then
				.Write "<a href=""?" & strPrimaryKey & "="
				.Write aRecord(0, I)
				.Write "&action=moveup"" class=""actionlink"">up</A> . <a href=""?" & strPrimaryKey & "="
				.Write aRecord(0, I)
				.Write "&action=movedown"" class=""actionlink"">down</A> . "
			End If
			.Write "<a href=""?" & strPrimaryKey & "="
			.Write aRecord(0, I)
			.Write "&action=edit"" class=""actionlink"">edit</A> . <a href=""?" & strPrimaryKey & "="
			.Write aRecord(0, I)
			.Write "&action=delete"" class=""actionlink"">delete</A></td>"
			.Write vbCrLf
			' .Write Request.ServerVariables("SCRIPT_NAME")
			.Write "</tr>"
			.Write vbCrLf

			' check for any child options
			nRecNo = nRecNo + 1
		Next
	End With
End Sub

'--------------------------------------------------------------------
' Build a list admin page

Sub lstListAdmin
	Dim sStat, rs, aRecord, aDisplay, aHeader, sCriteria, I


	' fix the critera (where clause) for the select
	If Trim(strCriteria) <> "" And Not InStr(1, strCriteria, "AND ") Then
		sCriteria = " AND " & strCriteria
	Else
		sCriteria = " " & strCriteria
	End If

	' build the display array and trim whitespace
	aDisplay = Split(strDisplayFields, ",")
	For I = 0 To UBound(aDisplay)
		aDisplay(I) = Trim(aDisplay(I))
	Next

	aHeader = Split(strDisplayLabels, ",")
	For I = 0 To UBound(aHeader)
		aHeader(I) = Trim(aHeader(I))
	Next

	' build the select query for the list
	sStat = "SELECT	" & strPrimaryKey & ", " & strDisplayFields & " " &_
			"FROM " & strTableName & " " &_
			"WHERE	Archive = " & intArchive & " "
	If boolHasActive Then sStat = sStat & "AND	Active = " & intActive & " "
	sStat = sStat &	sCriteria
	If strOrderField <> "" Then
		sStat = sStat & " ORDER BY " & strOrderField
	End If
	Set rs = adoOpenRecordset(sStat)

	' output the admin table
	With Response
		.Write "<h3>" & strObjectName & " List</h3>" & vbCrLf

		If strStatusMsg <> "" Then
		.Write "<p><b class=""error"">"
		.Write strStatusMsg 
		.Write "</b></p>" & vbCrLf
		End If

		If rs.EOF Then
			rs.Close
			Set rs = Nothing
			.Write "<b class=""error"">Nothing has been defined</b>"
		Else
			aRecord = rs.GetRows
			rs.Close
			Set rs = Nothing
			' display the table header
			.Write "<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 CLASS=""list"">"
			.Write vbCrLf
	
			Call lstAdminHeaders(aHeader)
	
	
			Call lstAdminRows(aRecord, aDisplay, 0, intSelected)
	
			.Write "</TABLE>"
			.Write vbCrLf
		End If

		' display the link to add a new item
		.Write "<p align=""center"">" & vbCrLf
		.Write "	<a href="""
		.Write Request.ServerVariables("SCRIPT_NAME")
		.Write "?archive=" & intArchive & "&active=" & intActive & "&action=add"" class=""adminlink"">Add New "
		.Write strObjectName
		.Write "</A>" & vbCrLf
		.Write "</p><br>" & vbCrLf
	End With
End Sub

'--------------------------------------------------------------------
' Process the list "add form" here

Sub lstDoAdd
	Dim query, sValues, rsOrder, nOrder, I

	' determine the OrderNo field (if nec)
	nOrderNo = 1
	If strOrderField <> "" Then
		' retrieve the new order field
		query = "SELECT Coalesce(Max(OrderNo) + 1, 1) AS OrderNo FROM " & strTableName
		Set rsOrder = adoOpenRecordset(query)
		If Not rsOrder.EOF Then nOrderNo = rsOrder.Fields("OrderNo").Value
		rsOrder.Close
		Set rsOrder = Nothing
	End If

	query = "INSERT INTO " & strTableName & " ("

	' build the insert fields / values for the insert query
	For I = 0 To UBound(arrField)
		If I > 0 Then
			query = query & ","
			sValues = sValues & ","
		End If
		query = query & arrField(I)
		If arrField(I) = strOrderField Then
			sValues = sValues & CStr(nOrderNo)
		Else
			sValues = sValues & "'" & Replace(steForm(arrField(I)), "'", "''") & "'"
		End If
	Next
	' add the order field to the query / values (if nec)
	If strOrderField <> "" Then
		If Not (InStr(1, ","&strOrderField&",", ","&strEditFields&",", vbTextCompare) > 0) Then
			query = query & "," & strOrderField
			sValues = sValues & "," & nOrderNo
		End If
	End If
	' add the creation date to the query (if nec)
	If Not (InStr(1, ",Created,", ","&strEditFields&",", vbTextCompare) > 0) Then
		query = query & ",Created"
		sValues = sValues & "," & adoGetDate
	End If
	query = query & ") VALUES (" & sValues & ")"

	' execute the insert statement here
	Call adoExecute(query)
	strStatusMsg = "The " & strObjectName & " record has been added"
	strAction = ""
End Sub

'--------------------------------------------------------------------
' Process the list "update form" here

Sub lstDoUpdate
	Dim query, sValues, I

	query = "UPDATE " & strTableName & " SET "

	For I = 0 To UBound(arrField)
		If I > 0 Then query = query & ", "
		query = query & arrField(I) & " = '" & Replace(steForm(arrField(I)), "'", "''") & "'"
	Next
	query = query & " WHERE " & strPrimaryKey & " = " & steNForm(strPrimaryKey)

	' execute the update statement here
	Call adoExecute(query)
	strStatusMsg = "The " & strObjectName & " record has been updated"
	strAction = ""
End Sub

'--------------------------------------------------------------------
' Process the list "delete form" here

Sub lstDoDelete
	Dim query, sValues, I

	If steNForm("Confirm") = 0 Then
		strErrorMsg = "You must confirm the delete operation first"
		strAction = "DELETE"
		Exit Sub
	End If
	query = "DELETE FROM " & strTableName & " WHERE " & strPrimaryKey & " = " & steNForm(strPrimaryKey)

	' execute the delete statement here
	Call adoExecute(query)
	strStatusMsg = "The " & strObjectName & " record has been deleted"
	strAction = ""
End Sub

'--------------------------------------------------------------------
' Retrieve the record to edit (for edit mode)
' Assumes an integer primary key

Function listditRS
	Dim query, sSelect, I

	For I = 0 To UBound(arrField)
		If I > 0 Then sSelect = sSelect & ","
		sSelect = sSelect & arrField(I)
	Next
	query = "SELECT " & sSelect & " " &_
			"FROM " & strTableName & " " &_
			"WHERE	" & strPrimaryKey & " = " & strPrimaryKeyValue
	Set listditRS = adoOpenRecordset(query)
End Function

'--------------------------------------------------------------------
' Build an "Add" or "Edit" database form 

Sub lstForm
	Dim aField, aLabel, aType, aSize, nSize, nCol, I
	Dim rsEdit, sActionVerb, sFormAction

	aField = Split(strEditFields, ",")
	aLabel = Split(strEditLabels, ",")
	aType = Split(strEditTypes, ",")
	aSize = Split(strEditSizes, ",")

	' get the record to edit (if nec)
	If strPrimaryKeyValue <> 0 Then
		Set rsEdit = listditRS
		sActionVerb = "Edit"
		sFormAction = "doupdate"
	Else
		sActionVerb = "Add"
		sFormAction = "doadd"
	End If

	With Response
	.Write "<h3>" & sActionVerb & " " & strObjectName & "</h3>" & vbCrLf

	If strErrorMsg <> "" Then
	.Write "<p><b class=""error"">"
	.Write strErrorMsg
	.Write "</b></p>" & vbCrLf
	End If

	.Write "<FORM METHOD=""post"" ACTION=""?action=" & sFormAction & """>" & vbCrLf
	.Write "<INPUT TYPE=""hidden"" NAME=""" & strPrimaryKey & """ VALUE=""" & strPrimaryKeyValue & """>" & vbCrLf
	.Write "<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>" & vbCrLf

	' build all of the individual edit fields
	nCol = 0
	For I = 0 To UBound(aField)
		iF aSize(I) > 44 Then nSize = 44 Else nSize = aSize(I)
		' If nCol mod 2 = 0 And nCol > 0 Then
			.Write "</tr><tr>"
		' End If
		.Write "<td valign=""top"" class=""forml"">"
		.Write aLabel(I) & "</td><td></td><td class=""formd"">"
		Select Case aType(I)
			Case "T"
				.Write "<input type=""text"" name=""" & aField(I) & """ value=""" &_
					steRecordEncValue(rsEdit, aField(I)) & """ size=""" &_
					nSize & """ maxlength=""" & aSize(I) & """ class=""form"">"
			Case "A"
				If nCol mod 2 = 1 Then
					.Write "</td></tr><tr><td colspan=""3"">" & vbCrLf
				Else ' force a new row after this
					nCol = nCol + 1
				End If
				.Write "<textarea name=""" & aField(I) & """ cols=""38"" rows=""10"" class=""form"" style=""width:420px"">" &_
					steRecordEncValue(rsEdit, aField(I)) & "</textarea>"
			Case "C"
				.Write "<input type=""checkbox"" name=""" & aField(I) & """ value=""" &_
					steRecordEncValue(rsEdit, aField(I)) & """ class=""form"">"
			Case "R"
				If steNForm(aField(I)) <> 0 Then sChecked = " CHECKED" Else sChecked = ""
				.Write "<input type=""radio"" name=""" & aField(I) & """ value=""1""" & sChecked & """ class=""formradio"">"
			Case "Y"
				If steNForm(aField(I)) <> 0 Then sChecked = " CHECKED" Else sChecked = ""
				.Write "<input type=""radio"" name=""" & aField(I) & """ value=""1""" & sChecked & """ class=""formradio""> Yes "
				If steNForm(aField(I)) = 0 Then sChecked = " CHECKED" Else sChecked = ""
				.Write "<input type=""radio"" name=""" & aField(I) & """ value=""0""" & sChecked & """ class=""formradio""> No "
		End Select
		.Write "</td>"
		nCol = nCol + 1
	Next
	' build the submit button here
	.Write "</tr><tr>" & vbCrLf
	.Write "	<td colspan=""3"" align=""right"">" & vbCrLf
	.Write "	<input type=""reset"" name=""lst_reset"" class=""form"" value="" Reset ""> &nbsp; " & vbCrLf
	If strPrimaryKeyValue <> 0 Then
		.Write "	<input type=""submit"" name=""lst_submit"" class=""form"" value="" Update "">" & vbCrLf
	Else
		.Write "	<input type=""submit"" name=""lst_submit"" class=""form"" value="" Add "">" & vbCrLf
	End If
	.Write "</tr>" & vbCrLf
	.Write "</table>" & vbCrLf
	End With
End Sub

'--------------------------------------------------------------------
' Build a "Delete" database form 

Sub lstDeleteForm
	Dim aField, aLabel, aType, aSize, nSize, nCol, I
	Dim rsEdit, sActionVerb, sFormAction

	aField = Split(strEditFields, ",")
	aLabel = Split(strEditLabels, ",")
	aType = Split(strEditTypes, ",")
	aSize = Split(strEditSizes, ",")

	' get the record to edit (if nec)
	If strPrimaryKeyValue <> 0 Then
		Set rsEdit = listditRS
	End If

	With Response
	.Write "<h3>" & sActionVerb & " " & strObjectName & "</h3>" & vbCrLf

	If strErrorMsg <> "" Then
	.Write "<p><b class=""error"">"
	.Write strErrorMsg
	.Write "</b></p>" & vbCrLf
	End If

	.Write "<FORM METHOD=""post"" ACTION=""?action=dodelete"">" & vbCrLf
	.Write "<INPUT TYPE=""hidden"" NAME=""" & strPrimaryKey & """ VALUE=""" & strPrimaryKeyValue & """>" & vbCrLf
	.Write "<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>" & vbCrLf

	' build all of the individual edit fields
	nCol = 0
	For I = 0 To UBound(aField)
		iF aSize(I) > 44 Then nSize = 44 Else nSize = aSize(I)
		.Write "<tr>" & vbCrLf
		.Write "<td valign=""top"" class=""forml"">"
		.Write aLabel(I) & "</td><td></td><td class=""formd"">"
		Select Case aType(I)
			Case "T"
				.Write steRecordEncValue(rsEdit, aField(I))
			Case "A"
				If nCol mod 2 = 1 Then
					.Write "</td></tr><tr>" & vbCrLf
				Else ' force a new row after this
					nCol = nCol + 1
				End If
				.Write "<td colspan=""3"" class=""formd"">" & vbCrLf
				.Write steRecordEncValue(rsEdit, aField(I))
			Case "C"
				If steRecordBoolValue(rsEdit, aField(I)) Then .Write "Y" Else .Write "N"
			Case "R"
				If steRecordBoolValue(rsEdit, aField(I)) Then .Write "Y" Else .Write "N"
			Case "Y"
				If steRecordBoolValue(rsEdit, aField(I)) Then .Write "Y" Else .Write "N"
		End Select
		.Write "</td>" & vbCrLf
		.Write "</tr>" & vbCrLf
		nCol = nCol + 1
	Next
	' build the confirmation buttons here
	.Write "<tr>" & vbCrLf
	.Write "<td valign=""top"" class=""forml"">"
	.Write "Confirm Delete?</td><td></td><td class=""formd"">"
	.Write "<input type=""radio"" name=""confirm"" value=""1"" class=""formradio""> Yes" & vbCrLf
	.Write "<input type=""radio"" name=""confirm"" value=""0"" class=""formradio""> No" & vbCrLf
	.Write "</td>" & vbCrLf
	.Write "</tr>"

	' build the submit button here
	.Write "<tr>" & vbCrLf
	.Write "	<td colspan=""2"" align=""right"">" & vbCrLf
	.Write "	<input type=""reset"" name=""lst_reset"" class=""form"" value="" Reset ""> &nbsp; " & vbCrLf
	.Write "	<input type=""submit"" name=""lst_submit"" class=""form"" value="" Delete "">" & vbCrLf
	.Write "</tr>" & vbCrLf
	.Write "</table>" & vbCrLf
	End With
End Sub

'--------------------------------------------------------------------
' Move an item up in the order (by updating the OrderNo field)

Sub lstMoveUp
	Dim rsOrder, nOrderNo, nPrevOrder

	' retrieve the OrderNo for the field
	Set rsOrder = adoOpenRecordset("SELECT OrderNo FROM " & strTableName & " WHERE " & strPrimaryKey & " = " & strPrimaryKeyValue)
	If Not rsOrder.EOF Then
		nOrderNo = rsOrder.Fields("OrderNo").Value
		Set rsOrder = adoOpenRecordset("SELECT	" & adoTop(1) & " OrderNo FROM " & strTableName & " WHERE OrderNo < " & nOrderNo & " ORDER BY OrderNo DESC " & adoTop2(1))
		If Not rsOrder.EOF Then nPrevOrder = rsOrder.Fields("OrderNo").Value
	End If
	rsOrder.Close
	Set rsOrder = Nothing
	If Not IsNumeric(nOrderNo) Or CStr(nOrderNo) = "" Then Exit Sub
	If Not IsNumeric(nPrevOrder) Or CStr(nPrevOrder) = "" Then Exit Sub

	' increment orders above the new order no (to make room)
	sStat = "UPDATE	" & strTableName & " " &_
			"SET	OrderNo = OrderNo + 1, Modified = " & adoGetDate & " " &_
			"WHERE	OrderNo = " & nPrevOrder
	Call adoExecute(sStat)

	sStat = "UPDATE	" & strTableName & " " &_
			"SET	OrderNo = " & nPrevOrder & ", Modified = " & adoGetDate & " " &_
			"WHERE	" & strPrimaryKey & " = " & strPrimaryKeyValue
	Call adoExecute(sStat)
End Sub

'--------------------------------------------------------------------
' Move an item up in the order (by updating the OrderNo field)

Sub lstMoveDown
	Dim rsOrder, nOrderNo, nNextOrder

	' retrieve the OrderNo for the field
	Set rsOrder = adoOpenRecordset("SELECT OrderNo FROM " & strTableName & " WHERE " & strPrimaryKey & " = " & strPrimaryKeyValue)
	If Not rsOrder.EOF Then
		nOrderNo = rsOrder.Fields("OrderNo").Value
		Set rsOrder = adoOpenRecordset("SELECT	" & adoTop(1) & " OrderNo FROM " & strTableName & " WHERE OrderNo > " & nOrderNo & " ORDER BY OrderNo " & adoTop2(1))
		If Not rsOrder.EOF Then nNextOrder = rsOrder.Fields("OrderNo").Value
	End If
	rsOrder.Close
	Set rsOrder = Nothing
	If Not IsNumeric(nOrderNo) Or CStr(nOrderNo) = "" Then Exit Sub
	If Not IsNumeric(nNextOrder) Or CStr(nNextOrder) = "" Then Exit Sub

	' increment orders above the new order no (to make room)
	sStat = "UPDATE	" & strTableName & " " &_
			"SET	OrderNo = OrderNo - 1, Modified = " & adoGetDate & " " &_
			"WHERE	OrderNo = " & nNextOrder
	Call adoExecute(sStat)

	sStat = "UPDATE	" & strTableName & " " &_
			"SET	OrderNo = " & nNextOrder & ", Modified = " & adoGetDate & " " &_
			"WHERE	" & strPrimaryKey & " = " & strPrimaryKeyValue
	Call adoExecute(sStat)
End Sub
%>