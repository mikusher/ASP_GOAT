<%
'--------------------------------------------------------------------
' listinput.asp
'	Class for building drop-list selection controls.  Options for
'	the list may be pulled from the database or built staticly.
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

Class clsListInput
	Private mstrInputName		' name for the input control
	Private mstrTableName		' databaes table to get options from
	Private mstrWhereClause		' where clause to limit option list
	Private mstrOptionLabelFields	' CSV list of option label fields
	Private mstrOptionValueField	' field to use for the option value
	Private mstrOrderClause		' how to order the option listing
	Private mstrQueryString		' query string to use for links
	Private mstrError			' error to report
	Private mintSelected		' selected option
	Private mstrInputWidth		' CSS width style (eg: "width:150px")
	Private mboolIsDynamic		' Picking an option reloads the page?
	Private mboolShowChoose		' show the "choose one" option?
	Private mstrChooseOptionValue ' value for option "choose one"
	Private mstrChooseOptionLabel ' label for option "choose one"
	Private mstrParentField		' parent field
	Private mstrPKeyField		' primary key field
	Private mobjDict			' dictionary object

	'------------------------------------------------------------------------
	' constructor for the date input control

	Private Sub Class_Initialize
		mintSelected = 0
		mboolShowChoose = True
		mstrChooseOptionValue = 0
		mstrChooseOptionLabel = "-- Choose --"
		Set mobjDict = Server.CreateObject("Scripting.Dictionary")
	End Sub

	'------------------------------------------------------------------------
	' validate all of the required properties

	Private Function Validate
		
	End Function

	'------------------------------------------------------------------------
	' output javascript where picking an option, makes the page reload itself

	Private Sub PickJS
		With Response
		.Write "<script language=""Javascript"" type=""text/javascript"">" & vbCrlf
		.Write "function pick" & mstrInputName & "(n" & mstrInputName & ")" & vbCrLf
		.Write "{" & vbCrLf
		.Write "	if (n" & mstrInputName & " != '0')" & vbCrLf
		.Write "		location.href='" & Request.ServerVariables("SCRIPT_NAME") & "?"
		If mstrQueryString <> "" Then .Write mstrQueryString & "&"
		.Write mstrInputName & "=' + n" & mstrInputName & ";" & vbCrLf
		.Write "}" & vbCrLf
		.Write "</script>" & vbCrLf
		End With
	End Sub

	'------------------------------------------------------------------------
	' build an individual option row for the hierarchical tree select

	Private Function BuildRow(rsList)
		Dim rsOpt, sStat, sWhere, sOrder, sFields, aLabel, sHTML, I

		sHTML = sHTML & "<option value=""" & Server.HTMLEncode(rsList.Fields(mstrOptionValueField).Value & "") & """"
		If CStr(mintSelected) = CStr(rsList.Fields(mstrOptionValueField).Value) Then sHTML = sHTML & " SELECTED"
		sHTML = sHTML & ">##indent##"
		' build the option label here
		aLabel = Split(Replace(mstrOptionLabelFields, " ", ""), ",")
		For I = 0 To UBound(aLabel)
			If I > 0 Then sHTML = sHTML & " "
			sHTML = sHTML & Server.HTMLEncode(rsList.Fields(aLabel(I)).Value)
		Next
		BuildRow = sHTML & "</option>" & vbCrLf
	End Function

	'--------------------------------------------------------------
	' build the hierarchical structure in the dictionary object

	Private Sub BuildTree(rsList)
		Dim sRow, nParentID

		' make sure we have an empty dictionary
		If mobjDict Is Nothing Then
			Set mobjDict = Server.CreateObject("Scripting.Dictionary")
		Else
			mobjDict.RemoveAll
		End If
		' make sure the parent / child fields are defined
		If mstrPKeyField = "" Then
			mstrError = "clsListInput - Must define the PrimaryKey property"
			Exit Sub
		End If
		If mstrParentField = "" Then
			mstrError = "clsListInput - Must define the ParentField property"
			Exit Sub
		End If
		' build the dictionary of hierarchical options
		Do Until rsList.EOF
			' build the list of primary keys for the parent
			nParentID = CStr(rsList.Fields(mstrParentField).Value)
			If mobjDict.Exists(nParentID) Then
				mobjDict.Item(nParentID) = mobjDict.Item(nParentID) & "," & rsList.Fields(mstrPKeyField).Value
			Else
				mobjDict.Item(nParentID) = rsList.Fields(mstrPKeyField).Value
			End If
			' Response.Write "parent(" & nParentID & ") = " & mobjDict.Item(nParentID) & "<Br><Br>" & vbCrLf

			' add the data row for the current record
			sRow = BuildRow(rsList)
			' Response.Write "row" & CStr(rsList.Fields(mstrPKeyField).Value) & " = " & Server.HTMLEncode(sRow) & "<BR><BR>" & vbCrLf
			mobjDict.Item("row" & CStr(rsList.Fields(mstrPKeyField).Value)) = sRow
			rsList.MoveNext
		Loop
	End Sub

	'--------------------------------------------------------------
	' display the options at a specific level (hierarchical list)

	Private Sub DisplayLevel(nParentID, nLevelNo)
		Dim aKey, sRow, sIndent, I

		If Not mobjDict.Exists(CStr(nParentID)) Then Exit Sub
		If mobjDict.Item(CStr(nParentID)) = "" Then Exit Sub
		' build the array of primary keys at this level
		aKey = Split(mobjDict.Item(CStr(nParentID)), ",")

		' build the indent code for this level
		For I = 1 To nLevelNo
			sIndent = sIndent & " &nbsp; &nbsp;"
		Next
		' build each success record in the list (from the row data)
		For I = 0 To UBound(aKey)
			Response.Write Replace(mobjDict.Item("row" & aKey(I)), "##indent##", sIndent, 1, -1, vbTextCompare)

			' call the child level (if exists)
			If mobjDict.Exists(aKey(I)) Then
				If mobjDict.Item(aKey(I)) <> "" Then DisplayLevel aKey(I), nLevelNo + 1
			End If
		Next
	End Sub

	'------------------------------------------------------------------------
	' output the option list from the database

	Private Sub DBOptions
		Dim rsOpt, sStat, sWhere, sOrder, sFields, aLabel, I

		' build the "where" and "order by" clauses for the option list
		If mstrWhereClause <> "" And Left(mstrWhereClause, 6) <> "where " Then
			sWhere = "WHERE " & mstrWhereClause
		End If
		If mstrOrderClause <> "" And Left(mstrOrderClause, 9) <> "order by " Then
			sOrder = "ORDER BY " & mstrOrderClause
		End If
		aLabel = Split(Replace(mstrOptionLabelFields, " ", ""), ",")
		sFields = Replace(mstrOptionLabelFields, " ", "") & "," & mstrOptionValueField
		' make sure the parent / child fields exist for hierarchical lists
		If mstrPKeyField <> "" And Not InStr(1, ","&sFields&",", ","&mstrPKeyField&",") > 0 Then
			sFields = sFields & "," & mstrPKeyField
		End If
		If mstrParentField <> "" And Not InStr(1, ","&sFields&",", ","&mstrParentField&",") > 0 Then
			sFields = sFields & "," & mstrParentField
		End If
		sStat = "SELECT	" & sFields & " " &_
				"FROM	" & mstrTableName & " " &_
				sWhere & " " &_
				sOrder
		Set rsOpt = adoOpenRecordset(sStat)
		If mstrPKeyField <> "" And mstrParentField <> "" Then
			' hierarchical list select goes here
			Call BuildTree(rsOpt)
			rsOpt.Close
			Set rsOpt = Nothing
			Call DisplayLevel(0, 0)
		Else
			' output the options here
			With Response
			Do Until rsOpt.EOF
				.Write "<option value="""
				.Write rsOpt.Fields(mstrOptionValueField).Value
				.Write """"
				If CStr(mintSelected) = CStr(rsOpt.Fields(mstrOptionValueField).Value) Then Response.Write " SELECTED"
				.Write ">"
				' build the option label here
				For I = 0 To UBound(aLabel)
					If I > 0 Then .Write " "
					.Write rsOpt.Fields(aLabel(I)).Value
				Next
				.Write "</option>"
				.Write vbCrLf
				rsOpt.MoveNext
			Loop
			End With
			rsOpt.Close
			Set rsOpt = Nothing
		End If
	End Sub

	'------------------------------------------------------------------------
	' display the list input

	Public Sub ListInput
		Dim sWidth, sOnClick

		' output the javascript pick function
		If mboolIsDynamic Then
			Call PickJS
			sOnClick = " onChange=""pick" & mstrInputName & "(this.options[this.selectedIndex].value)"""
		End If

		' define the CSS style for the input
		If mstrInputWidth <> "" Then
			sWidth = " STYLE=""" & mstrInputWidth & """"
		End If
		With Response
		.Write "<select name=""" & mstrInputName & """ class=""form""" & sWidth & sOnClick & ">" & vbCrLf
		' show the "choose one" option (if nec)
		If mboolShowChoose Then
			.Write "<option value=""" & mstrChooseOptionValue & """>"
			.Write Server.HTMLEncode(mstrChooseOptionLabel)
			.Write "</option>"
			.Write vbCrLf
		End If
		Call DBOptions
		.Write "</select>" & vbCrLf
		End With
	End Sub

	'------------------------------------------------------------------------
	' display a hierarchical list input

	Public Sub TreeListInput(sInputName, sTableName, sPKeyField, sParentField, sWhereClause, _
		sOrderClause, sOptionValueField, sOptionLabelFields, nSelected, sQueryString, bIsDynamic)
		mstrInputName = sInputName
		mstrTableName = sTableName
		mstrPKeyField = sPKeyField
		mstrParentField = sParentField
		mstrWhereClause = sWhereClause
		mstrOrderClause = sOrderClause
		mstrOptionLabelFields = sOptionLabelFields
		mstrOptionValueField = sOptionValueField
		mintSelected = nSelected
		mstrQueryString = sQueryString
		mboolIsDynamic = bIsDynamic
		Call ListInput
	End Sub

	'------------------------------------------------------------------------
	' get/set the input name property

	Public Property Let InputName(strValue)
		mstrInputName = strValue
	End Property

	Public Property Get InputName
		InputName = mstrInputName
	End Property

	'------------------------------------------------------------------------
	' get/set the table name property

	Public Property Let TableName(strValue)
		mstrTableName = strValue
	End Property

	Public Property Get TableName
		TableName = mstrTableName
	End Property

	'------------------------------------------------------------------------
	' get/set the where clause property

	Public Property Let WhereClause(strValue)
		mstrWhereClause = strValue
	End Property

	Public Property Get WhereClause
		WhereClause = mstrWhereClause
	End Property

	'------------------------------------------------------------------------
	' get/set the option label fields property (CSV list: "FName,MI,LName")

	Public Property Let OptionLabelFields(strValue)
		mstrOptionLabelFields = strValue
	End Property

	Public Property Get OptionLabelFields
		OptionLabelFields = mstrOptionLabelFields
	End Property

	'------------------------------------------------------------------------
	' get/set the option value field property

	Public Property Let OptionValueField(strValue)
		mstrOptionValueField = strValue
	End Property

	Public Property Get OptionValueField
		OptionValueField = mstrOptionValueField
	End Property

	'------------------------------------------------------------------------
	' get/set the selected value property

	Public Property Let Selected(intValue)
		mintSelected = intValue
	End Property

	Public Property Get Selected
		Selected = mintSelected
	End Property

	'------------------------------------------------------------------------
	' get/set the SQL order clause (for options list) property

	Public Property Let OrderClause(strValue)
		mstrOrderClause = strValue
	End Property

	Public Property Get OrderClause
		OrderClause = mstrOrderClause
	End Property

	'------------------------------------------------------------------------
	' get/set the CSS input width property

	Public Property Let InputWidth(strValue)
		mstrInputWidth = strValue
	End Property

	Public Property Get InputWidth
		InputWidth = mstrInputWidth
	End Property

	'------------------------------------------------------------------------
	' get/set the Choose Option Value property

	Public Property Let ChooseOptionValue(strValue)
		mstrChooseOptionValue = strValue
	End Property

	Public Property Get ChooseOptionValue
		ChooseOptionValue = mstrChooseOptionValue
	End Property

	'------------------------------------------------------------------------
	' get/set the Choose Option Label property

	Public Property Let ChooseOptionLabel(strValue)
		mstrChooseOptionLabel = strValue
	End Property

	Public Property Get ChooseOptionLabel
		ChooseOptionLabel = mstrChooseOptionLabel
	End Property

	'------------------------------------------------------------------------
	' get/set the Show Choose property

	Public Property Let ShowChoose(boolValue)
		mboolShowChoose = boolValue
	End Property

	Public Property Get ShowChoose
		ShowChoose = mboolShowChoose
	End Property

	'------------------------------------------------------------------------
	' get/set the Is Dynamic property

	Public Property Let IsDynamic(boolValue)
		mboolIsDynamic = boolValue
	End Property

	Public Property Get IsDynamic
		IsDynamic = mboolIsDynamic
	End Property

	'------------------------------------------------------------------------
	' get/set the QueryString property

	Public Property Let QueryString(strValue)
		mstrQueryString = strValue
	End Property

	Public Property Get QueryString
		QueryString = mstrQueryString
	End Property

	'------------------------------------------------------------------------
	' get/set the PrimaryKey (field name) property

	Public Property Let PrimaryKey(strValue)
		mstrPKeyField = strValue
	End Property

	Public Property Get PrimaryKey
		PrimaryKey = mstrPKeyField
	End Property

	'------------------------------------------------------------------------
	' get/set the ParentField (hierarchical list) property

	Public Property Let ParentField(strValue)
		mstrParentField = strValue
	End Property

	Public Property Get ParentField
		ParentField = mstrParentField
	End Property

	'------------------------------------------------------------------------
	' get the error message property

	Public Property Get ErrorMsg
		ErrorMsg = mstrError
	End Property
End Class
%>