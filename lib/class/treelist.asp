<%
'--------------------------------------------------------------------
' treelist.asp
'	Class for building hierarchical (tree) admin lists
'	TODO - Efficient paging of the result sets
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

Const adDate = 7
Const adDBDate = 133
Const adDBTimeStamp = 135

Class clsTreeList

	Private mstrQuery		' SQL statement to retrieve recordset
	Private mstrLabel		' labels to display (column headings)
	Private mstrError		' error message
	Private mstrActionLink	' HTML template for the action links
	Private mintPageSize	' max records to show on one page
	Private mintPageNo		' current page no being displayed
	Private mintResultCount ' total no. of results found
	Private marrDisplay()	' array of fields (columns) to display
	Private marrJust()		' justification for column
	Private marrDateFmt()	' date format type to perform
	Private mintColumns		' total number of columns defined
	Private mintDefDateFmt	' default date formatting for dates
	Private mstrParentField ' field containing parent ID
	Private mstrPKeyField	' primary key field
    Private mstrQueryString ' addl. querystring params to pass
	Private mobjDict		' dictionary hash for the data rows

	'--------------------------------------------------------------
	' Constructor

	Private Sub Class_Initialize
		mintPageSize = 25
		If Request.QueryString("pageno") <> "" Then
			mintPageNo = CInt(Request.QueryString("pageno"))
		Else
			mintPageNo = CInt(Request.Form("pageno"))
		End If
		If Request.QueryString("ResultCount") <> "" Then
			mintResultCount = Request.QueryString("ResultCount")
		ElseIf Request.Form("ResultCount") <> "" Then
			mintResultCount = CInt(Request.Form("ResultCount"))
		Else
			mintResultCount = 0
		End If
		' setup the array of column definitions
		ReDim marrDisplay(0)
		ReDim marrJust(0)
		ReDim marrDateFmt(0)
		mintColumns = 0
		mintDefDateFmt = vbShortDate
		Set mobjDict = Nothing 
	End Sub

	'--------------------------------------------------------------
	' Open a new table for the admin list and display the header
	' row (column headings)

	Private Sub DisplayTableHead
		Dim aLabel, I

		' get the array of column headings
		If mstrLabel <> "" Then
			aLabel = Split(mstrLabel, ",")
		ElseIf mstrDisplay <> "" Then
			aLabel = marrDisplay
		Else
			mstrError = "clsTreeList.DisplayTableHead - Must define display property"
			Exit Sub
		End If

		With Response
		.Write "<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 CLASS=""list"">" & vbCrLf
		.Write "<TR>" & vbCrLf
		For I = 0 To UBound(aLabel)
			.Write "	<TD"
			If LCase(marrJust(I)) = "center" Or LCase(marrJust(I)) = "middle" Then
				.Write " align=""center"""
			ElseIf LCase(marrJust(I)) = "right" Then
				.Write " align=""right"""
			End If
			.Write " class=""listhead"">"
			.Write Server.HTMLEncode(aLabel(I))
			.Write "</TD>" & vbCrLf
		Next
		If mstrActionLink <> "" Then
			.Write "	<TD class=""listhead"" ALIGN=""right"">Action</TD>" & vbCrLf
		End If
		.Write "</TR>" & vbCrLf
		End With
	End Sub

	'--------------------------------------------------------------
	' count total results (since OLEDB doesn't support RecordCount)

	Public Sub CountResults(rsList)
		mintResultCount = 0
		If Not rsList.EOF Then
			Do Until rsList.EOF
				mintResultCount = mintResultCount + 1
				rsList.MoveNext
			Loop
			rsList.MoveFirst
		End If
	End Sub

	'--------------------------------------------------------------
	' Display the page navigation

	Private Sub PageNav(bShowPosition)
		Dim nPages, I

		nPages = ((mintResultCount - 1) \ mintPageSize) + 1
		With Response
			.Write "<p align=""center"">" 
			.Write vbCrLf
			If bShowPosition Then
				.Write "<i>Displaying Results "
				.Write (mintPageNo * mintPageSize) + 1
				.Write " to "
				If mintPageNo = nPages - 1 Then
					.Write mintResultCount
				Else
					.Write ((mintPageNo + 1) * mintPageSize) - 1
				End If
				.Write " of "
				.Write mintResultCount
				.Write "</i><br>"
			End If
			' now show the "jump-to" navigation
			.Write "<B>Page:</B> "
			' If nPages > 20 Then
			' Else
				For I = 1 To nPages
					If I > 1 Then .Write " | "
					If I = mintPageNo + 1 Then
						.Write "<b>" & I & "</b>"
					Else
						.Write "<a href="""
						.Write Request.ServerVariables("SCRIPT_NAME")
						If mstrQueryString <> "" Then
							.Write "?" & mstrQueryString & "&"
						Else
							.Write "?"
						End If
						.Write "resultcount="
						.Write mintResultCount
						.Write "&pageno="
						.Write I - 1
						.Write """ class=""actionlink"">"
						.Write I
						.Write "</a>"
					End If
				Next
			' End If
			.Write "</p>"
			.Write vbCrLf
		End With
	End Sub

	'--------------------------------------------------------------
	' macro substition for fields (in the form ##fieldname##) that
	' are embedded within a template string
	' RETURNS: string with macros substituted

	Private Function MacroSub(rs, sTemplate)
		Dim sResult, oField

		sResult = sTemplate
		For Each oField In rs.Fields
			sResult = Replace(sResult, "##" & oField.Name & "##", oField.Value & "", 1, -1, vbTextCompare)
		Next
		MacroSub = sResult
	End Function

	'--------------------------------------------------------------
	' build HTML code for an individual row of the tree admin

	Private Function BuildRow(rsList)
		Dim sHTML, J

		sHTML = "<tr class=""list##listno##"">" & vbCrLf
		For J = 0 To UBound(marrDisplay)
			' set the alignment for the column
			sHTML = sHTML & "	<td"
			If LCase(marrJust(J)) = "center" Or LCase(marrJust(J)) = "middle" Then
				sHTML = sHTML & " align=""center"""
			ElseIf LCase(marrJust(J)) = "right" Then
				sHTML = sHTML & " align=""right"""
			End If
			If J = 0 Then
				sHTML = sHTML & ">##indentstart##"
			Else
				sHTML = sHTML & ">"
			End If

			' build the indentation for this level (if nec)
			'If J = 0 And nLevel > 0 Then
			'	sHTML = sHTML & "<table border=0 cellpadding=0 cellspacing=0><tr><td><img src=""" & Application("ASPNukeBasePath") & """images/pixel.gif"" width=""" & (nLevel * 15) & """ height=""1""></td><td>"
			'End If

			' output the column data
			If InStr(1, marrDisplay(J) & "", "##") > 0 Then
				' perform macro substitution on the field template
				sHTML = sHTML & MacroSub(rsList, marrDisplay(J))
			Else
				' perform special data conversions on field data
				Select Case rsList.Fields(marrDisplay(J)).Type
					Case adDate, adDBDate, adDBTimeStamp
						sHTML = sHTML & adoFormatDateTime(rsList.Fields(marrDisplay(J)).Value, mintDefDateFmt)
					Case Else
						sHTML = sHTML & (rsList.Fields(marrDisplay(J)).Value & "")
				End Select
			End if

			' close the indentation table for this level (if nec)
			'If J = 0 And nLevel > 0 Then
			'	sHTML = sHTML & "</td></tr></table>" & vbCrLf
			'End If

			If J = 0 Then
				sHTML = sHTML & "##indentend##</td>" & vbCrLf
			Else
				sHTML = sHTML & "</td>" & vbCrLf
			End If
		Next
		sAction = MacroSub(rsList, mstrActionLink)
		If sAction <> "" Then
			sHTML = sHTML & "	<td>" & sAction & "</td>" & vbCrLf
		End If
		BuildRow = sHTML & "</tr>" & vbCrLf
	End Function

	'--------------------------------------------------------------
	' build the hierarchical structure in the dictionary object

	Private Sub BuildTree(rsList)
		Dim sRow, nParentID

		If mobjDict Is Nothing Then Set mobjDict = Server.CreateObject("Scripting.Dictionary")
		If mstrPKeyField = "" Then
			mstrError = "clsTreeList - Must define the PrimaryKey property"
			Exit Sub
		End If
		If mstrParentField = "" Then
			mstrError = "clsTreeList - Must define the ParentField property"
			Exit Sub
		End If
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
	' display the data rows from the results set

	Private Sub DisplayRows(rsList)
		Dim aDisplay, sAction, oRE, I, J

		If mintColumns <= 0 Then
			mstrError = "clsTreeList.DisplayRows - Must define display Display"
			Exit Sub
		End If
		With Response
		For I = 0 To mintPageSize - 1
			If rsList.EOF Then Exit Sub
			.Write "<tr class=""list"
			.Write I mod 2
			.Write """>" & vbCrLf
			For J = 0 To UBound(marrDisplay)
				' set the alignment for the column
				.Write "	<td"
				If LCase(marrJust(J)) = "center" Or LCase(marrJust(J)) = "middle" Then
					.Write " align=""center"""
				ElseIf LCase(marrJust(J)) = "right" Then
					.Write " align=""right"""
				End If
				.Write ">"
				' output the column data
				If InStr(1, marrDisplay(J) & "", "##") > 0 Then
					' perform macro substitution on the field template
					.Write MacroSub(rsList, marrDisplay(J))
				Else
					' perform special data conversions on field data
					Select Case rsList.Fields(marrDisplay(J)).Type
						Case adDate, adDBDate, adDBTimeStamp
							.Write FormatDateTime(rsList.Fields(marrDisplay(J)).Value, mintDefDateFmt)
						Case Else
							.Write rsList.Fields(marrDisplay(J)).Value & ""
					End Select
				End if
				.Write "</td>" & vbCrLf
			Next
			sAction = MacroSub(rsList, mstrActionLink)
			If sAction <> "" Then
				.Write "	<td>"
				.Write sAction
				.Write "</td>" & vbCrLf
			End If
			.Write "</tr>" & vbCrLf
			rsList.MoveNext
		Next
		End With
	End Sub

	'--------------------------------------------------------------
	' display the records from the query for the specific page

	Private Sub DisplayLevel(nParentID, nRowNo, nLevelNo)
		Dim aKey, sRow, I

		If Not mobjDict.Exists(CStr(nParentID)) Then Exit Sub
		If mobjDict.Item(CStr(nParentID)) = "" Then Exit Sub
		' build the array of primary keys at this level
		aKey = Split(mobjDict.Item(CStr(nParentID)), ",")

		' build each success record in the list (from the row data)
		For I = 0 To UBound(aKey)
			sRow = mobjDict.Item("row" & aKey(I))
			If nLevelNo > 0 Then
				sRow = Replace(sRow, "##indentstart##", _
					"<table border=0 cellpadding=0 cellspacing=0><tr><td><img src=""" &_
					Application("ASPNukeBasePath") & "images/pixel.gif"" width=""" &_
					(nLevelNo * 15) & """ height=""1""></td><td>", 1, -1, vbTextCompare)
				' sRow = Replace(sRow, "##col1style##", " style=""{padding-left:" & (nLevelNo * 15) & "px}""")
				sRow = Replace(sRow, "##indentend##", "</td></tr></table>", 1, -1, vbTextCompare)
			Else
				'sRow = Replace(sRow, "##col1style##", "")
				sRow = Replace(sRow, "##indentstart##", "", 1, -1, vbTextCompare)
				sRow = Replace(sRow, "##indentend##", "", 1, -1, vbTextCompare)
			End If
			Response.Write Replace(sRow, "##listno##", CStr(nRowNo Mod 2), 1, -1, vbTextCompare) & vbCrLf

			nRowNo = nRowNo + 1
			' call the child level (if exists)
			If mobjDict.Exists(aKey(I)) Then
				If mobjDict.Item(aKey(I)) <> "" Then DisplayLevel aKey(I), nRowNo, nLevelNo + 1
			End If
		Next
	End Sub

	'--------------------------------------------------------------
	' display the records from the query for the specific page

	Public Sub Display
		Dim rsList

		' abort if the query is not defined
		If mstrQuery = "" Then
			mstrError = "clsTreeList.Display - Must define Query property first"
			Exit Sub
		End If
		' retrieve the list of records from the database
		Set rsList = adoOpenRecordset(mstrQuery)
		If rsList.EOF Then
			With Response
			.Write "<p><b class=""error"">"
			.Write "Sorry, no records were found to display here"
			.Write "</b></p>"
			End With
			Exit Sub
		End If
		' count the results and display the page navigation
		'If mintResultCount = 0 Then Call CountResults(rsList)
		'Call PageNav(True)

		' build the tree
		Call BuildTree(rsList)
		If mstrError <> "" Then Exit Sub
		' display the table header
		Call DisplayTableHead
		If mstrError <> "" Then Exit Sub

		DisplayLevel 0, 0, 0

		'If mintPageNo > 0 Then
		'	rsList.Move mintPageNo * mintPageSize
		'End If

		' display the current page of results here
		' Call PageNav(False)

		Response.Write "</table>" & vbCrLf
	End Sub

	'--------------------------------------------------------------
	' add a date format to a column

	Public Sub AddDateFormat(nColumnNo, nFormat)
		' add the format field or field template
		If UBound(marrDateFmt) < nColumnNo - 1 Then
			ReDim Preserve marrDateFmt(nColumnNo - 1)
		End If
		marrDateFmt(nColumnNo - 1) = nFormat
	End Sub

	'--------------------------------------------------------------
	' add a column to the display list

	Public Sub AddColumn(strFieldTemplate, strLabel, strJust)
		' add the display field or field template
		If UBound(marrDisplay) < mintColumns Then
			ReDim Preserve marrDisplay(UBound(marrDisplay) + 1)
		End If
		marrDisplay(mintColumns) = strFieldTemplate
		' add the label for the field
		If mstrLabel <> "" Then mstrLabel = mstrLabel & ","
		mstrLabel = mstrLabel & strLabel
		' add the justification (align) for this field 
		If UBound(marrJust) < mintColumns Then
			ReDim Preserve marrJust(UBound(marrJust) + 1)
		End If
		marrJust(mintColumns) = strJust
		mintColumns = mintColumns + 1
	End Sub

	'--------------------------------------------------------------
	' get/set the recordset Query property

	Public Property Let Query(strValue)
		mstrQuery = StrValue
	End Property

	Public Property Get Query
		Query = mstrQuery
	End Property

	'--------------------------------------------------------------
	' get/set the primary key field property

	Public Property Let PrimaryKey(strValue)
		mstrPKeyField = strValue
	End Property

	Public Property Get PrimaryKey
		PrimaryKey = mstrPKeyField
	End Property

	'--------------------------------------------------------------
	' get/set the parent field property

	Public Property Let ParentField(strValue)
		mstrParentField = strValue
	End Property

	Public Property Get ParentField
		ParentField = mstrParentField
	End Property

	'--------------------------------------------------------------
	' get/set the action link property (field template)

	Public Property Let ActionLink(strValue)
		mstrActionLink = StrValue
	End Property

	Public Property Get ActionLink
		ActionLink = mstrActionLink
	End Property

	'--------------------------------------------------------------
	' get/set the querystring property (field template)

	Public Property Let QueryString(strValue)
		mstrQueryString = StrValue
	End Property

	Public Property Get QueryString
		QueryString = mstrQueryString
	End Property

	'--------------------------------------------------------------
	' get the error msg property

	Public Property Get ErrorMsg
		ErrorMsg = mstrError
	End Property

End Class
%>