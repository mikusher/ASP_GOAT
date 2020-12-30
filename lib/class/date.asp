<%
'--------------------------------------------------------------------
' date.asp
'	Class for building HTML date inputs
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

Class clsDate

	Private mboolAbbrevMonth	' abbreviate month names in drop-list?
	Private mdatSelected		' selected date
	Private mstrSeparator		' separator string for the date parts
	Private mintStartYear		' starting year for year drop-list
	Private mintEndYear			' ending year for year drop-list

	'------------------------------------------------------------------------
	' constructor for the date input control

	Private Sub Class_Initialize
		mboolAbbrevMonth = true
		mdatSelected = ""
		mstrSeperator = "/"
		mintStartYear = Year(DateAdd("y", -12, Now()))
		mintEndYear = Year(Now())
	End Sub

	'------------------------------------------------------------------------
	' build the year drop-list input for the date control

	Private Sub BuildYear(sInpName)
		Dim nSelected, I

		If IsDate(mdatSelected) Then nSelected = Year(CDate(mdatSelected)) Else nSelected = 0
		With Response
		.Write "<select name=""" & sInpName & """ class=""form"">" & vbCrLf
		For I = mintStartYear To mintEndYear
			.Write "<option value="""
			.Write I
			.Write """"
			If nSelected = I Then .Write " SELECTED"
			.Write ">"
			.Write I
			.Write "</option>"
			.Write vbCrLf
			If I - mintStartYear > 150 Then Exit For
		Next
		.Write "</select>" & vbCrLf
		End With
	End Sub

	'------------------------------------------------------------------------
	' build the day drop-list input for the date control

	Private Sub BuildDay(sInpName)
		Dim nSelected, I

		If IsDate(mdatSelected) Then nSelected = Day(CDate(mdatSelected)) Else nSelected = 0
		With Response
		.Write "<select name=""" & sInpName & """ class=""form"">" & vbCrLf
		For I = 1 To 31
			.Write "<option value="""
			.Write I
			.Write """"
			If nSelected = I Then .Write " SELECTED"
			.Write ">"
			.Write I
			.Write "</option>"
			.Write vbCrLf
		Next
		.Write "</select>" & vbCrLf
		End With
	End Sub

	'------------------------------------------------------------------------
	' build the month drop-list input for the date control

	Private Sub BuildMonth(sInpName)
		Dim nSelected, I

		If IsDate(mdatSelected) Then nSelected = Month(CDate(mdatSelected)) Else nSelected = 0
		With Response
		.Write "<select name=""" & sInpName & """ class=""form"">" & vbCrLf
		For I = 1 To 12
			.Write "<option value="""
			.Write I
			.Write """"
			If nSelected = I Then .Write " SELECTED"
			.Write ">"
			.Write MonthName(I, mboolAbbrevMonth)
			.Write "</option>"
			.Write vbCrLf
		Next
		.Write "</select>" & vbCrLf
		End With
	End Sub

	'------------------------------------------------------------------------
	' display a simple date input control (mm/dd/yyyy)

	Public Sub DateInput(sPrefix)
		' 
		With Response
		Call BuildMonth(sPrefix & "_mon")
		.Write mstrSeparator
		Call BuildDay(sPrefix & "_day")
		.Write mstrSeparator
		Call BuildYear(sPrefix & "_yr")
		End With
	End Sub

	'------------------------------------------------------------------------
	' get/set the abbreviated month property

	Public Property Let AbbrevMonth(boolValue)
		mboolAbbrevMonth = boolValue
	End Property

	Public Property Get AbbrevMonth
		AbbrevMonth = mboolAbbrevMonth
	End Property

	'------------------------------------------------------------------------
	' get/set the selected date property

	Public Property Let Selected(datValue)
		If datValue & "" = "" Then
			mdatSelected = ""
		Else
			mdatSelected = datValue
		End If
	End Property

	Public Property Get Selected
		Selected = mdatSelected
	End Property

	'------------------------------------------------------------------------
	' get/set the starting year property

	Public Property Let StartYear(intValue)
		mintStartYear = intValue
	End Property

	Public Property Get StartYear
		StartYear = mintStartYear
	End Property

	'------------------------------------------------------------------------
	' get/set the ending year property

	Public Property Let EndYear(intValue)
		mintEndYear = intValue
	End Property

	Public Property Get EndYear
		EndYear = mintEndYear
	End Property
End Class

'------------------------------------------------------------------------
' recontruct a date value from the individual components

Function datForm(sPrefix)
	Dim sDate

	If Request.Form(sPrefix & "_mon") <> "" And Request.Form(sPrefix & "_day") <> "" _
		And Request.Form(sPrefix & "_yr") <> "" Then
		sDate = Request.Form(sPrefix & "_mon") & "/" & Request.Form(sPrefix & "_day") &_
			"/" & Request.Form(sPrefix & "_yr")
		' add the optional time information
		If Request.Form(sPrefix & "_hr") <> "" And Request.Form(sPrefix & "_min") <> "" Then
			sDate = sDate & " " & Request.Form(sPrefix & "_hr") & ":" & Request.Form(sPrefix & "_min")
		End If
	End If
	If IsDate(sDate) Then
		datForm = CDate(sDate)
	Else
		datForm = ""
	End If
End Function

'------------------------------------------------------------------------
' recontruct an "SQL quoted" date value from the individual components

Function datQForm(sPrefix)
	Dim dValue

	dValue = datForm(sPrefix)
	If CStr(dValue) = "" Then
		datQForm = "NULL"
	Else
		datQForm = "'" & Year(dValue) & "-" & Month(dValue) & "-" & Day(dValue) & " " & Hour(dValue) & ":" & Minute(dValue) & ":" & Second(dValue) & "'"
	End If
End Function
%>