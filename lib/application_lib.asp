<%
' -------------------------------------------------------------------
' application_lib.asp
'	Support routines for building the dynamic ASP Nuke application
'	configuration forms (/admin/configure.asp)
'
' AUTH:	Ken Richards
' DATE:	10/13/03
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

'--------------------------------------------------------------------
' validate all form values

Function appValidateForm(sErrorMsg)
	Dim bSuccess
	bSuccess = appCheckRequired(sErrorMsg)
	If bSuccess Then
		bSuccess = appRegExTest(sErrorMsg)
		If bSuccess Then
			bSuccess = appConversionTest(sErrorMsg)
		End If
	End If
	appValidateForm = bSuccess
End Function

'--------------------------------------------------------------------
' save the configuration form

Function appSave(nTabID, sErrorMsg)
	Dim aVar, I, sUpdate

	If appValidateForm(sErrorMsg) Then
		' save the variables here
		aVar = Split(Replace(steForm("appvariable"), " ", ""), ",")
		For I = 0 To UBound(aVar)
			' update this individual configuration
			sUpdate = "UPDATE tblApplicationVar SET VarValue = " & steQForm(aVar(I)) & ", Modified = " & adoGetDate & " " &_
				"WHERE	TabID = " & nTabID & " " &_
				"AND	VarName = '" & Replace(aVar(I), "'", "''") & "'"
			Call adoExecute(sUpdate)
		Next

		' reload the configuration immediately
		Call steLoadConfig
		appSave = True
	Else
		appSave = False
	End If
End Function

'--------------------------------------------------------------------
' build an option select here

Sub appFormDropList(nVarID, nTypeID, sVarName, sVarValue)
	Dim sStat, rsOpt

	sStat = "SELECT OptionValue, OptionLabel " &_
			"FROM	tblApplicationVarOption " &_
			"WHERE	VarID = " & nVarID & " " &_
			"AND	Archive = 0 " &_
			"AND	IsValid = 1 " &_
			"ORDER BY OrderNo"
	Set rsOpt = adoOpenRecordset(sStat)
	If rsOpt.EOF Then
		' try to find options associated with the type
		sStat = "SELECT OptionValue, OptionLabel " &_
				"FROM	tblApplicationVarOption " &_
				"WHERE	TypeID = " & nTypeID & " " &_
				"AND	Archive = 0 " &_
				"AND	IsValid = 1 " &_
				"ORDER BY OrderNo"
		Set rsOpt = adoOpenRecordset(sStat)
		If rsOpt.EOF Then
			Response.Write "<B CLASS=""error"">Unable to retrieve options for variable (ID = " & nVarID & " / TypeID = " & nTypeID & ")"
			Exit Sub
		End If
	End If
	With Response
	.Write "	<SELECT NAME=""" & sVarName & """ CLASS=""form"">" & vbCrLf
	.Write "	<OPTION VALUE=""""> -- Choose --" & vbCrLf
	Do Until rsOpt.EOF
		.Write "	<OPTION VALUE="""
		.Write Server.HTMLEncode(rsOpt.Fields("OptionValue").Value)
		.Write """"
		If sVarValue = rsOpt.Fields("OptionValue").Value Then .Write " SELECTED"
		.Write "> "
		.Write Server.HTMLEncode(rsOpt.Fields("OptionLabel").Value) & vbCrLf
		rsOpt.MoveNext
	Loop
	.Write "	</SELECT>"
	End With
	rsOpt.Close
	Set rsOpt = Nothing
End Sub

'--------------------------------------------------------------------
' build a normal select

Sub appFormInput(sHTMLInput, sVarName, sVarValue)
	With Response
	Select Case UCase(sHTMLInput)
		Case "TEXT"
			.Write "	<input type=""text"" name=""" & sVarName & """"
			.Write " value=""" & Server.HTMLEncode(sVarValue)
			.Write """ size=""32"" maxlength=""255"" class=""form"">" & vbCrLf
		Case "RADIO"
			.Write "	<input type=""radio"" name=""" & sVarName & """ value=""1"""
			If sVarValue = "1" Then .Write " CHECKED"
			.Write " class=""formradio""> Yes" & vbCrLf
			.Write "	<input type=""radio"" name=""" & sVarName & """ value=""0"""
			If sVarValue <> "1" Then .Write " CHECKED"
			.Write " class=""formradio""> No" & vbCrLf
		Case "TEXTAREA"
			.Write "<textarea name=""" & sVarName & """"
			.Write " cols=""52"" rows=""8"" class=""form"">"
			.Write Server.HTMLEncode(sVarValue)
			.Write "</textarea>" & vbCrLf
	End Select
	End With
End Sub

'--------------------------------------------------------------------
' build the configuration form from the database

Sub appConfigForm(nTabID)
	Dim sStat, rsVar, sVar, sRegExp, sConv, sReq, sConvMethod, sRegExpPattern

	sStat = "SELECT	v.VarID, v.Label, v.VarName, v.VarValue, v.MinValue, v.MaxValue, " &_
			"		v.HelpText, v.IsRequired, mt.LabelPos, mt.HTMLInputType, mt.HasOptions, " &_
			"		mt.TypeID, mt.ASPConvertFunction, mt.RegExValidate " &_
			"FROM	tblApplicationVar v " &_
			"INNER JOIN	tblApplicationVarType mt ON v.TypeID = mt.TypeID " &_
			"WHERE	v.TabID = " & nTabID & " " &_
			"AND	v.Archive = 0 " &_
			"ORDER BY v.OrderNo"
	Set rsVar = adoOpenRecordset(sStat)
	With Response
	Do Until rsVar.EOF
		sVar = sVar & "," & Server.HTMLEncode(rsVar.Fields("VarName").Value)
		If steRecordBoolValue(rsVar, "IsRequired") Then sReq = sReq & "," & Server.HTMLEncode(rsVar.Fields("VarName").Value)

		' build a list of regular expression validation patterns
		If Trim(rsVar.Fields("RegExValidate").Value & "") <> "" Then
			sRegEx = sRegEx & "," & rsVar.Fields("VarName").Value
			sRegExpPattern = sRegExpPattern & "&" & Server.URLEncode(rsVar.Fields("RegExValidate").Value)
		End If
		' build a list of ASP conversion functions to test values entered
		If Trim(rsVar.Fields("ASPConvertFunction").Value & "") <> "" Then
			sConv = sConv & "," & rsVar.Fields("VarName").Value
			sConvMethod = sConvMethod & "&" & Server.URLEncode(rsVar.Fields("ASPConvertFunction").Value)		
		End If
		' include the hidden vars
		.Write "<input type=""hidden"" name=""config" & rsVar.Fields("VarName").Value & """ value=""" & Server.HTMLEncode(rsVar.Fields("Label").Value) & """>" & vbCrLf

		If Trim(rsVar.Fields("MinValue").Value & "") <> "" Then
			.Write "<input type=""hidden"" name=""appvariableminvalue"" value=""" & Server.HTMLEncode(rsVar.Fields("VarName").Value) & """>" & vbCrLf
			.Write "<input type=""hidden"" name=""appvariableminvaluevalue"" value=""" & Server.HTMLEncode(rsVar.Fields("MinValue").Value) & """>" & vbCrLf
		End If
		If Trim(rsVar.Fields("MaxValue").Value & "") <> "" Then
			.Write "<input type=""hidden"" name=""appvariablemaxvalue"" value=""" & Server.HTMLEncode(rsVar.Fields("VarName").Value) & """>" & vbCrLf
			.Write "<input type=""hidden"" name=""appvariablemaxvaluevalue"" value=""" & Server.HTMLEncode(rsVar.Fields("MaxValue").Value) & """>" & vbCrLf
		End If

		' display the label first
		If rsVar.Fields("LabelPos").Value = "LEFT" Then
			.Write "<TR>" & vbCrLf
			.Write "	<TD CLASS=""forml"">"
			.Write Server.HTMLEncode(rsVar.Fields("Label").Value)
			.Write "</TD><TD>&nbsp;&nbsp;</TD>" & vbCrLf
			.Write "	<TD CLASS=""formd"">"
		Else
			.Write "<TR>" & vbCrLf
			.Write "	<TD COLSPAN=""3"" CLASS=""forml"">"
			.Write Server.HTMLEncode(rsVar.Fields("Label").Value)
			.Write "<BR>" & vbCrLf
			.Write "	<DIV CLASS=""formd"">"
		End If

		' display the form input
		If steRecordBoolValue(rsVar, "HasOptions") Then
			' input is a selection list - build that
			appFormDropList rsVar.Fields("VarID").Value, rsVar.Fields("TypeID").Value, _
				rsVar.Fields("VarName").Value, rsVar.Fields("VarValue").Value
		Else
			' input is a normal form input - build that
			appFormInput rsVar.Fields("HTMLInputType").Value, rsVar.Fields("VarName").Value, _
				rsVar.Fields("VarValue").Value
		End if

		' close off this variable input
		If rsVar.Fields("LabelPos").Value = "LEFT" Then
			.Write "	</TD>" & vbCrLf
			.Write "</TR>" & vbCrLf
		Else
			.Write "	</DIV>" & vbCrLf
			.Write "	</TD>" & vbCrLf
			.Write "</TR>" & vbCrLf
		End If
		rsVar.MoveNext
	Loop
	rsVar.Close
	Set rsVar = Nothing
	' include the hidden vars
	.Write "<input type=""hidden"" name=""appvariable"" value=""" & Mid(sVar, 2) & """>" & vbCrLf
	.Write "<input type=""hidden"" name=""appvariableregex"" value=""" & Mid(sRegEx, 2) & """>" & vbCrLf
	.Write "<input type=""hidden"" name=""appvariableregexpattern"" value=""" & Mid(sRegExpPattern, 2) & """>" & vbCrLf
	.Write "<input type=""hidden"" name=""appvariableconv"" value=""" & Mid(sConv, 2) & """>" & vbCrLf
	.Write "<input type=""hidden"" name=""appvariableconvmethod"" value=""" & Mid(sConvMethod, 2) & """>" & vbCrLf
	.Write "<input type=""hidden"" name=""appvariablerequired"" value=""" & Mid(sReq, 2) & """>" & vbCrLf
	End With
End Sub

'--------------------------------------------------------------------
' check for required fields in the form

Function appCheckRequired(sErrorMsg)
	Dim aReq, aLabel, bSuccess, I

	bSuccess = True
	aReq = Split(Replace(steForm("appvariablerequired"), " ", ""), ",")
	For I = 0 To UBound(aReq)
		If Trim(steForm(aReq(I))) = "" Then
			sErrorMsg = sErrorMsg & "Please enter the required value for """ & steForm("config" & aReq(I)) & """<BR>"
			bSuccess = False
		End If
	Next
	appCheckRequired = bSuccess
End Function

'--------------------------------------------------------------------
' regular expression test on the fields in the form

Function appRegExTest(sErrorMsg)
	Dim regEx, bSuccess, aRegEx, aPattern

	bSuccess = True
	aRegEx = Split(Replace(steForm("appvariableregex"), " ", ""), ",")
	aPattern = Split(Replace(steForm("appvariableregexpattern"), " ", ""), ",")
	If UBound(aRegEx) <> UBound(aPattern) Then
		sErrorMsg = sErrorMsg & "Number of RegEx fields (" & (UBound(aRegEx) + 1) &_
			") does not match number of RegEx patterns (" & (UBound(aPattern) + 1) & ")<BR>"
		appRegExTest = False
		Exit Function
	End If

	Set regEx = New RegExp
	regEx.Global = True
	regEx.IgnoreCase = True
	For I = 0 To UBound(aRegEx)
		regEx.Pattern = appURLDecode(aPattern(I))
		If Not regEx.Test(steForm(aRegEx(I))) Then
			sErrorMsg = sErrorMsg & "Invalid value entered for """ & steForm("config" & aRegEx(I)) & """<BR>"
			bSuccess = False
		End If
	Next
	Set regEx = Nothing
	appRegExTest = bSuccess
End Function

'--------------------------------------------------------------------
' ASP conversion function test for variable values

Function appConversionTest(sErrorMsg)
	Dim regEx, bSuccess, aConv, aConvMethod

	bSuccess = True
	' build the conversion test arrays
	aConv = Split(Replace(steForm("appvariableconv"), " ", ""), ",")
	aConvMethod = Split(Replace(steForm("appvariableconvmethod"), " ", ""), ",")
	If UBound(aConv) <> UBound(aConvMethod) Then
		sErrorMsg = sErrorMsg & "Number of conversion fields (" & (UBound(aConv) + 1) &_
			") does not match number of conversion methods (" & (UBound(aConvMethod) + 1) & ")<BR>"
		appConversionTest = False
		Exit Function
	End If

	' test each of the variable values
	On Error Resume Next
	For I = 0 To UBound(aConv)
		vDummy = Eval(appURLDecode(aConvMethod(I)) & "(""" & Replace(steForm(aConv(I)), """", """""") & """)")
		If Err.Number <> 0 Then		
			sErrorMsg = sErrorMsg & "Invalid value entered for """ & steForm("config" & aConv(I)) & """<BR>"
			Err.Clear
			bSuccess = False
		End If
	Next
	On Error Goto 0
	appConversionTest = bSuccess
End Function

'--------------------------------------------------------------------
' check for minimum / maximum values in the form

Function appCheckMinMaxValues(sErrorMsg)
	Dim aMin, aMinValue
	Dim aMax, aMaxValue, bSuccess, I

	bSuccess = True
	aMin = Split(Replace(steForm("appvariableminvalue"), " ", ""), ",")
	aMinValue = Split(Replace(steForm("appvariableminvaluevalue"), " ", ""), ",")
	For I = 0 To UBound(aMin)
		If Trim(steForm(aMin(I))) = "" Then
			sErrorMsg = sErrorMsg & "Please enter the required value for """ & steForm("config" & aMin(I)) & """<BR>"
			bSuccess = False
		End If
	Next
	appCheckRequired = bSuccess
End Function

' -------------------------------------------------------------------------
' perform a URL decode to retrieve the original value

Function appURLDecode(sConvert)
	Dim aSplit
	Dim sOutput
	Dim I
	If IsNull(sConvert) Then
	   lgnURLDecode = ""
	   Exit Function
	End If
	
	' convert all pluses to spaces
	sOutput = Replace(sConvert, "+", " ")
	
	' next convert %hexdigits to the character
	If InStr(1, sOutput, "%") > 0 Then
		aSplit = Split(sOutput, "%")
		
		If IsArray(aSplit) Then
		   sOutput = aSplit(0)
		   For I = LBound(aSplit) to UBound(aSplit) - 1
		      sOutput = sOutput & Chr("&H" & Left(aSplit(i + 1), 2)) & Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
		   Next
		End If
	End If
	
	appURLDecode = sOutput
End Function
%>