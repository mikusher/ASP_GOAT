<%
' -------------------------------------------------------------------
' config_lib.asp
'	Support routines for building the dynamic module configuration
'	forms (/admin/configure.asp)
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

Function cfgValidateForm(sErrorMsg)
	Dim bSuccess
	bSuccess = cfgCheckRequired(sErrorMsg)
	If bSuccess Then
		bSuccess = cfgRegExTest(sErrorMsg)
		If bSuccess Then
			bSuccess = cfgConversionTest(sErrorMsg)
		End If
	End If
	cfgValidateForm = bSuccess
End Function

'--------------------------------------------------------------------
' save the configuration form

Function cfgSave(nModuleID, sErrorMsg)
	Dim aParam, I, sUpdate

	If cfgValidateForm(sErrorMsg) Then
		' save the parameters here
		aParam = Split(Replace(steForm("configparam"), " ", ""), ",")
		For I = 0 To UBound(aParam)
			sUpdate = "UPDATE tblModuleParam SET ParamValue = " & steQForm(aParam(I)) & ", Modified = " & adoGetDate & " " &_
				"WHERE	ModuleID = " & nModuleID & " " &_
				"AND	ParamName = '" & Replace(aParam(I), "'", "''") & "'"
			Call adoExecute(sUpdate)
		Next

		Call modParamCache(nModuleID, "")
		cfgSave = True
	Else
		cfgSave = False
	End If
End Function

'--------------------------------------------------------------------
' build an option select here

Sub cfgFormDropList(nParamID, nTypeID, sParamName, sParamValue)
	Dim sStat, rsOpt

	sStat = "SELECT OptionValue, OptionLabel " &_
			"FROM	tblModuleParamOption " &_
			"WHERE	ParamID = " & nParamID & " " &_
			"AND	Archive = 0 " &_
			"AND	IsValid = 1 " &_
			"ORDER BY OrderNo"
	Set rsOpt = adoOpenRecordset(sStat)
	If rsOpt.EOF Then
		' try to find options associated with the type
		sStat = "SELECT OptionValue, OptionLabel " &_
				"FROM	tblModuleParamOption " &_
				"WHERE	TypeID = " & nTypeID & " " &_
				"AND	Archive = 0 " &_
				"AND	IsValid = 1 " &_
				"ORDER BY OrderNo"
		Set rsOpt = adoOpenRecordset(sStat)
		If rsOpt.EOF Then
			Response.Write "<B CLASS=""error"">Unable to retrieve options for parameter (ID = " & nParamID & " / TypeID = " & nTypeID & ")"
			Exit Sub
		End If
	End If
	With Response
	.Write "	<SELECT NAME=""" & sParamName & """ CLASS=""form"">" & vbCrLf
	.Write "	<OPTION VALUE=""""> -- Choose --" & vbCrLf
	Do Until rsOpt.EOF
		.Write "	<OPTION VALUE="""
		.Write Server.HTMLEncode(rsOpt.Fields("OptionValue").Value)
		.Write """"
		If sParamValue = rsOpt.Fields("OptionValue").Value Then .Write " SELECTED"
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

Sub cfgFormInput(sHTMLInput, sParamName, sParamValue)
	With Response
	Select Case UCase(sHTMLInput)
		Case "TEXT"
			.Write "	<input type=""text"" name=""" & sParamName & """"
			.Write " value=""" & Server.HTMLEncode(sParamValue)
			.Write """ size=""32"" maxlength=""255"" class=""form"">" & vbCrLf
		Case "RADIO"
			.Write "	<input type=""radio"" name=""" & sParamName & """ value=""1"""
			If sParamValue = "1" Then .Write " CHECKED"
			.Write " class=""formradio""> Yes" & vbCrLf
			.Write "	<input type=""radio"" name=""" & sParamName & """ value=""0"""
			If sParamValue <> "1" Then .Write " CHECKED"
			.Write " class=""formradio""> No" & vbCrLf
		Case "TEXTAREA"
			.Write "<textarea name=""" & sParamName & """"
			.Write " cols=""52"" rows=""8"" class=""form"">"
			.Write Server.HTMLEncode(sParamValue)
			.Write "</textarea>" & vbCrLf
	End Select
	End With
End Sub

'--------------------------------------------------------------------
' build the configuration form from the database

Sub cfgConfigForm(nModuleID)
	Dim sStat, rsParam, sParam, sRegEx, sConv, sReq, sConvMethod, sRegExpPattern

	sStat = "SELECT	p.ParamID, p.Label, p.ParamName, p.ParamValue, p.MinValue, p.MaxValue, " &_
			"		p.HelpText, p.IsRequired, mt.LabelPos, mt.HTMLInputType, mt.HasOptions, " &_
			"		mt.TypeID, mt.ASPConvertFunction, mt.RegExValidate " &_
			"FROM	tblModuleParam p " &_
			"INNER JOIN	tblModuleParamType mt ON p.TypeID = mt.TypeID " &_
			"WHERE	p.ModuleID = " & nModuleID & " " &_
			"AND	p.Archive = 0 " &_
			"ORDER BY p.OrderNo"
	Set rsParam = adoOpenRecordset(sStat)
	With Response
	Do Until rsParam.EOF
		sParam = sParam & "," & Server.HTMLEncode(rsParam.Fields("ParamName").Value)
		If  steRecordBoolValue(rsParam, "IsRequired") Then sReq = sReq & "," & Server.HTMLEncode(rsParam.Fields("ParamName").Value)

		' build a list of regular expression validation patterns
		If Trim(rsParam.Fields("RegExValidate").Value & "") <> "" Then
			sRegEx = sRegEx & "," & rsParam.Fields("ParamName").Value
			sRegExpPattern = sRegExpPattern & "&" & Server.URLEncode(rsParam.Fields("RegExValidate").Value)
		End If
		' build a list of ASP conversion functions to test values entered
		If Trim(rsParam.Fields("ASPConvertFunction").Value & "") <> "" Then
			sConv = sConv & "," & rsParam.Fields("ParamName").Value
			sConvMethod = sConvMethod & "&" & Server.URLEncode(rsParam.Fields("ASPConvertFunction").Value)		
		End If
		' include the hidden vars
		.Write "<input type=""hidden"" name=""config" & rsParam.Fields("ParamName").Value & """ value=""" & Server.HTMLEncode(rsParam.Fields("Label").Value) & """>" & vbCrLf

		If Trim(rsParam.Fields("MinValue").Value & "") <> "" Then
			.Write "<input type=""hidden"" name=""configminvalue"" value=""" & Server.HTMLEncode(rsParam.Fields("ParamName").Value) & """>" & vbCrLf
			.Write "<input type=""hidden"" name=""configminvaluevalue"" value=""" & Server.HTMLEncode(rsParam.Fields("MinValue").Value) & """>" & vbCrLf
		End If
		If Trim(rsParam.Fields("MaxValue").Value & "") <> "" Then
			.Write "<input type=""hidden"" name=""configmaxvalue"" value=""" & Server.HTMLEncode(rsParam.Fields("ParamName").Value) & """>" & vbCrLf
			.Write "<input type=""hidden"" name=""configmaxvaluevalue"" value=""" & Server.HTMLEncode(rsParam.Fields("MaxValue").Value) & """>" & vbCrLf
		End If

		' display the label first
		If rsParam.Fields("LabelPos").Value = "LEFT" Then
			.Write "<TR>" & vbCrLf
			.Write "	<TD CLASS=""forml"">"
			.Write Server.HTMLEncode(rsParam.Fields("Label").Value)
			.Write "</TD><TD>&nbsp;&nbsp;</TD>" & vbCrLf
			.Write "	<TD CLASS=""formd"">"
		Else
			.Write "<TR>" & vbCrLf
			.Write "	<TD COLSPAN=""3"" CLASS=""forml"">"
			.Write Server.HTMLEncode(rsParam.Fields("Label").Value)
			.Write "<BR>" & vbCrLf
			.Write "	<DIV CLASS=""formd"">"
		End If

		' display the form input
		If  steRecordBoolValue(rsParam, "HasOptions") Then
			' input is a selection list - build that
			cfgFormDropList rsParam.Fields("ParamID").Value, rsParam.Fields("TypeID").Value, _
				rsParam.Fields("ParamName").Value, rsParam.Fields("ParamValue").Value
		Else
			' input is a normal form input - build that
			cfgFormInput rsParam.Fields("HTMLInputType").Value, rsParam.Fields("ParamName").Value, _
				rsParam.Fields("ParamValue").Value
		End if

		' close off this parameter input
		If rsParam.Fields("LabelPos").Value = "LEFT" Then
			.Write "	</TD>" & vbCrLf
			.Write "</TR>" & vbCrLf
		Else
			.Write "	</DIV>" & vbCrLf
			.Write "	</TD>" & vbCrLf
			.Write "</TR>" & vbCrLf
		End If
		rsParam.MoveNext
	Loop
	rsParam.Close
	Set rsParam = Nothing
	' include the hidden vars
	.Write "<input type=""hidden"" name=""configparam"" value=""" & Mid(sParam, 2) & """>" & vbCrLf
	.Write "<input type=""hidden"" name=""configregex"" value=""" & Mid(sRegEx, 2) & """>" & vbCrLf
	.Write "<input type=""hidden"" name=""configregexpattern"" value=""" & Mid(sRegExpPattern, 2) & """>" & vbCrLf
	.Write "<input type=""hidden"" name=""configconv"" value=""" & Mid(sConv, 2) & """>" & vbCrLf
	.Write "<input type=""hidden"" name=""configconvmethod"" value=""" & Mid(sConvMethod, 2) & """>" & vbCrLf
	.Write "<input type=""hidden"" name=""configrequired"" value=""" & Mid(sReq, 2) & """>" & vbCrLf
	End With
End Sub

'--------------------------------------------------------------------
' check for required fields in the form

Function cfgCheckRequired(sErrorMsg)
	Dim aReq, aLabel, bSuccess, I

	bSuccess = True
	aReq = Split(Replace(steForm("configrequired"), " ", ""), ",")
	For I = 0 To UBound(aReq)
		If Trim(steForm(aReq(I))) = "" Then
			sErrorMsg = sErrorMsg & "Please enter the required value for """ & steForm("config" & aReq(I)) & """<BR>"
			bSuccess = False
		End If
	Next
	cfgCheckRequired = bSuccess
End Function

'--------------------------------------------------------------------
' regular expression test on the fields in the form

Function cfgRegExTest(sErrorMsg)
	Dim regEx, bSuccess, aRegEx, aPattern

	bSuccess = True
	aRegEx = Split(Replace(steForm("configregex"), " ", ""), ",")
	aPattern = Split(Replace(steForm("configregexpattern"), " ", ""), "&")
	If UBound(aRegEx) <> UBound(aPattern) Then
		sErrorMsg = sErrorMsg & "Number of RegEx fields (" & (UBound(aRegEx) + 1) &_
			") does not match number of RegEx patterns (" & (UBound(aPattern) + 1) & ")<BR>"
		cfgRegExTest = False
		Exit Function
	End If

	Set regEx = New RegExp
	regEx.Global = True
	regEx.IgnoreCase = True
	For I = 0 To UBound(aRegEx)
		regEx.Pattern = cfgURLDecode(aPattern(I))
		If Not regEx.Test(steForm(aRegEx(I))) Then
			sErrorMsg = sErrorMsg & "Invalid value entered for """ & steForm("config" & aRegEx(I)) & """ - failed regex test { " & regEx.Pattern & "}<BR>"
			bSuccess = False
		End If
	Next
	Set regEx = Nothing
	cfgRegExTest = bSuccess
End Function

'--------------------------------------------------------------------
' ASP conversion function test for parameter values

Function cfgConversionTest(sErrorMsg)
	Dim regEx, bSuccess, aConv, aConvMethod

	bSuccess = True
	' build the conversion test arrays
	aConv = Split(Replace(steForm("configconv"), " ", ""), ",")
	aConvMethod = Split(Replace(steForm("configconvmethod"), " ", ""), "&")
	If UBound(aConv) <> UBound(aConvMethod) Then
		sErrorMsg = sErrorMsg & "Number of conversion fields (" & (UBound(aConv) + 1) &_
			") does not match number of conversion methods (" & (UBound(aConvMethod) + 1) & ")<BR>"
		cfgConversionTest = False
		Exit Function
	End If

	' test each of the parameter values
	On Error Resume Next
	For I = 0 To UBound(aConv)
		vDummy = Eval(cfgURLDecode(aConvMethod(I)) & "(""" & Replace(steForm(aConv(I)), """", """""") & """)")
		If Err.Number <> 0 Then		
			sErrorMsg = sErrorMsg & "Invalid value entered for """ & steForm("config" & aConv(I)) & """<BR>"
			Err.Clear
			bSuccess = False
		End If
	Next
	On Error Goto 0
	cfgConversionTest = bSuccess
End Function

'--------------------------------------------------------------------
' check for minimum / maximum values in the form

Function cfgCheckMinMaxValues(sErrorMsg)
	Dim aMin, aMinValue
	Dim aMax, aMaxValue, bSuccess, I

	bSuccess = True
	aMin = Split(Replace(steForm("configminvalue"), " ", ""), ",")
	aMinValue = Split(Replace(steForm("configminvaluevalue"), " ", ""), ",")
	For I = 0 To UBound(aMin)
		If Trim(steForm(aMin(I))) = "" Then
			sErrorMsg = sErrorMsg & "Please enter the required value for """ & steForm("config" & aMin(I)) & """<BR>"
			bSuccess = False
		End If
	Next
	cfgCheckRequired = bSuccess
End Function

' -------------------------------------------------------------------------
' perform a URL decode to retrieve the original value

Function cfgURLDecode(sConvert)
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
	
	cfgURLDecode = sOutput
End Function
%>