<%
' action_lib.asp
'	Defines a class clsAction to build the action links
'	that appear next to an admin list.

Class clsAction
	Public sTemplate		' action link template

	Public Sub Class_Initialize
	End Sub

	'------------------------------------------------------
	' LoadConfiguration

	Private Sub LoadConfiguration
		Dim query, rs

		query = "SELECT 
	End Sub

	'------------------------------------------------------
	' BuildTemplate

	Public Sub BuildTemplate(sOperations, sQueryStr)
		' make sure the query string begins with "&"
		If sQueryStr <> "" Then
			If Left(sQueryStr, 1) = "?" Then
				sQueryStr = "&" & Mid(sQueryStr, 2)
			ElseIf Left(sQueryString, 1) <> "&" Then
				sQueryStr = "&" & sQueryStr
			End If
		End If
		sTemplate = ""
		If InStr(1, sOperations, "edit") > 0 Then
			sTemplate = sTemplate & "<A href=""" & Request.ServerVariables("SCRIPT_NAME") &_
				"?action=edit" & sQueryStr & """>" & steGetText("edit") & "</a>"
		End If
		If InStr(1, sOperations, "delete") > 0 Then
			If sTemplate <> "" Then sTemplate = sTemplate & " . "
			sTemplate = sTemplate & "<A href=""" & Request.ServerVariables("SCRIPT_NAME") &_
				"?action=delete" & sQueryStr & """>" & steGetText("delete") & "</a>"
		End If
		If InStr(1, sOperations, "moveup") > 0 Then
			If sTemplate <> "" Then sTemplate = sTemplate & " . "
			sTemplate = sTemplate & "<A href=""" & Request.ServerVariables("SCRIPT_NAME") &_
				"?action=moveup" & sQueryStr & """>" & steGetText("up") & "</a>"
		End If
		If InStr(1, sOperations, "movedown") > 0 Then
			If sTemplate <> "" Then sTemplate = sTemplate & " . "
			sTemplate = sTemplate & "<A href=""" & Request.ServerVariables("SCRIPT_NAME") &_
				"?action=moveup" & sQueryStr & """>" & steGetText("down") & "</a>"
		End If
	End Sub

	'------------------------------------------------------
	' Replace1

	Public Function Replace1(sName1, sValue1)
		Replace1 = Replace(sTemplate, "##"&sName1&"##", sValue1)
	End Function

	Public Function Replace2(sName1, sValue1, sName2, sValue2)
		Replace2 = Replace(Replace(sTemplate, "##"&sName1&"##", sValue1), "##"&sName2&"##", sValue2)
	End Function

	Public Function Replace3(sName1, sValue1, sName2, sValue2, sName3, sValue3)
		Replace3 = Replace(Replace(Replace(sTemplate, "##"&sName1&"##", sValue1), "##"&sName2&"##", sValue2), "##"&sName3&"##", sValue3)
	End Function

	Public Function Replace4(sName1, sValue1, sName2, sValue2, sName3, sValue3, sName4, sValue4)
		Replace4 = Replace(Replace(Replace(Replace(sTemplate, "##"&sName1&"##", sValue1), "##"&sName2&"##", sValue2), "##"&sName3&"##", sValue3), "##"&sName4&"##", sValue4)
	End Function

	'------------------------------------------------------
	' Template Property

	Public Property Let Template(sValue)
		sTemplate = sValue
	End Property

	Public Property Get Template
		Template = sTemplate 
	End Property
End Class
		<A HREF="userright_list.asp?RightID=<%= aRight(0, I) %>&ParentID=<%= aRight(1, I) %>&orderno=<%= aRight(5, I) %>&action=moveup" class="actionlink"><img src="<%= Application("ASPNukeBasePath") %>img/moveup.gif" alt="<% steTxt "up" %>"></A>

%>