<%
'--------------------------------------------------------------------
' module_lib.asp
'	Contains the code necessary for caching / updating the module
'	configuration as defined in the site administration.
'
' REQ: /lib/ado_lib.asp
' AUTH:	Ken Richards
' DATE:	07/25/2001
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


Const MOD_REFRESH_INTERVAL = 15

Dim modError

' initialize the module configuration when application is restarted
' DO NOT refresh every 15 minutes as INTERVAL above indicates
' cache will be refreshed automatically as needed by module admin
If Not IsDate(Application("MODULEREFRESH")) Then modRefresh(True)


'--------------------------------------------------------------------
' modRefresh
'	Refresh the modulce configuration cache

Sub modRefresh(bForce)
	Dim nCount, sGroupList, sLastGroup

	If Not bForce And IsDate(Application("MODULEREFRESH")) Then
		If DateDiff("n", Application("MODULEREFRESH"), Now()) < MOD_REFRESH_INTERVAL Then Exit Sub
	End If

	' retrieve all of the configured modules here
	sStat = "SELECT mg.GroupCode, m.Title, mc.FolderName AS CategoryFolder, m.FolderName, " &_
			"		mg.HasSize140Module, m.Size140Module, mg.HasSizeFullModule, m.SizeFullModule " &_
			"FROM	tblModule m " &_
			"INNER JOIN	tblModuleCategory mc on mc.CategoryID = m.CategoryID " &_
			"INNER JOIN	tblModuleGroupPos mgp on mgp.ModuleID = m.ModuleID " &_
			"INNER JOIN	tblModuleGroup mg on mgp.GroupID = mg.GroupID " &_
			"WHERE	mgp.Active <> 0 " &_
			"AND	mgp.Archive = 0 " &_
			"AND	m.Archive = 0 " &_
			"ORDER BY mg.OrderNo, mgp.OrderNo"
	Set rsMod = adoOpenRecordset(sStat)
	Application.Lock
	Do Until rsMod.EOF
		If sLastGroup <> rsMod.Fields("GroupCode").Value Then
			' store the count for the last group (if nec)
			sGroupList = sGroupList & "," & rsMod.Fields("GroupCode").Value
			sLastGroup = rsMod.Fields("GroupCode").Value
			nCount = 1
		End If
		If modRecordBoolValue(rsMod, "HasSize140Module") Then
			Application("MODULE" & sLastGroup & nCount) = "module/" & rsMod.Fields("CategoryFolder").Value & "/" & rsMod.Fields("FolderName") & "/" & rsMod.Fields("Size140Module").Value		
		End If
		If modRecordBoolValue(rsMod, "HasSizeFullModule") Then
			Application("MODULE" & sLastGroup & nCount) = "module/" & rsMod.Fields("CategoryFolder").Value & "/" & rsMod.Fields("FolderName") & "/" & rsMod.Fields("SizeFullModule").Value
		End If
		rsMod.MoveNext
		nCount = nCount + 1
	Loop
	rsMod.Close
	Set rsMod = Nothing

	' store the ordered list of groups here
	Application("MODULEREFRESH") = Now()
	Application("MODULEGROUPLIST") = sGroupList
	Application.UnLock
End Sub

'--------------------------------------------------------------------
' modShowGroup
'	Display a module group defined in the application cache
'	Admin for this is "Modules" in the nuke administration panel.

Sub modShowGroup(sGroupCode)
	Dim nCount

	nCount = 1
	Do Until Application("MODULE" & sGroupCode & nCount) = ""
		On Error Resume Next
		Server.Execute(Application("ASPNukeBasePath") & Application("MODULE" & sGroupCode & nCount))
		If Err.Number <> 0 Then
			Response.Write "<p><b class=""error"">Unable to display module """ & Application("MODULE" & sGroupCode & nCount) & """<br>"
			Response.Write Err.Number & " - " & Err.Description & "</b></p>" & vbCrLf
		End If
		On Error Goto 0
		nCount = nCount + 1
	Loop
End Sub

'--------------------------------------------------------------------
' modParamCache
'	Cache the configuration parameters for the module identified
'	by the module name ("Title") supplied 

Function modParamCache(nModuleID, sModuleName)
	Dim sStat, rsMod, rsParam

	' determine the module id (if nec)
	If nModuleID = 0 And sModuleName <> "" Then
		sStat = "SELECT ModuleID FROM tblModule WHERE Title = '" & Replace(sModuleName, "'", "''") & "'"
		Set rsMod = adoOpenRecordset(sStat)
		If Not rsMod.EOF Then
			nModuleID = rsMod.Fields("ModuleID").Value
		Else
			rsMod.Close
			modParamCache = False
			modError = "Unable to locate module with the name """ & sModuleName & """"
			Exit Function
		End If
		rsMod.Close
		Set rsMod = Nothing
	End If

	' determine the module name (if nec)
	If sModuleName = "" And nModuleID > 0 Then
		sStat = "SELECT Title FROM tblModule WHERE ModuleID = " & nModuleID
		Set rsMod = adoOpenRecordset(sStat)
		If Not rsMod.EOF Then
			sModuleName = rsMod.Fields("Title").Value
		Else
			rsMod.Close
			modParamCache = False
			modError = "Unable to locate module with the ID = """ & nModuleID & """"
			Exit Function
		End If
		rsMod.Close
		Set rsMod = Nothing
	End If

	' retrieve all of the parameters for the module
	sStat = "SELECT	mp.ParamName, mp.ParamValue, mt.ASPConvertFunction " &_
			"FROM	tblModuleParam mp " &_
			"INNER JOIN	tblModuleParamType mt ON mp.TypeID = mt.TypeID " &_
			"WHERE	mp.ModuleID = " & nModuleID & " " &_
			"AND	mp.Archive = 0"	
	Set rsParam = adoOpenRecordset(sStat)
	Application.Lock
	Application("MODNUM" & UCase(sModuleName)) = nModuleID
	Do Until rsParam.EOF
		' determine if we need to do a datatype conversion
		If Trim(rsParam.Fields("ASPConvertFunction").Value & "") <> "" Then
			Application("MODPARAM" & nModuleID & "_" & rsParam.Fields("ParamName").Value) = Eval(rsParam.Fields("ASPConvertFunction").Value & "(" & rsParam.Fields("ParamValue").Value & ")")
		Else
			Application("MODPARAM" & nModuleID & "_" & rsParam.Fields("ParamName").Value) = rsParam.Fields("ParamValue").Value
		End If
		rsParam.MoveNext
	Loop
	Application.UnLock
	modParamCache = True
End Function

'--------------------------------------------------------------------
' modParam
'	Retrieve a configuration parameter for the module identified
'	by the module name ("Title") supplied and the parameter name

Function modParam(sModuleName, sParamName)
	' make sure the module has been cached
	If CStr(Application("MODNUM" & UCase(sModuleName))) = "" Then
		' module params not cached - do it
		If Not modParamCache(0, sModuleName) Then
			modParam = ""
			Exit Function
		End If
	End If
	modParam = Application("MODPARAM" & Application("MODNUM" & UCase(sModuleName)) & "_" & sParamName)
End Function

'--------------------------------------------------------------------
' modModuleID
'	Determine the module ID for the admin area the user is in based
'	strictly upon the "SCRIPT_NAME" server variable

Function modModuleID
	Dim rsMod, nModCount, sScript, I

	' update the module ID cache
	If CStr(Application("MODULEIDCOUNT")) = "" Then
		' build the cache of module paths
		sStat = "SELECT	mc.FolderName As CategoryFolder, m.FolderName " &_
				"FROM	tblModule m " &_
				"INNER JOIN	tblModuleCategory mc ON m.CategoryID = mc.CategoryID " &_
				"WHERE	m.Archive = 0"
		Set rsMod = adoOpenRecordset(sStat)
		nModCount = 0
		Application.Lock
		Do Until rsMod.EOF
			nModCount = nModCount + 1
			Application("MODULEID" & CStr(nModCount)) = rsMod.Fields("CategoryFolder").Value & "/" & rsMod.Fields("FolderName").Value
			rsMod.MoveNext
		Loop
		Application("MODULEIDCOUNT") = nModCount
		Application.UnLock
	Else
		nModCount = CInt(Application("MODULEIDCOUNT"))
	End If

	' determine the module we are currently in
	sScript = Mid(Request.ServerVariables("SCRIPT_NAME"), Len(Application("ASPNukeBasePath") & "module/") + 1)
	For I = 1 To nModCount
		' Response.Write "Checking: *" & Application("MODULEID" & CStr(I)) & "*<BR>"
		If Left(sScript, Len(Application("MODULEID" & CStr(I)))) = Application("MODULEID" & CStr(I)) Then
			modModuleID = I
			Exit Function
		End If
	Next
	' Response.Write "Script = *" & sScript & "*<BR>"
	If Right(Request.ServerVariables("SCRIPT_NAME"), 14) = "/configure.asp" And Request("ModuleID") <> "" Then
		modModuleID = CInt(Request("ModuleID"))
	Else
		modModuleID = 0
	End If
End Function

'--------------------------------------------------------------------
' modUpdateCategoryCounts
'	Convert a datbase bit / tinyint value to a boolean value

Function modUpdateCategoryCounts
	Dim sStat, rsCount

	' update the counts for the categories
	sStat = "SELECT	Count(ModuleID) AS ModuleCount, CategoryID " &_
			"FROM	tblModule " &_
			"WHERE	Archive = 0 " &_
			"GROUP BY CategoryID"
	Set rsCount = adoOpenRecordset(sStat)
	Do Until rsCount.EOF
		sStat = "UPDATE	tblModuleCategory SET " &_
				"ModuleCount = " & rsCount.Fields("ModuleCount").Value & " " &_
				"WHERE	CategoryID = " & rsCount.Fields("CategoryID").Value
		Call adoExecute(sStat)
		rsCount.MoveNext
	Loop
	rsCount.Close : Set rsCount = Nothing
End Function

'--------------------------------------------------------------------
' modRecordBoolValue
'	Convert a datbase bit / tinyint value to a boolean value

Function modRecordBoolValue(rs, sField)
	Dim sValue
	sValue = CStr(rs.Fields(sField).Value)
	Select Case sValue
		Case "True", "1" : modRecordBoolValue = True
		Case Else : modRecordBoolValue = False
	End Select
End Function
%>