<%
'--------------------------------------------------------------------
' nav_lib.asp
'	Manages the top-level navigation that each admin user sees when
'	they login to the site admin area.
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

Dim navRight		' right to be checked
					' change this value before login_lib.asp is included
					' to check an access right other than the one determined
					' by the folder the user is currently in.

Const LOGIN_NAV_REFRESH = 15		' check for new nav content
Const LOGIN_NAV_COLUMNS = 6			' columns in the nav

If Request.ServerVariables("SCRIPT_NAME") <> Application("ASPNukeBasePath") & "admin/index.asp" Then
	' make sure access is allowed for this member
	Call navCheckAccess
End If

' rebuild the navigation items
Call navRefreshNav(False)

' build the navigation for the logged-in user
Call navMain

'--------------------------------------------------------------------
' create a new key for the navigation item.  This should be done
' everytime user rights (permissions) are changed to provide some
' measure of security to the admin areas.
' Store this value in tblUserRight.AccessKey
' To provide an extra measure of security, we can update all of
' the access keys on a regular interval (every 24 hours)

Function navNewKey
	Dim sCode, sChar, I

	Randomize
	For I = 1 To 20
		sChar = Int(36 * Rnd())
		If sChar < 10 Then sCode = sCode & Chr(sChar + 48) Else sCode = sCode & Chr(sChar + 55)
	Next
	navNewKey = sCode
End Function

'--------------------------------------------------------------------
' navUpdateAllKeys
'	Update all of the access keys

Function navUpdateAllKeys
	Dim sStat, rsNav, I

	sStat = "SELECT	RightID " &_
			"FROM	tblUserRight " &_
			"ORDER BY OrderNo"
	Set rsNav = adoOpenRecordset(sStat)
	If Not rsNav.EOF Then
		aNav = rsNav.GetRows
		rsNav.Close
		Set rsNav = Nothing
		sStat = ""
		For I = 0 To UBound(aNav, 2)
			sStat = sStat & "UPDATE tblUserRight " &_
					"SET	AccessKey = '" & navNewKey & "' " &_
					"WHERE	RightID = " & aNav(0, I) & "; "
		Next
		' execute all of the changes in one batch
		Call adoExecute(sStat)
		navUpdateAllKeys = True
	Else
		navError = "No user rights were found to update"
		navUpdateAllKeys = False
	End If
End Function

'--------------------------------------------------------------------
' check if access to this page is allowed (based on the folder name)
' TODO - calls Request.ServerVariables which is inefficient

Sub navCheckAccess
	Dim sFolder, nPos, bAllow

	If navRight <> "" Then
		bAllow = navCanAccess(navRight)
	Else
		nPos = InStrRev(Request.ServerVariables("SCRIPT_NAME"), "/")
		If nPos > 0 Then
			sFolder = Left(Request.ServerVariables("SCRIPT_NAME"), nPos)
			' strip off the protocol header (if exists)
			sFolder = Replace(Replace(sFolder, "https://", ""), "http://", "")
			' strip off domain name as well (if exists)
			sFolder = Mid(sFolder, InStr(1, sFolder, "/"))
		Else
			sFolder = "/"
		End If
		bAllow = InStr(1, Request.Cookies("AdminNav"), Application("ADMIN" & sFolder)) > 0
	End If
	If Not bAllow Then
		Response.Redirect Application("ASPNukeBasePath") & "admin/index.asp?error=" & Server.URLEncode("Access to this admin page is forbidden")
	End If
End Sub

'--------------------------------------------------------------------
' navCanAccess
'	Determine if current user has access to the right name given by
'	the parameter.

Function navCanAccess(sRightName)
	' build the right name
	Set oRegExp = New RegExp
	oRegExp.Pattern="[^\w]"
	oRegExp.IgnoreCase=True
	oRegExp.Global=True
	sRightName = oRegExp.Replace(sRightName, "_")

	navCanAccess = InStr(1, Request.Cookies("AdminNav"), Application("ADMINR" & sRightName)) > 0
End Function

'--------------------------------------------------------------------
' build the access nav for an individual user based on their assigned
' rights as defined by the cookie "AdminNav"

Function navMain
	Dim I, aCol(), nCount, aKey

	ReDim aCol(LOGIN_NAV_COLUMNS)
	aKey = Split(Request.Cookies("AdminNav"), "-")
	nCount = 0
	For I = 0 To UBound(aKey)
		aCol(nCount Mod LOGIN_NAV_COLUMNS) = aCol(nCount Mod LOGIN_NAV_COLUMNS) &_
			Application(aKey(I)) & "<BR>"
		nCount = nCount + 1
	Next
	aCol(nCount Mod LOGIN_NAV_COLUMNS) = aCol(nCount Mod LOGIN_NAV_COLUMNS) &_
		"<a href=""" & Application("ASPNukeBasePath") & "module/admin/logoff.asp"" class=""adminmenu"">Logoff</A><BR>"
	' rebuild the nav here
	With Response
		.Write "<TABLE BORDER=0 CELLPADDING=1 CELLSPACING=0 BGCOLOR=""#A08060"" WIDTH=""100%"">" & vbCrLf
		.Write "<TR><TD>" & vbCrLf
		.Write "	<TABLE BORDER=0 CELLPADDING=5 CELLSPACING=0 BGCOLOR=""#F0D090"" WIDTH=""100%"">" & vbCrLf
		.Write "	<TR>" & vbCrLf
		For I = 0 To LOGIN_NAV_COLUMNS - 1
			.Write "<TD VALIGN=""top"">" & aCol(I) & "</TD>"
		Next
		.Write "	</TR>" & vbCrLf
		.Write "	</TABLE>" & vbCrLf
		.Write "</TD></TR>" & vbCrLf
		.Write "</TABLE>" & vbCrLf
	End With
End Function

'--------------------------------------------------------------------
' build the list of navigation items for the main admin area
' from the database

Sub navRefreshNav(bForceRefresh)
	Dim sStat, rsNav, sHTML, nPos, I
	Dim sRight, sFolder, oRegExp

	If IsDate(Application("ADMIN_NAV_REFRESH")) And Not bForceRefresh Then
		If DateDiff("n", Application("ADMIN_NAV_REFRESH"), Now()) < LOGIN_NAV_REFRESH Then Exit Sub
	End If

	' build the navigation HTML
	Application.Unlock
	sStat = "SELECT	RightName, Hyperlink, AccessKey " &_
			"FROM	tblUserRight " &_
			"ORDER BY OrderNo"
	Set rsNav = adoOpenRecordset(sStat)

	Do Until rsNav.EOF
		nPos = InStrRev(rsNav.Fields("Hyperlink").Value, "/")
		If nPos > 0 Then
			sFolder = Left(rsNav.Fields("Hyperlink").Value, nPos)
		Else
			sFolder = "/"
		End If
		' build the right name
		Set oRegExp = New RegExp
		oRegExp.Pattern="[^\w]"
		oRegExp.IgnoreCase=True
		oRegExp.Global=True
		sRight = oRegExp.Replace(rsNav.Fields("RightName").Value, "_")

		' this is for building the main nav from "NavAdmin" cookie
		Application(rsNav.Fields("AccessKey").Value) = _
			"<A HREF=""" & Replace(Application("ASPNukeBasePath") & rsNav.Fields("Hyperlink").Value, "//", "/") & """ class=""adminmenu"">" & rsNav.Fields("RightName").Value & "</A>"
		' this is for checking access for a particular script in the admin
		' Application(sFolder) = sRight
		' Application(sRight) = rsNav.Fields("AccessKey").Value
		Application("ADMINR" & sRight) = rsNav.Fields("AccessKey").Value
		Application("ADMIN" & sFolder) = rsNav.Fields("AccessKey").Value
		rsNav.MoveNext
	Loop
	rsNav.Close
	Set rsNav = Nothing

	Application("ADMIN_NAV_REFRESH") = Now()
	Application.Lock
End Sub

'--------------------------------------------------------------------
' navLogin
'	Perform a login for an admin user - assign the access keys for
'	the assigned user rights to the user's cookie.

Function navLogin(nUserID)
	Dim sStat, rsNav, sKeys

	sStat = "SELECT	ur.RightName, ur.Hyperlink, ur.AccessKey " &_
			"FROM	tblUserRight ur " &_
			"INNER JOIN	tblUserToRight utr ON utr.RightID = ur.RightID " &_
			"WHERE	utr.UserID = " & nUserID & " " &_
			"ORDER BY ur.OrderNo"
	Set rsNav = adoOpenRecordset(sStat)
	sKeys = ""
	Do Until rsNav.EOF
		sKeys = sKeys & "-" & rsNav.Fields("AccessKey").Value
		rsNav.MoveNext
	Loop
	Response.Cookies("AdminNav") = Mid(sKeys, 2)
	navLogin = True
End Function
%>