<!-- #include file="../../../lib/ado_lib.asp" -->
<%
'--------------------------------------------------------------------
' capsule.asp
'	Create the dynamic drop-down menu capsule.
'	Depends on style-sheet definitions in the style.css file.
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

' rebuild the cached menu init string (as needed)
Call locMenuInitString
%>
<DIV CLASS="menubar">
<script language="Javascript" type="text/javascript">
<!-- // hide
<% If Trim(Application("MENUDYNAMIC")) <> "" Then %>
// Call the method that will build all of the menus
menuBuildAll('<%= Replace(Application("MENUDYNAMIC"), "'", "\'") %>');
<% End If %>
// unhide -->
</script>
</DIV>
<%
'----------------------------------------------------------------------------
' build the string used to initialize the dynamic menus

Sub locMenuInitString
	Dim sStat, rs, oItem, aItemID, sName, aName, sInit, sURL

	If IsDate(Application("MENUDYNAMICREFRESH")) Then
		If DateDiff("n", Application("MENUDYNAMICREFRESH"), Now()) < 30 Then Exit Sub
	End if

	sStat = "SELECT	ItemID, ParentItemID, MenuName, URL " &_
			"FROM	tblMenuItem " &_
			"WHERE	Active = 1 " &_
			"AND	Archive = 0 " &_
			"ORDER BY OrderNo"
	Set rs = adoOpenRecordset(sStat)
	Set oItem = Server.CreateObject("Scripting.Dictionary")
	Do Until rs.EOF
		' add the child list
		oItem("CHILD" & rs.Fields("ParentItemID").Value) = oItem("CHILD" & rs.Fields("ParentItemID").Value) &_
			"," & rs.Fields("ItemID").Value
		' build the URL for the menu item
		If Trim(rs.Fields("URL").Value & "") = "" Then
			sURL = "/content.asp?ID=" & Server.URLEncode(rs.Fields("MenuName").Value)
		Else
			sURL = rs.Fields("URL").Value
		End If
		If rs.Fields("ParentItemID").Value <> 0 Then
			oItem("MENU" & rs.Fields("ParentItemID").Value) = oItem("MENU" & rs.Fields("ParentItemID").Value) &_
				"^" & rs.Fields("MenuName").Value & "|" & sURL
		Else
			sName = sName & "^" & rs.Fields("MenuName").Value & "|" & sURL
		End If
		rs.MoveNext
	Loop
	rs.Close
	rs= Empty
	' Response.Write Mid(oItem("MENU4"), 2) & "<BR>"
	' Response.End
	If oItem.Exists("CHILD0") Then
		aItemID = Split(Mid(oItem("CHILD0"), 2), ",")
		aName = Split(Mid(sName, 2), "^")
		For I = 0 To UBound(aItemID)
			If I > 0 Then sInit = sInit & "~"
			sInit = sInit & aName(I) & "^" & Mid(oItem("MENU" & aItemID(I)), 2)
		Next
	End If

	Application("MENUDYNAMICREFRESH") = Now()
	Application("MENUDYNAMIC") = sInit
End Sub
%>