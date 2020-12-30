<%
'--------------------------------------------------------------------
' order_lib.asp
'	Library for ordering the items in a hierarchical list
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
' Move a hierarchical item down in the order (keeping it at the same
' level that it's at now.

Sub ordMoveDown(sObjectName, sTableName, sPKey, sParentField, nParentID, nOrderNo, nID, sErrorMsg)
	Dim rsNext, nNextID, nNextOrder

	' check for right below this one
	sStat = "SELECT " & sPKey & ", OrderNo " &_
			"FROM	" & sTableName & " " &_
			"WHERE	" & sParentField & " = " & nParentID & " " &_
			"AND	OrderNo >= " & nOrderNo & " " &_
			"AND	" & sPKey & " <> " & nID & " " &_
			"ORDER BY OrderNo"
	Set rsNext = adoOpenRecordset(sStat)
	If Not rsNext.EOF Then
		nNextID = rsNext.Fields(sPKey).Value
		nNextOrder = rsNext.Fields("OrderNo").Value
		If nNextOrder = steNForm("OrderNo") Then
			Call ordShiftAll(sTableName, sPKey, sParentField, nParentID, nOrderNo, nID)
			' better to exit now then make an error in changing the order
			sErrorMsg = "Ordering conflict detected - you may need to move the item down again"
			Exit Sub
		End If
	Else
		nNextID = 0
	End If
	If nNextID > 0 Then
		' increment orders above the new order no (to make room)
		sStat = "UPDATE	" & sTableName & " " &_
				"SET	OrderNo = OrderNo - 1, Modified = " & adoGetDate & " " &_
				"WHERE	OrderNo = " & nNextOrder
		Call adoExecute(sStat)

		sStat = "UPDATE	" & sTableName & " " &_
				"SET	OrderNo = " & nNextOrder & ", Modified = " & adoGetDate & " " &_
				"WHERE	" & sPKey & " = " & nID
		Call adoExecute(sStat)
	Else
		sErrorMsg = sErrorMsg & "Cannot move the " & sObjectName & " down in the order - already at the bottom<br>"
	End If
End Sub

'--------------------------------------------------------------------
' Move a hierarchical item up in the order (keeping it at the same
' level that it's at now.

Sub ordMoveUp(sObjectName, sTableName, sPKey, sParentField, nParentID, nOrderNo, nID, sErrorMsg)
	Dim rsPrev, nPrevID, nPrevOrder

	' check for right above this one
	sStat = "SELECT " & sPKey & ", OrderNo " &_
			"FROM	" & sTableName & " " &_
			"WHERE	" & sParentField & " = " & nParentID & " " &_
			"AND	OrderNo <= " & nOrderNo & " " &_
			"AND	" & sPKey & " <> " & nID & " " &_
			"ORDER BY OrderNo DESC"
	Set rsPrev = adoOpenRecordset(sStat)
	If Not rsPrev.EOF Then
		nPrevID = rsPrev.Fields(sPKey).Value
		nPrevOrder = rsPrev.Fields("OrderNo").Value
		If nPrevOrder = steNForm("OrderNo") Then
			Call ordShiftAll(sTableName, sPKey, sParentField, nParentID, nOrderNo, nID)
			' better to exit now then make an error in changing the order
			sErrorMsg = "Ordering conflict detected - you may need to move the item up again"
			Exit Sub
		End If
	Else
		nPrevID = 0
	End If
	If nPrevID > 0 Then
		' increment orders above the new order no (to make room)
		sStat = "UPDATE	" & sTableName & " " &_
				"SET	OrderNo = OrderNo + 1, Modified = " & adoGetDate & " " &_
				"WHERE	OrderNo = " & nPrevOrder
		Call adoExecute(sStat)

		sStat = "UPDATE	" & sTableName & " " &_
				"SET	OrderNo = " & nPrevOrder & ", Modified = " & adoGetDate & " " &_
				"WHERE	" & sPKey & " = " & nID
		Call adoExecute(sStat)
	Else
		sErrorMsg = "Cannot move the " & sObjectName & " up in the order - already at the top"
	End If
End Sub

'--------------------------------------------------------------------
' Resolve an ordering conflict by increasing the order no for all
' records at this level

Sub ordShiftAll(sTableName, sPKey, sParentField, nParentID, nOrderNo, nID)
	sStat = "UPDATE " & sTableName & " " &_
			"SET	OrderNo = OrderNo + 1 " &_
			"WHERE	" & sParentField & " = " & nParentID & " " &_
			"AND	OrderNo >= " & nOrderNo & " " &_
			"AND	" & sPKey & " <> " & nID
	Call adoExecute(sStat)
End Sub
%>