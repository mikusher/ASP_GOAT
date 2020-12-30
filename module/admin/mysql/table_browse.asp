<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' table_browse.asp
'	Browse the specified database table (in the MySQL database)
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

Dim sAction
Dim sTable
Dim nPageNo
Dim nLastRec
Dim nTotalRecs
Dim sPKeyList	' CSV list of primary keys
Dim sPKeyIndex	' CSV list of primary key indices
Dim sStat
Dim rsTab
Dim aTab

Const DEF_RECSPERPAGE = 25

sTable = steForm("table")
nPageNo = steNForm("pageno")
nTotalRecs = steNForm("totalrecs")

If nPageNo < 0 Then nPageNo = 0

If nTotalRecs = 0 Then
	' count all of the records in the table
	Set rsTab = adoOpenRecordset("select count(*) from " & sTable)
	If Not rsTab.EOF Then nTotalRecs = rsTab.Fields(0).Value
End If
nLastRec = DEF_RECSPERPAGE * (nPageNo + 1)
If nLastRec > nTotalRecs Then nLastRec = nTotalRecs

If nPageNo > 0 Then
	Set rsTab = adoOpenRecordset("select * from " & sTable & " limit " & (nPageNo * DEF_RECSPERPAGE) & ", " & DEF_RECSPERPAGE & ";")
Else
	Set rsTab = adoOpenRecordset("select * from " & sTable & " limit " & DEF_RECSPERPAGE & ";")
End If
If Not rsTab.EOF Then aTab = rsTab.GetRows
rsTab.Close
Set rsTab = Nothing
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Tables" %>
<!-- #include file="pagetabs_inc.asp" -->

<h4>Browsing Table <%= sTable %></h4>

<% If IsArray(aTab) Then %>

<p align="center">
<i>Displaying records <b><%= DEF_RECSPERPAGE * nPageNo + 1 %> - <%= nLastRec %></b> of <b><%= nTotalRecs %></b></i><br><br>
<% Call locPageNav(nTotalRecs, nPageNo) %>
</p>

<table border=0 cellpadding=2 cellspacing=0 class="list">
<tr>
<% ' build the list of column headers here
	aList = Split(locColumnList(sTable, sPKeyList, sPKeyIndex), ",")
	For I = 0 To UBound(aList)
		Response.Write "<td class=""listhead"">" & Server.HTMLEncode(aList(I)) & "</td>"
	Next %>
	<td class="listhead">Action</td>
</tr>
<%	' list the records here
	Dim aPKey, aPKI, sQuery
	aPKey = Split(sPKeyList, ",")
	aPKI = Split(sPKeyIndex, ",")
	For I = 0 To UBound(aTab, 2) %>
<tr class="list<%=I mod 2 %>">
<%	' display the individual field values
	' build the pkey field list
	sQuery = ""
	For J = 0 To UBound(aPKey)
		sQuery = sQuery & "&" & aPKey(J) & "=" & Server.URLEncode(aTab(aPKI(J), I)&"")
	Next
	For J = 0 To UBound(aTab, 1)
		If IsNull(aTab(J, I)) Then Response.Write "<td><i>null</i></td>" Else Response.Write "<td>" & Server.HTMLEncode(aTab(J, I)) & "</td>"
	Next %>
	<td><a href="record_delete.asp?table=<%= Server.URLEncode(sTable) %><%= sQuery %>" class="actionlink">delete</a> . <a href="record_edit.asp?table=<%= Server.URLEncode(sTable) %><%= sQuery %>" class="actionlink">edit</a></td>
</tr>
<% Next %>
</table>

<% Else %>

<p>
<b class="error">No records could be found in the table "<%= sTable %>"</b>
</p>

<% End If %>

<p align="center">
	<a href="table_list.asp" class="adminlink"><% steTxt "Table List" %></A>&nbsp;
	<a href="record_add.asp?table=<%= Server.URLEncode(sTable) %>" class="adminlink"><% steTxt "Add New Record" %></a>
</p>

<!-- #include file="../../../footer.asp" -->
<%
Function locColumnList(sTable, sPKeyList, sPKeyIndex)
	Dim rsCol, sList, I

	sPKeyList = ""
	sPKeyIndex = ""
	I = 0
	Set rsCol = adoOpenRecordset("describe " & sTable)
	Do Until rsCol.EOF
		If sList <> "" Then
			sList = sList & "," & rsCol.Fields("Field").Value
		Else
			sList = rsCol.Fields("Field").Value
		End If
		If rsCol.Fields("Key").Value = "PRI" Then
			sPKeyList = sPKeyList & "," & rsCol.Fields("Field").Value
			sPKeyIndex = sPKeyIndex & "," & I
		End If
		rsCol.MoveNext
		I = I + 1
	Loop
	rsCol.Close
	If sPKeyList <> "" And sPKeyIndex <> "" Then
		sPKeyList = Mid(sPKeyList, 2)
		sPKeyIndex = Mid(sPKeyIndex, 2)
	End If
	locColumnList = sList
End Function

Function locPageNav(nTotalRecs, nPageNo)
	Dim nPages, nMidStart, nMidEnd, I

	nPages = nTotalRecs \ DEF_RECSPERPAGE + 1
	With Response
	.Write "<i>Page:</i> "
	If nPages < 20 Then
		For I = 1 To nPages
			If I > 1 Then .Write " . "
			If nPageNo+1 <> I Then
				.Write "<a href=""table_browse.asp?table="
				.Write Server.URLEncode(sTable)
				.Write "&pageno="
				.Write (I - 1)
				.Write "&totalrecs="
				.Write nTotalRecs
				.Write """>"
				.Write I
				.Write "</A>"
			Else
				.Write "<B>"
				.Write I
				.Write "</B>"
			End If
		Next
	Else
		For I = 1 To 3
			If I > 1 Then .Write " . "
			If nPageNo+1 <> I Then
				.Write "<a href=""table_browse.asp?table="
				.Write Server.URLEncode(sTable)
				.Write "&pageno="
				.Write (I - 1)
				.Write "&totalrecs="
				.Write nTotalRecs
				.Write """>"
				.Write I
				.Write "</A>"
			Else
				.Write "<B>"
				.Write I
				.Write "</B>"
			End If
		Next
		If Not (nPageNo < 3 Or nPageNo > nPages - 3) Then
			nMidStart = nPageNo - 1
			If nMidStart < 4 Then nMidStart = 4
			nMidEnd = nPageNo + 3
			If nMidEnd > nPages - 3 Then nMidEnd = nPages - 3
			If nMidStart > 4 Then .Write " ... "
			For I = nMidStart To nMidEnd
				If I > 1 Then .Write " . "
				If nPageNo+1 <> I Then
					.Write "<a href=""table_browse.asp?table="
					.Write Server.URLEncode(sTable)
					.Write "&pageno="
					.Write (I - 1)
					.Write "&totalrecs="
					.Write nTotalRecs
					.Write """>"
					.Write I
					.Write "</A>"
				Else
					.Write "<B>"
					.Write I
					.Write "</B>"
				End If
			Next
			If nMidEnd < nPages - 3 Then .Write " ... "
		Else
			.Write " ... "
		End If
		For I = nPages - 2 To nPages
			If I > 1 Then .Write " . "
			If nPageNo+1 <> I Then
				.Write "<a href=""table_browse.asp?table="
				.Write Server.URLEncode(sTable)
				.Write "&pageno="
				.Write (I - 1)
				.Write "&totalrecs="
				.Write nTotalRecs
				.Write """>"
				.Write I
				.Write "</A>"
			Else
				.Write "<B>"
				.Write I
				.Write "</B>"
			End If
		Next
	End If
	End With
End Function
%>