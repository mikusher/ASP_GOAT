<!-- #include file="../../../lib/ado_lib.asp" -->
<!-- #include file="../../../lib/lang_lib.asp" -->
<%
'--------------------------------------------------------------------
' capsule.asp
'	Create the hit counter capsule which will appear on all pages of
'	the site.
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

%>

<%= Replace(Application("ModCapHeader"), "<CAPSULETITLE/>", steGetText("Hit Counter")) %>
<%= Application("ModCapLeft") %>

<font class="tinytext"><% steTxt "Total page views since" %><br>Sept 19, 2003<BR><BR></font>

	<% locCounter %>
	<table border=0 cellpadding=0 cellspacing=0 align="center">
	<tr>
		<TD CLASS="hitcounter"><%= String(6 - Len(Application("HitCount")), "0") %><%= Application("HitCount") %></TD>
	</tr>
	</table>

<%= Application("ModCapRight") %>
<%= Application("ModCapFooter") %>

<%
'----------------------------------------------------------------------------
' locCounter
'	Load or store the hit counter variable in the database

Sub locCounter
	Dim sStat, rs, nCount

	' cache the counter in the application object
	If IsDate(Application("HitUpdated")) Then
		If DateDiff("n", Application("HitUpdated"), Now()) < 15 Then
			Application("HitCount") = Application("HitCount") + 1
			Exit Sub
		End If
	End If

	' check for hit counter here
	If Application("HitCount") = "" Then
		sStat = "SELECT	HitCount FROM tblSiteStat"
		Set rs = adoOpenRecordset(sStat)
		If rs.EOF Then
			' create a new hit count here
			sStat = "INSERT INTO tblSiteStat (HitCount) VALUES (1)"
			Call adoExecute(sStat)
			nCount = 1
		Else
			nCount = rs.Fields("HitCount").Value + 1
		End If
	Else
		' store the existing hit counter
		sStat = "UPDATE	tblSiteStat " &_
				"SET	HitCount = " & CStr(Application("HitCount") + 1) & "," &_
				"		Modified = " & adoGetDate
		Call adoExecute(sStat)
		nCount = Application("HitCount") + 1
	End If

	' store the counter in the app (so we don't kill the DB)
	Application.Lock
	Application("HitCount") = nCount
	Application("HitUpdated") = Now()
	Application.UnLock
End Sub
%>