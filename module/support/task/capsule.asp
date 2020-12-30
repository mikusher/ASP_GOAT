<!-- #include file="../../../lib/ado_lib.asp" -->
<!-- #include file="../../../lib/lang_lib.asp" -->
<%
'--------------------------------------------------------------------
' capsule.asp
'	Capsule for the task manager
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

<%= Replace(Application("ModCapHeader"), "<CAPSULETITLE/>", steGetText("Task List")) %>
<%= Application("ModCapLeft") %>

<font class="tinytext">
<p align="center">
<% steTxt "ASP Nuke Tasks" %>:
</p><br>

<div class="forumcapsule" align="center">

<% Call locCacheTasks
	If Application("TASKCAPSULE") <> "" Then %>

<%= Application("TASKCAPSULE") %>

<% Else %>

<P><B CLASS="Error"><% steTxt "No tasks are defined yet" %></B></P>

<% End If %>
</div>
</font>

<%= Application("ModCapRight") %>
<%= Application("ModCapFooter") %>

<%
'----------------------------------------------------------------------------
' locCacheTasks
'	Cache the task capsule content to eliminate excessive database hits

Sub locCacheTasks
	Dim sStat, sHTML, rsTask, bFirst

	' check to see if we need to refresh (every 15 mins)
	If IsDate(Application("TASKCAPSULEREFRESH")) Then
		If DateDiff("n", Application("TASKCAPSULEREFRESH"), Now()) < 15 Then Exit Sub
	End If

	' retrieve the list of tasks from the database
	sStat = "SELECT	" & adoTop(5) & " t.TaskID, t.Title, t.PctComplete, p.ColorCode, t.Modified " &_
			"FROM	tblTask t " &_
			"INNER JOIN	tblTaskPriority p on p.PriorityID = t.PriorityID " &_
			"WHERE	t.Active <> 0 " &_
			"AND	t.Archive = 0 " &_
			"ORDER BY p.OrderNo DESC, t.Modified DESC" & adoTop2(5)
	Set rsTask = adoOpenRecordset(sStat)
	bFirst = False
	Do Until rsTask.EOF
		If bFirst Then sHTML = sHTML & "<hr class=""forumcapsulesep"">" & vbCrLf
		sHTML = sHTML & "<a href=""" & Application("ASPNukeBasePath") & "module/support/task/detail.asp?taskid=" & rsTask.Fields("TaskID").Value & """ class=""forumtopic"">" &_
			rsTask.Fields("Title").Value & "</a> <font class=""tinytext"">(" &_
			FormatNumber(rsTask.Fields("PctComplete").Value, 2) & "%)</font>" & vbCrLf
			' "&nbsp;&nbsp;<font class=""tinytext"">(" & rsTask.Fields("Threads").Value & "/" & rsTask.Fields("Messages").Value & ")</font>" & vbCrLf
		rsTask.MoveNext
		bFirst = True
	Loop
	rsTask.Close
	rsTask= Empty
	Application("TASKCAPSULEREFRESH") = Now()
	Application("TASKCAPSULE") = sHTML
End Sub
%>