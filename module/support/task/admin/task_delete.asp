<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' task_delete.asp
'	Update an existing task from the task manager
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

Dim sStat
Dim sAction
Dim rsStatus
Dim rsPriority
Dim sErrorMsg

sAction = steForm("Action")

' validate the form first
If Trim(UCase(sAction)) = "DELETE" Then
	If steNForm("Confirm") = 0 Then
		sErrorMsg = sErrorMsg & steGetText("You must Confirm the deletion of the task") & "<br>"
	End If
	If sErrorMsg = "" Then
		' add the new task to the database
		sStat = "DELETE FROM tblTask WHERE	TaskID = " & steNForm("TaskID")
		Call adoExecute(sStat)
	End If
End If

' retrieve the task to delete
Set rsTask = adoOpenRecordset("select * from tblTask where TaskID = " & steNForm("TaskID"))
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Task" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("Action") <> "Edit" Or sErrorMsg <> "" Then %>

<h3><% steTxt "Delete Task" %></h3>

<p>
<% steTxt "Please confirm that you would like to delete the task shown below." %>&nbsp;
<% steTxt "Once a task has been deleted, it will be gone forever." %>
</p>

<% If sErrorMsg <> "" Then %>
<p><b class="error"><%= sErrorMsg %></b></p>
<% End If %>

<form method="post" action="task_delete.asp">
<input type="hidden" name="TaskID" value="<%= steNForm("TaskID") %>">
<input type="hidden" name="action" value="delete">

<table border=0 cellpadding=4 cellspacing=0>
<tr>
	<td class="forml" valign="top"><% steTxt "Title" %></td><td>&nbsp;&nbsp;</td>
	<td class="formd"><%= steRecordEncValue(rsTask, "Title") %></td>
</tr><tr>
	<td class="forml" valign="top"><% steTxt "Priority" %></td><td></td>
	<td class="formd">
<%	' build the priority list to choose from
	sStat = "SELECT	PriorityID, PriorityName " &_
			"FROM	tblTaskPriority " &_
			"WHERE	Active <> 0 " &_
			"AND	Archive = 0 " &_
			"ORDER BY OrderNo"
	Set rsPriority = adoOpenRecordset(sStat)
	Do Until rsPriority.EOF
		If CStr(rsPriority.Fields("PriorityID").Value) = steRecordEncValue(rsTask, "PriorityID") Then
			Response.Write rsPriority.Fields("PriorityName").Value
			Exit Do
		End If
		rsPriority.MoveNext
	Loop
	rsPriority.Close
	Set rsPriority = Nothing %>
</tr><tr>
	<td class="forml" valign="top"><% steTxt "Status" %></td><td></td>
	<td class="formd">
<%	' build the status list to choose from
	sStat = "SELECT	StatusID, StatusName " &_
			"FROM	tblTaskStatus " &_
			"WHERE	Active <> 0 " &_
			"AND	Archive = 0 " &_
			"ORDER BY OrderNo"
	Set rsStatus = adoOpenRecordset(sStat)
	Do Until rsStatus.EOF
		If CStr(rsStatus.Fields("StatusID").Value) = steRecordEncValue(rsTask, "StatusID") Then
			Response.Write rsStatus.Fields("StatusName").Value
			Exit Do
		End If
		rsStatus.MoveNext
	Loop
	rsStatus.Close
	Set rsStatus = Nothing %>
</tr><tr>
	<td class="forml"><% steTxt "Comments" %></td><td></td>
	<td class="formd"><%= Replace(steRecordEncValue(rsTask, "Comments"), vbCrLf, "<br>") %></td>
</tr><tr>
	<td class="forml"><% steTxt "Percent Complete" %></td><td></td>
	<td class="formd"><%= steRecordEncValue(rsTask, "PctComplete") %></td>
</tr><tr>
	<td class="forml"><% steTxt "Confirm Delete?" %></td><td></td>
	<td class="formd">
		<input type="radio" name="Confirm" value="1" class="formradio"> <% steTxt "Yes" %>
		<input type="radio" name="Confirm" value="0" class="formradio"> <% steTxt "No" %>
	</td>
</tr><tr>
	<td colspan="3" align="right"><br>
		<input type="submit" name="_submit" value="<% steTxt "Delete Task" %>" class="form">
	</td>
</tr>
</table>
</form>

<% Else %>

<h3><% steTxt "Task Deleted" %></h3>

<p>
<% steTxt "The task was deleted successfully from the database." %>&nbsp;
<% steTxt "Please use the tab navigation to continue administering your ASP Nuke site." %>
</p>

<% End If %>

<!-- #include file="../../../../footer.asp" -->
