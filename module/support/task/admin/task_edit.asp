<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' task_edit.asp
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
If Trim(UCase(sAction)) = "EDIT" Then
	If Trim(steForm("Title")) = "" Then
		sErrorMsg = steGetText("Please enter the title for the task")
	End If
	If steNForm("StatusID") = 0 Then
		sErrorMsg = sErrorMsg & steGetText("Please enter the Status for the task") & "<br>"
	End If
	If steNForm("PriorityID") = 0 Then
		sErrorMsg = sErrorMsg & steGetText("Please enter the Priority for the task") & "<br>"
	End If
	If sErrorMsg = "" Then
		' add the new task to the database
		sStat = "UPDATE tblTask SET " &_
				"	UserID = " & Request.Cookies("AdminUserID") & ", " &_
				"	StatusID = " &  steNForm("StatusID") & ", " &_
				"	PriorityID = " & steNForm("PriorityID") & ", " &_
				"	Title = " & steQForm("Title") & ", " &_
				"	Comments = " & steQForm("Comments") & ", " &_
				"	PctComplete = " & steFForm("PctComplete") & " " &_
				"WHERE	TaskID = " & steNForm("TaskID")
		Call adoExecute(sStat)
	End If
End If

' retrieve the task to edit
Set rsTask = adoOpenRecordset("select * from tblTask where TaskID = " & steNForm("TaskID"))
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Task" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("Action") <> "edit" Or sErrorMsg <> "" Then %>

<h3><% steTxt "Update Task" %></h3>

<p>
<% steTxt "Please make your changes to the task using the form below." %>&nbsp;
<% steTxt "When you are finished, click the <i>Update Task</i> button to finalize your changes." %>
</p>

<% If sErrorMsg <> "" Then %>
<p><b class="error"><%= sErrorMsg %></b></p>
<% End If %>

<form method="post" action="task_edit.asp">
<input type="hidden" name="TaskID" value="<%= steNForm("TaskID") %>">
<input type="hidden" name="action" value="edit">

<table border=0 cellpadding=4 cellspacing=0>
<tr>
	<td class="forml"><% steTxt "Title" %></td><td>&nbsp;&nbsp;</td>
	<td class="formd"><input type="text" name="title" value="<%= steRecordEncValue(rsTask, "Title") %>" size="32" maxlength="50" class="form"></td>
</tr><tr>
	<td class="forml"><% steTxt "Priority" %></td><td></td>
	<td class="formd">
	<select name="PriorityID" class="form">
	<option value=""> -- <% steTxt "Choose" %> --
<%	' build the priority list to choose from
	sStat = "SELECT	PriorityID, PriorityName " &_
			"FROM	tblTaskPriority " &_
			"WHERE	Active <> 0 " &_
			"AND	Archive = 0 " &_
			"ORDER BY OrderNo"
	Set rsPriority = adoOpenRecordset(sStat)
	Do Until rsPriority.EOF %>
	<option value="<%= rsPriority.Fields("PriorityID").Value %>"<% If CStr(rsPriority.Fields("PriorityID").Value) = steRecordEncValue(rsTask, "PriorityID") Then Response.Write " SELECTED" %>> <%= rsPriority.Fields("PriorityName").Value %>
<%		rsPriority.MoveNext
	Loop
	rsPriority.Close
	Set rsPriority = Nothing %>
	</select>
</tr><tr>
	<td class="forml"><% steTxt "Status" %></td><td></td>
	<td class="formd">
	<select name="StatusID" class="form">
	<option value=""> -- <% steTxt "Choose" %> --
<%	' build the status list to choose from
	sStat = "SELECT	StatusID, StatusName " &_
			"FROM	tblTaskStatus " &_
			"WHERE	Active <> 0 " &_
			"AND	Archive = 0 " &_
			"ORDER BY OrderNo"
	Set rsStatus = adoOpenRecordset(sStat)
	Do Until rsStatus.EOF %>
	<option value="<%= rsStatus.Fields("StatusID").Value %>"<% If CStr(rsStatus.Fields("StatusID").Value) = steRecordEncValue(rsTask, "StatusID") Then Response.Write " SELECTED" %>> <%= rsStatus.Fields("StatusName").Value %>
<%		rsStatus.MoveNext
	Loop
	rsStatus.Close
	Set rsStatus = Nothing %>
	</select>
</tr><tr>
	<td class="forml"><% steTxt "Comments" %></td><td></td>
	<td><textarea name="comments" cols="42" rows="12" class="form" style="width:420px"><%= steRecordEncValue(rsTask, "Comments") %></textarea></td>
</tr><tr>
	<td class="forml"><% steTxt "Percent Complete" %></td><td></td>
	<td class="formd"><input type="text" name="PctComplete" value="<%= steRecordEncValue(rsTask, "PctComplete") %>" size="8" maxlength="8" class="form"></td>
</tr><tr>
	<td colspan="3" align="right"><br>
		<input type="submit" name="_submit" value="<% steTxt "Update Task" %>" class="form">
	</td>
</tr>
</table>
</form>

<% Else %>

<h3><% steTxt "Task Updated" %></h3>

<p>
<% steTxt "The task was updated successfully in the database." %>&nbsp;
<% steTxt "Please use the tab navigation to continue administering your ASP Nuke site." %>
</p>

<% End If %>

<P ALIGN="center">
	<A HREF="task_list.asp" class="adminlink"><% steTxt "Task List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->
