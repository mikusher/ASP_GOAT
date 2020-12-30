<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' task_list.asp
'	Display a list of all of the tasks that have been entered into
'	the task manager.
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
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->
<!-- #include file="../../../../lib/class/adminlist.asp" -->


<% sCurrentTab = "Task" %>
<!-- #include file="pagetabs_inc.asp" -->

<h3><% steTxt "Current Task List" %></h3>
<%
Dim oList
Set oList = New clsAdminList
oList.Query = "SELECT	t.TaskID, t.Title, u.Username, p.ColorCode, p.PriorityName, s.StatusName, " &_
			"		t.Modified " &_
			"FROM	tblTask t " &_
			"INNER JOIN	tblUser u on t.UserID = u.UserID " &_
			"INNER JOIN	tblTaskPriority p on p.PriorityID = t.PriorityID " &_
			"INNER JOIN	tblTaskStatus s on s.StatusID = t.StatusID " &_
			"WHERE	t.Active <> 0 " &_
			"AND	t.Archive = 0 " &_
			"ORDER BY p.OrderNo DESC, t.Modified DESC"
Call oList.AddColumn("Title", steGetText("Task"), "")
Call oList.AddColumn("Username", steGetText("Posted By"), "center")
Call oList.AddColumn("StatusName", steGetText("Status"), "center")
Call oList.AddColumn("PriorityName", steGetText("Priority"), "center")
Call oList.AddColumn("Modified", steGetText("Modified"), "right")
oList.ActionLink = "<a href=""task_comments.asp?taskid=##taskid##"" class=""actionlink"">" & steGetText("comments") & "</a> . <a href=""task_edit.asp?taskid=##taskid##"" class=""actionlink"">" & steGetText("edit") & "</a> . <a href=""task_delete.asp?taskid=##taskid##"" class=""actionlink"">" & steGetText("delete") & "</a>"
oList.HTMLColorColumn = "ColorCode"

' show the list here
Call oList.Display
%>

<p align="center">
	<a href="task_add.asp" class="adminlink"><% steTxt "Add New Task" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->
