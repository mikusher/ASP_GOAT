<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' detail.asp
'	Display an individual task from the database.
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
Dim rsArt
Dim sStat

sStat = "SELECT	tsk.TaskID, tsk.Title, tsk.Comments, usr.FirstName, usr.LastName, " &_
		"		pri.PriorityName, sta.StatusName, 0 As CommentCount, " &_
		"		tsk.Created " &_
		"FROM	tblTask tsk " &_
		"INNER JOIN	tblUser usr ON tsk.UserID = usr.UserID " &_
		"INNER JOIN	tblTaskPriority pri ON pri.PriorityID = tsk.PriorityID " &_
		"INNER JOIN	tblTaskStatus sta ON sta.StatusID = tsk.StatusID " &_
		"WHERE	tsk.TaskID = " & steForm("taskid") & " " &_
		"AND	tsk.Active <> 0 " &_
		"AND	tsk.Archive = 0"
Set rsArt = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->

<% If Not rsArt.EOF Then %>

<P>
<FONT CLASS="articlehead"><%= rsArt.Fields("Title").Value %></FONT><BR>
<FONT CLASS="tinytext">by <%= rsArt.Fields("FirstName").Value & " " & rsArt.Fields("LastName").Value %> - <%= adoFormatDateTime(rsArt.Fields("Created").Value, vbLongDate) %></FONT><BR>
<FONT CLASS="tinytext"><%= rsArt.Fields("PriorityName").Value %> / <%= rsArt.Fields("StatusName").Value %></FONT>
</P>

<P>
<%= Replace(rsArt.Fields("Comments").Value, vbCrLf, "<BR>") %>
</P>
<div><A HREF="comments.asp?taskid=<%= steForm("taskid") %>" class="articlelink"><% steTxt "Comments" %> (<%= rsArt.Fields("CommentCount").Value %>)</A></div>
<% Else %>

<H3><% steTxt "Task No Longer Available" %></H3>

<P>
<% steTxt "Sorry, but the task that you requested is no longer available." %>&nbsp;
<% steTxt "Although we try to maintain an archive of all of our old tasks, sometimes it becomes necessary to remove a task from our site." %>&nbsp;
<% steTxt "Please update your bookmarks accordingly." %>
</P>

<% End If %>

<!-- #include file="../../../footer.asp" -->
