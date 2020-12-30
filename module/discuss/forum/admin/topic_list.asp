<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/class/adminlist.asp" -->
<%
'--------------------------------------------------------------------
' topic_list.asp
'	Displays a list of the current forum topics for the site
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
Dim rsTopic
Dim I

sStat = "SELECT	TopicID, Title, Threads, LastPost, Modified " &_
		"FROM	tblMessageTopic " &_
		"ORDER BY OrderNo"
Set rsTopic = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Topic" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "Forum Topic List" %></H3>

<P>
<% steTxt "Shown below are all of the forum topics defined in the database." %>
</P>

<%
sStat = "SELECT	TopicID, Title, Threads, LastPost, Modified " &_
		"FROM	tblMessageTopic " &_
		"ORDER BY OrderNo"
Set oList = New clsAdminList
oList.Query = sStat
oList.AddColumn "<a href=""thread_list.asp?topicid=##TopicID##"">##Title##</A>", steGetText("Topic"), ""
oList.AddColumn "Threads", steGetText("Threads"), "center"
oList.AddColumn "LastPost", steGetText("Last Post"), ""
oList.AddColumn "Modified", steGetTExt("Modified"), ""
oList.ActionLink = "<A HREF=""topic_edit.asp?TopicID=##TopicID##"" class=""actionlink"">" & steGetText("edit") &_
	"</A> . <A HREF=""topic_delete.asp?TopicID=##TopicID##"" class=""actionlink"">" & steGetText("delete") & "</A>"
Call oList.Display
%>

<P ALIGN="center">
	<A HREF="topic_add.asp" class="adminlink"><% steTxt "Add New Forum Topic" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->