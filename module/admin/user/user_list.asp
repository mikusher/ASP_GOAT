<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/class/adminlist.asp" -->
<%
'--------------------------------------------------------------------
' user_list.asp
'	Displays a list of the current admins for the site
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
Dim oList
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Users" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "Admin User List" %></H3>

<P>
<% steTxt "Shown below are all of the administrators defined in the database." %>
</P>

<%
' retrieve the list of articles to display
sStat = "SELECT	UserID, FirstName, MiddleName, LastName, Username, Created, Modified " &_
		"FROM	tblUser " &_
		"ORDER BY Lastname, FirstName"
Set oList = New clsAdminList
oList.query = sStat
Call oList.AddColumn("<font class=""tinytext"">##FirstName## ##MiddleName## ##LastName##</font>", "Title / Author", "")
Call oList.AddColumn("Username", "Username", "")
Call oList.AddColumn("Modified", "Modified", "right")
oList.ActionLink = "<A HREF=""user_edit.asp?userid=##UserID##"" class=""actionlink"">" & steGetText("edit") & "</A> . <A HREF=""userright_edit.asp?userid=##UserID##"" class=""actionlink"">" & steGetText("rights") & "</A> . <A HREF=""user_delete.asp?userid=##UserID##"" class=""actionlink"">" & steGetText("delete") & "</A>"
Call oList.Display
%>

<P ALIGN="center">
	<A HREF="user_add.asp" class="adminlink"><% steTxt "Add New Admin User" %></A>
</P>

<!-- #include file="../../../footer.asp" -->