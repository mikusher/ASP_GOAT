<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/class/treelist.asp" -->
<%
'--------------------------------------------------------------------
' folder_list.asp
'	Displays a list of the document folders for the documentation
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
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Folder" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "Document Folder List" %></H3>

<P>
<% steTxt "Shown below are all of the administrators defined in the database." %>
</P>

<%
' retrieve the list of articles to display
sStat = "SELECT	fld.FolderID, fld.ParentFolderID, fld.FolderName, usr.Username, fld.DocumentCount, fld.Created, fld.Modified " &_
		"FROM	tblDocFolder fld " &_
		"LEFT JOIN tblUser usr ON usr.UserID = fld.CreatedByUserID " &_
		"ORDER BY fld.OrderNo"
Set oList = New clsTreeList
oList.query = sStat
oList.PrimaryKey = "FolderID"
oList.ParentField = "ParentFolderID"
Call oList.AddColumn("FolderName", "Folder Name", "")
Call oList.AddColumn("Username", "Username", "")
Call oList.AddColumn("DocumentCount", "Docs", "center")
Call oList.AddColumn("Modified", "Modified", "right")
oList.ActionLink = "<A HREF=""folder_edit.asp?folderid=##FolderID##"" class=""actionlink"">" & steGetText("edit") & "</A> . <A HREF=""folder_delete.asp?folderid=##FolderID##"" class=""actionlink"">" & steGetText("delete") & "</A>"
Call oList.Display
%>

<P ALIGN="center">
	<A HREF="folder_add.asp" class="adminlink"><% steTxt "Add New Document Folder" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->