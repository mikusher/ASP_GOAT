<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/class/adminlist.asp" -->
<%
'--------------------------------------------------------------------
' text_list.asp
'	Display a list of all of the english text words/phrases that
'	require translation.
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

<% sCurrentTab = "Text" %>
<!-- #include file="pagetabs_inc.asp" -->

<h3><% steTxt "Language Text List" %></h3>

<%
Dim oList
Set oList = New clsAdminList
oList.Query = "SELECT	t.TextID, t.EnglishText, t.Modified " &_
			"FROM	tblLangText t " &_
			"WHERE	t.Archive = 0 " &_
			"ORDER BY t.EnglishText"
Call oList.AddColumn("EnglishText", steGetText("Word / Phrase"), "")
Call oList.AddColumn("Modified", steGetText("Modified"), "right")
oList.ActionLink = "<a href=""text_edit.asp?textid=##textid##"" class=""actionlink"">" & steGetText("edit") & "</a> . <a href=""text_delete.asp?textid=##textid##"" class=""actionlink"">" & steGetText("delete") & "</a>"

' show the list here
Call oList.Display
%>

<p align="center">
	<a href="text_add.asp" class="adminlink"><% steTxt "Add Language Text" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->
