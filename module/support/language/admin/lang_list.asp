<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/class/adminlist.asp" -->
<%
'--------------------------------------------------------------------
' lang_list.asp
'	Display a list of all of the languages that are defined by the
'	the language manager.
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

<% sCurrentTab = "Language" %>
<!-- #include file="pagetabs_inc.asp" -->

<h3><% steTxt "Language List" %></h3>

<%
Dim oList
Set oList = New clsAdminList
oList.Query = "SELECT	l.LangCode, l.CountryName, l.NativeLanguage, l.FlagIcon, l.PctComplete, u.Username, " &_
			"		l.Modified " &_
			"FROM	tblLang l " &_
			"LEFT JOIN	tblUser u on l.UserID = u.UserID " &_
			"WHERE	l.Archive = 0 " &_
			"ORDER BY l.CountryName"
Call oList.AddColumn("<img src=""" & Application("Siteroot") & "##FlagIcon##"">", steGetText("Flag"), "")
Call oList.AddColumn("LangCode", steGetText("Code"), "")
Call oList.AddColumn("CountryName", steGetText("Country"), "center")
Call oList.AddColumn("PctComplete", steGetText("Complete"), "right")
Call oList.AddColumn("NativeLanguage", steGetText("Language Name"), "center")
Call oList.AddColumn("Username", steGetText("Maintainer"), "center")
Call oList.AddColumn("Modified", steGetText("Modified"), "right")
oList.ActionLink = "<a href=""lang_edit.asp?langcode=##langcode##"" class=""actionlink"">" & steGetText("edit") & "</a> . <a href=""lang_delete.asp?langcode=##langcode##"" class=""actionlink"">" & steGetText("delete") & "</a>"

' show the list here
Call oList.Display
%>

<p align="center">
	<a href="lang_add.asp" class="adminlink"><% steTxt "Add Language" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->
