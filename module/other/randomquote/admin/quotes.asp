<% Option Explicit %>
<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../../../../lib/class/adminlist.asp" -->
<%
'--------------------------------------------------------------------
' quotes.asp
'	Administer the list of random quotes.
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
Dim rsList
Dim nArchive
Dim nSelected

nArchive = steNForm("Archive")
nSelected = steNForm("FolderID")
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Quote" %>
<!-- #include file="pagetabs_inc.asp" -->

<h3>Random Quotes</h3>

<p>
<% steTxt "The following quotes will be randomly displayed wherever the Random Quote module is placed on your site." %>

<%
	Dim oList
	Set oList = New clsAdminList
	oList.Query = "SELECT	QuoteID, Quote, Author, Modified " &_
				"FROM	tblQuote " &_
				"WHERE	Active <> 0 " &_
				"AND	Archive = 0 " &_
				"ORDER BY Quote"
	Call oList.AddColumn("Quote", steGetText("Quote"), "")
	Call oList.AddColumn("Author", steGetText("Attributed To"), "")
	Call oList.AddColumn("Modified", steGetText("Modified"), "right")
	oList.ActionLink = "<nobr><a href=""quote_edit.asp?quoteid=##quoteid##"" class=""actionlink"">" & steGetText("edit") & "</a> . <a href=""quote_delete.asp?quoteid=" &_
		"##quoteid##"" class=""actionlink"">" & steGetText("delete") & "</a></nobr>"
	
	' show the list here
	Call oList.Display
%>

<p align="center">
	<A href="quote_add.asp" class="adminlink"><% steTxt "Add New Quote" %></A>
</p>

<!-- #include file="../../../../footer.asp" -->
