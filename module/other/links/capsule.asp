<!-- #include file="../../../lib/ado_lib.asp" -->
<!-- #include file="../../../lib/module_lib.asp" -->
<!-- #include file="../../../lib/lang_lib.asp" -->
<%
'--------------------------------------------------------------------
' capsule.asp
'	Create the links capsule which will appear on all pages of
'	the site.
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

Dim query
Dim rsLinkCat
Dim sPollStat
Dim nPollPollID
Dim rsLink			' links to be displayed
Dim aLink			' links to be displayed
Dim aCat			' array of categories to choose from
Dim sCategory		' links for a given category

nPollPollID = 0
query = "SELECT	CategoryID, CategoryName " &_
		"FROM	tblLinkCategory " &_
		"WHERE	tblLinkCategory.Active <> 0 " &_
		"AND	tblLinkCategory.Archive = 0 " &_
		"ORDER BY tblLinkCategory.OrderNo"
Set rsLinkCat = adoOpenRecordset(query)
If Not rsLinkCat.EOF Then
	aCat = rsLinkCat.GetRows
End If
rsLinkCat.Close
rsLinkCat = Empty

' retrieve the list of links to display
If IsArray(aCat) Then
	query = "SELECT CategoryID, URL, Label " &_
			"FROM	tblLink " &_
			"WHERE	Active <> 0 " &_
			"AND	Archive = 0 " &_
			"ORDER BY CategoryID, OrderNo"
	Set rsLink = adoOpenRecordset(query)
	If Not rsLink.EOF Then
		aLink = rsLink.GetRows
	End If
	rsLink.Close
	rsLink = Empty
End If
%>

<%= Replace(Application("ModCapHeader"), "<CAPSULETITLE/>", steGetText("Web Links")) %>
<%= Application("ModCapLeft") %>
<% If IsArray(aLink) Then %>

<% For I = 0 To UBound(aCat, 2)
	sCategory = ""
	For J = 0 To UBound(aLink, 2)
		If aLink(0, J) = aCat(0, I) Then
			sCategory = sCategory & modParam("Links", "BulletHTML") & "&nbsp;<a href=""" & aLink(1, J) & """ target=""_new"" class=""linklink"">" & aLink(2, J) & "</a><br>"
		End If
	Next
	If sCategory <> "" Then %>
	<DIV class="linkgroup">
		<%= aCat(1, I)  %><br>
		<DIV class="linklist">
			<%= sCategory %>
		</DIV>
	</DIV>
<%	End if
   Next %>

<% Else %>

<P><B CLASS="Error"><% steTxt "No links are defined yet" %></B></P>

<% End If %>
<%= Application("ModCapRight") %>
<%= Application("ModCapFooter") %>
