<!-- #include file="../../../lib/ado_lib.asp" -->
<!-- #include file="rss_lib.asp" -->
<%
'--------------------------------------------------------------------
' capsule.asp
'	Create the RSS feed capsule which will appear on all pages of
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

Dim nFeedID
Dim sTitle
Dim sHTML

nFeedID = Request.QueryString("feedID")
If IsNumeric(nFeedID) And CStr(nFeedID) <> "" Then nFeedID = CInt(nFeedID) Else nFeedID = 1
sHTML = rssCapsule(nFeedID, sTitle)
%>

<%= Replace(Application("ModCapHeader"), "<CAPSULETITLE/>", sTitle) %>
<%= Application("ModCapLeft") %>

<%= sHTML %>

<P ALIGN="center">
	<A HREF="../../../news/archive.asp" class="actionlink">(more...)</A>
</P>

<%= Application("ModCapRight") %>
<%= Application("ModCapFooter") %>