﻿<!-- #include file="../../../lib/lang_lib.asp" -->
<%
'--------------------------------------------------------------------
' capsule.asp
'	Create the suggestions capsule which will appear on
'	all pages of the site.
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

<%= Replace(Application("ModCapHeader"), "<CAPSULETITLE/>", steGetText("Suggestion Box")) %>
<%= Application("ModCapLeft") %>

<P align="center">
<font class="tinytext">
<% steTxt "Have a suggestion on how we can improve this application?" %><br><br>

<% steTxt "Please" %> <a href="<%= Application("ASPNukeBasePath") %>module/support/suggest/suggestions.asp" class="tinytext"><% steTxt "send us your suggestions" %></a>,
<% steTxt "we would love to hear what you have to say!" %>
</font>
</P>

<%= Application("ModCapRight") %>
<%= Application("ModCapFooter") %>