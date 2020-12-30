<!-- #include file="../../../lib/ado_lib.asp" -->
<!-- #include file="../../../lib/lang_lib.asp" -->
<!-- #include file="cache.asp" -->
<%
'--------------------------------------------------------------------
' capsule.asp
'	Create the polling capsule which will appear on all pages of
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
Dim nPollID

Call modCapsuleCache(False)
%>

<%= Replace(Application("ModCapHeader"), "<CAPSULETITLE/>", steGetText("Poll")) %>
<%= Application("ModCapLeft") %>

<% If Application("POLLCACHE") <> "" Then %>

<FORM METHOD="post" ACTION="<%= Application("ASPNukeBasePath") %>module/survey/poll/detail.asp">
<INPUT TYPE="hidden" NAME="PollID" VALUE="<%= Application("POLLID") %>">
<INPUT TYPE="hidden" NAME="action" VALUE="vote">

<%= Application("POLLCACHE") %>

<P ALIGN="center">
	<INPUT TYPE="submit" name="_vote" value="<% steTxt "VOTE" %>" class="form">
</P>
<p align="center">
	<A HREF="<%= Application("ASPNukeBasePath") %>module/survey/poll/detail.asp?PollID=<%= Application("POLLID") %>" class="actionlink"><% steTxt "View Results" %></A> .
	<A HREF="<%= Application("ASPNukeBasePath") %>module/survey/poll/archive.asp" class="actionlink"><% steTxt "Archive" %></A>
</p>
</FORM>

<% Else %>

<P><B CLASS="Error"><% steTxt "No poll has been defined yet" %></B></P>

<% End If %>
<%= Application("ModCapRight") %>
<%= Application("ModCapFooter") %>
