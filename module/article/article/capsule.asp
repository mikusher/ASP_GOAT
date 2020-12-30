<!-- #include file="../../../lib/lang_lib.asp" -->
<%
'--------------------------------------------------------------------
' capsule.asp
'	Create the news article search capsule which will appear on
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

<%= Replace(Application("ModCapHeader"), "<CAPSULETITLE/>", steGetText("Article Search")) %>
<%= Application("ModCapLeft") %>
<form method="post" action="<%= Application("ASPNukeBasePath") %>module/article/article/search.asp">
<input type="hidden" name="action" value="search">

<table border=0 cellpadding=2 cellspacing=0>
<tr>
	<td><input type="text" name="Keywords" value="" size="16" maxlength="50" class="form" style="width:115px"></td>
</tr><tr>
	<td align="right"><input type="submit" name="_submit" value=" <% steTxt "Find Articles" %> " class="form"></td>
</tr>
</table>

</form>
<%= Application("ModCapRight") %>
<%= Application("ModCapFooter") %>