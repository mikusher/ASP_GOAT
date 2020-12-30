<!-- #include file="../../../lib/lang_lib.asp" -->
<%
'--------------------------------------------------------------------
' capsule.asp
'	Create the donation capsule which will appear on all pages of
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

%>

<%= Replace(Application("ModCapHeader"), "<CAPSULETITLE/>", steGetText("Help ASP Nuke")) %>
<%= Application("ModCapLeft") %>

	<p align="center">
		<font class="tinytext"><% steTxt "Help support ""ASP Nuke"" and free software by donating now" %></font><br><br>

		<form action="https://www.paypal.com/cgi-bin/webscr" method="post">
		<input type="hidden" name="cmd" value="_xclick">
		<input type="hidden" name="business" value="info@orvado.com">
		<input type="hidden" name="item_name" value="<% steTxt "ASP Nuke Content Management System" %>">
		<input type="hidden" name="no_note" value="1">
		<input type="hidden" name="currency_code" value="USD">
		<input type="hidden" name="tax" value="0">
		<input type="image" src="https://www.paypal.com/en_US/i/btn/x-click-but04.gif" name="submit" alt="<% steTxt "Make payments with PayPal - it's fast, free and secure!" %>">
		</form>

<%= Application("ModCapRight") %>
<%= Application("ModCapFooter") %>