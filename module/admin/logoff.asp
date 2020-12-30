<!-- #include file="../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' loggoff.asp
'	Forces a logoff of the admin user who is current logged into the
'	ASP Nuke application.
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

' unconditionally log-off the user here
Response.Cookies("AdminUserID") = ""
Response.Cookies("AdminUserID").Expires = Now()
Response.Cookies("AdminUsername") = ""
Response.Cookies("AdminUsername").Expires = Now()
Response.Cookies("AdminFullname") = ""
Response.Cookies("AdminFullname").Expires = Now()
%>
<!-- #include file="../../header.asp" -->

<h3><% steTxt "User Logoff" %></h3>

<p>
<% steTxt "You have successfully logged off the ASP Nuke control panel." %>&nbsp;
<% steTxt "For additional security, you may wish to close your browser window." %>&nbsp;
<% steTxt "If you wish to, you may" %> <a href="index.asp"><% steTxt "Log back in" %></a>.
</p>

<p>
<% steTxt "Thank you for using the ASP Nuke control panel!" %>
</p>

<!-- #include file="../../footer.asp" -->
