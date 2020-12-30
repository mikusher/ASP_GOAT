<!-- #include file="../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' login.asp
'	Perform a login of a user to the admin area of the site.
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
<!-- #include file="../../header.asp" -->
<!-- #include file="../../lib/admin/login_lib.asp" -->

<H3><%= Application("CompanyName") %>&nbsp;<% steTxt "Administration" %></h3>

<% If steForm("error") <> "" Then %>
<p><b class="error"><%= steEncForm("error") %></b></p>
<% End If %>

<% steTxt "Welcome" %>&nbsp;<%= Request.Cookies("AdminFullname") %>,

<p>
<% steTxt "Please select an option from the menu above to begin administering your web site." %>&nbsp;
<% steTxt "Nearly every aspect of the web site can be configured." %>&nbsp;
<% steTxt "If there is any application you are missing, you should check the module list for ASP Nuke." %>
</p>

<p>
<% steTxt "Help will soon be available by clicking on the help link at the top of the page." %>&nbsp;
<% steTxt "For now, please refer to the documentation that you received with this software." %>&nbsp;
<% steTxt "If you have any trouble, please don't hesitate to contact" %> <%= Application("CompanyName") %>&nbsp;
<% steTxt "at" %>
<a href="mailto:<%= Application("SupportEmail") %>"><%= Application("SupportEmail") %></a>.
</p>

<!-- #include file="../../footer.asp" -->