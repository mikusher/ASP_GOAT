<!-- #include file="../../lib/site_lib.asp" -->
<!-- #include file="../../lib/tab_lib.asp" -->
<!-- #include file="../../lib/config_lib.asp" -->
<%
'--------------------------------------------------------------------
' configure.asp
'	Configure module parameters.  Builds a dynamic form for making
'	configuration changes to the ASP Nuke modules.
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
Dim sAction
Dim sReferrer
Dim nModuleID
Dim sTitle
Dim sErrorMsg

sAction = steForm("action")
nModuleID = steNForm("ModuleID")
sReferrer = steForm("Referrer")

' setup the page we need to return to
If sReferrer = "" Then
	sReferrer = Request.ServerVariables("HTTP_REFERER")
End If

If sAction = "save" Then
	Call cfgSave(nModuleID, sErrorMsg)
End If

' retrieve the module we are working with
sStat = "SELECT Title " &_
		"FROM	tblModule " &_
		"WHERE	ModuleID = " & nModuleID
Set rsMod = adoOpenRecordset(sStat)
If Not rsMod.EOF Then sTitle = rsMod.Fields("Title").Value
rsMod.Close : Set rsMod = Nothing
%>
<!-- #include file="../../header.asp" -->
<!-- #include file="../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Configuration" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "save" Or sErrorMsg <> "" Then %>

<h3><%= Server.HTMLEncode(sTitle) %>&nbsp;<% steTxt "Configuration" %></h3>

<p>
<% steTxt "Shown below is the module configuration for" %> "<%= Server.HTMLEncode(sTitle) %>".
<% steTxt "These parameters control how the module will function." %>&nbsp;
<% steTxt "Please make your changes to the configuration and hit the <B>Save Configuration</B> button." %>
</p>

<% If sErrorMsg <> "" Then %>
<p><b class="error"><%= sErrorMsg %></b></p>
<% End If %>

<form method="post" action="configure.asp">
<input type="hidden" name="action" value="save">
<input type="hidden" name="referrer" value="<%= Server.HTMLEncode(sReferrer) %>">
<input type="hidden" name="moduleid" value="<%= nModuleID %>">

<table border=0 cellpadding=4 cellspacing=0>
<% cfgConfigForm nModuleID %>
<tr>
	<td colspan="3" align="right"><BR>
		<input type="submit" name="_save" value="<% steTxt "Save Configuration" %>" class="form">
	</td>
</tr>
</table>

</form>

<% Else %>

<h3><%= steForm("Module") %>&nbsp;<% steTxt "Configuration Saved" %></h3>

<p>
<% steTxt "The module configuration has been saved." %>&nbsp;
<% steTxt "The configuration changes made to the module should take effect immediately." %>&nbsp;
<% steTxt "Use the button below to return to the administration area." %>
</p>

<% End If %>

<P ALIGN="center">
	<A HREF="<%= sReferrer %>" class="adminlink"><%= sTitle %>&nbsp;<% steTxt "Admin" %></A>
</P>

<!-- #include file="../../footer.asp" -->