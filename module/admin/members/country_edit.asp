﻿<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' country_edit.asp
'	Update an existing country in the database (for member registrations)
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
Dim nCountryID
Dim rsCountry

nCountryID = steNForm("CountryID")

If LCase(steForm("action")) = "update" Then
	If Trim(steForm("CountryName")) = "" Then
		sErrorMsg = steGetText("Please enter the name for the Country")
	Else
		sStat = "UPDATE tblCountry " &_
				"SET	CountryName = " & steQForm("CountryName") & " " &_
				"WHERE CountryID = " & nCountryID
		Call adoExecute(sStat)
	End If
End If

sStat = "SELECT * FROM tblCountry WHERE CountryID = " & nCountryID
Set rsCountry = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp"-->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Country" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If LCase(steForm("action")) <> "update" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Country" %></H3>

<P>
<% steTxt "Please enter your changes for the country using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="country_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="update">
<input type="hidden" name="countryid" value="<%= nCountryID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml"><% steTxt "Country Name" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><input type="text" name="CountryName" value="<%= steRecordEncValue(rsCountry, "CountryName") %>" size="32" maxlength="50" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br>
		<INPUT TYPE="submit" NAME="_sub" ACTION=" <% steTxt "Update Country" %> " class="form">
	</TD>
</TR>
</TABLE>

</FORM>

<% Else %>

<H3><% steTxt "Country Updated" %></H3>

<P>
<% steTxt "The country was successfully updated in the database." %>&nbsp;
<% steTxt "You may use the admin links at the top of the page to continue administering the site." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="country_list.asp" class="adminlink"><% steTxt "Country List" %></A>
</P>

<!-- #include file="../../../footer.asp"-->