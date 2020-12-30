<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' option_delete.asp
'	Delete an existing module parameter option to the database
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

Dim sErrorMsg
Dim sStat
Dim rsCat

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If steNForm("Confirm") = 0 Then
		sErrorMsg = steGetText("Please confirm that you would like to delete this option")
	Else
		' create the new module option in the database
		sStat = "DELETE FROM tblModuleParamOption WHERE OptionID = " & steNForm("OptionID")
		Call adoExecute(sStat)
	End If
End If

' retrieve the option to delete here
sStat = "SELECT * FROM tblModuleParamOption WHERE OptionID = " & steNForm("OptionID")
Set rsDelete = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Options" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Module Option" %></H3>

<P>
<% steTxt "Please confirm that you would like to delete the module option shown below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="option_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<input type="hidden" name="typeid" value="<%= steNForm("TypeID") %>">
<input type="hidden" name="paramid" value="<%= steNForm("ParamID") %>">
<input type="hidden" name="optionid" value="<%= steNForm("OptionID") %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Option Value" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD class="formd"><%= steRecordEncValue(rsDelete, "OptionValue") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Option Label" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsDelete, "OptionLabel") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Confirm Delete?" %></TD><TD></TD>
	<TD class="formd">
		<input type="radio" name="Confirm" value="1" class="formradio"> <% steTxt "Yes" %>
		<input type="radio" name="Confirm" value="0" checked class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Option" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Parameter Option Deleted" %></H3>

<P>
<% steTxt "The module parameter option was successfully deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
<% If steNForm("TypeID") > 0 Then %>
	<A HREF="type_list.asp" class="adminlink"><% steTxt "Type List" %></A> &nbsp;
<% ElseIf steNForm("ParamID") > 0 Then %>
	<A HREF="param_list.asp?paramid=<%= steNForm("paramid") %>" class="adminlink"><% steTxt "Param List" %></A> &nbsp;
<% End If %>
	<A HREF="option_list.asp?typeid=<%= steNForm("typeid") %>&paramid=<%= steNForm("paramid") %>" class="adminlink"><% steTxt "Option List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->
