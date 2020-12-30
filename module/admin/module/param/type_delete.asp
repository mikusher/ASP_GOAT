<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' type_delete.asp
'	Delete a module parameter type from the database
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
		sErrorMsg = steGetText("Please confirm that you wish to delete this module parameter type")
	Else
		' create the new module type in the database
		sStat = "DELETE FROM tblModuleParamType WHERE TypeID = " & steNForm("TypeID")
		Call adoExecute(sStat)
	End If
End If

' retrieve the recordset to delete
sStat = "SELECT	* FROM tblModuleParamType WHERE TypeID = " & steNForm("TypeID")
Set rsType = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Types" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Module Param Type" %></H3>

<P>
<% steTxt "Please verify that you would like to delete the module parameter type shown below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="type_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<input type="hidden" name="TypeID" value="<%= steNForm("TypeID") %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Type Code" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD class="formd"><%= steRecordEncValue(rsType, "TypeCode") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Type Name" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsType, "TypeName") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "ASP Convert Function" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsType, "ASPConvertFunction") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "HTML Input Type" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsType, "HTMLInputType") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "RegEx Validation" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsType, "RegExValidate") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Label Position" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsType, "LabelPos") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has Options?" %></TD><TD></TD>
	<TD class="formd"><% If steRecordBoolValue(rsType, "HasOptions") Then Response.Write steGetText("Yes") Else Response.Write steGetText("No") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Minimum Value" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsType, "MinValue") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Maximum Value" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsType, "MaxValue") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Is Numeric?" %></TD><TD></TD>
	<TD class="formd"><% If steRecordBoolValue(rsType, "IsNumeric") Then Response.Write steGetText("Yes") Else Response.Write steGetText("No") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "DB Quote Character" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsType, "QuoteChar") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top">Confirm Delete?</TD><TD></TD>
	<TD class="formd">
		<INPUT TYPE="radio" NAME="Confirm" VALUE="1" class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="Confirm" VALUE="0" CHECKED class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Type" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Module Parameter Deleted" %></H3>

<P>
<% steTxt "The module parameter type has been deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="type_list.asp" class="adminlink"><% steTxt "Type List" %></A> &nbsp;
	<A HREF="param_list.asp" class="adminlink"><% steTxt "Param List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->
