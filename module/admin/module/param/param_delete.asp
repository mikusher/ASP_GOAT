<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' param_delete.asp
'	Delete a module param to the database
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
		sErrorMsg = steGetText("Please confirm that you would like to delete this module parameter")
	Else
		' create the new module param in the database
		sStat = "DELETE FROM tblModuleParam WHERE ParamID = " & steNForm("ParamID")
		Call adoExecute(sStat)
	End If
End If

' retrieve the record to delete
sStat = "SELECT mp.*, mpt.TypeName " &_
		"FROM	tblModuleParam mp " &_
		"INNER JOIN	tblModuleParamType mpt ON mpt.TypeID = mp.TypeID " &_
		"WHERE	ParamID = " & steNForm("ParamID")
Set rsParam = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Param" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Module Param" %></H3>

<P>
<% steTxt "Please make your changes to the module param using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="param_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<input type="hidden" name="ModuleID" value="<%= steNForm("ModuleID") %>">
<input type="hidden" name="ParamID" value="<%= steNForm("ParamID") %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Parameter Name" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsParam, "ParamName") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Parameter Value" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsParam, "ParamValue") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Data Type" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsParam, "TypeName") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Label" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsParam, "Label") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Minimum Value" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsParam, "MinValue") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Maximum Value" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsParam, "MaxValue") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Is Required?" %></TD><TD></TD>
	<TD class="formd">
		<% If steRecordBoolValue(rsParam, "IsRequired") Then Response.Write steGetText("Yes") Else Response.Write steGetText("No") %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Help Text" %></TD><TD></TD>
	<TD class="formd"><%= Replace(steRecordEncValue(rsParam, "HelpText"), vbCrLf, "<BR>") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Confirm Delete?" %></TD><TD></TD>
	<TD class="formd">
		<INPUT TYPE="radio" NAME="Confirm" VALUE="1" class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="Confirm" VALUE="0" CHECKED class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="center"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Param" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Module Param Deleted" %></H3>

<P>
<% steTxt "The module param has been updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="param_list.asp?moduleid=<%= steForm("ModuleID") %>" class="adminlink"><% steTxt "Param List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->
