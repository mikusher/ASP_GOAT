<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' type_edit.asp
'	Edit a application variable type in the database
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

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("TypeCode")) = "" Then
		sErrorMsg = steGetText("Please enter the Type Code for this variable type")
	ElseIf Trim(steForm("TypeName")) = ""	Then
		sErrorMsg = steGetText("Please enter the Type Name for this variable type")
	ElseIf Trim(steForm("LabelPos")) = ""	Then
		sErrorMsg = steGetText("Please enter the Label Position for this variable type")
	Else
		' create the new variable type in the database
		sStat = "UPDATE tblApplicationVarType SET " &_
				"	TypeCode = " & steQForm("TypeCode") & "," &_
				"	TypeName = " & steQForm("TypeName") & "," &_
				"	ASPConvertFunction = " & steQForm("ASPConvertFunction") & "," &_
				"	HTMLInputType = " & steQForm("HTMLInputType") & "," &_
				"	LabelPos = " & steQForm("LabelPos") & "," &_
				"	RegExValidate = " & steQForm("RegExValidate") & "," &_
				"	HasOptions = " & steNForm("HasOptions") & "," &_
				"	MinValue = " & steQForm("MinValue") & "," &_
				"	MaxValue = " & steQForm("MaxValue") & "," &_
				"	IsNumeric = " & steNForm("IsNumeric") & "," &_
				"	QuoteChar = " & steQForm("QuoteChar") & "," &_
				"	Modified = " & adoGetDate & " " &_
				"WHERE	TypeID = " & steNForm("TypeID")

		Call adoExecute(sStat)
	End If
End If

' retrieve the recordset to edit
sStat = "SELECT	* FROM tblApplicationVarType WHERE TypeID = " & steNForm("TypeID")
Set rsType = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Types" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Application Variable Type" %></H3>

<P>
<% steTxt "Please enter the properties for the application variable type using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="type_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<input type="hidden" name="TypeID" value="<%= steNForm("TypeID") %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Type Code" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="TypeCode" VALUE="<%= steRecordEncValue(rsType, "TypeCode") %>" SIZE="6" MAXLENGTH="4" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Type Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="TypeName" VALUE="<%= steRecordEncValue(rsType, "TypeName") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "ASP Convert Function" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="ASPConvertFunction" VALUE="<%= steRecordEncValue(rsType, "ASPConvertFunction") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "HTML Input Type" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="HTMLInputType" VALUE="<%= steRecordEncValue(rsType, "HTMLInputType") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "RegEx Validation" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="RegExValidate" VALUE="<%= steRecordEncValue(rsType, "RegExValidate") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Label Position" %></TD><TD></TD>
	<TD><SELECT NAME="LabelPos" class="form">
		<option value=""> --
		<option value="LEFT"<% If steRecordEncValue(rsType, "LabelPos") = "LEFT" Then Response.Write " SELECTED" %>> LEFT
		<option value="TOP"<% If steRecordEncValue(rsType, "LabelPos") = "TOP" Then Response.Write " SELECTED" %>> TOP
		</SELECT>
	</td>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has Options?" %></TD><TD></TD>
	<TD class="formd">
		<INPUT TYPE="radio" NAME="HasOptions" VALUE="1"<% If steRecordBoolValue(rsType, "HasOptions") Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="HasOptions" VALUE="0"<% If Not steRecordBoolValue(rsType, "HasOptions") Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Minimum Value" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="MinValue" VALUE="<%= steRecordEncValue(rsType, "MinValue") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Maximum Value" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="MaxValue" VALUE="<%= steRecordEncValue(rsType, "MaxValue") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Is Numeric?" %></TD><TD></TD>
	<TD class="formd">
		<INPUT TYPE="radio" NAME="IsNumeric" VALUE="1"<% If steRecordBoolValue(rsType, "IsNumeric") Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="IsNumeric" VALUE="0"<% If steRecordBoolValue(rsType, "IsNumeric") Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "DB Quote Character" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="QuoteChar" VALUE="<%= steRecordEncValue(rsType, "QuoteChar") %>" SIZE="4" MAXLENGTH="1" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update Type" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Application Variable Updated" %></H3>

<P>
<% steTxt "The application variable type has been updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="variable_list.asp" class="adminlink"><% steTxt "Variable List" %></A> &nbsp;
	<A HREF="type_list.asp" class="adminlink"><% steTxt "Type List" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
