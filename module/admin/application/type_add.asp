<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' type_add.asp
'	Add a new application variable type to the database
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
Dim rsOrder
Dim nOrderNo

If steForm("action") = "add" Then
	' make sure the required fields are present
	If Trim(steForm("TypeCode")) = "" Then
		sErrorMsg = steGetText("Please enter the Type Code for this variable type")
	ElseIf Trim(steForm("TypeName")) = ""	Then
		sErrorMsg = steGetText("Please enter the Type Name for this variable type")
	ElseIf Trim(steForm("LabelPos")) = ""	Then
		sErrorMsg = steGetText("Please enter the Label Position for this variable type")
	Else
		' retrieve the new order no
		sStat = "SELECT Coalesce(Max(OrderNo) + 1, 1) As OrderNo FROM tblApplicationVarType"
		Set rsOrder = adoOpenRecordset(sStat)
		If Not rsOrder.EOF Then nOrderNo = rsOrder.Fields("OrderNo").Value Else nOrderNo = 1

		' create the new variable type in the database
		sStat = "INSERT INTO tblApplicationVarType (" &_
				"	OrderNo, TypeCode, TypeName, ASPConvertFunction, HTMLInputType, LabelPos, RegExValidate, " &_
				"	HasOptions, MinValue, MaxValue, IsNumeric, QuoteChar, Created" &_
				") VALUES (" &_
				nOrderNo & "," &_
				steQForm("TypeCode") & "," &_
				steQForm("TypeName") & "," &_
				steQForm("ASPConvertFunction") & "," &_
				steQForm("HTMLInputType") & "," &_
				steQForm("LabelPos") & "," &_
				steQForm("RegExValidate") & "," &_
				steNForm("HasOptions") & "," &_
				steQForm("MinValue") & "," &_
				steQForm("MaxValue") & "," &_
				steNForm("IsNumeric") & "," &_
				steQForm("QuoteChar") & "," &_
				adoGetDate &_
				")"
		Call adoExecute(sStat)
	End If
End If
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Types" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New Variable Type" %></H3>

<P>
<% steTxt "Please enter the properties for the new variable type using the form below." %>&nbsp;
<% steTxt "You should only create new types if you understand the concept of adding layout types to the site templates." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="type_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Type Code" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="TypeCode" VALUE="<%= steEncForm("TypeCode") %>" SIZE="6" MAXLENGTH="4" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Type Name" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="TypeName" VALUE="<%= steEncForm("TypeName") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "ASP Convert Function" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="ASPConvertFunction" VALUE="<%= steEncForm("ASPConvertFunction") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "HTML Input Type" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="HTMLInputType" VALUE="<%= steEncForm("HTMLInputType") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "RegEx Validation" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="RegExValidate" VALUE="<%= steEncForm("RegExValidate") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Label Position" %></TD><TD></TD>
	<TD><SELECT NAME="LabelPos" class="form">
		<option value=""> --
		<option value="LEFT"<% If steEncForm("LabelPos") = "LEFT" Then Response.Write " SELECTED" %>> LEFT
		<option value="TOP"<% If steEncForm("LabelPos") = "TOP" Then Response.Write " SELECTED" %>> TOP
		</SELECT>
	</td>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has Options?" %></TD><TD></TD>
	<TD class="formd">
		<INPUT TYPE="radio" NAME="HasOptions" VALUE="1"<% If steNForm("HasOptions") <> 0 Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="HasOptions" VALUE="0"<% If steNForm("HasOptions") = 0 Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Minimum Value" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="MinValue" VALUE="<%= steEncForm("MinValue") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Maximum Value" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="MaxValue" VALUE="<%= steEncForm("MaxValue") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Is Numeric?" %></TD><TD></TD>
	<TD class="formd">
		<INPUT TYPE="radio" NAME="IsNumeric" VALUE="1"<% If steNForm("IsNumeric") <> 0 Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="IsNumeric" VALUE="0"<% If steNForm("IsNumeric") = 0 Then Response.Write " CHECKED" %> class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "DB Quote Character" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="QuoteChar" VALUE="<%= steEncForm("QuoteChar") %>" SIZE="4" MAXLENGTH="1" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Type" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "New Application Variable Type Added" %></H3>

<P>
<% steTxt "The new application variable type has been added to the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="variable_list.asp" class="adminlink"><% steTxt "Variable List" %></A> &nbsp;
	<A HREF="type_list.asp" class="adminlink"><% steTxt "Type List" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
