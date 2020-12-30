<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' variable_delete.asp
'	Delete an existing application variable to the database
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
Dim rsVar
Dim rsTab
Dim aTab
Dim nVarID
Dim I

nVarID = steNForm("varid")

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If steNForm("Confirm") <> 1	Then
		sErrorMsg = steGetText("Please confirm the deletion of this variable")
	Else
		' create the new application variable in the database
		sStat = "DELETE FROM tblApplicationVar WHERE VarID = " & nVarID
		Call adoExecute(sStat)
	End If
End If

sStat = "SELECT * FROM tblApplicationVar WHERE VarID = " & nVarID
Set rsVar = adoOpenRecordset(sStat)

' retrieve the list of tabs to choose from
sStat = "SELECT TabID, TabName FROM tblApplicationVarTab WHERE Archive = 0 ORDER BY OrderNo"
Set rsTab = adoOpenRecordset(sStat)
If Not rsTab.EOF Then aTab = rsTab.GetRows
rsTab.Close
Set rsTab = Nothing

sStat = "SELECT TypeID, TypeName FROM tblApplicationVarType WHERE Archive = 0 ORDER BY OrderNo"
Set rsType = adoOpenRecordset(sStat)
If Not rsType.EOF Then aType = rsType.GetRows
rsType.Close
Set rsType = Nothing
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Configure" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Application Variable" %></H3>

<P>
<% steTxt "Please confirm that you would like to delete this application variable by clicking <I>Yes</I> next to <B>Confirm</B> below." %>&nbsp;
<% steTxt "Once the application variable has been deleted, it may not be recovered." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="variable_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="varid" VALUE="<%= nVarID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Tab Group" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD CLASS="formd">
		<% If IsArray(aTab) Then
			For I = 0 To UBound(aTab, 2)
				If steRecordEncValue(rsVar, "TabID") = CStr(aTab(0, I)) Then Response.Write Server.HTMLEncode(aTab(1, I))
			Next
		   End If %>
		</SELECT>
	</TD>
</TR><TR>
	<TD class="forml"><% steTxt "Name" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD class="formd"><%= steRecordEncValue(rsVar, "VarName") %></TD>
</TR><TR>
	<TD class="forml" VALIGN="top"><% steTxt "Value" %></TD><TD></TD>
	<TD><%= steRecordEncValue(rsVar, "VarValue") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Data Type" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD class="formd">
		<% If IsArray(aType) Then
			For I = 0 To UBound(aType, 2)
				If steRecordEncValue(rsVar, "TypeID") = CStr(aType(0, I)) Then Response.Write Server.HTMLEncode(aType(1, I))
			Next
		   End If %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Is Required?" %></TD><TD></TD>
	<TD CLASS="formd">
		<% If steRecordBoolValue(rsVar, "IsRequired") Then Response.Write "Yes" Else Response.Write "No" %>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top">Help Text</TD><TD></TD>
	<TD CLASS="formd"><%= Replace(steRecordEncValue(rsVar, "HelpText"), vbCrLf, "<br>") %></TD>
</TR><TR>
	<TD class="forml">Confirm Delete</TD><TD></TD>
	<TD><INPUT TYPE="radio" NAME="confirm" VALUE="1" class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="confirm" VALUE="0" CHECKED class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Variable" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Application Variable Deleted" %></H3>

<P>
<% steTxt "The application variable was permanently deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="variable_list.asp" class="adminlink"><% steTxt "Variable List" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
