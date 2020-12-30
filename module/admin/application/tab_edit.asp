<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' tab_edit.asp
'	Update existing application variable tab in the database
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
Dim nTabID
Dim I

nTabID = steNForm("TabID")

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("TabName")) = ""	Then
		sErrorMsg = steGetText("Please enter the Name for this variable tab")
	ElseIf Trim(steForm("Title")) = ""	Then
		sErrorMsg = steGetText("Please enter the Title for this variable tab")
	Else
		' create the new application variable tab in the database
		sStat = "UPDATE tblApplicationVarTab SET " &_
				"	TabName = " & steQForm("TabName") & "," &_
				"	Title = " & steQForm("Title") & "," &_
				"	Introduction = " & steQForm("Introduction") & "," &_
				"	Summary = " & steQForm("Summary") & " " &_
				"WHERE	TabID = " & nTabID
		Call adoExecute(sStat)
	End If
End If


sStat = "SELECT * FROM tblApplicationVarTab WHERE TabID = " & nTabID
Set rsVar = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Tabs" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Application Variable Tab" %></H3>

<P>
<% steTxt "Please make your changes to the application variable tab using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="tab_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="TabID" VALUE="<%= nTabID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Tab Name" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="TabName" VALUE="<%= steRecordEncValue(rsVar, "TabName") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Title" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Title" VALUE="<%= steRecordEncValue(rsVar, "Title") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Introduction" %></TD><TD></TD>
	<TD><TEXTAREA NAME="Introduction" COLS="52" ROWS="8" WRAP="virtual" class="form"><%= steRecordEncValue(rsVar, "Introduction") %></TEXTAREA></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Summary" %></TD><TD></TD>
	<TD><TEXTAREA NAME="Summary" COLS="52" ROWS="8" WRAP="virtual" class="form"><%= steRecordEncValue(rsVar, "Summary") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Update Tab" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Application Variable Tab Updated" %></H3>

<P>
<% steTxt "The application variable tab was successfully updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="tab_list.asp" class="adminlink"><% steTxt "Tab List" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
