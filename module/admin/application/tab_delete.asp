<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' tab_delete.asp
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

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If steNForm("Confirm") = 0	Then
		sErrorMsg = steGetText("Please confirm the deletion of the application variable tab")
	Else
		' create the new application variable tab in the database
		sStat = "DELETE FROM tblApplicationVarTab WHERE TabID = " & nTabID
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

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Application Variable Tab" %></H3>

<P>
<% steTxt "Please confirm you would like to permanently delete the application variable tab shown below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="tab_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="TabID" VALUE="<%= nTabID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Tab Name" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsVar, "TabName") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Title" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsVar, "Title") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Introduction" %></TD><TD></TD>
	<TD CLASS="formd"><%= Replace(steRecordEncValue(rsVar, "Introduction"), vbCrLf, "<Br>") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Summary" %></TD><TD></TD>
	<TD CLASS="formd"><%= Replace(steRecordEncValue(rsVar, "Summary"), vbCrLf, "<br>") %></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Confirm Delete?" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" name="confirm" value="1" class="form"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" name="confirm" value="0" checked class="form"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Tab" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Application Variable Tab Deleted" %></H3>

<P>
<% steTxt "The application variable tab was successfully deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="tab_list.asp" class="adminlink"><% steTxt "Tab List" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
