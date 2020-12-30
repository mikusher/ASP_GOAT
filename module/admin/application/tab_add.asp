<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/tab_lib.asp" -->
<%
'--------------------------------------------------------------------
' tab_add.asp
'	Add a new application variable tab to the database
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
Dim rsTab
Dim rsOrder
Dim nOrderNo

If steForm("action") = "add" Then
	' make sure the required fields are present
	If Trim(steForm("TabName")) = ""	Then
		sErrorMsg = steGetText("Please enter the Name for this variable")
	ElseIf Trim(steForm("Title")) = ""	Then
		sErrorMsg = steGetText("Please enter the Value for this variable")
	Else
		' determine the new order no
		sStat = "SELECT Coalesce(Max(OrderNo) + 1, 1) As OrderNo FROM tblApplicationVarTab"
		Set rsOrder = adoOpenRecordset(sStat)
		If Not rsOrder.EOF Then nOrderNo = rsOrder.Fields("OrderNo").Value
		rsOrder.Close
		Set rsOrder = Nothing

		' create the new application variable in the database
		sStat = "INSERT INTO tblApplicationVarTab (" &_
				"	OrderNo, TabName, Title, Introduction, Summary, Created " &_
				") VALUES (" &_
				nOrderNo & "," &_
				steQForm("TabName") & "," &_
				steQForm("Title") & "," &_
				steQForm("Introduction") & "," &_
				steQForm("Summary") & "," & adoGetDate &_
				")"
		Call adoExecute(sStat)
	End If
End If
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Tabs" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New Application Variable Tab" %></H3>

<P>
<% steTxt "Please enter the new properties for the new application variable tab using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="tab_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Name" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="TabName" VALUE="<%= steEncForm("TabName") %>" SIZE="32" MAXLENGTH="32" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Title" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Title" VALUE="<%= steEncForm("Title") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Introduction" %></TD><TD></TD>
	<TD><TEXTAREA NAME="Introduction" COLS="52" ROWS="8" WRAP="virtual" class="form"><%= steEncForm("Introduction") %></TEXTAREA></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Summary" %></TD><TD></TD>
	<TD><TEXTAREA NAME="Summary" COLS="52" ROWS="8" WRAP="virtual" class="form"><%= steEncForm("Summary") %></TEXTAREA></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Tab" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "New Application Variable Tab Added" %></H3>

<P>
<% steTxt "The new application variable tab has been added to the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="tab_list.asp" class="adminlink"><% steTxt "Tab List" %></A>
</P>

<!-- #include file="../../../footer.asp" -->
