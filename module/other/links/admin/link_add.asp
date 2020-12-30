<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' link_add.asp
'	Allows the addition of current links for the site
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
Dim rsCat		' list of categories to choose from
Dim sCatList	' list of currently selected categories
Dim rsOrder		' generate the new order no
Dim nOrderNo	' order no for inserting records
Dim sErrorMsg	' error message to display to user

If UCase(steForm("action")) = "ADD" Then
	' check for required fields here
	If steNForm("categoryid") = 0 Then
		sErrorMsg = steGetText("Please select the caegory for this link")
	ElseIf Trim(steForm("URL")) = "" Then
		sErrorMsg = steGetText("Please enter the URL for this link")
	ElseIf Trim(steForm("Label")) = "" Then
		sErrorMsg = steGetText("Plese enter the Label for this link")
	Else
		' determine the new order no
		sStat = "SELECT COALESCE(Max(OrderNo) + 1, 1) As OrderNo " &_
				"FROM	tblLink " &_
				"WHERE	CategoryID = " & steNForm("CategoryID")
		Set rsOrder = adoOpenRecordset(sStat)
		If Not rsOrder.EOF Then
			nOrderNo = rsOrder.Fields("OrderNo").Value
		Else
			nOrderNo = 1
		End If

		' insert the new article into the database
		sStat = "INSERT INTO tblLink (" &_
				"	CategoryID, URL, Label, OrderNo, Created" &_
				") VALUES (" &_
				steNForm("CategoryID") & "," &_
				steQForm("URL") & "," & steQForm("Label") & "," & nOrderNo &_
				"," & adoGetDate &_
				")"
		Call adoExecute(sStat)
	End If
End If

' build the list of categories to choose from
sStat = "SELECT	CategoryID, CategoryName " &_
		"FROM	tblLinkCategory " &_
		"WHERE	Active <> 0 " &_
		"AND	Archive = 0 " &_
		"ORDER BY OrderNo"
Set rsCat = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Link" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If UCase(steForm("action")) <> "ADD" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New Link" %></H3>

<P>
<% steTxt "Please enter the information for the new link in the form below." %>
</P>

<FORM METHOD="post" ACTION="link_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Category" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><SELECT NAME="CategoryID" class="form">
	<OPTION VALUE="0"> -- <% steTxt "Choose" %> --
	<% Do Until rsCat.EOF %>
	<OPTION VALUE="<%= rsCat.Fields("CategoryID").Value %>"<% If CStr(rsCat.Fields("CategoryID").Value) = CStr(steForm("CategoryID")) Then Response.Write " SELECTED" %>> <%= rsCat.Fields("CategoryName").Value %>
	<%	rsCat.MoveNext
	   Loop %>
	</SELECT>
	</TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "URL" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="URL" VALUE="<%= steEncForm("URL") %>" SIZE="32" MAXLENGTH="100" class="form"></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Label" %></TD><TD></TD>
	<TD><INPUT TYPE="TEXT" NAME="Label" value="<%= steEncForm("Label") %>" size="32" maxlength="50" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Link" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Link Added" %></H3>

<P>
<% steTxt "The new link was successfully added." %>&nbsp;
<% steTxt "You may use the admin links above to continue administering the site." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="link_list.asp" class="adminlink"><% steTxt "Link List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->