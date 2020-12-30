<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' userright_delete.asp
'	Delete an existing user right to the database
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
Dim rsRight
Dim nRightID

nRightID = steNForm("rightid")

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If Trim(steNForm("Confirm")) <> 1	Then
		sErrorMsg = steGetText("Please confirm the deletion of this user right")
	Else
		' create the new user right in the database
		sStat = "DELETE FROM tblUserRight WHERE RightID = " & nRightID
		Call adoExecute(sStat)
	End If
End If

sStat = "SELECT ur1.*, ur2.RightName As ParentRight " &_
		"FROM tblUserRight ur1 " &_
		"LEFT JOIN	tblUserRight ur2 ON ur2.RightID = ur1.ParentRightID " &_
		"WHERE	ur1.RightID = " & nRightID
Set rsRight = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Rights" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete User Right" %></H3>

<P>
Please confirm that you would like to delete this user right by
clicking <I>Yes</I> next to <B>Confirm</B> below.  Once the right
has been deleted, it can not be recovered.
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="userright_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="rightid" VALUE="<%= nRightID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD class="forml"><% steTxt "Parent Right Name" %></TD><TD>&nbsp;&nbsp;&nbsp;</TD>
	<TD class="formd"><% If steRecordEncValue(rsRight, "ParentRight") = "" Then %>TOP-LEVEL RIGHT<% Else %><%= steRecordEncValue(rsRight, "ParentRight") %><% End If %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Right Name" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsRight, "RightName") %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Menu Name" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsRight, "AdminMenuName") %></TD>
</TR><TR>
	<TD class="forml" VALIGN="top"><% steTxt "Hyperlink" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsRight, "Hyperlink") %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has Add?" %></TD><TD></TD>
	<TD CLASS="formd"><% If CStr(steRecordEncValue(rsRight, "HasAdd")) = "1" Then Response.Write "Yes" Else Response.Write "No" %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has Edit?" %></TD><TD></TD>
	<TD CLASS="formd"><% If CStr(steRecordEncValue(rsRight, "HasEdit")) = "1" Then Response.Write "Yes" Else Response.Write "No" %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has Delete?" %></TD><TD></TD>
	<TD CLASS="formd"><% If CStr(steRecordEncValue(rsRight, "HasDelete")) = "1" Then Response.Write "Yes" Else Response.Write "No" %></TD>
</TR><TR>
	<TD CLASS="forml" VALIGN="top"><% steTxt "Has View?" %></TD><TD></TD>
	<TD CLASS="formd"><% If CStr(steRecordEncValue(rsRight, "HasView")) = "1" Then Response.Write "Yes" Else Response.Write "No" %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Confirm Delete" %></TD><TD></TD>
	<TD><INPUT TYPE="radio" NAME="confirm" VALUE="1" class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="confirm" VALUE="0" class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete User Right" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "User Right Deleted" %></H3>

<P>
The user right was permanently deleted from the database.
Please proceed administering
the site by using the menu shown at the top of the screen.
</P>

<p align="center">
	<a href="userright_list.asp" class="adminlink"><% steTxt "User Right List" %></a>
</p>

<% End If %>

<!-- #include file="../../../footer.asp" -->
