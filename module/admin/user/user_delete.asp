<!-- #include file="../../../lib/site_lib.asp" -->
<!-- #include file="../../../lib/admin/login_lib.asp" -->
<%
'--------------------------------------------------------------------
' user_delete.asp
'	Delete an existing admin user from the database
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
Dim sErrorMsg

If steForm("action") = "delete" Then
	If steForm("confirm") <> "yes" Then
		sErrorMsg = steGetText("You must type ""yes"" to delete this member")
	Else
		' delete the member from the database
		sStat = "DELETE FROM tblUser WHERE UserID = " & steForm("UserID")
		Call adoExecute(sStat)
	End If
End If

' retrieve the user information to display (if nec)
If steForm("action") <> "delete" Or sErrorMsg <> "" Then
	sStat = "SELECT * FROM tblUser WHERE UserID = " & steForm("UserID")
	Set rsUser = adoOpenRecordset(sStat)
End If
%>
<!-- #include file="../../../header.asp" -->

<% sCurrentTab = "Users" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Admin User" %></H3>

<P>
<% steTxt "The details for the member you chose to delete are shown below." %>&nbsp;
<% steTxt "The member will be deleted once you type ""yes"" in the <I>Confirm</I> field and submit this form." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="user_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="UserID" VALUE="<%= steForm("UserID") %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD class="forml"><% steTxt "Username" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD class="formd"><%= steRecordEncValue(rsUser, "Username") %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Password" %></TD><TD></TD>
	<TD class="formd"><I>* <% steTxt "hidden" %> *</I></TD>
</TR><TR>
	<TD class="forml"><% steTxt "First Name" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsUser, "FirstName") %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Middle Name" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsUser, "MiddleName") %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Last Name" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsUser, "LastName") %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Email Address" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsUser, "EmailAddress") %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Daytime Phone" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsUser, "DayPhone") %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Evening Phone" %></TD><TD></TD>
	<TD class="formd"><%= steRecordEncValue(rsUser, "EvePhone") %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Confirm Delete (type ""yes"")" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Confirm" VALUE="<%= steEncForm("confirm") %>" SIZE="16" MAXLENGTH="3" CLASS="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><BR><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete User" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Admin User Deleted" %></H3>

<P>
<% steTxt "The admin user has been permanently deleted from the database along with all associated records." %>
</P>

<% End If %>

<!-- #include file="../../../footer.asp" -->
