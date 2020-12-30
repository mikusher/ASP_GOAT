<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' state_add.asp
'	Add a new state for the member registration process.
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

If LCase(steForm("action")) = "add" Then
	If Len(Trim(steForm("statecode"))) <> 2 Then
		sErrorMsg = steGetText("Invalid State Code Entered")
	ElseIf Trim(steForm("StateName")) = "" Then
		sErrorMsg = steGetText("Invalid State Name Entered")
	Else
		' check to see if the state code is unique
		sStat = "SELECT	* FROM	tblState WHERE StateCode = " & steQForm("statecode")
		Set rsState = adoOpenRecordset(sStat)
		If Not rsState.EOF Then
			sErrorMsg = steGetText("The state code you entered is already in use, please choose another")
		Else
			sStat = "INSERT INTO tblState (" &_
					"		StateCode, StateName, Created" &_
					") VALUES (" &_
					steQForm("StateCode") & ", " & steQForm("StateName") & "," & adoGetDate &_
					")"
			Call adoExecute(sStat)
		End If
	End If
End If
%>
<!-- #include file="../../../header.asp"-->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "State" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If LCase(steForm("action")) <> "add" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add New State" %></H3>

<P>
<% steTxt "Please enter the information for the new state in the form below." %>&nbsp;
<% steTxt "When you are finished click the <I>Add State</I> button to add the new State." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="state_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml"><% steTxt "State Code" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="StateCode" VALUE="<%= steEncForm("StateCode") %>" SIZE="16" MAXLENGTH="2" class="form"></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "State Name" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="StateName" VALUE="<%= steEncForm("StateName") %>" SIZE="32" MAXLENGTH="50" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br>
		<INPUT TYPE="submit" NAME="_sub" ACTION=" <% steTxt "Add State" %> " class="form">
	</TD>
</TR>
</TABLE>

</FORM>

<% Else %>

<H3><% steTxt "New State Added" %></H3>

<P>
<% steTxt "The brand new state was successfully added to the database." %>&nbsp;
<% steTxt "You may use the admin links at the top of the page to continue administering the site." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="state_list.asp" class="adminlink"><% steTxt "State List" %></A>
</P>

<!-- #include file="../../../footer.asp"-->