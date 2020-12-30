<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' state_delete.asp
'	Delete a state
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
Dim sStateCode
Dim rsState

sStateCode = steForm("StateCode")

If LCase(steForm("action")) = "delete" Then
	If Trim(steNForm("Confirm")) = "0" Then
		sErrorMsg = steGetText("Please confirm you would like to delete the state")
	Else
		' check to see if the state code is unique
		sStat = "DELETE	FROM	tblState WHERE StateCode = " & steQForm("statecode")
		Call adoExecute(sStat)
	End If
End If

' retrieve the state to delete
sStat = "SELECT * FROM tblState WHERE StateCode = '" & Replace(sStateCode, "'", "''") & "'"
Set rsState = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp"-->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "State" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If LCase(steForm("action")) <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete State" %></H3>

<P>
<% steTxt "Please confirm that you would like to delete the state shown below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="state_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml"><% steTxt "State Code" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsState, "StateCode") %></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "State Name" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsState, "StateName") %></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "Confirm Delete?" %></TD><TD></TD>
	<TD CLASS="formd">
		<INPUT TYPE="radio" NAME="Confirm" VALUE="1" CLASS="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="Confirm" VALUE="0" CHECKED CLASS="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br>
		<INPUT TYPE="submit" NAME="_sub" ACTION=" <% steTxt "Delete State" %> " class="form">
	</TD>
</TR>
</TABLE>

</FORM>

<% Else %>

<H3><% steTxt "State Deleted" %></H3>

<P>
<% steTxt "The state was successfully deleted from the database." %>&nbsp;
<% steTxt "You may use the admin links at the top of the page to continue administering the site." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="state_list.asp" class="adminlink"><% steTxt "State List" %></A>
</P>

<!-- #include file="../../../footer.asp"-->