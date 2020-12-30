<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' state_edit.asp
'	Update a state used for the member registration.
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
Dim sErrorMsg

sStateCode = steForm("StateCode")

If LCase(steForm("action")) = "update" Then
	If Trim(steForm("StateCode")) = "" Then
		sErrorMsg = steGetText("Please enter the State Code")
	ElseIf Trim(steForm("StateName")) = "" Then
		sErrorMsg = steGetText("Please enter the State Name")
	Else
		' check to see if the state code is unique
		If steForm("StateCode") <> steForm("OldStateCode") Then
			sStat = "SELECT * FROM tblState WHERE StateCode = " & steQForm("statecode")
			Set rsState = adoOpenRecordset(sStat)
			If Not rsState.EOF Then
				sErrorMsg = steGetText("The State Code you entered is already in use")
			End If
		End If
		If sErrorMsg = "" Then
			sStat = "UPDATE tblState " &_
					"SET	StateCode = " & steQForm("StateCode") & ", " &_
					"		StateName = " & steQForm("StateName") & " " &_
					"WHERE	StateCode = " & steQForm("OldStateCode")
			Call adoExecute(sStat)
		End If
	End If
End If

' retrieve the state to update
sStat = "SELECT * FROM tblState WHERE StateCode = '" & Replace(sStateCode, "'", "''") & "'"
Set rsState = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp"-->
<!-- #include file="../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "State" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If LCase(steForm("action")) <> "update" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Update State" %></H3>

<P>
<% steTxt "Please make your changes to the state using the form provided below." %>&nbsp;
<% steTxt "When you are done click the <i>Update State</i> button to finalize your changes." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="state_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="update">
<input type="hidden" name="OldStateCode" value="<%= rsState.Fields("StateCode").Value %>">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD CLASS="forml"><% steTxt "State Code" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><input type="text" name="StateCode" value="<%= steRecordEncValue(rsState, "StateCode") %>" size="4" maxlength="2" class="form"></TD>
</TR><TR>
	<TD CLASS="forml"><% steTxt "State Name" %></TD><TD></TD>
	<TD><input type="text" name="StateName" value="<%= steRecordEncValue(rsState, "StateName") %>" size="32" maxlength="32" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 ALIGN="right"><br>
		<INPUT TYPE="submit" NAME="_sub" ACTION=" <% steTxt "Update State" %> " class="form">
	</TD>
</TR>
</TABLE>

</FORM>

<% Else %>

<H3><% steTxt "State Updated" %></H3>

<P>
<% steTxt "The state was successfully updated in the database." %>&nbsp;
<% steTxt "You may use the admin links at the top of the page to continue administering the site." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="state_list.asp" class="adminlink"><% steTxt "State List" %></A>
</P>

<!-- #include file="../../../footer.asp"-->