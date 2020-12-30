<!-- #include file="../../../../lib/site_lib.asp" -->
<!-- #include file="../cache.asp" -->
<%
'--------------------------------------------------------------------
' poll_delete.asp
'	Delete an existing poll question from the database
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
Dim rsPoll
Dim nPollID

nPollID = steNForm("pollid")

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If steNForm("Confirm") <> 1	Then
		sErrorMsg = steGetText("Please confirm the deletion of this poll question")
	Else
		' delete the existing poll in the database
		sStat = "DELETE FROM tblPoll WHERE PollID = " & nPollID
		Call adoExecute(sStat)

		Call modCapsuleCache(True)
	End If
End If

sStat = "SELECT * FROM tblPoll WHERE PollID = " & nPollID
Set rsPoll = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Poll" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Poll Question" %></H3>

<P>
<% steTxt "Please confirm that you would like to delete this poll question by clicking <I>Yes</I> next to <B>Confirm</B> below." %>&nbsp;
<% steTxt "Once the question has been deleted, it may not be recovered." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="poll_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="pollid" VALUE="<%= nPollID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD class="forml"><% steTxt "Name" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD class="formd"><%= steRecordEncValue(rsPoll, "Question") %></TD>
</TR><TR>
	<TD class="forml">Confirm Delete</TD><TD></TD>
	<TD><INPUT TYPE="radio" NAME="confirm" VALUE="1" class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="confirm" VALUE="0" checked class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 align="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Question" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Poll Question Deleted" %></H3>

<P>
<% steTxt "The poll question was permanently deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<P ALIGN="center">
	<A HREF="poll_list.asp" class="adminlink"><% steTxt "Poll List" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->
