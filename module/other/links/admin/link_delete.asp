<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' link_delete.asp
'	Delete an existing link from the database
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
Dim rsLink
Dim nLinkID

nLinkID = steNForm("LinkID")

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If steNForm("confirm") <> 1 Then
		sErrorMsg = steGetText("Please confirm the deletion of this article author")
	Else
		' create the new author in the database
		sStat = "DELETE FROM tblLink " &_
				"WHERE	LinkID = " & nLinkID
		Call adoExecute(sStat)
	End If
End If

' retrieve the link to be deleted
Set rsLink = adoOpenRecordset("SELECT * FROM tblLink WHERE LinkID = " & nLinkID)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Link" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Link" %></H3>

<P>
<% steTxt "Please confirm the deletion of the link by clicking the yes button next to <I>Confirm</I>." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="link_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="LinkID" VALUE="<%= nLinkID %>">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD class="forml"><% steTxt "URL" %></TD><TD></TD>
	<TD class="formd"><%= rsLink.Fields("URL").Value %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Label" %></TD><TD></TD>
	<TD class="formd"><%= rsLink.Fields("Label").Value %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Confirm Delete" %></TD><TD></TD>
	<TD><INPUT TYPE="radio" NAME="confirm" VALUE="1" class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="confirm" VALUE="0" class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 align="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Link" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Link Deleted" %></H3>

<P>
<% steTxt "The link was successfully deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align="center">
	<a href="link_list.asp" class="adminlink"><% steTxt "Link List" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->
