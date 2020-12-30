<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' quote_delete.asp
'	Delete an existing random quote from the database
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
Dim rsQuote
Dim nQuoteID
Dim nUserID

nQuoteID = steForm("QuoteID")

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If steNForm("confirm") <> 1 Then
		sErrorMsg = steGetText("Please confirm the deletion of this random quote")
	Else
		' update the random quote in the database
		' delete the existing random quote in the database
		sStat = "DELETE FROM tblQuote " &_
				"WHERE	QuoteID = " & nQuoteID
		Call adoExecute(sStat)
	End If
End If

' retrieve the random quote to delete
sStat = "SELECT	* " &_
		"FROM	tblQuote " &_
		"WHERE QuoteID = " & nQuoteID
Set rsQuote = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Quote" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Random Quote" %></H3>

<P>
<% steTxt "Please confirm the deletion of the ramdom quote by clicking the yes button next to <I>Confirm</I>." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="quote_delete.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">
<INPUT TYPE="hidden" NAME="QuoteID" VALUE="<%= nQuoteID %>">
<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" nowrap><% steTxt "Quote" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsQuote, "Quote") %></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Attributed To" %></TD><TD></TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsQuote, "Author") %></TD>
</TR><TR>
	<TD class="forml"><% steTxt "Confirm Delete" %></B></TD><TD></TD>
	<TD class="formd">
		<INPUT TYPE="radio" NAME="confirm" VALUE="1" class="formradio"> <% steTxt "Yes" %>
		<INPUT TYPE="radio" NAME="confirm" VALUE="0" checked class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 align="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt("Delete Quote") %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Random Quote Updated" %></H3>

<P>
<% steTxt "The random quote was successfully deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align="center">
	<a href="quotes.asp" class="adminlink"><% steTxt "Quote List" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->
