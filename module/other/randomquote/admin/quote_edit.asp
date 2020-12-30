<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' quote_edit.asp
'	Edit an existing random quote from the database
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

If steForm("action") = "edit" Then
	' make sure the required fields are present
	If Trim(steForm("Quote")) = ""	Then
		sErrorMsg = steGetText("Please enter the Quote for this random quote")
	ElseIf Trim(steForm("Author")) = "" Then
		sErrorMsg = steGetText("Please enter the Attributed To for this quote")
	Else
		' update the random quote in the database
			sStat = "UPDATE tblQuote " &_
					"SET	Quote = " & steQForm("Quote") & "," &_
					"		Author = " & steQForm("Author") & "," &_
					"		Modified = " & adoGetDate & " " &_
					"WHERE	QuoteID = " & steQForm("QuoteID")
			Call adoExecute(sStat)
	End If
End If

' retrieve the random quote to edit
sStat = "SELECT	* " &_
		"FROM	tblQuote " &_
		"WHERE QuoteID = " & nQuoteID
Set rsQuote = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Quote" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "edit" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Edit Random Quote" %></H3>

<P>
<% steTxt "Please make your changes to the random quote using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="quote_edit.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="edit">
<INPUT TYPE="hidden" NAME="QuoteID" VALUE="<%= nQuoteID %>">
<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" nowrap><% steTxt "Quote" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Quote" VALUE="<%= steRecordEncValue(rsQuote, "Quote") %>" SIZE="42" MAXLENGTH="255" class="form" style="{width:400px}"></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Attributed To" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Author" VALUE="<%= steRecordEncValue(rsQuote, "Author") %>" SIZE="32" MAXLENGTH="50" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 align="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt("Update Quote") %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Random Quote Updated" %></H3>

<P>
<% steTxt "The random quote was successfully updated in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<p align="center">
	<a href="quotes.asp" class="adminlink"><% steTxt "Quote List" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->
