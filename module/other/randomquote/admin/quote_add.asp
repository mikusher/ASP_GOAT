<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' quote_add.asp
'	Add a new random quote in the database
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
Dim nUserID


If steForm("action") = "add" Then
	' make sure the required fields are present
	If Trim(steForm("Quote")) = ""	Then
		sErrorMsg = steGetText("Please enter the Quote for this random quote")
	ElseIf Trim(steForm("Author")) = "" Then
		sErrorMsg = steGetText("Please enter the Attributed To for this quote")
	Else
		' create the author in the database
		sStat = "INSERT INTO tblQuote (" &_
				"	Quote, Author, Created, Modified" &_
				") VALUES (" &_
				steQForm("Quote") & "," & steQForm("Author") & "," &_
				adoGetDate & "," & adoGetDate &_
				")"
		Call adoExecute(sStat)
	End If
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Quote" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add Random Quote" %></H3>

<P>
<% steTxt "Please enter the new random quote using the form below." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="quote_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR>
	<TD CLASS="forml" nowrap><% steTxt "Quote" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="Quote" VALUE="<%= steEncForm("Quote") %>" SIZE="48" MAXLENGTH="255" class="form" style="{width:400px}"></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Attributed To" %></TD><TD></TD>
	<TD><INPUT TYPE="text" NAME="Author" VALUE="<%= steEncForm("Author") %>" SIZE="32" MAXLENGTH="50" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 align="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Quote" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Random Quote Added" %></H3>

<P>
<% steTxt "The new random quote was successfully created in the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<P ALIGN="center">
	<A HREF="quote_add.asp" class="adminlink"><% steTxt "Add Another" %></A>
</P>

<% End If %>

<p align="center">
	<a href="quotes.asp" class="adminlink"><% steTxt "Quote List" %></a>
</p>

<!-- #include file="../../../../footer.asp" -->
