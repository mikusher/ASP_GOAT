<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' text_add.asp
'	Add new language text to the database
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
Dim rsText
Dim nTextID
Dim nUserID

nTextID = steNForm("TextID")

If steForm("action") = "add" Then
	' make sure the required fields are present
	If Trim(steForm("EnglishText")) = ""	Then
		sErrorMsg = steGetText("Please enter the English Text for this language text")
	Else
		' create the author in the database
		sStat = "INSERT INTO tblLangText (" &_
				"	EnglishText, Created" &_
				") VALUES (" &_
				steQForm("EnglishText") & "," & adoGetDate &_
				")"
		Call adoExecute(sStat)
	End If
End If
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Text" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add Language Text"  %></H3>

<P>
<% steTxt "Please enter the new language text in the form below." %>&nbsp;
<% steTxt "This text will allow translations to be done based on the translation languages defined." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="text_add.asp">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" nowrap><% steTxt "English Text" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="EnglishText" VALUE="<%= steEncForm("EnglishText") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 align="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Language Text" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Language Text Added" %></H3>

<P>
<% steTxt "The new language text was successfully added to the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<!-- #include file="../../../../footer.asp" -->
