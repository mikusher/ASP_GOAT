<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' text_delete.asp
'	Delete an existing language text from the database
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

If steForm("action") = "delete" Then
	' make sure the required fields are present
	If steNForm("Confirm") = 0	Then
		sErrorMsg = steGetText("Please confirm the deletion of this language text")
	Else
		' create the author in the database
		sStat = "DELETE FROM tblLangText WHERE	TextID = " & nTextID
		Call adoExecute(sStat)
	End If
End If

' retrieve the record to delete from the database
Set rsText = adoOpenRecordset("select EnglishText from tblLangText where TextID = " & nTextID)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Text" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "delete" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Delete Language Text" %></H3>

<P>
<% steTxt "Please confirm that you would like to delete the language text shown below." %>&nbsp;
<% steTxt "Once text has been deleted, it will be lost forever." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="text_delete.asp">
<input type="hidden" name="textid" value="<%= nTextID %>">
<INPUT TYPE="hidden" NAME="action" VALUE="delete">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" nowrap><% steTxt "English Text" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD CLASS="formd"><%= steRecordEncValue(rsText, "EnglishText") %></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Confirm Delete?" %></TD><TD></TD>
	<TD CLASS="formd">
		<input type="radio" name="confirm" value="1" class="formradio"> <% steTxt "Yes" %>
		<input type="radio" name="confirm" value="0" class="formradio"> <% steTxt "No" %>
	</TD>
</TR><TR>
	<TD COLSPAN=3 align="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Delete Language Text" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Language Text Deleted" %></H3>

<P>
<% steTxt "The language text was successfully deleted from the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<!-- #include file="../../../../footer.asp" -->
