<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' tran_add.asp
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
Dim sEnglishText
Dim sLangCode
Dim nTranslationID
Dim nUserID

sLangCode = steForm("LangCode")
nTranslationID = steNForm("TranslationID")

If steForm("action") = "add" Then
	' make sure the required fields are present
	If Trim(steForm("Translation")) = ""	Then
		sErrorMsg = steGetText("Please enter the Translation for this language text")
	Else
		' create the author in the database
		sStat = "INSERT INTO tblLangTranslation (" &_
				"	LangCode, TextID, Translation, Created" &_
				") VALUES (" &_
					"'" & sLangCode & "', " & steNForm("TextID") & ", " & steQForm("Translation") &_
					"," & adoGetDate &_
				")"
		Call adoExecute(sStat)
	End If
End If

' retrieve the language name
Set rsLang = adoOpenRecordset("select NativeLanguage, CountryName from tblLang WHERE LangCode = '" & sLangCode & "' AND Archive = 0")
If Not rsLang.EOF Then
	sNativeLang = rsLang.Fields("NativeLanguage").Value
	sCountryName = rsLang.Fields("CountryName").Value
End If
rsLang.Close
Set rsLang = Nothing

' retrieve the english text to translate
Set rsText = adoOpenRecordset("select EnglishText from tblLangText WHERE TextID = " & nTextID)
If Not rsText.EOF Then sEnglishText = rsText.Fields("EnglishText").Value
rsText.Close
Set rsText = Nothing
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Translation" %>
<!-- #include file="pagetabs_inc.asp" -->

<% If steForm("action") <> "add" Or sErrorMsg <> "" Then %>

<H3><% steTxt "Add Language Translation" %></H3>

<P>
<% steTxt "Please enter the new language translation in the form below." %>&nbsp;
<% steTxt "This translation will only be used for the specific language indicated." %>
</P>

<% If sErrorMsg <> "" Then %>
<P><B CLASS="error"><%= sErrorMsg %></B></P>
<% End If %>

<FORM METHOD="post" ACTION="tran_add.asp">
<input type="hidden" name="langcode" value="<%= sLangCode %>">
<input type="hidden" name="TextID" value="<%= steNForm("TextID") %>">
<INPUT TYPE="hidden" NAME="action" VALUE="add">

<TABLE BORDER=0 CELLPADDING=2 CELLSPACING=0>
<TR>
	<TD CLASS="forml" nowrap><% steTxt "Translation Language" %></TD><TD></TD>
	<TD CLASS="formd"><%= sNativeLanguage %>&nbsp;(<%= sCountryName %>)</TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "English Text" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><%= Server.HTMLEncode(sEnglishText) %></TD>
</TR><TR>
	<TD CLASS="forml" nowrap><% steTxt "Translation" %></TD><TD>&nbsp;&nbsp;</TD>
	<TD><INPUT TYPE="text" NAME="Translation" VALUE="<%= steEncForm("Translation") %>" SIZE="32" MAXLENGTH="255" class="form"></TD>
</TR><TR>
	<TD COLSPAN=3 align="right"><br><INPUT TYPE="submit" NAME="_submit" VALUE=" <% steTxt "Add Translation" %> " class="form"></TD>
</TR>
</TABLE>
</FORM>

<% Else %>

<H3><% steTxt "Language Translation Added" %></H3>

<P>
<% steTxt "The new language translation was successfully added to the database." %>&nbsp;
<% steTxt "Please proceed administering the site by using the menu shown at the top of the screen." %>
</P>

<% End If %>

<!-- #include file="../../../../footer.asp" -->
