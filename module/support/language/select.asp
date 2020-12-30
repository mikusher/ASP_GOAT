<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' select.asp
'	Select a language to use ASP Nuke in
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

Dim query
Dim rsLang
Dim sLangCode
Dim sErrorMsg

sLangCode = steForm("code")

If sLangCode = "US" Then
	Response.Cookies("LANGUAGE") = sLangCode
	If InStr(Request.ServerVariables("HTTP_REFERRER"), Application("SiteRoot")) > 0 Then
		Response.Redirect Request.ServerVariables("HTTP_REFERRER")
	Else
		Response.Redirect Application("SiteRoot")
	End If
End If

query = "SELECT	l.LangCode, l.CountryName, l.NativeLanguage, l.Published, l.PctComplete, l.Archive " &_
			"FROM	tblLang l " &_
			"LEFT JOIN	tblUser u on l.UserID = u.UserID " &_
			"WHERE	l.LangCode = " & steQForm("code")
Set rsLang = adoOpenRecordset(query)
If Not rsLang.EOF Then
	' check to see if it is published
	If Not (CStr(rsLang.Fields("Published").Value&"") = "True" Or CStr(rsLang.Fields("Published").Value&"") = "1") Then
		sErrorMsg = steGetText("Sorry, the language you chose is not published yet")
	ElseIf CStr(rsLang.Fields("Archive").Value&"") = "True" Or CStr(rsLang.Fields("Archive").Value&"") = "1" Then
		sErrorMsg = steGetText("Sorry, the language you chose has been archived and is no longer available.")
	Else
		Response.Cookies("LANGUAGE") = sLangCode
		If InStr(Request.ServerVariables("HTTP_REFERRER"), Application("SiteRoot")) > 0 Then
			Response.Redirect Request.ServerVariables("HTTP_REFERRER")
		Else
			Response.Redirect Application("SiteRoot")
		End If
	End If
Else
	sErrorMsg = steGetText("Invalid language code specified") & " (""" & sLangCode & """)"
End If
rsLang.Close
%>
<!-- #include file="../../../header.asp" -->

<h3><% steTxt "Select Language" %></h3>

<p>
<b class="error"><%= sErrorMsg %></b>
</p>

<p>
<% steTxt "Sorry, but the language code you selected is invalid." %>
<% steTxt "Please choose a language from the list below:" %>
</p>
<%
' retrieve the published languages
query = "SELECT	l.LangCode, l.CountryName, l.NativeLanguage, l.FlagIcon, l.PctComplete, u.Username, " &_
		"		l.Modified " &_
		"FROM	tblLang l " &_
		"LEFT JOIN	tblUser u on l.UserID = u.UserID " &_
		"WHERE	l.Published = 1 " &_
		"AND	l.Archive = 0 " &_
		"ORDER BY l.CountryName"
Set rsLang = adoOpenRecordset(query)
If Not rsLang.EOF Then %>
<div align=""center""><table border=0 cellpadding=2 cellspacing=0>
<%	Do Until rsLang.EOF
		Response.Write "<tr><td nowrap onclick=""location.href='" & Application("SiteRoot") & "/module/support/language/select.asp?code=" & rsLang.Fields("langcode").Value & "'""><img src=""" & Application("SiteRoot") & rsLang.Fields("FlagIcon").Value & """ align=""left"" hspace=""10"">" & rsLang.Fields("NativeLanguage").Value & "</td></tr>" & vbCrLf
		rsLang.MoveNext
	Loop
	rsLang.Close
	Set rsLang = Nothing %>
</table></div>
<% Else %>
	<div align="center"><b class="error">none complete</b></div>
<% End If %>

<!-- #include file="../../../footer.asp" -->