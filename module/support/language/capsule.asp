<!-- #include file="../../../lib/ado_lib.asp" -->
<!-- #include file="../../../lib/lang_lib.asp" -->
<%
'--------------------------------------------------------------------
' capsule.asp
'	Create the translations capsule which will appear on
'	all pages of the site.
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
Const LANG_CAPSULE_INTERVAL = 15		' minutes between cache reloads

Call locLanguageCache
%>

<%= Replace(Application("ModCapHeader"), "<CAPSULETITLE/>", steGetText("Other Languages")) %>
<%= Application("ModCapLeft") %>

<%= Application("LANG_CAPSULE_HTML") %>

<font class="tinytext">
<P align="center">
<a href="<%= Application("ASPNukeBasePath") %>module/support/language/language_select.asp" class="tinytext">
<% steTxt "Join the language translation project" %></a>&nbsp;
<% steTxt "to translate ASP Nuke into other languages" %><br><br>
</P>
</font>

<%= Application("ModCapRight") %>
<%= Application("ModCapFooter") %>
<%
Sub locLanguageCache
	Dim query, rsLang, sHTML

	If IsDate(Application("LANG_CAPSULE_UPDATED")) Then
		If DateDiff("n", Application("LANG_CAPSULE_UPDATED"), Now()) < LANG_CAPSULE_INTERVAL Then Exit Sub
	End If

	' retrieve the published languages
	query = "SELECT	l.LangCode, l.CountryName, l.NativeLanguage, l.FlagIcon, l.PctComplete, u.Username, " &_
			"		l.Modified " &_
			"FROM	tblLang l " &_
			"LEFT JOIN	tblUser u on l.UserID = u.UserID " &_
			"WHERE	l.Published = 1 " &_
			"AND	l.Archive = 0 " &_
			"ORDER BY l.CountryName"
	Set rsLang = adoOpenRecordset(query)
	If Not rsLang.EOF Then
		Do Until rsLang.EOF
			sHTML = sHTML & "<tr><td nowrap onclick=""location.href='" & Application("SiteRoot") & "/module/support/language/select.asp?code=" & rsLang.Fields("langcode").Value & "'"">" &_
				"<img src=""" & Application("SiteRoot") & rsLang.Fields("FlagIcon").Value & """ align=""left"" hspace=""10"" alt="""">" &_
				"<a href=""" & Application("SiteRoot") & "/module/support/language/select.asp?code=" & rsLang.Fields("langcode").Value & """>" & rsLang.Fields("NativeLanguage").Value & "</a></td></tr>" & vbCrLf
			rsLang.MoveNext
		Loop
		Application("LANG_CAPSULE_HTML") = "<div align=""center""><table border=0 cellpadding=2 cellspacing=0>" & vbCRLf &_
			"<tr><td nowrap onclick=""location.href='" & Application("SiteRoot") & "/module/support/language/select.asp?code=US'"">" &_
			"<img src=""" & Application("SiteRoot") & "/icon/flags/us.png"" align=""left"" hspace=""10"" alt="""">" &_
			"<a href=""" & Application("SiteRoot") & "/module/support/language/select.asp?code=US"">English</a></td></tr>" & vbCrLf &_
			sHTML &_
			"</table></div>" & vbCrLf
	Else
		Application("LANG_CAPSULE_HTML") = "<div align=""center""><b class=""error"">none complete</b></div>"
	End If
	rsLang.Close
	Application("LANG_CAPSULE_UPDATED") = Now()
End Sub
%>