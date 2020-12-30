<!-- #include file="../lib/site_lib.asp" -->
<!-- #include file="../lib/tab_lib.asp" -->
<!-- #include file="../lib/class/help.asp" -->
<%
'--------------------------------------------------------------------
' help_popup.asp
'	Display the integrated help for the current page being
'	displayed.
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

Dim sCurrentSection
Dim sPathInfo
Dim nLastSlash

sCurrentSection = steForm("section")
If sCurrentSection = "" Then sCurrentSection = "Overview"
sPathInfo = steForm("pathinfo")

nLastSlash = InStrRev(sPathInfo, "/")
If nLastSlash > 0 And nLastSlash <> Len(sPathInfo) Then
	sPathInfo = Left(sPathInfo, nLastSlash)
End If
%>
<!-- #include file="../header_popup.asp" -->

<%
' display the contents of the help here
Dim oHelp

Set oHelp = New clsHelp
If sPathInfo <> "" Then oHelp.PathInfo = sPathInfo

' display the help contents here
If oHelp.RetrieveHelp(sCurrentSection, True) Then
	' display the tab control for the various help sections available
	If Not oHelp.TabControl(sCurrentSection) Then
		Response.Write "<p><b class=""error"">" & oHelp.ErrorMsg & "</b></p>"
	End If

	Response.Write oHelp.HelpSection
Else
	Response.Write "<p><b class=""error"">" & oHelp.ErrorMsg & "</b></p>"
End If
%>

<!-- #include file="../footer_popup.asp" -->