<!-- #include file="../lib/site_lib.asp" -->
<!-- #include file="setup_lib.asp" -->
<!-- #include file="../lib/application_lib.asp" -->
<!-- #include file="../lib/tab_lib.asp" -->
<%
' -------------------------------------------------------------------
' setup3.asp
'	Build the tabbed interface used for initializing the site-wide
'	configuration variables stored in tblApplicationVar
'
' AUTH:	Ken Richards
' DATE:	10/23/03
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

Dim nTabID
Dim sErrorMsg
Dim sAction
Dim sIntroduction
Dim sSummary
Dim sTitle

nTabID = steNForm("TabID")
sAction = Trim(LCase(steForm("action")))
sTitle = steForm("SetupTitle")
If nTabID = 0 Then nTabID = 1

If sAction = "save" Then
	If Not appSave(nTabID, sErrorMsg) Then
	End if
End If

%>
<html>
<head>
	<title>ASP Nuke Setup</title>
	<meta name="author" content="Ken Richards">
	<meta http-equiv="content-type" content="text/html; charset=iso-8859-1">
	<meta name="robots" content="all">
	<meta http-equiv="Pragma" content="no-cache">
	<meta http-equiv="Expires" content="-1">
	<style>
	BODY { font: 10pt/10pt tahoma,helvetica,sans-serif; font-weight: normal; }
	P { font: 10pt/10pt tahoma,helvetica,sans-serif; font-weight: normal; }
	B { font: 10pt/10pt tahoma,helvetica,sans-serif; font-weight: bold; }
	H2 { font: 14pt/14pt tahoma,helvetica,sans-serif; font-weight:bold; color: #805030 }
	H3 { font: 12pt/12pt tahoma,helvetica,sans-serif; font-weight:bold; color: #805030 }
	H4 { font: 10pt/10pt tahoma,helvetica,sans-serif; font-weight:bold; color: #805030 }
	HR { height: 1px; margin-top: 2px; margin-bottom: 2px; padding: 0px; width:100%; color: #F0C0A0; }
	.error { font-family: Tahoma,Helvetica,sans-serif; font-size: 10pt; font-weight: bold; color: #FF0000 }
	A { font: 10pt/10pt Tahoma,Helvetica,sans-serif; color: #A06030; font-weight: bolder; text-decoration: none; }
	A:hover { font: 10pt/10pt Tahoma,Helvetica,sans-serif; color: #8080F0;; font-weight: bolder; text-decoration: none; }
	.form {  font-family: Tahoma,Helvetica,sans-serif; font-size: 9pt; font-weight: normal; background-color: #FFF8F0 }
	.formd {  font-family: Tahoma,Helvetica,sans-serif; font-size: 9pt; font-weight: normal; }
	.forml {  font-family: Tahoma,Helvetica,sans-serif; font-size: 9pt; font-weight: bolder; }
	.formradio { font-family: Tahoma,Helvetica,sans-serif; font-size: 9pt; font-weight: bolder;  border-color: black;  border-width: 1px; }
	</style>
</head>
<body>
<% locTabs nTabID, sTitle, sIntroduction, sSummary %>

<% If sAction <> "save" Or sErrorMsg <> "" Then %>

<% If sTitle <> "" Then %>
<H2><%= sTitle %> Configuration</H2>
<% Else %>
<H2>*Unknown* Configuration</H2>
<% End If %>

<% If sIntroduction & "" <> "" Then %>
<p>
<%= Replace(sIntroduction, vbCrLf, "<br>") %>
</p>
<% End If %>

<form method="post" action="setup3.asp">
<input type="hidden" name="TabID" value="<%= steEncform("TabID") %>">
<input type="hidden" name="SetupTitle" value="<%= Server.HTMLEncode(sTitle) %>">

<table border=0 cellpadding=2 cellspacing=0>
<% appConfigForm nTabID %>
<tr>
	<Td colsan="3" align="right"><br>
		<input type="submit" name="action" value=" Save " class="form">
	</td>
</tr>
</table>
</form>

<% If sSummary & "" <> "" Then %>
<p>
<%= Replace(sSummary, vbCrLf, "<br>") %>
</p>
<% End If %>

<% Else %>

<H2><%= sTitle %> Configuration Saved</H2>

<p>
The changes to the <%= sTitle %> were saved successfully in the database.  Please use the
tab navigation provided at the top of the page to continue configuring the ASP Nuke web
application.
</p>

<p>
When you are finished, you may <a href="../index.asp" target="_new">View Your ASP Nuke Site</a>.
You may also log into the <a href="../module/admin/index.asp" target="_new">Login as an Adminstrator</a>.
</p>

<% End If %>
<%
' show the wizard buttons only after no errors occurred
Call setWizardButtons(bAllowForward) 
%>

<p><b>AFTER YOUR ASP NUKE HAS BEEN CONFIGURED, CHOOSE ONE:</b></p>

</p>
<ul>
<li><a href="<%= Application("ASPNukeBasePath") %>" target="_new" alt="View ASP Nuke Home">View ASP Nuke Home</a>
<li><a href="<%= Replace(Application("ASPNukeBasePath") & "/module/admin/", "//", "/") %>" target="_new" alt="Goto ASP Nuke Admin">View ASP Nuke Admin Login</a>
</ul>
</body>
</html>
<%
Sub locTabs(nTabID, sTitle, sIntroduction, sSummary)
	Dim sStat, rsTab, sTab, sURL, sSelected, I

	sStat = "SELECT TabID, TabName, Title, Introduction, Summary FROM tblApplicationVarTab WHERE Archive = 0 ORDER BY OrderNo"
	Set rsTab = adoOpenRecordset(sStat)
	I = 0
	Do Until rsTab.EOF
		If I > 0 Then
			sTab = sTab & ","
			sURL = sURL & ","
		End If
		sTab = sTab & Replace(rsTab.Fields("TabName").Value, ",", ";")
		sURL = sURL & "setup3.asp?tabid=" & rsTab.Fields("TabID").Value
		If nTabID = rsTab.Fields("TabID").Value Then
			sSelected = rsTab.Fields("TabName").Value
			sTitle = rsTab.Fields("title").Value
			sIntroduction = rsTab.Fields("Introduction").Value
			sSummary = rsTab.Fields("Summary").Value
		End If
		rsTab.MoveNext
		I = I + 1
	Loop
	rsTab.Close
	Set rsTab = Nothing
	' display all of the tabs for this page
	Call tabShow(sTab, sURL, sSelected)
End Sub
%>