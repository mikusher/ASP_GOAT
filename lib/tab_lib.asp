<%
'--------------------------------------------------------------------
' tab_lib.asp
'	Displays a series of tabs with clickable links for the labels
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

'--------------------------------------------------------------------
' tabShow
'	Builds a list of tabs using the graphical tab images found under
'	Application("ASPNukeBasePath") & /img/tab

Sub tabShow(aLabel, aLink, sActive)
	Dim I

	' validate argument 1
	If Not IsArray(aLabel) Then
		aLabel = Split(aLabel, ",")
	End If

	' validate argument 2
	If Not IsArray(aLink) Then
		aLink = Split(aLink, ",")
	End If

	' make sure the number of labels = number of links
	If UBound(aLabel) <> UBound(aLink) Then
		tabShowError "Number of Labels (" & UBound(aLabel) & ") Should Equal Number of Links (" & UBound(aLink) & ")"
		Exit Sub
	End If

	' build the open table and top row for the tab display
	With Response
		.Write "<TABLE BORDER=0 CELLPADDING=0 CELLSPACING=0 BGCOLOR=""#A0A0A0"" WIDTH=""100%"">"
		.Write vbCrLf
		.Write "<TR>"
		For I = 0 To UBound(aLabel)
			.Write "<TD></TD>"
			.Write "<TD BGCOLOR=""#000000""><IMG SRC=""" & Application("ASPNukeBasePath") & "img/pixel.gif"" WIDTH=1 HEIGHT=1></TD>"
			.Write vbCrLf
		Next
		.Write "<TD></TD>"
		.Write vbCrLf
		.Write "</TR>"
		.Write vbCrLf
		.Write "<TR>"
		.Write vbCrLf
	End With

	' display the tabs here
	With Response
		For I = 0 To UBound(aLabel)
			' first create the graphic to the left of the link
			If I = 0 Then
				If sActive = aLabel(I) Then
					.Write "<TD WIDTH=16><IMG SRC=""" & Application("ASPNukeBasePath") & "img/tabs/lefton.gif"" WIDTH=16 HEIGHT=16></TD>"
					.Write vbCrLf
				Else
					.Write "<TD WIDTH=16><IMG SRC=""" & Application("ASPNukeBasePath") & "img/tabs/leftoff.gif"" WIDTH=16 HEIGHT=16></TD>"
					.Write vbCrLf
				End If
			Else ' we are inbetween tabs here
				If sActive = aLabel(I) Then
					.Write "<TD WIDTH=16><IMG SRC=""" & Application("ASPNukeBasePath") & "img/tabs/offon.gif"" WIDTH=16 HEIGHT=16></TD>"
					.Write vbCrLf
				ElseIf sActive = aLabel(I-1) Then
					.Write "<TD WIDTH=16><IMG SRC=""" & Application("ASPNukeBasePath") & "img/tabs/onoff.gif"" WIDTH=16 HEIGHT=16></TD>"
					.Write vbCrLf
				Else
					.Write "<TD WIDTH=16><IMG SRC=""" & Application("ASPNukeBasePath") & "img/tabs/offoff.gif"" WIDTH=16 HEIGHT=16></TD>"
					.Write vbCrLf
				End If
			End If

			' add the link to the tab item here
			.Write "<TD"
			If sActive = aLabel(I) Then .Write " BGCOLOR=""#FFFFFF""" Else .Write " BGCOLOR=""#D0D0D0"""
			.Write "><A HREF="""
			.Write aLink(I)
			.Write """ CLASS="""
			If sActive = aLabel(I) Then .Write "tabactive"">" Else .Write "tabinactive"">"
			.Write aLabel(I)
			.Write "</A></TD>"
		Next

		' build the right-most tab graphic
		If sActive = aLabel(UBound(aLabel)) Then
			.Write "<TD WIDTH=16><IMG SRC=""" & Application("ASPNukeBasePath") & "img/tabs/righton.gif"" WIDTH=16 HEIGHT=16></TD>"
			.Write vbCrLf
		Else
			.Write "<TD WIDTH=16><IMG SRC=""" & Application("ASPNukeBasePath") & "img/tabs/rightoff.gif"" WIDTH=16 HEIGHT=16></TD>"
			.Write vbCrLf
		End If

		' display the setup and help icons (if nec)
		If Request.ServerVariables("SCRIPT_NAME") <> "/module/help_popup.asp" And Right(Request.ServerVariables("SCRIPT_NAME"), 9) <> "/help.asp" Then
			' .Write "<TD WIDTH=16><IMG SRC=""" & Application("ASPNukeBasePath") & "img/pixel.gif"" WIDTH=16 HEIGHT=1></TD>"
			' .Write vbCrLf
			Dim nModuleID
			nModuleID = modModuleID
			If nModuleID > 0 Then
				If Right(Request.ServerVariables("SCRIPT_NAME"), 14) = "/configure.asp" Then
					' link to the param configuration
					.Write "<TD WIDTH=16><a href=""" & Application("ASPNukeBasePath") & "module/admin/module/param/param_list.asp?moduleid=" & nModuleID & """>"
					.Write "<IMG SRC=""" & Application("ASPNukeBasePath") & "img/setupicon.gif"" WIDTH=16 HEIGHT=16 BORDER=0 alt=""" & steGetText("Modify Configuration Parameters") & """></A></TD>"
				Else
					' link to the module configuration
					.Write "<TD WIDTH=16><a href=""" & Application("ASPNukeBasePath") & "module/admin/configure.asp?moduleid=" & nModuleID & """>"
					.Write "<IMG SRC=""" & Application("ASPNukeBasePath") & "img/setupicon.gif"" WIDTH=16 HEIGHT=16 BORDER=0 alt=""" & steGetText("Configure Module") & """></A></TD>"
				End If
				.Write vbCrLf
			End If
			.Write "<TD WIDTH=16><a href=""javascript:void(0)"" onclick=""window.open('" & Application("ASPNukeBasePath") & "module/help_popup.asp?pathinfo=" & Request.ServerVariables("PATH_INFO") & "&section=" & Server.URLEncode(sActive) & "' , '_new', 'menubar=no,scrollbars=yes,toolbar=no,statusbar=no,width=600,height=460')"">"
			.Write "<IMG SRC=""" & Application("ASPNukeBasePath") & "img/helpicon.gif"" WIDTH=16 HEIGHT=16 BORDER=0 alt=""Help Information""></A></TD>"
			.Write vbCrLf
		End If
		
		' close the table here
		.Write "</TR>"
		.Write vbCrLf
		.Write "</TABLE>"
	End With
End Sub

'--------------------------------------------------------------------
' tabShowError
'	Builds a list of tabs using the graphical tab images found under
'	/img/tab

Sub tabShowError(sError)
	With Response
		.Write "<P><B CLASS=""error"">"
		.Write "tab_lib.asp - "
		.Write sError
		.Write "</B></P>"
		.Write vbCrLf
	End With
End Sub
%>