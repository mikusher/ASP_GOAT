<%
'--------------------------------------------------------------------
' nukemail.asp
'	Class to send e-mails out from the server.
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

Class NukeMail
	Private mstrFromName		' who it's sent from
	Private mstrFromAddress		' address we're sending from
	Private mstrToName			' recipient name
	Private mstrToAddress		' recipient e-mail address
	Private mstrSubject			' subject for the mail message
	Private mstrHTML			' HTML body for the mail
	Private mstrText			' text body for the mail
	Private mstrError			' error message to report

	'----------------------------------------------------------------------
	' attempt to send an e-mail using CDO (new method)

	Function SendCDO
		Dim objMailConf, oMail

		On Error Resume Next
		' first update the CDO configuration
		Set objMailConf = Server.CreateObject("CDO.Configuration")
		If Err.Number <> 0 Then
			mstrError = mstrError & "Unable to create an instance of the CDO.Configuration object<br>"
			SendCDO = False
			Exit Function
		End If
		objMailConf.Fields.item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1
		objMailConf.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = "c:\inetpub\mailroot\pickup"
		objMailConf.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "http://localhost"
		objMailConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
		objMailConf.Fields.item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
		objMailConf.Fields.Update 
		If Err.Number <> 0 Then
			mstrError = mstrError & "Unable to update the CDO Configuration<br>" & Err.Number & " - " & Err.Description & "<br>"
			SendCDO = False
			Exit Function
		End If

		Set oMail = Server.CreateObject("CDO.Message")
		If Err.Number <> 0 Then
			mstrError = mstrError & "Unable to create an instance of the CDO.Message object<br>"
			SendCDO = False
			Exit Function
		End If
		Set oMail.Configuration = objMailConf
		oMail.Subject = mstrSubject
		oMail.From = mstrFromAddress
		oMail.To = mstrToAddress
		If Trim(mstrHTML) <> "" Then oMail.HTMLBody = mstrHTML
		If Trim(mstrText) <> "" Then oMail.TextBody = mstrText
		oMail.Send
		If Err.Number <> 0 Then
			mstrError = mstrError & "Unable to send message using the CDO.Message object<br>" & Err.Number & " - " & Err.Description & "<br>"
			SendCDO = False
			Exit Function
		End If
		Set oMail = Nothing
		On Error Goto 0
		SendCDO = True
	End Function

	'----------------------------------------------------------------------
	' attempt to send an e-mail using CDONTS (old method)

	Function SendCDONTS
		Dim oMail
	
		On Error Resume Next
		Set oMail = Server.CreateObject("CDONTS.NewMail")
		If Err.Number <> 0 Then
			mstrError = mstrError & "Unable to create an instance of the CDONTS.NewMail object<br>"
			SendCDONTS = False
			Exit Function
		End If
		oMail.SendUsing = "SMTP"
		oMail.From = mstrFromAddress ' & " (" & sFromName & ")"
		oMail.To = mstrToAddress ' & " (" & sToName & ")"
		oMail.Subject = mstrSubject
		If mstrHTML <> "" Then
			oMail.Body = mstrHTML
			oMail.BodyFormat = 0
			oMail.MailFormat = 0
		ElseIf mstrText <> "" Then
			oMail.Body = mstrText
			oMail.BodyFormat = 1
		End If
		oMail.Send
		If Err.Number <> 0 Then
			mstrError = mstrError & "Unable to send message using the CDONTS.NewMail object<br>" & Err.Number & " - " & Err.Description & "<br>"
			SendCDONTS = False
			Exit Function
		End If
		On Error Goto 0
		Set oMail = Nothing
		SendCDONTS = True
	End Function

	'----------------------------------------------------------------------
	' attempt to send an e-mail using ASPMail (3rd party component)

	Function SendASPMail
		Dim oMail
	
		On Error Resume Next
		Set oMail = Server.CreateObject("SMTPsvg.Mailer")
		If Err.Number <> 0 Then
			mstrError = mstrError & "Unable to create an instance of the SMTPsvg.Mailer (ASPMail) object<br>"
			SendASPMail = False
			Exit Function
		End If
		oMail.FromName   = mstrFromName
		oMail.FromAddress= mstrFromAddress
		' oMail.RemoteHost = "mailhost.localisp.net"
		oMail.AddRecipient mstrToName, mstrToAddress
		oMail.Subject    = mstrSubject
		If mstrHTML <> "" Then
			oMail.ContentType = "text/html"
			oMail.BodyText = mstrHTML
		Else
			oMail.BodyText = mstrText
		End If
		oMail.SendMail
		If Err.Number <> 0 Then
			mstrError = mstrError & "Unable to send message using the SMTPsvg.Mailer (ASPMail) object<br>" & Err.Number & " - " & Err.Description & "<br>"
			SendASPMail = False
			Exit Function
		End If
		On Error Goto 0
		Set oMail = Nothing
		SendASPMail = True
	End Function

	'----------------------------------------------------------------------
	' attempt to send an e-mail using ASPEmail (3rd party component)
	' http://www.aspemail.com/

	Function SendASPEmail
		Dim oMail
	
		On Error Resume Next
		Set oMail = Server.CreateObject("Persits.MailSender")
		If Err.Number <> 0 Then
			mstrError = mstrError & "Unable to create an instance of the Persits.MailSender (ASPEmail) object<br>"
			SendASPEmail = False
			Exit Function
		End If
		oMail.FromName   = mstrFromName
		oMail.From = mstrFromAddress
		' oMail.RemoteHost = "mailhost.localisp.net"
		oMail.AddAddress mstrToAddress, mstrToName
		oMail.Subject    = mstrSubject
		If mstrHTML <> "" Then
			oMail.Body = mstrHTML
			oMail.IsHTML = True
		Else
			oMail.Body = mstrText
		End If
		oMail.Send
		If Err.Number <> 0 Then
			mstrError = mstrError & "Unable to send message using the Persits.MailSender (ASPEmail) object<br>" & Err.Number & " - " & Err.Description & "<br>"
			SendASPEmail = False
			Exit Function
		End If
		On Error Goto 0
		Set oMail = Nothing
		SendASPEmail = True
	End Function

	'----------------------------------------------------------------------
	' check to see if the supplied e-mail address is valid
	' RETURNS: True if it is valid, False otherwise

	Private Function IsValidEmail(sEmail)
		Dim oRE

		' make sure a subject and body are present
		Set oRE = New RegExp
		oRE.Pattern = "(@.*@)|(\.\.)|(@\.)|(^\.)"
		If Not oRE.Test(sEmail) Then
			oRE.Pattern = "^.+\@(\[?)[a-zA-Z0-9\-\.]+\.([a-zA-Z]{2,3}|[0-9]{1,3})(\]?)$"
			If Not oRE.Test(sEmail) Then
				IsValidEmail = False
				Exit Function
			End If
		End If
		IsValidEmail = True
	End Function

	'----------------------------------------------------------------------
	' Send
	'	Attempt to send an e-mail using any of the methods available

	Function Send
		' check the mail parameters
		If Trim(mstrFromAddress) = "" Then
			mstrError = mstrError & "NukeMail.Send() - No valid from address found<br>"
			Send = False : Exit Function
		ElseIf Not IsValidEmail(mstrFromAddress) Then
			mstrError = mstrError & "NukeMail.Send() - Invalid ""FromAddress"" specified (" & mstrFromAddress & ")<br>"
			Send = False : Exit Function
		End If
		If Trim(mstrToAddress) = "" Then
			mstrError = mstrError & "NukeMail.Send() - No valid to address found<br>"
			Send = False : Exit Function
		ElseIf Not IsValidEmail(mstrToAddress) Then
			mstrError = mstrError & "NukeMail.Send() - Invalid ""ToAddress"" specified (" & mstrToAddress & ")<br>"
			Send = False : Exit Function
		End If
		If Trim(mstrSubject) = "" Then
			mstrError = mstrError & "NukeMail.Send() - No valid message subject found<br>"
			Send = False : Exit Function
		End If
		If Trim(mstrHTML) = "" And Trim(mstrText) = "" Then
			mstrError = mstrError & "NukeMail.Send() - No valid message body found<br>"
			Send = False : Exit Function
		End If

		' attempt to send using the primary method
		bSuccess = False
		Select Case Application("NukeMailMethod")
			Case "SendCDO" : bSuccess = SendCDO
			Case "SendCDONTS" : bSuccess = SendCDONTS
			Case "SendASPMail" : bSuccess = SendASPMail
			Case "SendASPEmail" : bSuccess = SendASPEmail
		End Select

		' if operation failed - attempt to use other methods
		If Not bSuccess And Application("NukeMailMethod") <> "SendCDO" Then
			bSuccess = SendCDO
			If bSuccess And Application("NukeMailMethod") <> "SendCDO" Then
				Application.Lock
				Application("NukeMailMethod") = "SendCDO"
				Application.Unlock
			End If
		End If
		
		If Not bSuccess And Application("NukeMailMethod") <> "SendCDONTS" Then
			bSuccess = SendCDONTS
			If bSuccess And Application("NukeMailMethod") <> "SendCDONTS" Then
				Application.Lock
				Application("NukeMailMethod") = "SendCDONTS"
				Application.Unlock
			End If
		End If

		If Not bSuccess And Application("NukeMailMethod") <> "SendASPMail" Then
			bSuccess = SendASPMail
			If bSuccess And Application("NukeMailMethod") <> "SendASPMail" Then
				Application.Lock
				Application("NukeMailMethod") = "SendASPMail"
				Application.Unlock
			End If
		End If

		If Not bSuccess And Application("NukeMailMethod") <> "SendASPEmail" Then
			bSuccess = SendASPEmail
			If bSuccess And Application("NukeMailMethod") <> "SendASPEmail" Then
				Application.Lock
				Application("NukeMailMethod") = "SendASPEmail"
				Application.Unlock
			End If
		End If
		If Not bSuccess Then
			mstrError = mstrError & "NukeMail.Send() - Unable to send mail using available methods<br>"
			Send = False : Exit Function
		End If
		Send = True
	End Function

	'----------------------------------------------------------------------
	' sender name and address properties

	Public Property Let FromAddress(sValue)
		mstrFromAddress = sValue
	End Property

	Public Property Get FromAddress
		FromAddress = mstrFromAddress
	End Property

	Public Property Let FromName(sValue)
		mstrFromName = sValue
	End Property
	
	Public Property Get FromName
		FromName = mstrFromName
	End Property

	'----------------------------------------------------------------------
	' recipient name and address properties

	Public Property Let ToAddress(sValue)
		 mstrToAddress = sValue
	End Property

	Public Property Get ToAddress
		ToAddress = mstrToAddress
	End Property

	Public Property Let ToName(sValue)
		mstrToName = sValue
	End Property

	Public Property Get ToName
		ToName = mstrToName
	End Property

	'----------------------------------------------------------------------
	' subject and content of the e-mail properties

	Public Property Let Subject(sValue)
		mstrSubject = sValue
	End Property

	Public Property Get Subject
		Subject = mstrSubject
	End Property

	Public Property Let HTMLBody(sValue)
		mstrHTML = sValue
	End Property

	Public Property Get HTMLBody
		Body = mstrHTML
	End Property

	Public Property Let TextBody(sValue)
		mstrText = sValue
	End Property

	Public Property Get TextBody
		Body = mstrText
	End Property

	'----------------------------------------------------------------------
	' error message generated by this class

	Public Property Get ErrorMsg
		ErrorMsg = mstrError
	End Property
End Class
%>