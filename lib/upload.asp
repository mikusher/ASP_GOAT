﻿<!-- #include file="../../lib/site_lib.asp" -->
<!-- #include file="../../lib/graphics/gfx_lib.asp" -->
<%
'--------------------------------------------------------------------
' upload.asp
'	Perform the upload of files to the server and then re-construct
'	the post to the original target (which will save the information
'	in the database)
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

' define the variables
Dim oObjContentType		' dictionary of form object content-types
Dim oObjFilenames		' dictionary of uploaded files
Dim oObjBlob			' dictionary of blobs (binary large objects)
Dim oObjForm			' dictionary of form variables
Dim bWasUpload			' boolean - has the form been posted & parsed?
Dim sErrorMsg			' error message generated by library
Dim sReplaceFind		' string in path to replace
Dim sReplaceWith		' string in path to replace with
Dim bOverwrite			' overwrite the existing file (if any)
Dim sPrependPrefix		' prefix to prepend to uploaded filenames
Dim sUploadMime			' acceptable mime types for upload
Dim sUploadExt			' restrict filename extensions for upload
Dim sUploadMax			' maximum bytes for a single uploaded file

bWasUpload = False
bOverwrite = False
sUploadMime = "image/jpg; image/gif; image/png;"
sUploadExt = "jpg; gif; png;"
sUploadMax = "8192"		' 8K maximum uploaded file size

' try to prevent caching pages in Internet Explorer
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

' save uploaded files and repost form to the referring script
Call FormUploadAndRepost

'--------------------------------------------------------------------
' FormPost
'	Saves all of the uploaded files to the path given in the method
'	parameter (sPath)
' RETURNS: True if a POST occurred, false otherwise

Function FormPost(sPath)
	Dim oFSO		' file system object
	Dim oFSOFile	' file object
	Dim sFilename	' filename being uploaded
	Dim nPos
	Dim sFile

	If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
		' parse the form data / file uploads into the dictionary objects
		ParseRequest sErrorMsg

		' create the necessary folders for this file
		CreateDirs sPath

		Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
		For Each sFile In oObjFilenames.Keys
			If Trim(oObjFilenames(sFile)) <> "" Then
				' build the filename on the local server (to save as)
				sFilename = Replace(Server.MapPath("/") & oFSO.BuildPath(sPath, oObjFilenames(sFile)), "/", "\")
				If (sReplaceFind <> "") Then sFilename = Replace(sFilename, sReplaceFind, sReplaceWith)

				If oFSO.FileExists(sFilename) And Not bOverwrite _
					And (sPath & "/" & oObjFilenames(sFile)) <> oObjForm("steallowoverwrite") Then
					sErrorMsg = sErrorMsg & "Sorry, File """ & sFilename & """ Already Exists (" & (sPath & "/" & oObjFilenames(sFile)) & ")"
				Else
					On Error Resume Next
					Set oFSOFile = oFSO.CreateTextFile(sFilename)
					If Err.Number <> 0 Then
						sErrorMsg = sErrorMsg & "Unable to create new file (" & Err.Description & "): " &_
							sFilename & "<BR>Key = " & sFile & "</B><BR>"
						FormPost = False
						Exit Function
					End If
					On Error Goto 0
					For nPos = 1 to LenB(oObjBlob(sFile))
						oFSOFile.Write Chr(AscB(MidB(oObjBlob(sFile), nPos, 1)))
					Next
					oFSOFile.Close
				End If
			End If
		Next
		Set oFSO = Nothing
		FormPost = True
	Else
		FormPost = False
	End If
End Function

' validate the image uploaded (if any)
If steForm("action") = "update" Then
	Dim nWidth, nHeight, nColors, sImgType

	' define the category icon
	If steForm("iconimage") <> "" Then
		sIconImage = Application("ASPNukeBasePath") & "img/articles/category/" & steForm("iconimagefile")
		' check the size of the icon image
		If gfxSpex(Server.MapPath(sIconImage), nWidth, nHeight, nColors, sImgType) Then
			If nWidth <> modParam("Articles", "IconImageWidth") Or nHeight <> modParam("Articles", "IconImageHeight") Then
				sErrorMsg = "Invalid Icon Image Size (" & nWidth & "x" & nHeight &_
					") - Should be (" & modParam("Articles", "IconImageWidth") & "x" & modParam("Articles", "IconImageHeight") & ")<br>"
			End If
		Else
			sErrorMsg = "File is corrupt (" & steForm("iconimagefile") & ") - Expected GIF, JPG or PNG image<br>"
		End If
	Else
		sIconImage = steForm("iconimage")
	End If
End If

'----------------------------------------------------------------
' URL decode a string

Function URLDecode(sText)
	Dim sSource, sTemp, sResult

	Dim nPos
	sSource = Replace(sText, "+", " ")
	For nPos = 1 To Len(sSource)
	    sTemp = Mid(sSource, nPos, 1)
	    If sTemp = "%" Then
			If nPos + 2 < Len(sSource) Then
				sResult = sResult & Chr(CInt("&H" & Mid(sSource, nPos + 1, 2)))
				nPos = nPos + 2
			End If
		Else
			sResult = sResult & sTemp
		End If
	Next
	URLDecode = sResult
End Function

'----------------------------------------------------------------
'Unicode string to Byte string conversion

Function UStr2Bstr(UStr)
	Dim nPos, sChar

	UStr2Bstr = ""
	For nPos = 1 to Len(UStr)
		sChar = Mid(UStr, nPos, 1)
		UStr2Bstr = UStr2Bstr & ChrB(AscB(sChar))
	Next
End Function

'----------------------------------------------------------------
'Byte string to Unicode string conversion

Function BStr2UStr(BStr)
	Dim nPos
	BStr2UStr = ""
	For nPos = 1 to LenB(BStr)
		BStr2UStr = BStr2UStr & Chr(AscB(MidB(BStr, nPos, 1)))
	Next
End Function

'----------------------------------------------------------------
' Parse the contents of a request (posted with the attribute:
' enctype="multipart/form-data")

Sub ParseRequest(sErrorMsg)
	Dim nTotalBytes, nPosBeg, nPosEnd
	Dim nPosBoundary, nPosTmp, nPosFileName
	Dim sBRequest, sBBoundary, sBContent
	Dim sName, sFileName, sContentType
	Dim strValue, sTemp
	Dim objFile

	' create the hashtables to store form information
	Set oObjContentType = Server.CreateObject("Scripting.Dictionary")
	Set oObjFilenames = Server.CreateObject("Scripting.Dictionary")
	Set oObjBlob = Server.CreateObject("Scripting.Dictionary")
	Set oObjForm = Server.CreateObject("Scripting.Dictionary")

	'Grab the entire contents of the Request as a Byte string
	nTotalBytes = Request.TotalBytes
	sBRequest = Request.BinaryRead(nTotalBytes)
	bWasUpload = True

	'Find the first Boundary
	nPosBeg = 1
	nPosEnd = _
	    InStrB(nPosBeg, sBRequest, UStr2Bstr(Chr(13)))
	If nPosEnd > 0 Then
		sBBoundary = _
	        MidB(sBRequest, nPosBeg, nPosEnd - nPosBeg)
		nPosBoundary = InStrB(1, sBRequest, sBBoundary)
	End If
	' was form submitted without ENCTYPE="multipart/form-data"?
	If sBBoundary = "" Then
		' YES - we can no longer access the Request.Form collection,
		' parse the request and populate the form collection
		nPosBeg = 1
		nPosEnd = InStrB(nPosBeg, sBRequest, UStr2Bstr("&"))
		Do While nPosBeg < LenB(sBRequest)
			' parse the element and add it to the collection
			sTemp = BStr2UStr(MidB(sBRequest, nPosBeg, nPosEnd - nPosBeg))
			nPosTmp = InStr(1, sTemp, "=")
			sName = URLDecode(Left(sTemp, nPosTmp - 1))
			strValue = URLDecode(Right(sTemp, Len(sTemp) - nPosTmp))
			oObjForm.Add sName, strValue

			' find the next element
			nPosBeg = nPosEnd + 1
			nPosEnd = InStrB(nPosBeg, sBRequest, UStr2Bstr("&"))
			If nPosEnd = 0 Then
				nPosEnd = LenB(sBRequest) + 1
			End If
		Loop
	Else
		' NO - assume form submitted with ENCTYPE="multipart/form-data"
		' parse all boundaries, place values into the form or file dictionary.
		Do Until (nPosBoundary = InStrB(sBRequest, sBBoundary & UStr2Bstr("--")))
			' get the post data properties
			nPosTmp = InStrB(nPosBoundary, sBRequest, UStr2Bstr("Content-Disposition"))
			nPosTmp = InStrB(nPosTmp, sBRequest, UStr2Bstr("name="))
			nPosBeg = nPosTmp + 6
			nPosEnd = InStrB(nPosBeg, sBRequest, UStr2Bstr(Chr(34)))
			sName = BStr2UStr(MidB(sBRequest, nPosBeg, nPosEnd - nPosBeg))
			' test for an element named 'filename'
			nPosFileName = InStrB(nPosBoundary, sBRequest, UStr2Bstr("filename="))

			' if found, we have a file, otherwise it is a normal form element
		    If nPosFileName <> 0 And nPosFileName < InStrB(nPosEnd, sBRequest, sBBoundary) Then
				' it is a file. Get the FileName
				nPosBeg = nPosFileName + 10
				nPosEnd = InStrB(nPosBeg, sBRequest, UStr2Bstr(chr(34)))
				sFileName = BStr2UStr(MidB(sBRequest, nPosBeg, nPosEnd - nPosBeg))
				' get the ContentType
				nPosTmp = InStrB(nPosEnd, sBRequest, UStr2Bstr("Content-Type:"))
				nPosBeg = nPosTmp + 14
				nPosEnd = InstrB(nPosBeg, sBRequest, UStr2Bstr(chr(13)))
				sContentType = BStr2UStr(MidB(sBRequest, nPosBeg, nPosEnd - nPosBeg))
				' check the acceptable filename extensions here
				If sFilename <> "" And Not (InStr(1, sUploadExt, Right(sFileName, 3) & ";") > 0) Then
					' sorry, filename extension not allowed
					sErrorMsg = sErrorMsg & "Invalid file (" & Mid(sFileName, InStrRev(sFileName, "\") + 1) & ") - Acceptable filename extensions:<BR>" & sUploadExt & ".<BR>"
				Else
					' check for acceptable mime encodings here
					If sFileName <> "" And Not (InStr(1, sUploadMime, sContentType & ";") > 0) Then
						' sorry, mime type not allowed
						sErrorMsg = sErrorMsg & "Invalid file (" & Mid(sFileName, InStrRev(sFileName, "\") + 1) & ") - Acceptable file types:<BR>" & sUploadMime & ".<BR>"
					Else
						' get the Content
						nPosBeg = nPosEnd + 4
						nPosEnd = InStrB(nPosBeg, sBRequest, sBBoundary) - 2
						If nPosEnd - nPosBeg > sUploadMax Then
							' sorry, file is too large
							sErrorMsg = sErrorMsg & "Uploaded file (" &  Mid(sFileName, InStrRev(sFileName, "\") + 1) & ") too large (maximum " & sUploadMax & " bytes)<BR>"
						Else
							sBContent = MidB(sBRequest, nPosBeg, nPosEnd - nPosBeg)
							If sFileName <> "" And sBContent <> "" Then
								' create the file object and add it to the files collection
								oObjFilenames.Add sName, sPrependPrefix & Right(sFileName, Len(sFileName) - InStrRev(sFileName, "\"))
								oObjContentType.Add sName, sContentType
								oObjBlob.Add sName, sBContent
							End If
						End If
					End If
				End If
		    Else ' it is a form element
				' get the value of the form element
				nPosTmp = InStrB(nPosTmp, sBRequest, UStr2Bstr(chr(13)))
				nPosBeg = nPosTmp + 4
				nPosEnd = InStrB(nPosBeg, sBRequest, sBBoundary) - 2

				strValue = BStr2UStr(MidB(sBRequest, nPosBeg, nPosEnd - nPosBeg))
				' add the form element to the collection
				oObjForm.Add sName, strValue
			End If
			' move to next element
			nPosBoundary = InStrB(nPosBoundary + LenB(sBBoundary), sBRequest, sBBoundary)
		Loop
	End If
End Sub

'--------------------------------------------------------------------
' CreateDirs
'	Create all directories necessary in the file path.  The path
'	supplied should be the virtual path to the new directory

Sub CreateDirs(sPath)
	Dim sDir, oFSO, nPos

	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	nPos = InStr(2, sPath, "/")
	Do While nPos > 0
		sDir = Server.MapPath("/") & Left(sPath, nPos - 1)
		if Not oFSO.FolderExists(sDir) Then Call oFSO.CreateFolder(sDir)
		nPos = Instr(nPos + 1, sPath, "/")
	Loop
	sDir = Server.MapPath("/") & sPath
	if Not oFSO.FolderExists(sDir) Then Call oFSO.CreateFolder(sDir)
	Set oFSO = Nothing
End Sub

'--------------------------------------------------------------------
' Repost the original form data to the destination script

Sub FormUploadAndRepost(ByVal sPath)
	Dim bSuccess, sKey, sQuery

	' process the upload here
	bSuccess = FormPost(sPath)

	' strip leading "/" from path (if nec)
	If Left(sPath, 1) = "/" Then sPath = Mid(sPath, 2)
	' add trailing "/" to the path (if nec)
	If Right(sPath, 1) <> "/" Then sPath = sPath & "/"

	With Response
		.Write "<HTML><HEAD>Post Redirect</HEAD><BODY onLoad=""document.form1.submit()"">"
		.Write "<h4>Javascript Redirect</h4>"
		.Write "<P>This page should have redirected you to the final page using a Javascript event.  If you are seeing this page it means that your browser does not support javascript.  Please click on the button below to continue:</P>"
		.Write "<FORM METHOD=""post"" NAME=""form1"" ACTION=""" & Request.ServerVariables("HTTP_REFERER") & """>"
		' build the form variables to pass to the next page
		If IsObject(oObjForm) Then
			For Each sKey In oObjForm.Keys
				.Write "<input type=""hidden"" name=""" & sKey
				.Write """ value=""" & Server.HTMLEncode(oObjForm.Item(sKey))
				.Write ">" & vbCrLf
			Next
		End If
		' include the filenames (local to server) for uploads
		If IsObject(oObjFilenames) Then
			For Each sKey In oObjFilenames
				.Write "<input type=""hidden"" name=""" & sKey
				.Write """ value=""" & Server.HTMLEncode(oObjForm.Item(sKey))
				.Write ">" & vbCrLf
			Next
		End If
		' indicate upload and redirect was already called
		' .Write "<input type=""hidden"" name=""uploadandredirect"" value=""Y"">"
		.Write "<input type=""hidden"" name=""uploaderror"" value=""" & Server.HTMLEncode(sErrorMsg) & """>"
		' build the submit button here
		.Write "<input type=""submit"" name=""_submit"" value="" Continue... "" value="""">"
		.Write "</FORM>"
		.Write "</body></html>"
	End With
End Sub

'--------------------------------------------------------------------
' Validate the file type and image size after upload completes

Function UploadValidation(ByVal sPath)
	Dim nWidth, nHeight, nColors, sImgType

	' include the filenames (local to server) for uploads
	If IsObject(oObjFilenames) Then
		' strip leading "/" from path (if nec)
		If Left(sPath, 1) = "/" Then sPath = Mid(sPath, 2)
		' add trailing "/" to the path (if nec)
		If Right(sPath, 1) <> "/" Then sPath = sPath & "/"
		For Each sKey In oObjFilenames
			' define the category icon
			If steForm("iconimage") <> "" Then
				sIconImage = Application("ASPNukeBasePath") & sPath & steForm("iconimagefile")
				' check the size of the icon image (ext. and mime already checked)
				If gfxSpex(Server.MapPath(sIconImage), nWidth, nHeight, nColors, sImgType) Then
					If nWidth <> modParam("Articles", "IconImageWidth") Or nHeight <> modParam("Articles", "IconImageHeight") Then
						sErrorMsg = "Invalid Icon Image Size (" & nWidth & "x" & nHeight &_
							") - Should be (" & modParam("Articles", "IconImageWidth") & "x" & modParam("Articles", "IconImageHeight") & ")<br>"
					End If
				Else
					sErrorMsg = "File is corrupt (" & steForm("iconimagefile") & ") - Expected GIF, JPG or PNG image<br>"
				End If
			Else
				sIconImage = steForm("iconimage")
			End If
		Next
	End If
End Function
%>