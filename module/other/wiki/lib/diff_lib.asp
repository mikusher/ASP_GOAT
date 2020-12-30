<%
' diff_lib.asp
'	Class for performing diff calculations on two separate files.

Public Class clsDiff
	Private aFile1		' array of lines from file 1
	Private aFile2		' array of lines from file 2
	Private oSolutions	' dict object containing the solutions
	Private sAddStart		' prefix for lines that were added
	Private sDeleteStart	' prefix for lines that were deleted
	Private sChangeStart	' prefix for lines that were changed
	Private sAddEnd			' suffix for lines that were added
	Private sDeleteEnd		' suffix for lines that were deleted
	Private sChangeEnd		' suffix for lines that were changed

	Private Sub Class_Initialize
		Set oSolutions = Server.CreateObject("Scripting.Dictionary")
		sAddStart = "<font style=""color:green"">"
		sAddEnd = "<font>"
		sDeleteStart = "<font style=""color:red"">"
		sDeleteEnd = "</font>"
		sChangeStart = "<font style=""color:yellow"">"
		sChangeEnd = "</font>"
	End Sub

	'--------------------------------------------------------------
	' Array concatenation function

	Private Sub ArrayConcat(pa1, pa2)
		Dim I, nIndex

		nIndex = UBound(pa1) + 1
		ReDim Preserve pa1(UBound(pa1) + UBound(pa2) + 1)
		For I = 0 To UBound(pa2)
			pa1(nIndex + I) = pa2(I)
		Next
	End Sub

	'--------------------------------------------------------------
	' Recursive function for determining the longest common
	' sequence of lines from the DIFF files (sFile1 & sFile2)

	Private Function LCS(pnStart1, pnStart2)
		Dim aResult
		Dim aRemainder1
		Dim aRemainder2
		Dim sIndex

		' determine if a solution has already been found
		sIndex = pnStart1 & "," & pnStart2
		If (oSolutions.Exists(sIndex)) Then
		    LCS = oSolutions.Item(sIndex)
			Exit Function
		End If
		
		' If we're at the end of either list, then the longest subsequence is empty 
		If pnStart1 = UBound(aFile1) Or pnStart2 = UBound(aFile2) Then
		    aResult = Array(0)
		ElseIf (aFile1(pnStart1) = aFile2(pnStart2) Then
			' If the start element is the same in both, then it is on the LCS, so
			' we'll just recurse on the remainder of both lists.

            aResult = Array(0)
            aResult(0) = aFile1(pnStart1)
			Call ArrayConcat(aResult, LCS(pnStart1 + 1, pnStart2 + 1)
		Else
			' We don't know which list we should discard from.  
			' Try both ways, pick whichever is better.
		
			aRemainder1 = LCS(pnStart1 + 1, pnStart2);
			aRemainder2 = LCS(pnStart1, pnStart2 + 1);
			if (UBound(aRemainder1) > UBound(aRemainder2)) Then
				aResult = aRemainder1
			Else
				aResult = aRemainder2
			End IF
		End If		
		LCS = aResult
	End Function

	'--------------------------------------------------------------
	' Determine the longest common sequence of lines from the files
	' psFile1, psFile2 which are the contents of the files

	Private Function LongestCommonSubsequence(psSourceContents, psTargetContents)
		aFile1 = Split(psSourceContents, vbCrLf)
		aFile2 = Split(psTargetContents, vbCrLf)
	    LongestCommonSubsequence = LCS(0, 0)
	End Function 

	'--------------------------------------------------------------
	' Determine the longest common sequence of lines from the files
	' psFile1, psFile2 which are the contents of the files

	Public Sub DiffHTML(psSourceContents, psTargetContents)
		Dim aCommon			' common lines between the two
		Dim nIndex			' index within the common lines

		aCommon = LongestCommonSubsequence(psSourceContents, psTargetContents)
		nIndex = 0
		For I = 0 To UBound(aFile2)
			If nIndex <= UBound(aCommon) Then
				
			Else
			End If
		Next 
	End Sub
End Class

' create an instance of the diff object for the user
Dim oDiff
On Error Resume Next
Set oDiff = New clsDiff
If Err.Number <> 0 Then
	Response.Write "<p><b style=""color:red"">Unable to create the ""clsDiff"" object - aborting</b><br>"
	Response.Write "<b>" & Err.Number & " - " & Err.Description & "</p>" : Response.End
End If
On Error Goto 0
%>