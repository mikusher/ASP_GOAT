<%
'--------------------------------------------------------------------
' lzw.asp
'	A class to implement LZW compression of data (scripts or binaries)
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

'This is a 1 run method but we have to keep the whole contents
'in memory until some variables are saved wich are needed bij the decompressor

Class clsLZW
	Private m_lngDictPos ' As Long 'de positie waar de volgende karakters worden ingevoegd
	Private m_intMaxCharLength ' As Byte 'Maximum stringlengte in de dictionary
	Private m_lngMaxDictDeep ' As Long 'maximaal opgeslagen woorden per dictionary
	Private m_intTotBitDeep ' As Byte 'totale bitlengte per karakter of karaktervolgorde
	Private m_objHash 					' As oHashTable
	Private m_intDictionarySize			' maximum size of the dictionary

	Private m_lOnBits(30)			' and mask to turn on lower "n" bits
	Private m_l2Power(30)			' powers of 2 ("n" ^ 2)

	'--------------------------------------------------------------
	' constructor for the MD5 class

	Private Sub Class_Initialize
		m_intMaxCharLength = 16
		m_intDictionarySize = 16383

		m_lOnBits(0) = CLng(1)
		m_lOnBits(1) = CLng(3)
		m_lOnBits(2) = CLng(7)
		m_lOnBits(3) = CLng(15)
		m_lOnBits(4) = CLng(31)
		m_lOnBits(5) = CLng(63)
		m_lOnBits(6) = CLng(127)
		m_lOnBits(7) = CLng(255)
		m_lOnBits(8) = CLng(511)
		m_lOnBits(9) = CLng(1023)
		m_lOnBits(10) = CLng(2047)
		m_lOnBits(11) = CLng(4095)
		m_lOnBits(12) = CLng(8191)
		m_lOnBits(13) = CLng(16383)
		m_lOnBits(14) = CLng(32767)
		m_lOnBits(15) = CLng(65535)
		m_lOnBits(16) = CLng(131071)
		m_lOnBits(17) = CLng(262143)
		m_lOnBits(18) = CLng(524287)
		m_lOnBits(19) = CLng(1048575)
		m_lOnBits(20) = CLng(2097151)
		m_lOnBits(21) = CLng(4194303)
		m_lOnBits(22) = CLng(8388607)
		m_lOnBits(23) = CLng(16777215)
		m_lOnBits(24) = CLng(33554431)
		m_lOnBits(25) = CLng(67108863)
		m_lOnBits(26) = CLng(134217727)
		m_lOnBits(27) = CLng(268435455)
		m_lOnBits(28) = CLng(536870911)
		m_lOnBits(29) = CLng(1073741823)
		m_lOnBits(30) = CLng(2147483647)
		
		m_l2Power(0) = CLng(1)
		m_l2Power(1) = CLng(2)
		m_l2Power(2) = CLng(4)
		m_l2Power(3) = CLng(8)
		m_l2Power(4) = CLng(16)
		m_l2Power(5) = CLng(32)
		m_l2Power(6) = CLng(64)
		m_l2Power(7) = CLng(128)
		m_l2Power(8) = CLng(256)
		m_l2Power(9) = CLng(512)
		m_l2Power(10) = CLng(1024)
		m_l2Power(11) = CLng(2048)
		m_l2Power(12) = CLng(4096)
		m_l2Power(13) = CLng(8192)
		m_l2Power(14) = CLng(16384)
		m_l2Power(15) = CLng(32768)
		m_l2Power(16) = CLng(65536)
		m_l2Power(17) = CLng(131072)
		m_l2Power(18) = CLng(262144)
		m_l2Power(19) = CLng(524288)
		m_l2Power(20) = CLng(1048576)
		m_l2Power(21) = CLng(2097152)
		m_l2Power(22) = CLng(4194304)
		m_l2Power(23) = CLng(8388608)
		m_l2Power(24) = CLng(16777216)
		m_l2Power(25) = CLng(33554432)
		m_l2Power(26) = CLng(67108864)
		m_l2Power(27) = CLng(134217728)
		m_l2Power(28) = CLng(268435456)
		m_l2Power(29) = CLng(536870912)
		m_l2Power(30) = CLng(1073741824)
	End Sub

	Public Sub Compress_LZW_Static_Hash(aFileArray() As Byte)
	    Dim nByteValue 		' As Byte
	    Dim lTempByte 		' As Long
	    Dim nExtraBits		' As Integer
	    Dim sDictStr		' As String
	    Dim sNewStr			' As String
	    Dim nComPByte()		' As Byte
	    Dim lCompPos		' As Long
	    Dim lDictVal		' As Long
	    Dim lDictPosit		' As Long
	    Dim lDictPositOld	' As Long
	    Dim lFilePos As Long
	    Dim lFileLength As Long
	    Dim lTemp As Long
	    Dim lMaxDictPagesInBytes
	    Set m_objHash = Server.CreateObject("Scripting.Dictionary")	' New HashTable
	
	    lMaxDictPagesInBytes = CLng(1024) * m_intDictionarySize - 1
	    Call Init_Dict(lMaxDictPagesInBytes, 0)
	    lFileLength = UBound(aFileArray)
	    ReDim nComPByte(lFileLength + 10)
	    nComPByte(0) = m_intMaxCharLength
	    nComPByte(1) = m_lngMaxDictDeep - Int(m_lngMaxDictDeep / 256) * 256
	    nComPByte(2) = Int((m_lngMaxDictDeep - nComPByte(1)) / 256)
	    lTemp = lFileLength
	    nComPByte(6) = lTemp And 255 : lTemp = Int(lTemp / 256)
	    nComPByte(5) = lTemp And 255 : lTemp = Int(lTemp / 256)
	    nComPByte(4) = lTemp And 255 : lTemp = Int(lTemp / 256)
	    nComPByte(3) = lTemp And 255 : lTemp = Int(lTemp / 256)
	    lFilePos = 0
	    lCompPos = 7
	    sDictStr = ""
	    nExtraBits = 0
	    Do Until lFilePos > lFileLength
	        nByteValue = aFileArray(lFilePos)
	        lFilePos = lFilePos + 1
	        sNewStr = sDictStr & Chr(nByteValue)
	        lDictPosit = Search(sNewStr)
	        If lDictPosit <> m_lngMaxDictDeep + 1 Then
	            sDictStr = sNewStr
	            lDictPositOld = lDictPosit
	        Else
	            nExtraBits = nExtraBits + m_intTotBitDeep - 8
	            lDictVal = (lTempByte * (m_l2Power(m_intTotBitDeep))) + lDictPositOld
	            lTempByte = lDictVal And ((m_l2Power(nExtraBits)) - 1)
	            lDictVal = Int(lDictVal / (m_l2Power(nExtraBits)))
	            If lCompPos > UBound(nComPByte) Then ReDim Preserve nComPByte(lCompPos + 500)
	            nComPByte(lCompPos) = lDictVal
	            lCompPos = lCompPos + 1
	            If nExtraBits >= m_intTotBitDeep Then
	                nExtraBits = nExtraBits - 8
	                lDictVal = lTempByte
	                lTempByte = lDictVal And ((m_l2Power(nExtraBits)) - 1)
	                lDictVal = Int(lDictVal / (m_l2Power(nExtraBits)))
	                If lCompPos > UBound(nComPByte) Then ReDim Preserve nComPByte(lCompPos + 500)
	                nComPByte(lCompPos) = lDictVal
	                lCompPos = lCompPos + 1
	            End If
	            Call AddToDict(sNewStr, 1)
	            lDictPositOld = nByteValue
	            sDictStr = Chr(nByteValue)
	        End If
	    Loop
	    nExtraBits = nExtraBits + m_intTotBitDeep - 8
	    lDictVal = (lTempByte * (m_l2Power(m_intTotBitDeep))) + lDictPositOld
	    lTempByte = lDictVal And ((m_l2Power(nExtraBits)) - 1)
	    lDictVal = Int(lDictVal / (m_l2Power(nExtraBits)))
	    If lCompPos > UBound(nComPByte) Then ReDim Preserve nComPByte(lCompPos + 500)
	    nComPByte(lCompPos) = lDictVal
	    lCompPos = lCompPos + 1
	    Do While nExtraBits > 0
	        nExtraBits = nExtraBits - 8
	        lDictVal = lTempByte
	        lTempByte = lDictVal And ((m_l2Power(nExtraBits)) - 1)
	        lDictVal = Int(lDictVal / (m_l2Power(nExtraBits)))
	        If lCompPos > UBound(nComPByte) Then ReDim Preserve nComPByte(lCompPos + 500)
	        nComPByte(lCompPos) = lDictVal
	        lCompPos = lCompPos + 1
	    Loop
	    Set m_objHash = Nothing
	    ReDim aFileArray(lCompPos - 1)
	    Call CopyMem(aFileArray(0), nComPByte(0), lCompPos)
	End Sub
	
	Public Sub DeCompress_LZW_Static_Hash(aFileArray) '() As Byte
	    Dim nReadBits		' As Integer
	    Dim lDictVal		' As Long
	    Dim lTempByte		' As Long
	    Dim lOldKarValue	' As Long
	    Dim nDeComPByte()	' As Byte
	    Dim lDeCompPos		' As Long
	    Dim lFilePos		' As Long
	    Dim lFileLength		' As Long
	    Dim sChar			' As String
	    Dim sOldChar		' As String

	    Set m_objHash = Server.CreateObject("Scripting.Dictionary") ' New oHashTable
	    m_intMaxCharLength = aFileArray(0)
	    m_lngMaxDictDeep = aFileArray(1) + 256 * aFileArray(2)
	    lFileLength = aFileArray(3) * 256 + aFileArray(4)
	    lFileLength = lFileLength * 256 + aFileArray(5)
	    lFileLength = lFileLength * 256 + aFileArray(6)
	    Call Init_Dict(m_lngMaxDictDeep, 0)
	    ReDim nDeComPByte(lFileLength)
	    nReadBits = 0
	    lTempByte = 0
	    lDeCompPos = -1
	    lFilePos = 7
	    lDictVal = -1
	    sChar = ""
	    Do Until lDeCompPos >= lFileLength
	        lOldKarValue = lDictVal
	        sOldChar = sChar
	        lDictVal = lTempByte
	        Do While nReadBits < m_intTotBitDeep And lFilePos <= UBound(aFileArray)
	            nReadBits = nReadBits + 8
	            lDictVal = lDictVal * 256 + aFileArray(lFilePos)
	            lFilePos = lFilePos + 1
	        Loop
	        If nReadBits < m_intTotBitDeep Then lDictVal = lDictVal * (m_l2Power((m_intTotBitDeep - nReadBits))): nReadBits = m_intTotBitDeep
	        nReadBits = nReadBits - m_intTotBitDeep
	        lTempByte = (lDictVal And ((m_l2Power(nReadBits)) - 1))
	        If nReadBits > 0 Then lDictVal = Int(lDictVal / m_l2Power(nReadBits))
	        sChar = m_objHash.GetKey(lDictVal)
	        If sChar <> "" Then
	            Call AddASC2Array(nDeComPByte, lDeCompPos, sChar)
	            If lOldKarValue <> -1 Then Call AddToDict(sOldChar & Left(sChar, 1), 0)
	        Else
	            sChar = sOldChar & Left(sOldChar, 1)
	            Call AddToDict(sChar, 0)
	            Call AddASC2Array(nDeComPByte, lDeCompPos, sChar)
	        End If
	    Loop
	    Set m_objHash = Nothing
	    ReDim aFileArray(lDeCompPos)
	    Call CopyMem(aFileArray(0), nDeComPByte(0), lDeCompPos + 1)
	End Sub
	
	Private Sub Init_Dict(lMaxDictPagesInBytes, nStoreTilCharLength)
	    Dim X

		If lMaxDictPagesInBytes = 0 Then
			lMaxDictPagesInBytes = 512
		End If
		If nStoreTilCharLength = 0 Then
			nStoreTilCharLength = 50
		End If
	    If lMaxDictPagesInBytes > 65535 Then
	        lMaxDictPagesInBytes = 65535
	    ElseIf lMaxDictPagesInBytes < 255 Then
	        lMaxDictPagesInBytes = 255
	    End If
	    lMaxDictPagesInBytes = lMaxDictPagesInBytes - 1
	    For X = 0 To 16
	        If lMaxDictPagesInBytes < m_l2Power(X) Then
	            m_intTotBitDeep = X
	            Exit For
	        End If
	    Next
	    m_intMaxCharLength = nStoreTilCharLength
	    m_lngMaxDictDeep = lMaxDictPagesInBytes
	    Call Clean_Dictionary
	End Sub
	
	Private Sub Clean_Dictionary()
	    Dim X As Long
	    Dim Y As Long
	    m_objHash.SetSize (m_lngMaxDictDeep)
	    For X = 0 To 255
	        m_objHash.Add Chr(X), X
	    Next
	    m_lngDictPos = 256
	End Sub
	
	Private Function Search(sChar) ' As Long
	    Dim X As Variant
	    X = m_objHash.Item(sChar)
	    If Not IsEmpty(X) Then
	        Search = X
	    Else
	        Search = m_lngMaxDictDeep + 1
	    End If
	End Function
	
	Private Sub AddToDict(sChar, Comp1Decomp0 As Byte)
	    If Len(sChar) = 1 Or Len(sChar) - 2 > m_intMaxCharLength Then Exit Sub
	    If m_lngDictPos + Comp1Decomp0 >= m_lngMaxDictDeep Then Call Clean_Dictionary
	    m_objHash.Add sChar, m_lngDictPos
	    m_lngDictPos = m_lngDictPos + 1
	End Sub
	
	Private Sub AddASC2Array(WichArray() As Byte, StartPos As Long, Text As String)
	    Dim X As Long
	    For X = 1 To Len(Text)
	        WichArray(StartPos + X) = ASC(Mid(Text, X, 1))
	    Next
	    StartPos = StartPos + Len(Text)
	End Sub
End Class	
