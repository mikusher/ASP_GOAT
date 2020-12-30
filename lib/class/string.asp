<%
'--------------------------------------------------------------------
' string.asp
'	A class to implement efficient concatenation of strings
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

Class clsString 
	Private m_intGrowBy		' initial size & grow by size
   	Private m_intCounter	' array index to add new strings
	Private m_arrString()	' dynamic array of strings to concatenate

	'--------------------------------------------------------------
	' constructor for the string class

	Private Sub Class_Initialize()
		' dim an array and set position counter
   		m_intCounter = 1
   		m_intGrowBy = 100
   		Redim m_arrString(m_intGrowBy)
   	End Sub

	'--------------------------------------------------------------
	' empty the entire string (for reusing the class variables)

	Public Sub Reset
		'Erase current array and recreate
   		Erase m_arrString
   		Call Class_Initialize()
   	End Sub

	'--------------------------------------------------------------
	' get the value of the concatenated string

   	Public Property Get Value
   		' use the join function to create final string
   		Value = Join(m_arrString,"")
   	End Property 

	'--------------------------------------------------------------
	' add a new string to the buffer (dynamic array "m_arrString")

   	Public Sub Add(byval strValue)
   		' add a value to string array (doesn't have to be a string)
   		If m_intCounter > Ubound(m_arrString) Then _
   			Redim Preserve m_arrString(Ubound(m_arrString) + m_intGrowBy)
   		m_arrString(m_intCounter) = strValue

   		' increment position counter
   		m_intCounter = m_intCounter + 1	   		
   	End Sub
End Class
%>