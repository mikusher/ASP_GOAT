<!-- #include file="../../../lib/ado_lib.asp" -->
<%
'--------------------------------------------------------------------
' content.asp
'	Create the random quote content for the site.
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

Call locQuote

'--------------------------------------------------------------------------
' locQuote
'	Cache the random quotes for the site (in the application hash)
'	Output a single random quote on the page.

Sub locQuote
	Dim sStat, rsQuote, nPick, I

	If Application("QUOTECOUNT") = "" Then
		' cache all of the quotes here
		sStat = "SELECT	Quote, Author " &_
				"FROM	tblQuote " &_
				"WHERE	Active <> 0 " &_
				"AND	Archive = 0"
		Set rsQuote = adoOpenRecordset(sStat)
		I = 0
		Do Until rsQuote.EOF
			I = I + 1
			Application("QUOTE" & I) = "<TABLE BORDER=""0"" CELLPADDING=""2"" CELLSPACING=""0"" WIDTH=""100%"">" & vbCrLf &_
				"<TR><TD>&nbsp;&nbsp;</TD>" & vbCrLf &_
				"<TD><FONT CLASS=""quotetext"">" & rsQuote.Fields("Quote").Value &_
				"</FONT></TD>" & vbCrLf &_
				"<TD ALIGN=""right"" VALIGN=""bottom""><FONT CLASS=""quoteattr"">-- " & rsQuote.Fields("Author").Value & "</FONT></TD>" & vbCrLf &_
				"<TD>&nbsp;&nbsp;</TD>" & vbCrLf &_
				"</TR></TABLE>" & vbCrLf
			rsQuote.MoveNext
		Loop
		Application("QUOTECOUNT") = I
	End If
	Randomize
	nPick = Int(Rnd() * Application("QUOTECOUNT")) + 1
	Response.Write Application("QUOTE" & nPick)
End Sub
%>