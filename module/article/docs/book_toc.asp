<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' book_toc.asp
'	Displays the table-of-contents for a book.
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

Dim sStat
Dim rsBook
Dim nBookID
Dim bShowSectionNo		' show section numbers? (like 2.3.13)
Dim bHasSections		' did the book viewed have any sections

nBookID = steNForm("BookID")

' retrieve the properties for this book
sStat = "SELECT	db.BookID, db.Title, db.SubTitle, db.Version, " &_
		"		da.FirstName, da.MiddleName, da.LastName, " &_
		"		db.ShowSectionNo, db.Created, db.PublishDate, db.OrderNo " &_
		"FROM	tblDocBook db " &_
		"INNER JOIN	tblDocAuthor da ON db.AuthorID = da.AuthorID " &_
		"WHERE	db.BookID = " & nBookID
Set rsBook = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->

<% If Not rsBook.EOF Then
	bShowSectionNo = rsBook.Fields("ShowSectionNo").Value %>

	<h3><%= rsBook.Fields("Title").Value %></h3>
	<% If Trim(rsBook.Fields("SubTitle").Value&"") <> "" Then %>
	<font class="tinytext"><%= Server.HTMLEncode(rsBook.Fields("SubTitle").Value) %></font><br>
	<% End If %>
	<font class="tinytext"><%= rsBook.Fields("FirstName").Value & " " & Trim(rsBook.Fields("MiddleName").Value & " " & rsBook.Fields("LastName").Value) %></font><br>

	<!-- TABLE-OF-CONTENTS SHOWN HERE -->
	<% bHasSections = locTOCDisplay(bShowSectionNo) %>
	<% If Not bHasSections Then %>
		<p><b class=""error""><% steTxt "Table-of-Contents Not Available" %></b></p>

		<p>
		<% steTxt "The Table-of-Contents for this book is currently unavailable because no content sections have been written yet." %>
		<% steTxt "We apologize for this inconvenience." %>
		<% steTxt "Please use the link below to return to the book listing." %>
		</p>
	<% End If %>

<% Else %>

<h3><% steTxt "Book Not Available" %></h3>

<p>
<% steTxt "Sorry, but the book that you requested could not be located in the database." %>&nbsp;
<% steTxt "This may be because the book has been removed or because you used an invalid URL when accessing this page." %>&nbsp;
<% steTxt "Please check out the" %> <a href="book_list.asp"><% steTxt "book list" %></a>
<% steTxt "to view all of the available books published on the site." %>
</p>

<% End If %>

<p align="center">
	<a href="book_list.asp" class="footerlink"><% steTxt "Book Listing" %></a>
</p>

<!-- #include file="../../../footer.asp" -->

<%
'----------------------------------------------------------------------------
' Display one level of the table of contents (read from the hash oSect)

Sub locTOCShowLevel(oSect, nParentID, nLevelNo)
	Dim aDocID, sIndent, I

	If Not oSect.Exists("PAR" & nParentID) Then Exit Sub
	If Trim(oSect.Item("PAR" & nParentID)) = "" Then Exit Sub

	' build the indent code for this level
	For I = 1 To nLevelNo
		sIndent = sIndent & " &nbsp; &nbsp; &nbsp;"
	Next

	aDocID = Split(oSect.Item("PAR" & nParentID), ",")
	With Response
	For I = 0 To UBound(aDocID)
		' output the section here
		.Write sIndent
		.Write oSect.Item(aDocID(I))
		.Write "<br>" & vbCrLf

		' display the child entries for the TOC (if nec)
		If oSect.Exists("PAR" & aDocID(I)) Then
			Call locTOCShowLevel(oSect, aDocID(I), nLevelNo + 1)
		End If
	Next
	End With
End Sub

'----------------------------------------------------------------------------
' Display the table-of-contents for a book
' RETURNS: True if the TOC was displayed, False otherwise

Function locTOCDisplay(bShowSectionNo)
	Dim sStat, rsDoc, sLink, oSect

	' create the dictionary object for the book sections
	Set oSect = Server.CreateObject("Scripting.Dictionary")

	sStat = "SELECT	DocID, ParentDocID, SectionNo, Title " &_
			"FROM	tblDoc " &_
			"WHERE	Active <> 0 " &_
			"AND	Archive = 0 " &_
			"ORDER BY OrderNo"
	Set rsDoc = adoOpenRecordset(sStat)
	Do Until rsDoc.EOF
		' define the list of child sections for the parent
		If oSect.Exists(rsDoc.Fields("ParentID").Value&"") Then
			oSect.Item("PAR" & rsDoc.Fields("ParentID").Value) = oSect.Item("PAR" & rsDoc.Fields("ParentID").Value) & "," &_
				rsDoc.Fields("DocID").Value
		Else
			oSect.Item("PAR" & rsDoc.Fields("ParentID").Value) = oSect.Item("PAR" & rsDoc.Fields("DocID").Value) & ","
		End if

		' define the section HTML to be displayed
		sLink = "<a href=""book_section.asp?bookid=" & nBookID & "&docid=" &_
			rsDoc.Fields("DocID").Value & """ class=""booktocsection"">" &_
			Server.HTMLEncode(rsDoc.Fields("Title").Value) & "</a>"
		If bShowSectionNo Then
			oSect.Item(rsDoc.Fields("DocID").Value&"") = rsDoc.Fields("SectionNo").Value & " - " & sLink
		Else
			oSect.Item(rsDoc.Fields("DocID").Value&"") = sLink
		End If
		rsDoc.MoveNext
	Loop
	rsDoc.Close
	Set rsDoc = Nothing

	' display all of the book sections & return the status
	If Not oSect.Exists("PAR0") Then
		locTOCDisplay = False
	Else
		' sections exist - show them
		Call locTOCShowLevel(oSect, 0, 0)
		locTOCDisplay = True
	End If
End Function
%>