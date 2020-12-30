<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' book_list.asp
'	Displays a list of the document books for the site
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

Select Case LCase(steForm("action"))
	Case "moveup"
		Dim rsPrev, sPrevOrder

		sStat = "SELECT	OrderNo " &_
				"FROM	tblDocBook " &_
				"WHERE	OrderNo < " & steNForm("OrderNo") & " " &_
				"ORDER BY OrderNo DESC"
		Set rsPrev = adoOpenRecordset(sStat)
		If Not rsPrev.EOF Then
			sPrevOrder = rsPrev.Fields("OrderNo").Value
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblDocBook " &_
					"SET	OrderNo = OrderNo + 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & sPrevOrder
			' Response.Write "query = *" & sStat & "*<BR><BR>"
			Call adoExecute(sStat)

			sStat = "UPDATE	tblDocBook " &_
					"SET	OrderNo = " & sPrevOrder & ", Modified = " & adoGetDate & " " &_
					"WHERE	BookID = " & steNForm("BookID")
			' Response.Write "query = *" & sStat & "*<BR><BR>"
			Call adoExecute(sStat)
			modRefresh True
		End If
		rsPrev.Close
		Set rsPrev = Nothing
	Case "movedown"
		Dim rsNext, sNextOrder

		sStat = "SELECT	OrderNo " &_
				"FROM	tblDocBook " &_
				"WHERE	OrderNo > " & steNForm("OrderNo") & " " &_
				"ORDER BY OrderNo"
		Set rsNext = adoOpenRecordset(sStat)
		If Not rsNext.EOF Then
			sNextOrder = rsNext.Fields("OrderNo").Value
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblDocBook " &_
					"SET	OrderNo = OrderNo - 1, Modified = " & adoGetDate & " " &_
					"WHERE	OrderNo = " & sNextOrder
			Call adoExecute(sStat)

			sStat = "UPDATE	tblDocBook " &_
					"SET	OrderNo = " & sNextOrder & ", Modified = " & adoGetDate & " " &_
					"WHERE	BookID = " & steNForm("BookID")
			Call adoExecute(sStat)
			modRefresh True
		End If
		rsNext.Close
		Set rsNext = Nothing
	Case "activ"
			' archive a module
			sStat = "UPDATE	tblDocBook " &_
					"SET	Archive = 1 " &_
					"WHERE	BookID = " & steNForm("BookID")
			Call adoExecute(sStat)
			modRefresh True
	Case "deactiv"
			sStat = "UPDATE	tblDocBook " &_
					"SET	Archive = 0 " &_
					"WHERE	BookID = " & steNForm("BookID")
			Call adoExecute(sStat)
			modRefresh True
End Select

sStat = "SELECT	tblDocBook.BookID, tblDocBook.Title, tblDocBook.SubTitle, tblDocBook.Version, " &_
		"		tblDocAuthor.FirstName, tblDocAuthor.MiddleName, tblDocAuthor.LastName, " &_
		"		tblDocBook.Created, tblDocBook.PublishDate, tblDocBook.OrderNo " &_
		"FROM	tblDocBook " &_
		"INNER JOIN	tblDocAuthor ON tblDocBook.AuthorID = tblDocAuthor.AuthorID " &_
		"ORDER BY tblDocBook.OrderNo"
Set rsBook = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Book" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "Book List" %></H3>

<P>
<% steTxt "Shown below are all of the current books defined in the database." %>
</P>

<% If Not rsBook.EOF Then %>

<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 class="list">
<TR BGCOLOR="#E0C0A0">
	<TD class="listhead"><% steTxt "Order" %></TD>
	<TD class="listhead"><% steTxt "Title" %></TD>
	<TD class="listhead"><% steTxt "Author" %></TD>
	<TD class="listhead" ALIGN="center"><% steTxt "Version" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Published" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<%	I = 0
	Do Until rsBook.EOF %>
<TR class="list<%= I mod 2 %>">
	<TD><%= rsBook.Fields("OrderNo").Value %></TD>
	<TD nowrap><%= rsBook.Fields("Title").Value %><% If rsBook.Fields("SubTitle").Value & "" <> "" Then %><br><font class="tinytext"><%= rsBook.Fields("SubTitle").Value %></font><% End If %></TD>
	<TD nowrap><%= rsBook.Fields("FirstName").Value & " " & rsBook.Fields("MiddleName").Value & " " & rsBook.Fields("LastName").Value %></TD>
	<TD nowrap align="center"><%= rsBook.Fields("Version").Value %></TD>
	<TD ALIGN="right"><%= adoFormatDateTime(rsBook.Fields("PublishDate").Value, vbShortDate) %></TD>
	<TD>
		<A HREF="book_list.asp?bookid=<%= rsBook.Fields("BookID").Value %>&orderno=<%= rsBook.Fields("OrderNo").Value %>&action=moveup" class="actionlink"><% steTxt "up" %></A> .
		<A HREF="book_list.asp?bookid=<%= rsBook.Fields("BookID").Value %>&orderno=<%= rsBook.Fields("OrderNo").Value %>&action=movedown" class="actionlink"><% steTxt "down" %></A> .
		<A HREF="book_edit.asp?BookID=<%= rsBook.Fields("BookID").Value %>" class="actionlink"><% steTxt "edit" %></A> .
		<A HREF="book_delete.asp?BookID=<%= rsBook.Fields("BookID").Value %>" class="actionlink"><% steTxt "delete" %></A>
	</TD>
</TR>
<%	rsBook.MoveNext
	I = I + 1
   Loop %>
</TABLE>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No books exist in the database" %></B></P>

<% End If %>

<P ALIGN="center">
	<A HREF="book_add.asp" class="adminlink"><% steTxt "Add New Book" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->