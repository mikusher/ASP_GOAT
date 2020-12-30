<!-- #include file="../../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' doc_list.asp
'	Displays a list of the documents for the site
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
Dim rsDoc

Select Case LCase(steForm("action"))
	Case "moveup"
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblDoc " &_
					"SET	OrderNo = OrderNo + 1 " &_
					"WHERE	OrderNo = " & steNForm("OrderNo") - 1
			Call adoExecute(sStat)

			sStat = "UPDATE	tblDoc " &_
					"SET	OrderNo = " & (steNForm("OrderNo") - 1) & " " &_
					"WHERE	DocID = " & steNForm("DocID")
			Call adoExecute(sStat)
			modRefresh True
	Case "movedown"
			' increment orders above the new order no (to make room)
			sStat = "UPDATE	tblDoc " &_
					"SET	OrderNo = OrderNo - 1 " &_
					"WHERE	OrderNo = " & steNForm("OrderNo") + 1
			Call adoExecute(sStat)

			sStat = "UPDATE	tblDoc " &_
					"SET	OrderNo = " & (steNForm("OrderNo") + 1) & " " &_
					"WHERE	DocID = " & steNForm("DocID")
			Call adoExecute(sStat)
			modRefresh True
	Case "activ"
			' archive a module
			sStat = "UPDATE	tblDoc " &_
					"SET	Archive = 1 " &_
					"WHERE	DocID = " & steNForm("DocID")
			Call adoExecute(sStat)
			modRefresh True
	Case "deactiv"
			sStat = "UPDATE	tblDoc " &_
					"SET	Archive = 0 " &_
					"WHERE	DocID = " & steNForm("DocID")
			Call adoExecute(sStat)
			modRefresh True
End Select

sStat = "SELECT	tblDoc.DocID, tblDoc.Title, tblDoc.SubTitle, tblDocBook.Version, " &_
		"		tblDocAuthor.FirstName, tblDocAuthor.MiddleName, tblDocAuthor.LastName, " &_
		"		tblDoc.Created, tblDocBook.PublishDate, tblDoc.OrderNo " &_
		"FROM	tblDoc " &_
		"INNER JOIN	tblDocAuthor ON tblDoc.AuthorID = tblDocAuthor.AuthorID " &_
		"LEFT JOIN tblDocBook ON tblDocBook.BookID = tblDoc.BookID " &_
		"WHERE	tblDoc.Archive = 0 " &_
		"ORDER BY tblDoc.OrderNo"
Set rsDoc = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../../header.asp" -->
<!-- #include file="../../../../lib/admin/login_lib.asp" -->

<% sCurrentTab = "Document" %>
<!-- #include file="pagetabs_inc.asp" -->

<H3><% steTxt "Document List" %></H3>

<P>
<% steTxt "Shown below are all of the current documents defined in the database." %>
</P>

<% If Not rsDoc.EOF Then %>

<TABLE BORDER=0 CELLPADDING=4 CELLSPACING=0 class="list">
<TR>
	<TD class="listhead"><% steTxt "Order" %></TD>
	<TD class="listhead"><% steTxt "Title" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Author" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Version" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Published" %></TD>
	<TD class="listhead" ALIGN="right"><% steTxt "Action" %></TD>
</TR>
<%	I = 0
	Do Until rsDoc.EOF %>
<tr class="list<%= I mod 2 %>">
	<TD><%= rsDoc.Fields("OrderNo").Value %></TD>
	<TD nowrap><%= rsDoc.Fields("Title").Value %><% If rsDoc.Fields("SubTitle").Value & "" <> "" Then %><br><font class="tinytext"><%= rsDoc.Fields("SubTitle").Value %></font><% End If %></TD>
	<TD nowrap><%= rsDoc.Fields("FirstName").Value & " " & Trim(rsDoc.Fields("MiddleName").Value & " " & rsDoc.Fields("LastName").Value) %></TD>
	<TD nowrap align="right"><%= rsDoc.Fields("Version").Value %></TD>
	<TD ALIGN="right">
		<% If Not IsNull(rsDoc.Fields("PublishDate").Value) Then %>
		<%= adoFormatDateTime(rsDoc.Fields("PublishDate").Value, vbShortDate) %>
		<% Else %>
		<%= adoFormatDateTime(rsDoc.Fields("Created").Value, vbShortDate) %>
		<% End If %>
	</TD>
	<TD>
		<A HREF="doc_list.asp?DocID=<%= rsDoc.Fields("DocID").Value %>&orderno=<%= rsDoc.Fields("OrderNo").Value %>&action=moveup" class="actionlink"><% steTxt "up" %></A> .
		<A HREF="doc_list.asp?DocID=<%= rsDoc.Fields("DocID").Value %>&orderno=<%= rsDoc.Fields("OrderNo").Value %>&action=movedown" class="actionlink"><% steTxt "down" %></A> .
		<A HREF="doc_edit.asp?DocID=<%= rsDoc.Fields("DocID").Value %>" class="actionlink"><% steTxt "edit" %></A> .
		<A HREF="doc_delete.asp?DocID=<%= rsDoc.Fields("DocID").Value %>" class="actionlink"><% steTxt "delete" %></A>
	</TD>
</TR>
<%	rsDoc.MoveNext
   Loop %>
</TABLE>

<% Else %>

<P><B CLASS="error"><% steTxt "Sorry, No documents exist in the database" %></B></P>

<% End If %>

<P ALIGN="center">
<A HREF="doc_add.asp" class="adminlink"><% steTxt "Add New Document" %></A>
</P>

<!-- #include file="../../../../footer.asp" -->