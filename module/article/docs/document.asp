<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' article.asp
'	Display an individual article from the database.
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
Dim rsDoc
Dim sStat
Dim sParams

sParams = "action="&steForm("Action")&"&pageno="&steNForm("PageNo")&"&results="&steNForm("Results")&"&keywords="&Server.URLEncode(steForm("Keywords"))

sStat = "SELECT	doc.DocID, doc.Title, doc.SubTitle, b.Version, " &_
		"		a.FirstName, a.MiddleName, a.LastName, " &_
		"		doc.Body, Coalesce(b.PublishDate, doc.Created) AS PublishDate, doc.OrderNo " &_
		"FROM	tblDoc doc " &_
		"INNER JOIN	tblDocAuthor a ON doc.AuthorID = a.AuthorID " &_
		"LEFT JOIN tblDocBook b ON b.BookID = doc.BookID " &_
		"WHERE	doc.Archive = 0 " &_
		"AND	doc.Active = 1 " &_
		"AND	Coalesce(b.Archive, 0) = 0 " &_
		"AND	doc.DocID = " & steNForm("DocID")
Set rsDoc = adoOpenRecordset(sStat)
%>
<!-- #include file="../../../header.asp" -->

<% If Not rsDoc.EOF Then %>

<p>
<font class="articlehead"><%= rsDoc.Fields("Title").Value %></font><br>
<font class="tinytext"><% steTxt "by" %>&nbsp;<%= rsDoc.Fields("FirstName").Value & " " & rsDoc.Fields("MiddleName").Value & " " & rsDoc.Fields("LastName").Value %><% If IsDate(rsDoc.Fields("PublishDate").Value) Then %> - <%= adoFormatDateTime(rsDoc.Fields("PublishDate").Value, vbLongDate) %><% End If %></font><br>
</font>
</p>

<p>
<%= Replace(rsDoc.Fields("Body").Value, vbCrLf, "<BR>") %>
</p>
<% Else %>

<H3><% steTxt "Document No Longer Available" %></H3>

<p>
<% steTxt "Sorry, but the document that you requested is no longer available." %>&nbsp;
<% steTxt "Although we try to maintain an archive of all of our old documents," %>&nbsp;
<% steTxt "sometimes it becomes necessary to remove a document from our site." %>&nbsp;
<% steTxt "Please update your bookmarks accordingly." %>
</p>

<% End If %>

<% If steForm("src") = "search" Then %>
<p align="center">
	<a href="search.asp?<%= sParams %>" class="adminlink">&lt;&lt; <% steTxt "Back to Search" %></A>
</p>
<% End If %>
<!-- #include file="../../../footer.asp" -->
