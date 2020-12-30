<!-- #include file="../../../lib/site_lib.asp" -->
<%
'--------------------------------------------------------------------
' index.asp
'	Create an overview of all the FAQ documents in the database
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
Dim sArtStat

sArtStat = "SELECT	fd.DocumentID, fd.Title AS FaqTitle, fd.Synopsis, fd.Modified, fd.Created, " &_
			"		fa.Title, fa.FirstName, fa.MiddleName, fa.LastName " &_
			"FROM	tblFaqDocument fd " &_
			"INNER JOIN	tblFaqAuthor fa on fd.AuthorID = fa.AuthorID " &_
			"WHERE	fd.Active <> 0 " &_
			"AND	fd.Archive = 0 " &_
			"ORDER BY fd.OrderNo DESC"
Set rsDoc = adoOpenRecordset(sArtStat)
%>
<!-- #include file="../../../header.asp" -->

<h3><% steTxt "FAQ Documents" %></h3>

<p>
<% steTxt "Below are the currently available Frequently Asked Question documents." %>&nbsp;
<% steTxt "Click on a document title to view its contents or use the search box below to find the answer to your question." %>
</p>

<form method="post" action="search.asp">
<input type="hidden" name="action" value="GO">
<p align="center">
<table border=0 cellpadding=2 cellspacing=0>
<tr>
	<td class="forml"><% steTxt "Search For" %>:</td><td>&nbsp;&nbsp;</td>
	<td><input type="text" name="searchtext" value="<%= steEncForm("searchtext") %>" size="32" maxlength="100" class="form"></td><td>&nbsp;&nbsp;</td>
	<td><input type="submit" name="_action" value=" <% steTxt "GO" %> " class="form"></td>
</tr>
</table>
</p>
</form>

<% Do Until rsDoc.EOF %>
<A HREF="document.asp?documentid=<%= rsDoc.Fields("DocumentID").Value %>" CLASS="articlehead2"><%= rsDoc.Fields("FaqTitle").Value %></A><BR>
<FONT CLASS="articleauthor"><%= rsDoc.Fields("Created").Value %> - <%= Server.HTMLEncode(Trim(rsDoc.Fields("Title").Value & " " & rsDoc.Fields("FirstName").Value) & " " & Trim(rsDoc.Fields("MiddleName").Value & " " & rsDoc.Fields("LastName").Value)) %></FONT><BR>
<font class="articleleadin"><%= rsDoc.Fields("Synopsis").Value %></font>
<BR>
<div align="right" class="articlelink">
	<A HREF="document.asp?documentid=<%= rsDoc.Fields("DocumentID").Value %>" class="articlelink">...(<% steTxt "Read More" %>)</A>
	<!-- A HREF="../../news/comments.asp?documentid=<%= rsDoc.Fields("DocumentID").Value %>" class="articlelink">Comments (<= rsDoc.Fields("CommentCount").Value >)</A -->
</div>
<hr noshade width="100%" SIZE="1" style="color:#F8E8D8">
<%	rsDoc.MoveNext
   Loop %>

<!-- #include file="../../../footer.asp" -->